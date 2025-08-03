VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmRadarExport 
   Caption         =   "Radar Export"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "AffRadarExport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9615
   Begin VB.TextBox txtMaxSpots 
      Height          =   360
      Left            =   4005
      MaxLength       =   2
      TabIndex        =   16
      Text            =   "2"
      Top             =   1770
      Width           =   465
   End
   Begin VB.TextBox txtMaxRepeats 
      Height          =   360
      Left            =   8100
      MaxLength       =   2
      TabIndex        =   18
      Text            =   "6"
      Top             =   1770
      Width           =   465
   End
   Begin VB.TextBox txtIndicator 
      Height          =   360
      Index           =   1
      Left            =   5580
      TabIndex        =   8
      Top             =   735
      Width           =   1320
   End
   Begin VB.TextBox txtIndicator 
      Height          =   360
      Index           =   0
      Left            =   3150
      TabIndex        =   6
      Top             =   720
      Width           =   1320
   End
   Begin VB.CheckBox ckcOutput 
      Caption         =   "Csv"
      Height          =   195
      Index           =   1
      Left            =   6390
      TabIndex        =   4
      Top             =   255
      Width           =   615
   End
   Begin VB.CheckBox ckcOutput 
      Caption         =   "Prn"
      Height          =   195
      Index           =   0
      Left            =   5745
      TabIndex        =   3
      Top             =   255
      Value           =   1  'Checked
      Width           =   675
   End
   Begin VB.PictureBox pbcDaylight 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   120
      ScaleHeight     =   330
      ScaleWidth      =   5055
      TabIndex        =   9
      Top             =   1305
      Width           =   5055
      Begin VB.OptionButton rbcDaylight 
         Caption         =   "No"
         Height          =   195
         Index           =   1
         Left            =   3930
         TabIndex        =   12
         Top             =   0
         Width           =   630
      End
      Begin VB.OptionButton rbcDaylight 
         Caption         =   "Yes"
         Height          =   195
         Index           =   0
         Left            =   3105
         TabIndex        =   11
         Top             =   0
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Specified Week on Daylight-Saving Time"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   3075
      End
   End
   Begin VB.CheckBox chkAllVC 
      Caption         =   "All"
      Height          =   195
      Left            =   1680
      TabIndex        =   24
      Top             =   5445
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.ListBox lbcVehCodes 
      Height          =   2790
      ItemData        =   "AffRadarExport.frx":08CA
      Left            =   1680
      List            =   "AffRadarExport.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   2490
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtRadarNo 
      Height          =   360
      Left            =   6150
      MaxLength       =   2
      TabIndex        =   14
      Top             =   1260
      Width           =   495
   End
   Begin VB.CheckBox chkAllStation 
      Caption         =   "All"
      Height          =   195
      Left            =   3045
      TabIndex        =   27
      Top             =   5445
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.ListBox lbcStations 
      Height          =   2790
      ItemData        =   "AffRadarExport.frx":08CE
      Left            =   3015
      List            =   "AffRadarExport.frx":08D0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   26
      Top             =   2490
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CheckBox chkAllNC 
      Caption         =   "All"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   5445
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Enabled         =   0   'False
      Height          =   2790
      ItemData        =   "AffRadarExport.frx":08D2
      Left            =   5700
      List            =   "AffRadarExport.frx":08D4
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2490
      Width           =   3720
   End
   Begin VB.ListBox lbcNetworkCodes 
      Height          =   2790
      ItemData        =   "AffRadarExport.frx":08D6
      Left            =   120
      List            =   "AffRadarExport.frx":08D8
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   2490
      Width           =   1305
   End
   Begin VB.TextBox txtDate 
      Height          =   360
      Left            =   1830
      TabIndex        =   1
      Top             =   195
      Width           =   1320
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9420
      Top             =   4050
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6285
      FormDesignWidth =   9615
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   5925
      TabIndex        =   28
      Top             =   5610
      Width           =   1665
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7755
      TabIndex        =   29
      Top             =   5610
      Width           =   1665
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   195
      Left            =   5730
      TabIndex        =   33
      Top             =   5355
      Visible         =   0   'False
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label lacMaxSpots 
      Caption         =   "Max # of Spots per break to Export (Blank = All)"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   3645
   End
   Begin VB.Label lacMaxRepeats 
      Caption         =   "Max # Times Same Spot Exported"
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   1800
      Width           =   2565
   End
   Begin VB.Label lacIndicator 
      Caption         =   "End Date"
      Height          =   255
      Index           =   1
      Left            =   4635
      TabIndex        =   7
      Top             =   795
      Width           =   750
   End
   Begin VB.Label lacIndicator 
      Caption         =   "Intention Week Declaration:  Start Date"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   825
      Width           =   2895
   End
   Begin VB.Label lacOutpout 
      Caption         =   "Output Format"
      Height          =   225
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label lacTitleVC 
      Alignment       =   2  'Center
      Caption         =   "Vehicle Codes"
      Height          =   255
      Left            =   1710
      TabIndex        =   22
      Top             =   2235
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lacRunLetter 
      Caption         =   "Radar #"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   1305
      Width           =   990
   End
   Begin VB.Label lacTitleStation 
      Alignment       =   2  'Center
      Caption         =   "Stations"
      Height          =   255
      Left            =   3045
      TabIndex        =   25
      Top             =   2235
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   105
      TabIndex        =   32
      Top             =   5715
      Width           =   5580
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   6525
      TabIndex        =   30
      Top             =   2235
      Width           =   1965
   End
   Begin VB.Label lacTitleNC 
      Alignment       =   2  'Center
      Caption         =   "Network Codes"
      Height          =   255
      Left            =   135
      TabIndex        =   19
      Top             =   2235
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "Clearance Start Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1515
   End
End
Attribute VB_Name = "frmRadarExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmContact - allows for selection of station/vehicle/advertiser for contact information
'*
'*  Created January,2003 by Dick LeVine
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text

Private smDate As String     'Export Date
Private smIndicatorStartDate As String
Private smIndicatorEndDate As String
Private imVefCode As Integer
Private imAdfCode As Integer
Private imAllNCChk As Integer
Private imAllVCChk As Integer
Private imAllStationClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private hmToPRN As Integer
Private hmToCSV As Integer
Private hmAst As Integer
Private cprst As ADODB.Recordset
Private smMessage As String
Private smWarnFlag As Integer
Private tmRadarExportInfo() As RADAREXPORTINFO
Private imIncludeVehicleCodeInFileName As Integer
Private smRADARMultiAir As String
Private tmCPDat() As DAT
Private rst_rht As ADODB.Recordset
Private rst_ret As ADODB.Recordset
Private rst_att As ADODB.Recordset
Private rst_Shtt As ADODB.Recordset
Private rst_DAT As ADODB.Recordset
Private tmAstInfo() As ASTINFO
Private tmSvAstInfo() As ASTINFO
Private tmRet() As RETINFO




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
Private Function mOpenMsgFile(sMsgFileName As String) As Integer
'    Dim slToFile As String
'    Dim slDateTime As String
'    Dim slFileDate As String
'    Dim slNowDate As String
'    Dim ilRet As Integer
'
'    On Error GoTo mOpenMsgFileErr:
'    ilRet = 0
'    slNowDate = Format$(gNow(), sgShowDateForm)
'    slToFile = sgMsgDirectory & "ExptSchdSpots.Txt"
'    slDateTime = FileDateTime(slToFile)
'    If ilRet = 0 Then
'        slFileDate = Format$(slDateTime, sgShowDateForm)
'        If DateValue(gAdjYear(slFileDate)) = DateValue(gAdjYear(slNowDate)) Then  'Append
'            On Error GoTo 0
'            ilRet = 0
'            On Error GoTo mOpenMsgFileErr:
'            hmMsg = FreeFile
'            Open slToFile For Append As hmMsg
'            If ilRet <> 0 Then
'                Close hmMsg
'                hmMsg = -1
'                gMsgBox "Open File " & slToFile & " error #" & Str$(Err.Number), vbOKOnly
'                mOpenMsgFile = False
'                Exit Function
'            End If
'        Else
'            Kill slToFile
'            On Error GoTo 0
'            ilRet = 0
'            On Error GoTo mOpenMsgFileErr:
'            hmMsg = FreeFile
'            Open slToFile For Output As hmMsg
'            If ilRet <> 0 Then
'                Close hmMsg
'                hmMsg = -1
'                gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
'                mOpenMsgFile = False
'                Exit Function
'            End If
'        End If
'    Else
'        On Error GoTo 0
'        ilRet = 0
'        On Error GoTo mOpenMsgFileErr:
'        hmMsg = FreeFile
'        Open slToFile For Output As hmMsg
'        If ilRet <> 0 Then
'            Close hmMsg
'            hmMsg = -1
'            gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
'            mOpenMsgFile = False
'            Exit Function
'        End If
'    End If
'    On Error GoTo 0
'    'Print #hmMsg, "** Export of StarGuide: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
'    'Print #hmMsg, ""
'    'sMsgFileName = slToFile
'    'mOpenMsgFile = True
'    Exit Function
'mOpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function

Private Sub mFillNC()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llRow As Long
    On Error GoTo ErrHand
    
    lbcNetworkCodes.Clear
    lbcMsg.Clear
    chkAllNC.Value = vbUnchecked
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        SQLQuery = "SELECT * FROM rht WHERE (rhtVefCode = " & tgVehicleInfo(ilLoop).iCode & ")"
        Set rst_rht = gSQLSelectCall(SQLQuery)
        Do While Not rst_rht.EOF
            slStr = rst_rht!rhtRadarNetCode
            llRow = SendMessageByString(lbcNetworkCodes.hwnd, LB_FINDSTRING, -1, slStr)
            If llRow < 0 Then
                lbcNetworkCodes.AddItem slStr
            End If
            rst_rht.MoveNext
        Loop
    Next ilLoop
    rst_rht.Close
    On Error GoTo 0
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarExport-mFillNC"
End Sub

Private Sub chkAllNC_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllNCChk Then
        Exit Sub
    End If
    If chkAllNC.Value = vbChecked Then
        iValue = True
        If lbcNetworkCodes.ListCount > 1 Then
            lacTitleVC.Visible = False
            chkAllVC.Visible = False
            lbcStations.Visible = False
            lbcVehCodes.Clear
            lacTitleStation.Visible = False
            chkAllStation.Visible = False
            lbcStations.Visible = False
            lbcStations.Clear
        Else
            lacTitleVC.Visible = True
            chkAllVC.Visible = True
            lbcVehCodes.Visible = True
        End If
    Else
        iValue = False
    End If
    If lbcNetworkCodes.ListCount > 0 Then
        imAllNCChk = True
        lRg = CLng(lbcNetworkCodes.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcNetworkCodes.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllNCChk = False
    End If
    mSetVC
End Sub

Private Sub chkAllStation_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllStationClick Then
        Exit Sub
    End If
    If chkAllStation.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcStations.ListCount > 0 Then
        imAllStationClick = True
        lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllStationClick = False
    End If

End Sub



Private Sub chkAllVC_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllVCChk Then
        Exit Sub
    End If
    If chkAllVC.Value = vbChecked Then
        iValue = True
        If lbcVehCodes.ListCount > 1 Then
            lacTitleStation.Visible = False
            chkAllStation.Visible = False
            lbcStations.Visible = False
            lbcStations.Clear
        Else
            lacTitleStation.Visible = True
            chkAllStation.Visible = True
            lbcStations.Visible = True
        End If
    Else
        iValue = False
    End If
    If lbcVehCodes.ListCount > 0 Then
        imAllVCChk = True
        lRg = CLng(lbcVehCodes.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehCodes.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllVCChk = False
    End If
    mSetStations
End Sub

Private Sub cmdExport_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVCCount As Integer
    Dim ilNCCount As Integer
    Dim ilProcVCCount As Integer
    Dim ilProcNCCount As Integer
    Dim ilStationCount As Integer
    Dim ilNC As Integer
    Dim ilVC As Integer
    Dim ilStation As Integer
    Dim slNC As String
    Dim slVC As String
    Dim llPercent As Long
    Dim ilSelStation As Integer

    On Error GoTo ErrHand
    
    lbcMsg.Clear
    If lbcNetworkCodes.ListIndex < 0 Then
        Exit Sub
    End If
    If txtDate.Text = "" Then
        gMsgBox "Learance Date must be specified.", vbOKOnly
        txtDate.SetFocus
        Exit Sub
    End If
    If gIsDate(txtDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid Clearance date (m/d/yy).", vbCritical
        txtDate.SetFocus
        Exit Sub
    Else
        smDate = Format(txtDate.Text, sgShowDateForm)
    End If
    If Weekday(gAdjYear(smDate)) <> vbMonday Then
        Beep
        gMsgBox "Please enter a Monday Clearance date (m/d/yy).", vbCritical
        txtDate.SetFocus
        Exit Sub
    End If
    If txtIndicator(0).Text <> "" Then
        If gIsDate(txtIndicator(0).Text) = False Then
            Beep
            gMsgBox "Please enter a valid Intention Start date (m/d/yy).", vbCritical
            txtIndicator(0).SetFocus
            Exit Sub
        End If
        If txtIndicator(1).Text = "" Then
            gMsgBox "Intention End Date must be specified.", vbOKOnly
            txtIndicator(1).SetFocus
            Exit Sub
        End If
        If gIsDate(txtIndicator(1).Text) = False Then
            Beep
            gMsgBox "Please enter a valid Intention End date (m/d/yy).", vbCritical
            txtIndicator(1).SetFocus
            Exit Sub
        End If
    End If
    smIndicatorStartDate = txtIndicator(0).Text
    smIndicatorEndDate = txtIndicator(1).Text
    If txtRadarNo.Text = "" Then
        gMsgBox "Radar # must be specified.", vbOKOnly
        txtRadarNo.SetFocus
        Exit Sub
    End If
    If Val(txtMaxRepeats.Text) = 0 Then
        gMsgBox "Max # Times Same Spot Exported must be specified (1-99).", vbOKOnly
        txtRadarNo.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilNCCount = 0
    ilProcNCCount = 0
    For ilLoop = 0 To lbcNetworkCodes.ListCount - 1 Step 1
        If lbcNetworkCodes.Selected(ilLoop) Then
            ilNCCount = ilNCCount + 1
            'If ilNCCount > 1 Then
            '    Exit For
            'End If
        End If
    Next ilLoop
    imIncludeVehicleCodeInFileName = False
    ilVCCount = 0
    ilProcVCCount = 0
    If ilNCCount = 1 Then
        For ilLoop = 0 To lbcVehCodes.ListCount - 1 Step 1
            If lbcVehCodes.Selected(ilLoop) Then
                ilVCCount = ilVCCount + 1
                'If ilVCCount > 1 Then
                '    Exit For
                'End If
            End If
        Next ilLoop
        If (ilVCCount <> lbcVehCodes.ListCount) And (ilVCCount <> 0) Then
            imIncludeVehicleCodeInFileName = True
        End If
    End If
    plcGauge.Visible = True
    plcGauge.Value = 0
    smWarnFlag = False
    imExporting = True
    On Error GoTo 0
    bgTaskBlocked = False
    sgTaskBlockedName = "RADAR Export"
    lacResult.Caption = ""
    gLogMsg "Radar Export run for the week of " & smDate, "RadarExportLog.Txt", False
    For ilNC = 0 To lbcNetworkCodes.ListCount - 1 Step 1
        If imTerminate Then
            Exit For
        End If
        If lbcNetworkCodes.Selected(ilNC) Then
            slNC = lbcNetworkCodes.List(ilNC)
            If Not imIncludeVehicleCodeInFileName Then
                If Not mOpenRadarExportFile(slNC, "") Then
                    bgTaskBlocked = False
                    sgTaskBlockedName = ""
                    imExporting = False
                    cmdCancel.SetFocus
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            If ilVCCount = 0 Then
                mFillVC slNC
                chkAllVC.Value = vbChecked
            End If
            For ilVC = 0 To lbcVehCodes.ListCount - 1 Step 1
                If imTerminate Then
                    Exit For
                End If
                If lbcVehCodes.Selected(ilVC) Then
                    slVC = lbcVehCodes.List(ilVC)
                    If imIncludeVehicleCodeInFileName Then
                        If Not mOpenRadarExportFile(slNC, slVC) Then
                            bgTaskBlocked = False
                            sgTaskBlockedName = ""
                            imExporting = False
                            cmdCancel.SetFocus
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                    If (ilNCCount <> 1) Or (ilVCCount <> 1) Then
                        mFillStations slNC, slVC
                        chkAllStation.Value = vbChecked
                    Else
                        ilSelStation = 0
                        For ilStation = 0 To lbcStations.ListCount - 1 Step 1
                            If lbcStations.Selected(ilStation) Then
                                ilSelStation = ilSelStation + 1
                            End If
                        Next ilStation
                    End If
                    ilStationCount = 0
                    For ilStation = 0 To lbcStations.ListCount - 1 Step 1
                        If imTerminate Then
                            Exit For
                        End If
                        If lbcStations.Selected(ilStation) Then
                            ilStationCount = ilStationCount + 1
                            ilRet = mExportSpots(slNC, slVC, lbcStations.ItemData(ilStation), lbcStations.List(ilStation))
                            If (ilRet = False) Then
                                bgTaskBlocked = False
                                sgTaskBlockedName = ""
                                gCloseRegionSQLRst
                                gLogMsg "** Terminated - mExportSpots returned False **", "RadarExportLog.Txt", False
                                imExporting = False
                                Screen.MousePointer = vbDefault
                                cmdCancel.SetFocus
                                Exit Sub
                            End If
                            If (ilNCCount <= 1) And (ilVCCount <= 1) Then
                                llPercent = (ilStationCount * CSng(100)) / ilSelStation
                                If llPercent > 100 Then
                                    llPercent = 100
                                End If
                                plcGauge.Value = llPercent
                            End If
                        End If
                    Next ilStation
                    If (lbcStations.ListCount = 0) Or (chkAllStation.Value = vbChecked) Or (lbcStations.ListCount = lbcStations.SelCount) Then
                        gClearASTInfo True
                    Else
                        gClearASTInfo False
                    End If
                    ''Check for Stations only in future
                    'If ilStationCount = lbcStations.ListCount Then
                    '    mGetStationsNotReported slNC, slVC
                    'End If
                    If (ilNCCount <= 1) And (ilVCCount > 1) Then
                        ilProcVCCount = ilProcVCCount + 1
                        llPercent = (ilProcVCCount * CSng(100)) / ilVCCount
                        If llPercent > 100 Then
                            llPercent = 100
                        End If
                        plcGauge.Value = llPercent
                    End If
                End If
            Next ilVC
            If ckcOutput(0).Value = vbChecked Then
                Print #hmToPRN, ""
                Close #hmToPRN
                DoEvents
            End If
            If ckcOutput(1).Value = vbChecked Then
                Print #hmToCSV, ""
                Close #hmToCSV
                DoEvents
            End If
            ilProcNCCount = ilProcNCCount + 1
            If ilNCCount > 1 Then
                llPercent = (ilProcNCCount * CSng(100)) / ilNCCount
                If llPercent > 100 Then
                    llPercent = 100
                End If
                plcGauge.Value = llPercent
            End If
        End If
    Next ilNC
    gCloseRegionSQLRst
    If imTerminate Then
        gLogMsg "** User Terminated **", "RadarExportLog.Txt", False
        imExporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    On Error GoTo ErrHand:
    If bgTaskBlocked And igExportSource <> 2 Then
        gMsgBox "Some spots were blocked during the Export generation." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
    End If
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    imExporting = False
    'Print #hmMsg, "** Completed Export of StarGuide: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    gLogMsg "** Completed Export of Radar **", "RadarExportLog.Txt", False
    'Close #hmMsg
    lacResult.Caption = "Exports placed into: " & sgExportDirectory & ", and Results logged into data\messages\RadarExportLog.Txt"
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    plcGauge.Visible = False
    Screen.MousePointer = vbDefault
    gLogMsg "", "RadarExportLog.Txt", False
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmRadarExport-cmdExport"
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    txtDate.Text = ""
    Unload frmRadarExport
End Sub


Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Dim ilLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    frmRadarExport.Caption = "RADAR Export - " & sgClientName
    smDate = gObtainNextMonday(Format$(gNow(), sgShowDateForm))
    txtDate.Text = smDate
    imAllNCChk = False
    imAllStationClick = False
    imTerminate = False
    imExporting = False
    imAllNCChk = False
    imAllVCChk = False
    imAllStationClick = False
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
    '4/25/19: Get multi-airplay type
    smRADARMultiAir = "S"
    SQLQuery = "SELECT * From Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        smRADARMultiAir = rst!siteRADARMultiAir         'P=By Program code; S or Blank by Spot ID
        'P=Multi-Play By RADAR Program code; S= Multi-Play (by Spot ID);  A=Air Time (test for A)
    End If
    
    lbcStations.Clear
    lbcVehCodes.Clear
    mFillNC
    'txtFile.Text = sgExportDirectory & "MktSpots.txt"
    chkAllNC.Value = vbChecked
    gLogMsg "Radar Exported run on " & Format(gNow(), "m/d/yy"), "RadarExportLog.Txt", True
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    Erase tmRadarExportInfo
    Erase tmSvAstInfo
    Erase tmAstInfo
    Erase tmRet
    cprst.Close
    rst_rht.Close
    rst_ret.Close
    rst_att.Close
    rst_Shtt.Close
    rst_DAT.Close
    Set frmRadarExport = Nothing
End Sub


Private Sub lbcStations_Click()
    If imAllStationClick Then
        Exit Sub
    End If
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
    If chkAllStation.Value = vbChecked Then
        imAllStationClick = True
        chkAllStation.Value = vbUnchecked
        imAllStationClick = False
    End If
End Sub

Private Sub lbcNetworkCodes_Click()
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim slNC As String
    
    lbcStations.Clear
    lbcVehCodes.Clear
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
    If chkAllStation.Value = vbChecked Then
        chkAllStation.Value = vbUnchecked
    End If
    If chkAllVC.Value = vbChecked Then
        chkAllVC.Value = vbUnchecked
    End If
    If imAllNCChk Then
        Exit Sub
    End If
    If chkAllNC.Value = vbChecked Then
        imAllNCChk = True
        chkAllNC.Value = vbUnchecked
        imAllNCChk = False
    End If
    mSetVC
End Sub

Private Sub lbcVehCodes_Click()
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim slVC As String
    Dim slNC As String
    
    lbcStations.Clear
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
    If chkAllStation.Value = vbChecked Then
        chkAllStation.Value = vbUnchecked
    End If
    If imAllVCChk Then
        Exit Sub
    End If
    If chkAllVC.Value = vbChecked Then
        imAllVCChk = True
        chkAllVC.Value = vbUnchecked
        imAllVCChk = False
    End If
    mSetStations

End Sub

Private Sub txtDate_Change()
    lbcMsg.Clear
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
    If gIsDate(txtDate.Text) Then
        Screen.MousePointer = vbHourglass
        mSetStations
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Function mExportSpots(slNC As String, slVC As String, ilShttCode As Integer, slInCallLetters As String) As Integer
    Dim ilLoop As Integer
    Dim ilVefCode As Integer
    Dim slVehicleName As String
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim llSpotTime As Long
    Dim llDate As Long
    Dim llTime As Long
    Dim ilIncludeSpot As Integer
    Dim ilTest As Integer
    Dim slDayName As String
    Dim slDate As String
    Dim slRecordPRN As String
    Dim slRecordCSV As String
    Dim slRecordNC As String
    Dim slRecordVC As String
    Dim ilSdfCode As Integer
    Dim slCmmlUnit As String
    Dim slRepeatCount As String
    Dim slCallLetters As String
    Dim slBand As String
    Dim slCall As String
    Dim ilPos As Integer
    Dim ilUnitCount As Integer
    Dim ilRepeatCount As Integer
    Dim slZone As String
    Dim slClearType As String
    Dim slTime As String
    Dim slProgCode As String
    Dim ilVef As Integer
    Dim ilZone As Integer
    Dim ilLocalAdj As Integer
    Dim ilZoneFound As Integer
    Dim ilNumberAsterisk As Integer
    Dim ilDACode As Integer
    Dim ilWeekDay As Integer
    Dim ilNCVCFound As Integer
    Dim slIndicator As String
    Dim ilAst As Integer
    Dim ilMaxRepeats As Integer
    Dim ilMaxSpots As Integer
    Dim slCPTTDate As String
    Dim ilProgCodeRepeatCount As Integer
    '8/21/19
    Dim ilAirPlayCount As Integer
    Dim ilAirPlay As Integer
    '8/26/19
    Dim ilMaxProgCodeRepeatCount As Integer
    '3/13/20
    Dim slAirDate As String
    Dim slAirTime As String
    Dim slPrevAirDate As String
    Dim slPrevAirTime As String
    Dim ilAirSpotCount As Integer
    Dim ilAirBreakCount As Integer
    Dim slSortDate As String
    Dim slSortTime As String
    On Error GoTo ErrHand
    slCPTTDate = smDate
    ilMaxRepeats = Val(txtMaxRepeats.Text)
    If Trim$(txtMaxSpots.Text) <> "" Then
        ilMaxSpots = Val(txtMaxSpots.Text)
    Else
        ilMaxSpots = -1
    End If
    slRecordNC = slNC
    Do While Len(slRecordNC) < 2
        slRecordNC = slRecordNC & " "
    Loop
    slRecordVC = slVC
    Do While Len(slRecordVC) < 3
        slRecordVC = slRecordVC & " "
    Loop
    ilNCVCFound = False
    ReDim tmRadarExportInfo(0 To 0) As RADAREXPORTINFO
    ReDim tmSvAstInfo(0 To 0) As ASTINFO
    ilVefCode = 0
    ReDim tmRet(0 To 0) As RETINFO
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        DoEvents
        SQLQuery = "SELECT * FROM rht WHERE (rhtVefCode = " & tgVehicleInfo(ilLoop).iCode & ")"
        Set rst_rht = gSQLSelectCall(SQLQuery)
        Do While Not rst_rht.EOF
            DoEvents
            If (rst_rht!rhtRadarNetCode = slNC) And (rst_rht!rhtRadarVehCode = slVC) Then
                '5/21/12: Clear previous values
                ReDim tmRet(0 To 0) As RETINFO
                ilNCVCFound = True
                ilVef = ilLoop
                ilVefCode = rst_rht!rhtVefCode
                slVehicleName = Trim$(tgVehicleInfo(ilLoop).sVehicle)
                '4/25/19: Added Order by
                SQLQuery = "SELECT * FROM ret WHERE (retRhtCode = " & rst_rht!rhtCode & ")" & " Order By retProgCode, retStartTime "
                Set rst_ret = gSQLSelectCall(SQLQuery)
                Do While Not rst_ret.EOF
                    DoEvents
                    tmRet(UBound(tmRet)).sProgCode = rst_ret!retProgCode
                    tmRet(UBound(tmRet)).lStartTime = gTimeToLong(Format$(CStr(rst_ret!retStartTime), sgShowTimeWSecForm), False)
                    tmRet(UBound(tmRet)).lEndTime = gTimeToLong(Format$(CStr(rst_ret!retEndTime), sgShowTimeWSecForm), True)
                    tmRet(UBound(tmRet)).sDayType = rst_ret!retDayType
                    '4/25/19: Added setting repeat code
                    If UBound(tmRet) > 0 Then
                        If tmRet(UBound(tmRet) - 1).sProgCode = tmRet(UBound(tmRet)).sProgCode Then
                            tmRet(UBound(tmRet)).iRepeatCount = tmRet(UBound(tmRet) - 1).iRepeatCount + 1
                        Else
                            tmRet(UBound(tmRet)).iRepeatCount = 0
                        End If
                    Else
                        tmRet(UBound(tmRet)).iRepeatCount = 0
                    End If
                    ReDim Preserve tmRet(0 To UBound(tmRet) + 1) As RETINFO
                    rst_ret.MoveNext
                Loop
                rst_ret.Close
                DoEvents
                SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate, attRadarClearType, attPledgeType"
                SQLQuery = SQLQuery & " FROM cptt, shtt, att"
                SQLQuery = SQLQuery & " WHERE (cpttVefCode = " & ilVefCode
                SQLQuery = SQLQuery & " AND cpttShfCode = " & ilShttCode
                SQLQuery = SQLQuery & " AND ShttCode = cpttShfCode"
                SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
                '10/29/14: Bypass Service agreements
                SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"
                'SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(smDate, sgSQLDateForm) & "')"
                SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slCPTTDate, sgSQLDateForm) & "')"
                Set cprst = gSQLSelectCall(SQLQuery)
                If cprst.EOF Then
                    'Test Indicator date
                    If Trim$(smIndicatorStartDate) <> "" Then
                        slIndicator = "X"
                        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate, attRadarClearType, attPledgeType"
                        SQLQuery = SQLQuery & " FROM cptt, shtt, att"
                        SQLQuery = SQLQuery & " WHERE (cpttVefCode = " & ilVefCode
                        SQLQuery = SQLQuery & " AND cpttShfCode = " & ilShttCode
                        SQLQuery = SQLQuery & " AND ShttCode = cpttShfCode"
                        SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
                        SQLQuery = SQLQuery & " AND cpttStartDate >= '" & Format$(smIndicatorStartDate, sgSQLDateForm) & "')"
                        SQLQuery = SQLQuery & " ORDER BY cpttStartDate"
                        Set cprst = gSQLSelectCall(SQLQuery)
                        If Not cprst.EOF Then
                            Do While Not cprst.EOF
                                DoEvents
                                If DateValue(Format$(cprst!CpttStartDate, sgShowDateForm)) <= DateValue(smIndicatorEndDate) Then
                                    On Error GoTo ErrHand
                                    If (DateValue(Format$(cprst!attOnAir, sgShowDateForm)) >= DateValue(smIndicatorStartDate)) And (DateValue(Format$(cprst!attOnAir, sgShowDateForm)) <= DateValue(smIndicatorEndDate)) Then
                                        slCPTTDate = gObtainPrevMonday(Format$(cprst!CpttStartDate, sgShowDateForm))
                                        ReDim tgCPPosting(0 To 1) As CPPOSTING
                                        tgCPPosting(0).lCpttCode = cprst!cpttCode
                                        tgCPPosting(0).iStatus = cprst!cpttStatus
                                        tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                                        tgCPPosting(0).lAttCode = cprst!cpttatfCode
                                        tgCPPosting(0).iAttTimeType = cprst!attTimeType
                                        tgCPPosting(0).iVefCode = ilVefCode
                                        tgCPPosting(0).iShttCode = ilShttCode
                                        tgCPPosting(0).sZone = cprst!shttTimeZone
                                        tgCPPosting(0).sDate = slCPTTDate   'Format$(smDate, sgShowDateForm)
                                        tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                                        'SQLQuery = "SELECT * "
                                        'SQLQuery = SQLQuery + " FROM dat"
                                        'SQLQuery = SQLQuery + " WHERE (datatfCode= " & tgCPPosting(0).lAttCode & ")"
                                        'Set rst_dat = gSQLSelectCall(SQLQuery)
                                        'If Not rst_dat.EOF Then
                                        '    ilDACode = rst_dat!datDACode
                                        'Else
                                        '    ilDACode = -1
                                        'End If
                                        If cprst!attPledgeType = "D" Then
                                            ilDACode = 0
                                        ElseIf cprst!attPledgeType = "A" Then
                                            ilDACode = 1
                                        ElseIf cprst!attPledgeType = "C" Then
                                            ilDACode = 2
                                        Else
                                            ilDACode = -1
                                        End If
                                        'Create AST records
                                        'igTimes = 2 'By Week, sort by feed
                                        If smRADARMultiAir <> "A" Then
                                            igTimes = 2 'By Week, sort by feed date/time
                                        Else
                                            igTimes = 1 'By Week, sort by air date/time
                                        End If
                                        imAdfCode = -1
                                        DoEvents
                                        'Dan M 9/26/13 6442
                                        ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, False, True)
                                        'ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, False, False, True)
                                        If LBound(tmAstInfo) < UBound(tmAstInfo) Then
                                            slIndicator = "I"
                                            Exit Do
                                        End If
                                    End If
                                Else
                                    slIndicator = "X"
                                    Exit Do
                                End If
                                cprst.MoveNext
                            Loop
                        Else
                            slIndicator = "X"
                        End If
                    Else
                        slIndicator = "X"
                    End If
                Else
                    slIndicator = " "
                End If
                If slIndicator <> "X" Then
                    If cprst!attRadarClearType <> "E" Then
                        '7/6/12: Include None Aired as Completed
                        'If ((cprst!cpttStatus = 1) And (cprst!cpttPostingStatus = 2)) Or (slIndicator = "I") Then
                        If (((cprst!cpttStatus = 1) Or (cprst!cpttStatus = 2)) And (cprst!cpttPostingStatus = 2)) Or (slIndicator = "I") Then
                            On Error GoTo ErrHand
                            ReDim tgCPPosting(0 To 1) As CPPOSTING
                            tgCPPosting(0).lCpttCode = cprst!cpttCode
                            tgCPPosting(0).iStatus = cprst!cpttStatus
                            tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                            tgCPPosting(0).lAttCode = cprst!cpttatfCode
                            tgCPPosting(0).iAttTimeType = cprst!attTimeType
                            tgCPPosting(0).iVefCode = ilVefCode
                            tgCPPosting(0).iShttCode = ilShttCode
                            tgCPPosting(0).sZone = cprst!shttTimeZone
                            tgCPPosting(0).sDate = Format$(slCPTTDate, sgShowDateForm)  'Format$(smDate, sgShowDateForm)
                            tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                            'SQLQuery = "SELECT * "
                            'SQLQuery = SQLQuery + " FROM dat"
                            'SQLQuery = SQLQuery + " WHERE (datatfCode= " & tgCPPosting(0).lAttCode & ")"
                            'Set rst_dat = gSQLSelectCall(SQLQuery)
                            'If Not rst_dat.EOF Then
                            '    ilDACode = rst_dat!datDACode
                            'Else
                            '    ilDACode = -1
                            'End If
                            If cprst!attPledgeType = "D" Then
                                ilDACode = 0
                            ElseIf cprst!attPledgeType = "A" Then
                                ilDACode = 1
                            ElseIf cprst!attPledgeType = "C" Then
                                ilDACode = 2
                            Else
                                ilDACode = -1
                            End If
                            'Create AST records
                            If smRADARMultiAir <> "A" Then
                                igTimes = 2 'By Week, sort by feed date/time
                            Else
                                igTimes = 1 'By Week, sort by air date/time
                            End If
                            imAdfCode = -1
                            DoEvents
                            'Dan M 9/26/13 6442
                            ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, False, True)
                           ' ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, False, False, True)
                            ilAirPlayCount = 0
                            For ilAirPlay = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                                If smRADARMultiAir = "A" Then
                                    tmAstInfo(ilAirPlay).iAirPlay = 1
                                    'tmAstInfo(ilAirPlay).sFeedDate = tmAstInfo(ilAirPlay).sAirDate
                                    'tmAstInfo(ilAirPlay).sFeedTime = tmAstInfo(ilAirPlay).sAirTime
                                    slSortDate = gDateValue(tmAstInfo(ilAirPlay).sAirDate)
                                    Do While Len(slSortDate) < 6
                                        slSortDate = "0" & slSortDate
                                    Loop
                                    slSortTime = gTimeToLong(tmAstInfo(ilAirPlay).sAirTime, False)
                                    Do While Len(slSortTime) < 6
                                        slSortTime = "0" & slSortTime
                                    Loop
                                    tmAstInfo(ilAirPlay).sKey = slSortDate & slSortTime
                                End If
                                If tmAstInfo(ilAirPlay).iAirPlay > ilAirPlayCount Then
                                    ilAirPlayCount = tmAstInfo(ilAirPlay).iAirPlay
                                End If
                            Next ilAirPlay
                            If smRADARMultiAir = "A" Then
                                If UBound(tmAstInfo) - 1 >= 1 Then
                                    ArraySortTyp fnAV(tmAstInfo(), 0), UBound(tmAstInfo), 0, LenB(tmAstInfo(0)), 0, LenB(tmAstInfo(0).sKey), 0
                                End If
                            End If
                            'TTP 10324  - JW - 10/19/21 - RADAR export: spot ID method, a spot with a single spot ID that airs on multiple airing vehicles carried by a single station doesn't get the "# times spot repeated" count set correctly
                            If smRADARMultiAir <> "S" And Trim(smRADARMultiAir) <> "" Then
                                ReDim tmRadarExportInfo(0 To 0) As RADAREXPORTINFO
                                ReDim tmSvAstInfo(0 To 0) As ASTINFO
                            End If
                            For ilAirPlay = 1 To ilAirPlayCount Step 1
                                ilIndex = LBound(tmAstInfo)
                                Do While ilIndex < UBound(tmAstInfo)
                                    DoEvents
                                    If imTerminate Then
                                        mExportSpots = False
                                        Exit Function
                                    End If
                                    If (tmAstInfo(ilIndex).iAirPlay = ilAirPlay) Or ((tmAstInfo(ilIndex).iAirPlay <= 0) And (ilAirPlay = 1)) Then
                                        ilIncludeSpot = True
                                        llDate = DateValue(tmAstInfo(ilIndex).sFeedDate)
                                        llTime = gTimeToLong(tmAstInfo(ilIndex).sFeedTime, False)
                                        'Translate time based on zone
                                        'Select Case UCase$(Trim$(cprst!shttTimeZone))
                                        '    Case "EST"
                                        '        llSpotTime = llTime
                                        '    Case "CST"
                                        '        llSpotTime = llTime + 3600
                                        '    Case "MST"
                                        '        llSpotTime = llTime + 2 * 3600
                                        '    Case "PST"
                                        '        llSpotTime = llTime + 3 * 3600
                                        '    Case Else
                                        '        llSpotTime = llTime
                                        'End Select
                                        'If (llSpotTime >= 24 * CLng(3600)) Then
                                        '    'Adjust date
                                        '    llDate = llDate + 1
                                        '    llSpotTime = llSpotTime - 24 * CLng(3600)
                                        'End If
                                        slZone = UCase$(Trim$(cprst!shttTimeZone))
                                        ilLocalAdj = 0
                                        ilZoneFound = False
                                        ilNumberAsterisk = 0
                                        If smRADARMultiAir = "A" Then
                                            slZone = ""
                                        End If
                                        ' Adjust time zone properly.
                                        If Len(slZone) <> 0 Then
                                            'Get zone
                                            DoEvents
                                            For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                                                If Trim$(tgVehicleInfo(ilVef).sZone(ilZone)) = slZone Then
                                                    If tgVehicleInfo(ilVef).sFed(ilZone) <> "*" Then
                                                        slZone = tgVehicleInfo(ilVef).sZone(tgVehicleInfo(ilVef).iBaseZone(ilZone))
                                                        ilLocalAdj = tgVehicleInfo(ilVef).iLocalAdj(ilZone)
                                                        ilZoneFound = True
                                                    End If
                                                    Exit For
                                                End If
                                            Next ilZone
                                            For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                                                If tgVehicleInfo(ilVef).sFed(ilZone) = "*" Then
                                                    ilNumberAsterisk = ilNumberAsterisk + 1
                                                End If
                                            Next ilZone
                                        End If
                                        If (Not ilZoneFound) And (ilNumberAsterisk <= 1) Then
                                            slZone = ""
                                        End If
                                        ilLocalAdj = -1 * ilLocalAdj
                                        llSpotTime = llTime + 3600 * ilLocalAdj
                                        If llSpotTime < 0 Then
                                            llSpotTime = llSpotTime + 86400
                                            llDate = llDate - 1
                                        ElseIf llSpotTime > 86400 Then
                                            llSpotTime = llSpotTime - 86400
                                            llDate = llDate + 1
                                        End If
                                        If ilDACode = 2 Then  'Tape/CD
                                            llDate = DateValue(tmAstInfo(ilIndex).sFeedDate)
                                            llSpotTime = gTimeToLong(tmAstInfo(ilIndex).sFeedTime, False)
                                        End If
                                        'Test if within Program Schedule
                                        slDate = Format$(llDate, "m/d/yy")
                                        ilWeekDay = Weekday(slDate)
                                        ilProgCodeRepeatCount = 0
                                        If tgStatusTypes(gGetAirStatus(tmAstInfo(ilIndex).iPledgeStatus)).iPledged <> 2 Then
                                            For ilTest = 0 To UBound(tmRet) - 1 Step 1
                                                DoEvents
                                                ilIncludeSpot = True
                                                slProgCode = tmRet(ilTest).sProgCode
                                                '4/45/19
                                                ilProgCodeRepeatCount = tmRet(ilTest).iRepeatCount
                                                If (llSpotTime < tmRet(ilTest).lStartTime) Or (llSpotTime > tmRet(ilTest).lEndTime) Then
                                                    ilIncludeSpot = False
                                                End If
                                                If ilIncludeSpot Then
                                                    Select Case tmRet(ilTest).sDayType
                                                        Case "MF"
                                                            If (ilWeekDay = vbSaturday) Or (ilWeekDay = vbSunday) Then
                                                                ilIncludeSpot = False
                                                            End If
                                                        Case "Mo"
                                                            If ilWeekDay <> vbMonday Then
                                                                ilIncludeSpot = False
                                                            End If
                                                        Case "Tu"
                                                            If ilWeekDay <> vbTuesday Then
                                                                ilIncludeSpot = False
                                                            End If
                                                        Case "We"
                                                            If ilWeekDay <> vbWednesday Then
                                                                ilIncludeSpot = False
                                                            End If
                                                        Case "Th"
                                                            If ilWeekDay <> vbThursday Then
                                                                ilIncludeSpot = False
                                                            End If
                                                        Case "Fr"
                                                            If ilWeekDay <> vbFriday Then
                                                                ilIncludeSpot = False
                                                            End If
                                                        Case "Sa"
                                                            If ilWeekDay <> vbSaturday Then
                                                                ilIncludeSpot = False
                                                            End If
                                                        Case "Su"
                                                            If ilWeekDay <> vbSunday Then
                                                                ilIncludeSpot = False
                                                            End If
                                                    End Select
                                                End If
                                                If ilIncludeSpot Then
                                                    Exit For
                                                End If
                                            Next ilTest
                                            '8/26/19
                                            ilMaxProgCodeRepeatCount = 0
                                            For ilTest = 0 To UBound(tmRet) - 1 Step 1
                                                If slProgCode = tmRet(ilTest).sProgCode Then
                                                    ilMaxProgCodeRepeatCount = ilMaxProgCodeRepeatCount + 1
                                                End If
                                            Next ilTest
                                        Else
                                            ilIncludeSpot = False
                                        End If
                                        If ilIncludeSpot Then
                                            ilRepeatCount = 0
                                            For ilTest = 0 To UBound(tmRadarExportInfo) - 1 Step 1
                                                DoEvents
                                                If tmAstInfo(ilIndex).lSdfCode = tmRadarExportInfo(ilTest).lSdfCode Then
                                                    tmRadarExportInfo(ilTest).iRepeatCount = tmRadarExportInfo(ilTest).iRepeatCount + 1
                                                    ilRepeatCount = tmRadarExportInfo(ilTest).iRepeatCount
                                                    Exit For
                                                End If
                                            Next ilTest
                                            If ilRepeatCount >= ilMaxRepeats Then
                                                ilIncludeSpot = False
                                            End If
                                        End If
                                        If ilIncludeSpot Then
                                            'Write spot to file
                                            slRecordPRN = " " & slRecordNC
                                            slRecordCSV = slRecordNC & ","
                                            'Program code
                                            slRecordPRN = slRecordPRN & slProgCode
                                            slRecordCSV = slRecordCSV & slProgCode & ","
                                            'Cmml Unit
                                            ilUnitCount = -1
                                            For ilAst = 0 To UBound(tmSvAstInfo) - 1 Step 1
                                                DoEvents
                                                If gTimeToLong(tmAstInfo(ilIndex).sFeedTime, False) = gTimeToLong(tmSvAstInfo(ilAst).sFeedTime, False) And (DateValue(tmAstInfo(ilIndex).sFeedDate) = DateValue(tmSvAstInfo(ilAst).sFeedDate)) Then
                                                    For ilTest = 0 To UBound(tmRadarExportInfo) - 1 Step 1
                                                        If tmSvAstInfo(ilAst).lSdfCode = tmRadarExportInfo(ilTest).lSdfCode Then
                                                            If tmRadarExportInfo(ilTest).iInitUnitCount > ilUnitCount Then
                                                                ilUnitCount = tmRadarExportInfo(ilTest).iInitUnitCount
                                                            End If
                                                            Exit For
                                                        End If
                                                    Next ilTest
                                                End If
                                            Next ilAst
                                            If ilUnitCount = -1 Then
                                                ilUnitCount = 0
                                            Else
                                                ilUnitCount = ilUnitCount + 1
                                            End If
                                            'Reset for case where this spot is from the load factor
                                            For ilTest = 0 To UBound(tmRadarExportInfo) - 1 Step 1
                                                DoEvents
                                                If tmAstInfo(ilIndex).lSdfCode = tmRadarExportInfo(ilTest).lSdfCode Then
                                                    ilUnitCount = tmRadarExportInfo(ilTest).iInitUnitCount
                                            '        tmRadarExportInfo(ilTest).iRepeatCount = tmRadarExportInfo(ilTest).iRepeatCount + 1
                                                    ilRepeatCount = tmRadarExportInfo(ilTest).iRepeatCount
                                                    Exit For
                                                End If
                                            Next ilTest
                                            '4/25/19: Test if by progcode
                                            If smRADARMultiAir = "P" Then
                                                ilRepeatCount = ilProgCodeRepeatCount
                                                If ilIndex > LBound(tmAstInfo) Then
                                                    'Scan back counting number of spots
                                                    ilUnitCount = 0
                                                    ilAst = UBound(tmSvAstInfo) - 1
                                                    Do While ilAst >= LBound(tmSvAstInfo)
                                                        If (gDateValue(tmSvAstInfo(ilAst).sFeedDate) <> gDateValue(tmAstInfo(ilIndex).sFeedDate)) Or (gTimeToLong(tmSvAstInfo(ilAst).sFeedTime, False) <> gTimeToLong(tmAstInfo(ilIndex).sFeedTime, False)) Then
                                                            Exit Do
                                                        End If
                                                        ilUnitCount = ilUnitCount + 1
                                                        ilAst = ilAst - 1
                                                    Loop
                                                Else
                                                    ilUnitCount = 0
                                                End If
                                            ElseIf smRADARMultiAir = "A" Then
                                                slAirTime = Format$(tmAstInfo(ilIndex).sAirTime, "HH.MM A/P")
                                                If slAirTime <> slPrevAirTime Then
                                                    ilAirSpotCount = 0
                                                    ilAirBreakCount = ilAirBreakCount + 1
                                                Else
                                                    ilAirSpotCount = ilAirSpotCount + 1
                                                End If
                                                slPrevAirTime = slAirTime
                                                slAirDate = tmAstInfo(ilIndex).sAirDate
                                                If gDateValue(slAirDate) <> gDateValue(slPrevAirDate) Then
                                                    ilAirSpotCount = 0
                                                    ilAirBreakCount = 1
                                                End If
                                                ilUnitCount = ilAirSpotCount
                                                slPrevAirDate = slAirDate
                                            End If
                                            If (ilUnitCount < ilMaxSpots) Or (ilMaxSpots = -1) Then
                                                slCmmlUnit = Chr(Asc("A") + ilUnitCount)
                                                slRecordPRN = slRecordPRN & slCmmlUnit
                                                slRecordCSV = slRecordCSV & slCmmlUnit & ","
                                                'Clearance Designation
                                                Select Case ilRepeatCount
                                                    Case 0
                                                        'slRepeatCount = " "
                                                        slRepeatCount = "01"
                                                        tmRadarExportInfo(UBound(tmRadarExportInfo)).lSdfCode = tmAstInfo(ilIndex).lSdfCode
                                                        tmRadarExportInfo(UBound(tmRadarExportInfo)).iInitUnitCount = ilUnitCount
                                                        tmRadarExportInfo(UBound(tmRadarExportInfo)).iRepeatCount = 0
                                                        ReDim Preserve tmRadarExportInfo(0 To UBound(tmRadarExportInfo) + 1) As RADAREXPORTINFO
                                                    Case 1
                                                        'slRepeatCount = "S"
                                                        slRepeatCount = "02"
                                                    Case 2
                                                        'slRepeatCount = "T"
                                                        slRepeatCount = "03"
                                                    Case Else
                                                        'slRepeatCount = Chr(Asc("4") + ilRepeatCount - 3)
                                                        slRepeatCount = Trim$(Str(CInt("4") + ilRepeatCount - 3))
                                                        If ilRepeatCount < 9 Then
                                                            slRepeatCount = "0" & slRepeatCount
                                                        End If
                                                End Select
                                                '8/26/19
                                                If smRADARMultiAir = "P" Then
                                                    slRecordPRN = slRecordPRN & slRepeatCount + ilMaxProgCodeRepeatCount * (ilAirPlay - 1)
                                                    slRecordCSV = slRecordCSV & slRepeatCount + ilMaxProgCodeRepeatCount * (ilAirPlay - 1) & ","
                                                ElseIf smRADARMultiAir = "A" Then
                                                    slRecordPRN = slRecordPRN & ilAirBreakCount
                                                    slRecordCSV = slRecordCSV & ilAirBreakCount & ","
                                                Else
                                                    slRecordPRN = slRecordPRN & slRepeatCount
                                                    slRecordCSV = slRecordCSV & slRepeatCount & ","
                                                End If
                                                'Call Letters
                                                slCallLetters = Trim$(cprst!shttCallLetters)
                                                ilPos = InStr(1, slCallLetters, "-", vbTextCompare)
                                                If ilPos > 0 Then
                                                    slCall = Left$(slCallLetters, ilPos - 1)
                                                    slBand = Mid$(slCallLetters, ilPos + 1)
                                                    If Left$(slBand, 1) = "A" Then
                                                        slCallLetters = slCall
                                                    End If
                                                End If
                                                Do While Len(slCallLetters) < 7
                                                    slCallLetters = slCallLetters & " "
                                                Loop
                                                slRecordPRN = slRecordPRN & slCallLetters
                                                slRecordCSV = slRecordCSV & slCallLetters & ","
                                                Do While Len(slRecordPRN) < 19
                                                    slRecordPRN = slRecordPRN & " "
                                                Loop
                                                'Daylight Savine Time
                                                slZone = Left$(UCase$(Trim$(cprst!shttTimeZone)), 1)
                                                If rbcDaylight(0).Value Then
                                                    slZone = slZone & "D"
                                                Else
                                                    slZone = slZone & "S"
                                                End If
                                                slRecordPRN = slRecordPRN & slZone
                                                Do While Len(slRecordPRN) < 23
                                                    slRecordPRN = slRecordPRN & " "
                                                Loop
                                                slRecordCSV = slRecordCSV & slZone & ","
                                                'Schedule Day
                                                If rst_rht!rhtSchdDayType = "MF" Then
                                                    slRecordPRN = slRecordPRN & "MF1"
                                                    slRecordCSV = slRecordCSV & "MF1" & ","
                                                ElseIf rst_rht!rhtSchdDayType = "MS" Then
                                                    slRecordPRN = slRecordPRN & "MS1"
                                                    slRecordCSV = slRecordCSV & "MS1" & ","
                                                Else
                                                    Select Case Weekday(tmAstInfo(ilIndex).sFeedDate)
                                                        Case vbMonday
                                                            slRecordPRN = slRecordPRN & "MON"
                                                            slRecordCSV = slRecordCSV & "MON" & ","
                                                        Case vbTuesday
                                                            slRecordPRN = slRecordPRN & "TUE"
                                                            slRecordCSV = slRecordCSV & "TUE" & ","
                                                        Case vbWednesday
                                                            slRecordPRN = slRecordPRN & "WED"
                                                            slRecordCSV = slRecordCSV & "WED" & ","
                                                        Case vbThursday
                                                            slRecordPRN = slRecordPRN & "THU"
                                                            slRecordCSV = slRecordCSV & "THU" & ","
                                                        Case vbFriday
                                                            slRecordPRN = slRecordPRN & "FRI"
                                                            slRecordCSV = slRecordCSV & "FRI" & ","
                                                        Case vbSaturday
                                                            slRecordPRN = slRecordPRN & "SAT"
                                                            slRecordCSV = slRecordCSV & "SAT" & ","
                                                        Case vbSunday
                                                            slRecordPRN = slRecordPRN & "SUN"
                                                            slRecordCSV = slRecordCSV & "SUN" & ","
                                                    End Select
                                                End If
                                                Do While Len(slRecordPRN) < 37
                                                    slRecordPRN = slRecordPRN & " "
                                                Loop
                                                'Clearance type
                                                If tgStatusTypes(gGetAirStatus(tmAstInfo(ilIndex).iStatus)).iPledged = 2 Then
                                                    slClearType = "X"
                                                ElseIf rst_rht!rhtClearType = "A" Then
                                                    slClearType = cprst!attRadarClearType
                                                Else
                                                    If (cprst!attRadarClearType = "P") Or (cprst!attRadarClearType = "C") Then
                                                        slClearType = cprst!attRadarClearType
                                                    Else
                                                        slClearType = rst_rht!rhtClearType
                                                    End If
                                                End If
                                                slRecordPRN = slRecordPRN & slClearType
                                                Do While Len(slRecordPRN) < 39
                                                    slRecordPRN = slRecordPRN & " "
                                                Loop
                                                slRecordCSV = slRecordCSV & slClearType & ","
                                                'Declaration Indicator
                                                slRecordPRN = slRecordPRN & slIndicator
                                                slRecordCSV = slRecordCSV & slIndicator & ","
                                                'Time
                                                'Check if posted
                                                slTime = ""
                                                If tmAstInfo(ilIndex).iCPStatus = 1 Then
                                                    If (slClearType = "P") Or (slClearType = "C") Then
                                                        slTime = Format$(tmAstInfo(ilIndex).sAirTime, "HH.MM A/P")
                                                        If slTime = "12.00 A" Then
                                                            slTime = "12.00 M"
                                                        ElseIf slTime = "12.00 P" Then
                                                            slTime = "12.00 N"
                                                        End If
                                                        slRecordPRN = slRecordPRN & slTime
                                                    End If
                                                End If
                                                Do While Len(slRecordPRN) < 49
                                                    slRecordPRN = slRecordPRN & " "
                                                Loop
                                                slRecordCSV = slRecordCSV & slTime & ","
                                                If (slClearType = "P") Or (slClearType = "C") Then
                                                    Select Case Weekday(tmAstInfo(ilIndex).sAirDate)
                                                        Case vbMonday
                                                            slRecordPRN = slRecordPRN & "MON"
                                                            slRecordCSV = slRecordCSV & "MON" & ","
                                                        Case vbTuesday
                                                            slRecordPRN = slRecordPRN & "TUE"
                                                            slRecordCSV = slRecordCSV & "TUE" & ","
                                                        Case vbWednesday
                                                            slRecordPRN = slRecordPRN & "WED"
                                                            slRecordCSV = slRecordCSV & "WED" & ","
                                                        Case vbThursday
                                                            slRecordPRN = slRecordPRN & "THU"
                                                            slRecordCSV = slRecordCSV & "THU" & ","
                                                        Case vbFriday
                                                            slRecordPRN = slRecordPRN & "FRI"
                                                            slRecordCSV = slRecordCSV & "FRI" & ","
                                                        Case vbSaturday
                                                            slRecordPRN = slRecordPRN & "SAT"
                                                            slRecordCSV = slRecordCSV & "SAT" & ","
                                                        Case vbSunday
                                                            slRecordPRN = slRecordPRN & "SUN"
                                                            slRecordCSV = slRecordCSV & "SUN" & ","
                                                        Case Else
                                                            slRecordCSV = slRecordCSV & ","
                                                    End Select
                                                Else
                                                    slRecordCSV = slRecordCSV & ","
                                                End If
                                                Do While Len(slRecordPRN) < 70
                                                    slRecordPRN = slRecordPRN & " "
                                                Loop
                                                slRecordPRN = slRecordPRN & rst_rht!rhtRadarVehCode
                                                slRecordCSV = slRecordCSV & rst_rht!rhtRadarVehCode
                                                If ckcOutput(0).Value = vbChecked Then
                                                    Print #hmToPRN, slRecordPRN
                                                End If
                                                If ckcOutput(1).Value = vbChecked Then
                                                    Print #hmToCSV, slRecordCSV
                                                End If
                                                tmSvAstInfo(UBound(tmSvAstInfo)) = tmAstInfo(ilIndex)
                                                ReDim Preserve tmSvAstInfo(0 To UBound(tmSvAstInfo) + 1) As ASTINFO
                                            End If
                                        End If
                                    End If
                                    ilIndex = ilIndex + 1
                                    DoEvents
                                Loop
                            Next ilAirPlay
                        Else
                            'Posting not complete
                            gLogMsg "Posting Not Completed for: " & slVehicleName & " " & slNC & " " & slVC & Trim$(cprst!shttCallLetters), "RadarExportLog.Txt", False
                            lbcMsg.AddItem "Posting Not Completed for: " & slVehicleName & " " & slNC & " " & slVC & Trim$(cprst!shttCallLetters)
                        End If
                    End If
                    DoEvents
                Else
                    ''Error message
                    'gLogMsg "Week Not Found for: " & slVehicleName & " " & slNC & " " & slVC & " " & slInCallLetters, "RadarExportLog.Txt", False
                    'lbcMsg.AddItem "Week Not Found for: " & slVehicleName & " " & slNC & " " & slVC & " " & slInCallLetters
                End If
            End If
            rst_rht.MoveNext
        Loop
    Next ilLoop
    If Not ilNCVCFound Then
        'Error message
        gLogMsg "No Vehicle Programming Schedule found for: " & slNC & " " & slVC, "RadarExportLog.Txt", False
        lbcMsg.AddItem "No Vehicle Programming Schedule found for: " & slNC & " " & slVC
    End If
    rst_rht.Close
    mExportSpots = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarExport-mExportSpots"
    mExportSpots = False
    Exit Function
End Function

Private Sub mFillStations(slNC As String, slVC As String)
    Dim slDate As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilCloseATT As Integer
    Dim ilCloseShtt As Integer
    Dim ilVefCode As Integer
    Dim llRow As Long
    Dim slISDate As String
    Dim slIEDate As String
    
    On Error GoTo ErrHand
    lbcStations.Clear
    chkAllStation.Value = vbUnchecked
    If gIsDate(txtDate.Text) = False Then
        On Error GoTo 0
        Beep
        Exit Sub
    End If
    'Screen.MousePointer = vbHourglass
    slDate = txtDate.Text
    slISDate = txtIndicator(0).Text
    If slISDate <> "" Then
        If gIsDate(slISDate) = False Then
            On Error GoTo 0
            Beep
            Exit Sub
        End If
        slIEDate = txtIndicator(1).Text
        If slIEDate <> "" Then
            If gIsDate(slIEDate) = False Then
                On Error GoTo 0
                Beep
                Exit Sub
            End If
        Else
            On Error GoTo 0
            Beep
            Exit Sub
        End If
    End If
    ilCloseATT = False
    ilCloseShtt = False
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        SQLQuery = "SELECT * FROM rht WHERE (rhtVefCode = " & tgVehicleInfo(ilLoop).iCode & ")"
        Set rst_rht = gSQLSelectCall(SQLQuery)
        Do While Not rst_rht.EOF
            If (rst_rht!rhtRadarNetCode = slNC) And (rst_rht!rhtRadarVehCode = slVC) Then
                ilCloseATT = True
                ilVefCode = rst_rht!rhtVefCode
                SQLQuery = "SELECT * FROM att WHERE (attVefCode = " & ilVefCode
                SQLQuery = SQLQuery & " AND " & "(attOnAir <= '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "')"
                SQLQuery = SQLQuery & " AND " & "(attOffAir >= '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "') AND (attDropDate >= '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "')" & ")"
                Set rst_att = gSQLSelectCall(SQLQuery)
                Do While Not rst_att.EOF
                    ilCloseShtt = True
                    SQLQuery = "SELECT * FROM shtt WHERE (shttCode = " & rst_att!attshfcode & ")"
                    Set rst_Shtt = gSQLSelectCall(SQLQuery)
                    If Not rst_Shtt.EOF Then
                        slStr = rst_Shtt!shttCallLetters
                        llRow = SendMessageByString(lbcStations.hwnd, LB_FINDSTRING, -1, slStr)
                        If llRow < 0 Then
                            lbcStations.AddItem slStr
                            lbcStations.ItemData(lbcStations.NewIndex) = rst_Shtt!shttCode
                        End If
                    End If
                    rst_att.MoveNext
                Loop
                If slISDate <> "" Then
                    SQLQuery = "SELECT * FROM att WHERE (attVefCode = " & ilVefCode
                    SQLQuery = SQLQuery & " AND " & "(attOnAir <= '" & Format$(gAdjYear(slIEDate), sgSQLDateForm) & "')"
                    SQLQuery = SQLQuery & " AND " & "(attOffAir >= '" & Format$(gAdjYear(slISDate), sgSQLDateForm) & "') AND (attDropDate >= '" & Format$(gAdjYear(slISDate), sgSQLDateForm) & "')" & ")"
                    Set rst_att = gSQLSelectCall(SQLQuery)
                    Do While Not rst_att.EOF
                        ilCloseShtt = True
                        SQLQuery = "SELECT * FROM shtt WHERE (shttCode = " & rst_att!attshfcode & ")"
                        Set rst_Shtt = gSQLSelectCall(SQLQuery)
                        If Not rst_Shtt.EOF Then
                            slStr = rst_Shtt!shttCallLetters
                            llRow = SendMessageByString(lbcStations.hwnd, LB_FINDSTRING, -1, slStr)
                            If llRow < 0 Then
                                lbcStations.AddItem slStr
                                lbcStations.ItemData(lbcStations.NewIndex) = rst_Shtt!shttCode
                            End If
                        End If
                        rst_att.MoveNext
                    Loop
                End If
            End If
            rst_rht.MoveNext
        Loop
    Next ilLoop
    If ilCloseShtt Then
        rst_Shtt.Close
    End If
    If ilCloseATT Then
        rst_att.Close
    End If
    rst_rht.Close
    chkAllStation.Value = vbChecked
    On Error GoTo 0
    'Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarExport-mFillStation"
End Sub


Private Function mOutputSch(ilShttCode As Integer) As Integer
    Dim ilLoop As Integer
    Dim ilLoop1 As Integer
    Dim tlAstInfo As ASTINFO
    Dim slAdvt As String
    Dim slProd As String
    Dim slLen As String
    Dim slCart As String
    Dim slISCI As String
    Dim slCreative As String
    Dim slRCart As String
    Dim slRISCI As String
    Dim slRCreative As String
    Dim slRProd As String
    Dim llRCrfCsfCode As Long
    Dim ilRet As Integer
    Dim slRecord As String
    Dim ilLen As Integer
    Dim slDate As String
    Dim slTime As String
    Dim llTime As Long
    Dim llDate As Long
    Dim ilWindow As Integer
    Dim llVpf As Long
    Dim slShortTitle As String
    Dim slKey As String
    Dim ilFound As Integer
    Dim ilXRef As Integer
    Dim slVehicle As String
    Dim llVeh As Long
    
    On Error GoTo ErrHand:
'    If UBound(tmStarGuideAst) > 1 Then
'        ArraySortTyp fnAV(tmStarGuideAst(), 0), UBound(tmStarGuideAst), 0, LenB(tmStarGuideAst(0)), 0, LenB(tmStarGuideAst(0).sKey), 0
'    End If
'    ilLoop = 0
'    Do While ilLoop <= UBound(tmStarGuideAst) - 1
'        tlAstInfo = tmStarGuideAst(ilLoop).tAstInfo
'        slAdvt = "Missing"
'        slCart = ""
'        slISCI = ""
'        slCreative = ""
'        SQLQuery = "SELECT lstProd, lstCart, lstISCI, lstLen, adfName, cpfCreative"
'        SQLQuery = SQLQuery & " FROM (LST LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on lstCpfCode = cpfCode) LEFT OUTER JOIN ADF_Advertisers on lstadfCode = adfCode"
'        SQLQuery = SQLQuery & " WHERE lstCode =" & Str(tlAstInfo.lLstCode)
'        Set rst = gSQLSelectCall(SQLQuery)
'        If Not rst.EOF Then
'            If IsNull(rst!adfName) = True Then
'                slAdvt = "Missing"
'            Else
'                slAdvt = Trim$(rst!adfName)
'            End If
'            If IsNull(rst!lstProd) = True Then
'                slProd = ""
'            Else
'                slProd = Trim$(rst!lstProd)
'            End If
'            If IsNull(rst!lstCart) = True Then
'                slCart = ""
'            Else
'                slCart = Trim$(rst!lstCart)
'            End If
'            If IsNull(rst!lstISCI) = True Then
'                slISCI = ""
'            Else
'                slISCI = Trim$(rst!lstISCI)
'            End If
'            If IsNull(rst!cpfCreative) = True Then
'                slCreative = ""
'            Else
'                slCreative = Trim$(rst!cpfCreative)
'            End If
'            slLen = Trim$(Str$(rst!lstLen))
'        End If
'        '6/12/06- Check if any region copy defined for the spots
'        ilRet = gGetRegionCopy(tlAstInfo.iShttCode, tlAstInfo.lSdfCode, slRCart, slRProd, slRISCI, slRCreative, llRCrfCsfCode)
'        If ilRet Then
'            slCart = slRCart
'            slProd = slRProd
'            slISCI = slRISCI
'            slCreative = slRCreative
'        End If
'        'Get Short Title
'        slShortTitle = gGetShortTitle(tlAstInfo.lSdfCode)
'        slVehicle = ""
'        llVeh = gBinarySearchVef(CLng(tlAstInfo.iVefCode))
'        If llVeh <> -1 Then
'            slVehicle = Trim$(tgVehicleInfo(llVeh).sVehicle)
'        End If
'        ilFound = False
'        slKey = slShortTitle & "|" & slCart & "|" & slISCI & "|" & slCreative & "|" & slVehicle
'        For ilXRef = 0 To UBound(tmStarGuideXRef) - 1 Step 1
'            If StrComp(Trim$(tmStarGuideXRef(ilXRef).sKey), slKey, vbTextCompare) = 0 Then
'                ilFound = True
'                Exit For
'            End If
'        Next ilXRef
'        If Not ilFound Then
'            tmStarGuideXRef(UBound(tmStarGuideXRef)).sKey = slKey
'            tmStarGuideXRef(UBound(tmStarGuideXRef)).iVefCode = tlAstInfo.iVefCode
'            ReDim Preserve tmStarGuideXRef(0 To UBound(tmStarGuideXRef) + 1) As STARGUIDEXREF
'        End If
'        slISCI = gFileNameFilter(slISCI)
'        'Correct Date and Time
'        slDate = gAdjYear(tlAstInfo.sFeedDate)
'        slTime = tlAstInfo.sFeedTime
'        slDate = Format$(slDate, "yyyy-mm-dd")
'        slTime = Format$(slTime, "hh:mm:ss")
'        If slTime = "12M" Then
'            slTime = "00:00:00"
'        End If
'        slRecord = "EVENT: " & slDate & " " & slTime
'        'Window
'        llVpf = gBinarySearchVpf(CLng(tmStarGuideAst(ilLoop).iVefCode))
'        If llVpf <> -1 Then
'            ilWindow = tgVpfOptions(llVpf).lEDASWindow
'        Else
'            ilWindow = 400
'        End If
'        slRecord = slRecord & "," & Trim$(Str$(ilWindow))
'        If tmStarGuideAst(ilLoop).iBreakLen = 30 Then
'            slRecord = slRecord & ",0004"
'        ElseIf tmStarGuideAst(ilLoop).iBreakLen = 60 Then
'            slRecord = slRecord & ",0002"
'        Else
'            slRecord = slRecord & ",0001"
'        End If
'        slShortTitle = gFileNameFilter(slShortTitle)
'        slRecord = slRecord & "," & """" & slShortTitle & "(" & slISCI & ")" & ".mp2" & """"
'        ilLoop1 = ilLoop + 1
'        Do While ilLoop1 <= UBound(tmStarGuideAst) - 1
'            tlAstInfo = tmStarGuideAst(ilLoop1).tAstInfo
'            If DateValue(slDate) = DateValue(gAdjYear(tlAstInfo.sFeedDate)) Then
'                If gTimeToLong(slTime, False) = gTimeToLong(tlAstInfo.sFeedTime, False) Then
'                    slISCI = ""
'                    SQLQuery = "SELECT lstProd, lstCart, lstISCI, lstLen, adfName, cpfCreative"
'                    SQLQuery = SQLQuery & " FROM (LST LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on lstCpfCode = cpfCode) LEFT OUTER JOIN ADF_Advertisers on lstadfCode = adfCode"
'                    SQLQuery = SQLQuery & " WHERE lstCode =" & Str(tlAstInfo.lLstCode)
'                    Set rst = gSQLSelectCall(SQLQuery)
'                    If Not rst.EOF Then
'                        If IsNull(rst!lstCart) = True Then
'                            slCart = ""
'                        Else
'                            slCart = Trim$(rst!lstCart)
'                        End If
'                        If IsNull(rst!lstISCI) = True Then
'                            slISCI = ""
'                        Else
'                            slISCI = Trim$(rst!lstISCI)
'                        End If
'                        If IsNull(rst!cpfCreative) = True Then
'                            slCreative = ""
'                        Else
'                            slCreative = Trim$(rst!cpfCreative)
'                        End If
'                    End If
'                    '6/12/06- Check if any region copy defined for the spots
'                    ilRet = gGetRegionCopy(tlAstInfo.iShttCode, tlAstInfo.lSdfCode, slRCart, slRProd, slRISCI, slRCreative, llRCrfCsfCode)
'                    If ilRet Then
'                        slCart = slRCart
'                        slISCI = slRISCI
'                        slCreative = slRCreative
'                    End If
'                    slShortTitle = gGetShortTitle(tlAstInfo.lSdfCode)
'                    slVehicle = ""
'                    llVeh = gBinarySearchVef(CLng(tlAstInfo.iVefCode))
'                    If llVeh <> -1 Then
'                        slVehicle = Trim$(tgVehicleInfo(llVeh).sVehicle)
'                    End If
'                    ilFound = False
'                    slKey = slShortTitle & "|" & slCart & "|" & slISCI & "|" & slCreative & "|" & slVehicle
'                    For ilXRef = 0 To UBound(tmStarGuideXRef) - 1 Step 1
'                        If StrComp(Trim$(tmStarGuideXRef(ilXRef).sKey), slKey, vbTextCompare) = 0 Then
'                            ilFound = True
'                            Exit For
'                        End If
'                    Next ilXRef
'                    If Not ilFound Then
'                        tmStarGuideXRef(UBound(tmStarGuideXRef)).sKey = slKey
'                        tmStarGuideXRef(UBound(tmStarGuideXRef)).iVefCode = tlAstInfo.iVefCode
'                        ReDim Preserve tmStarGuideXRef(0 To UBound(tmStarGuideXRef) + 1) As STARGUIDEXREF
'                    End If
'                    slISCI = gFileNameFilter(slISCI)
'                    slShortTitle = gFileNameFilter(slShortTitle)
'                    slRecord = slRecord & "," & """" & slShortTitle & "(" & slISCI & ")" & ".mp2" & """"
'                    ilLoop = ilLoop + 1
'                    ilLoop1 = ilLoop1 + 1
'                Else
'                    Exit Do
'                End If
'            Else
'                Exit Do
'            End If
'        Loop
'        Print #hmToPRN, slRecord
'        DoEvents
'        ilLoop = ilLoop + 1
'    Loop
'    'Output EDAS
'    SQLQuery = "SELECT shttSerialNo1, shttSerialNo2"
'    SQLQuery = SQLQuery & " FROM SHTT"
'    SQLQuery = SQLQuery & " WHERE shttCode = " & Str(ilShttCode)
'    Set rst = gSQLSelectCall(SQLQuery)
'    If Not rst.EOF Then
'        If Trim$(rst!shttSerialNo1) <> "" Then
'            Print #hmToPRN, "ADDR: " & Trim$(rst!shttSerialNo1)
'        End If
'        If Trim$(rst!shttSerialNo2) <> "" Then
'            Print #hmToPRN, "ADDR: " & Trim$(rst!shttSerialNo2)
'        End If
'    End If
    mOutputSch = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarExport-mOutputSch"
    mOutputSch = False
End Function

Private Sub mFillVC(slNC As String)
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llRow As Long
    On Error GoTo ErrHand
    
    'Screen.MousePointer = vbHourglass
    lbcVehCodes.Clear
    chkAllVC.Value = vbUnchecked
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        SQLQuery = "SELECT * FROM rht WHERE (rhtVefCode = " & tgVehicleInfo(ilLoop).iCode & ")"
        Set rst_rht = gSQLSelectCall(SQLQuery)
        Do While Not rst_rht.EOF
            slStr = rst_rht!rhtRadarNetCode
            If slStr = slNC Then
                slStr = rst_rht!rhtRadarVehCode
                llRow = SendMessageByString(lbcVehCodes.hwnd, LB_FINDSTRING, -1, slStr)
                If llRow < 0 Then
                    lbcVehCodes.AddItem slStr
                End If
            End If
            rst_rht.MoveNext
        Loop
    Next ilLoop
    rst_rht.Close
    If lbcVehCodes.ListCount = 1 Then
        chkAllVC.Value = vbChecked
    End If
    On Error GoTo 0
    'Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarExport-mFillVC"
End Sub


Private Sub mSetVC()
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim slNC As String
    
    ilCount = 0
    For ilLoop = 0 To lbcNetworkCodes.ListCount - 1 Step 1
        If lbcNetworkCodes.Selected(ilLoop) Then
            ilCount = ilCount + 1
            If ilCount > 1 Then
                Exit For
            End If
            slNC = lbcNetworkCodes.List(ilLoop)
        End If
    Next ilLoop
    If ilCount = 1 Then
        lacTitleVC.Visible = True
        chkAllVC.Visible = True
        lbcVehCodes.Visible = True
        Screen.MousePointer = vbHourglass
        mFillVC slNC
        Screen.MousePointer = vbDefault
    Else
        lacTitleVC.Visible = False
        chkAllVC.Visible = False
        lbcVehCodes.Visible = False
        lacTitleStation.Visible = False
        chkAllStation.Visible = False
        lbcStations.Visible = False
    End If

End Sub

Private Sub mSetStations()
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim slVC As String
    Dim slNC As String
    
    ilCount = 0
    For ilLoop = 0 To lbcVehCodes.ListCount - 1 Step 1
        If lbcVehCodes.Selected(ilLoop) Then
            ilCount = ilCount + 1
            If ilCount > 1 Then
                Exit For
            End If
            slVC = lbcVehCodes.List(ilLoop)
        End If
    Next ilLoop
    If ilCount = 1 Then
        ilCount = 0
        For ilLoop = 0 To lbcNetworkCodes.ListCount - 1 Step 1
            If lbcNetworkCodes.Selected(ilLoop) Then
                ilCount = ilCount + 1
                If ilCount > 1 Then
                    Exit For
                End If
                slNC = lbcNetworkCodes.List(ilLoop)
            End If
        Next ilLoop
        If ilCount = 1 Then
            lacTitleStation.Visible = True
            chkAllStation.Visible = True
            lbcStations.Visible = True
            Screen.MousePointer = vbHourglass
            mFillStations slNC, slVC
            Screen.MousePointer = vbDefault
        Else
            lacTitleStation.Visible = False
            chkAllStation.Visible = False
            lbcStations.Visible = False
        End If
    Else
        lacTitleStation.Visible = False
        chkAllStation.Visible = False
        lbcStations.Visible = False
    End If

End Sub

Private Function mOpenRadarExportFile(slNC As String, slVC As String)
    Dim slToFile As String
    Dim ilRet As Integer
    Dim slName As String
    Dim slDateTime As String
    
    If imIncludeVehicleCodeInFileName Then
        slName = "R" & Trim$(txtRadarNo.Text) & "-" & slNC & "-" & slVC & "-" & Month(smDate) & "-" & Day(smDate)
    Else
        slName = "R" & Trim$(txtRadarNo.Text) & "-" & slNC & "-" & Month(smDate) & "-" & Day(smDate)
    End If
    If ckcOutput(0).Value = vbChecked Then
        slToFile = sgExportDirectory & slName & ".prn"
        ilRet = 0
        'On Error GoTo mOpenRadarExportFileErr:
        'slDateTime = FileDateTime(slToFile)
        ilRet = gFileExist(slToFile)
        If ilRet = 0 Then
            Kill slToFile
        End If
        'ilRet = 0
        'hmToPRN = FreeFile
        'Open slToFile For Output As hmToPRN
        ilRet = gFileOpen(slToFile, "Output", hmToPRN)
        If ilRet <> 0 Then
            Close #hmToPRN
            Screen.MousePointer = vbDefault
            gMsgBox "Open Error #" & Str$(Err.Numner) & slToFile, vbOKOnly, "Open Error"
            mOpenRadarExportFile = False
            Exit Function
        End If
    End If
    If ckcOutput(1).Value = vbChecked Then
        slToFile = sgExportDirectory & slName & ".csv"
        ilRet = 0
        'On Error GoTo mOpenRadarExportFileErr:
        'slDateTime = FileDateTime(slToFile)
        ilRet = gFileExist(slToFile)
        If ilRet = 0 Then
            Kill slToFile
        End If
        'ilRet = 0
        'hmToCSV = FreeFile
        'Open slToFile For Output As hmToCSV
        ilRet = gFileOpen(slToFile, "Output", hmToCSV)
        If ilRet <> 0 Then
            Close #hmToCSV
            Screen.MousePointer = vbDefault
            gMsgBox "Open Error #" & Str$(Err.Numner) & slToFile, vbOKOnly, "Open Error"
            mOpenRadarExportFile = False
            Exit Function
        End If
    End If
    mOpenRadarExportFile = True
    Exit Function
'mOpenRadarExportFileErr:
'    ilRet = Err
'    Resume Next
End Function

Private Sub mGetStationsNotReported(slNC As String, slVC As String)
    Dim slDate As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilCloseATT As Integer
    Dim ilCloseShtt As Integer
    Dim ilVefCode As Integer
    Dim llRow As Long
    Dim slAirDate As String
    Dim ilTitleLinePrinted As Integer
    
    On Error GoTo ErrHand
    slDate = txtDate.Text
    ilCloseATT = False
    ilCloseShtt = False
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        SQLQuery = "SELECT * FROM rht WHERE (rhtVefCode = " & tgVehicleInfo(ilLoop).iCode & ")"
        Set rst_rht = gSQLSelectCall(SQLQuery)
        Do While Not rst_rht.EOF
            If (rst_rht!rhtRadarNetCode = slNC) And (rst_rht!rhtRadarVehCode = slVC) Then
                ilTitleLinePrinted = False
                ilCloseATT = True
                ilVefCode = rst_rht!rhtVefCode
                SQLQuery = "SELECT * FROM att WHERE (attVefCode = " & ilVefCode
                SQLQuery = SQLQuery & " AND " & "(attOnAir > '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "')" & ")"
                Set rst_att = gSQLSelectCall(SQLQuery)
                Do While Not rst_att.EOF
                    ilCloseShtt = True
                    SQLQuery = "SELECT * FROM shtt WHERE (shttCode = " & rst_att!attshfcode & ")"
                    Set rst_Shtt = gSQLSelectCall(SQLQuery)
                    If Not rst_Shtt.EOF Then
                        slStr = rst_Shtt!shttCallLetters
                        llRow = SendMessageByString(lbcStations.hwnd, LB_FINDSTRING, -1, slStr)
                        If llRow < 0 Then
                            'Generate message
                            If Not ilTitleLinePrinted Then
                                gLogMsg "The following Stations Not Reported on: " & Trim$(tgVehicleInfo(ilLoop).sVehicle), "RadarExportLog.Txt", False
                                lbcMsg.AddItem "Stations on " & Trim$(tgVehicleInfo(ilLoop).sVehicle) & " Not Reported"
                                ilTitleLinePrinted = True
                            End If
                            slAirDate = Format$(rst_att!attOnAir, sgShowDateForm)
                            gLogMsg "          " & slStr & " with On Air date " & slAirDate, "RadarExportLog.Txt", False
                            lbcMsg.AddItem slStr & " starting " & slAirDate & " Not Reported"
                        End If
                    End If
                    rst_att.MoveNext
                Loop
            End If
            rst_rht.MoveNext
        Loop
    Next ilLoop
    If ilCloseShtt Then
        rst_Shtt.Close
    End If
    If ilCloseATT Then
        rst_att.Close
    End If
    rst_rht.Close
    chkAllStation.Value = vbChecked
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmRadarExport-mGetStationsNotReported"
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtIndicator_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtIndicator_KeyPress(Index As Integer, KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
