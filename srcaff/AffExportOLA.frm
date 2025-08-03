VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmExportOLA 
   Caption         =   "Export OLA"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "AffExportOLA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   9615
   Begin VB.TextBox edcMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   2325
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   510
      Visible         =   0   'False
      Width           =   4770
   End
   Begin VB.CheckBox ckcGenerate 
      Caption         =   "Generate OLA Export"
      Height          =   195
      Index           =   1
      Left            =   2025
      TabIndex        =   1
      Top             =   105
      Value           =   1  'Checked
      Width           =   2310
   End
   Begin VB.CheckBox ckcGenerate 
      Caption         =   "Import Station Info"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Value           =   1  'Checked
      Width           =   1740
   End
   Begin VB.PictureBox pbcArial 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9300
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   1005
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   7770
      TabIndex        =   5
      Top             =   465
      Width           =   1665
   End
   Begin VB.CheckBox ckcGenCSV 
      Caption         =   "Generate CSV file along with the XML file"
      Height          =   210
      Left            =   4500
      TabIndex        =   2
      Top             =   105
      Width           =   3405
   End
   Begin VB.TextBox txtStationInfo 
      Height          =   300
      Left            =   1500
      TabIndex        =   4
      Top             =   495
      Width           =   6060
   End
   Begin VB.TextBox txtNumberDays 
      Height          =   360
      Left            =   3915
      TabIndex        =   9
      Text            =   "1"
      Top             =   1170
      Width           =   405
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   2400
      ItemData        =   "AffExportOLA.frx":08CA
      Left            =   4650
      List            =   "AffExportOLA.frx":08CC
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1920
      Width           =   4755
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   2205
      ItemData        =   "AffExportOLA.frx":08CE
      Left            =   120
      List            =   "AffExportOLA.frx":08D0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox txtDate 
      Height          =   360
      Left            =   1560
      TabIndex        =   7
      Top             =   1170
      Width           =   1320
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8655
      Top             =   870
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5265
      FormDesignWidth =   9615
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   5910
      TabIndex        =   15
      Top             =   4710
      Width           =   1665
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7755
      TabIndex        =   16
      Top             =   4710
      Width           =   1665
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   855
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lacFileInfo 
      Caption         =   $"AffExportOLA.frx":08D2
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   120
      TabIndex        =   20
      Top             =   810
      Width           =   7440
   End
   Begin VB.Label lacStationInfo 
      Caption         =   "Station Info File"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   525
      Width           =   1245
   End
   Begin VB.Label lacDays 
      Caption         =   "# of Days"
      Height          =   255
      Left            =   3030
      TabIndex        =   8
      Top             =   1215
      Width           =   795
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   120
      TabIndex        =   17
      Top             =   4650
      Width           =   5580
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   6015
      TabIndex        =   13
      Top             =   1560
      Width           =   1965
   End
   Begin VB.Label lacTitle1 
      Alignment       =   2  'Center
      Caption         =   "Vehicles"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   3885
   End
   Begin VB.Label lacStartDate 
      Caption         =   "Export Start Date"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1215
      Width           =   1395
   End
End
Attribute VB_Name = "FrmExportOLA"
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

'  All the code taken from ExportWegener.
'  In Wegener, all spots within a Region Break are Exported.  In OLA, only the Region spots exported
'  Break #'s: In wegener, break number reset for each day.  In PLA, reset each hour
'  The XML is different
'
Option Explicit
Option Compare Text

Private hmFrom As Integer
'Private smFields(1 To 31) As String
Private smFields(0 To 30) As String

Private smAdminPhone As String
Private smAdminEMail As String

Private smDate As String     'Export Date
Private imNumberDays As Integer
Private imVefCode As Integer
Private imAdfCode As Integer
Private smVefName As String
Private imAllClick As Integer
Private imAllStationClick As Integer
Private imExporting As Integer
Private smExportPath As String
Private imTerminate As Integer
Private lmMaxWidth As Long
Private hmCSV As Integer
Private smGrpFileName As String
Private smVehicleGroupName As String
Private smCustomGroupName As String
Private imCustomGroupNo As Integer
Private tmRegionBreakSpots() As REGIONBREAKSPOTS
Private tmTempRegionBreakSpots() As REGIONBREAKSPOTS
Private tmRegionDefinition() As REGIONDEFINITION
Private tmSplitCategoryInfo() As SPLITCATEGORYINFO
Private tmMergeRegionDefinition() As REGIONDEFINITION
Private tmMergeSplitCategoryInfo() As SPLITCATEGORYINFO
Private tmCustomGroupNames() As OLACUSTOMGROUPNAMES
Private tmCategoryName() As OLACATEGORYNAME
Private tmUniqueGroupNames() As OLAUNIQUEGROUPNAMES
Private imStartHourNumber As Integer
'Private imShttCodes() As Integer
Private hmCsf As Integer
Private imCsfOpenStatus As Integer
Private cprst As ADODB.Recordset
Private lst_rst As ADODB.Recordset
Private vff_rst As ADODB.Recordset
Private rsf_rst As ADODB.Recordset
Private cpf_rst As ADODB.Recordset
Private adrst As ADODB.Recordset
Private err_rst As ADODB.Recordset
'Dan M 11/01/10 search for xml.ini once and store in new variable
Private smIniPathFileName As String







Private Sub mFillVehicle()
    Dim iLoop As Integer
    Dim llVpf As Long
    
    lbcVehicles.Clear
    lbcMsg.Clear
    chkAll.Value = 0
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        llVpf = gBinarySearchVpf(CLng(tgVehicleInfo(iLoop).iCode))
        If llVpf <> -1 Then
            If tgVpfOptions(llVpf).sOLAExport = "Y" Then
                lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
            End If
        End If
    Next iLoop
End Sub

Private Sub chkAll_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcVehicles.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehicles.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllClick = False
    End If

End Sub

Private Sub ckcGenerate_Click(Index As Integer)
    If Index = 0 Then
        If ckcGenerate(0).Value = vbChecked Then
            lacStationInfo.Enabled = True
            txtStationInfo.Enabled = True
            cmcBrowse.Enabled = True
        Else
            lacStationInfo.Enabled = False
            txtStationInfo.Enabled = False
            cmcBrowse.Enabled = False
        End If
    ElseIf Index = 1 Then
        If ckcGenerate(1).Value = vbChecked Then
            lacStartDate.Enabled = True
            txtDate.Enabled = True
            lacDays.Enabled = True
            txtNumberDays.Enabled = True
            ckcGenCSV.Enabled = True
            lbcVehicles.Enabled = True
        Else
            lacStartDate.Enabled = False
            txtDate.Enabled = False
            lacDays.Enabled = False
            txtNumberDays.Enabled = False
            ckcGenCSV.Enabled = False
            lbcVehicles.Enabled = False
        End If
    End If

End Sub

Private Sub ckcGenerate_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcBrowse_Click()
    'sgGetPath = txtStationInfo.text
    'frmGetPath.Show vbModal
    'If igGetPath = 0 Then
    '    txtStationInfo.text = sgGetPath
    'End If
    Dim slCurDir As String
    
    slCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    ' Set filters
    CommonDialog1.Filter = "CSV Files (*.csv)|*.csv"
    ' Specify default filter
    CommonDialog1.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    txtStationInfo.Text = Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub cmdExport_Click()
    Dim slNowDate As String
    Dim ilRet As Integer
    Dim slExportType As String
    Dim slXMLFileName As String
    Dim slOutputType As String
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    
    On Error GoTo ErrHand
    
    If imExporting = True Then
        Exit Sub
    End If
    imExporting = True
    
    lbcMsg.Clear
    slNowDate = Format$(gNow(), "m/d/yy")
    If ckcGenerate(1).Value = vbChecked Then
        If lbcVehicles.ListIndex < 0 Then
            imExporting = False
            Exit Sub
        End If
        If txtDate.Text = "" Then
            imExporting = False
            gMsgBox "Date must be specified.", vbOKOnly
            txtDate.SetFocus
            Exit Sub
        End If
        If gIsDate(txtDate.Text) = False Then
            imExporting = False
            Beep
            gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            txtDate.SetFocus
            Exit Sub
        Else
            smDate = Format(txtDate.Text, sgShowDateForm)
        End If
        imNumberDays = Val(txtNumberDays.Text)
        If imNumberDays <= 0 Then
            imExporting = False
            gMsgBox "Number of days must be specified.", vbOKOnly
            txtNumberDays.SetFocus
            Exit Sub
        End If
        Select Case Weekday(gAdjYear(smDate))
            Case vbMonday
                If imNumberDays > 7 Then
                    gMsgBox "Number of days can not exceed 7.", vbOKOnly
                    txtNumberDays.SetFocus
                End If
            Case vbTuesday
                If imNumberDays > 6 Then
                    gMsgBox "Number of days can not exceed 6.", vbOKOnly
                    txtNumberDays.SetFocus
                End If
            Case vbWednesday
                If imNumberDays > 5 Then
                    gMsgBox "Number of days can not exceed 5.", vbOKOnly
                    txtNumberDays.SetFocus
                End If
            Case vbThursday
                If imNumberDays > 4 Then
                    gMsgBox "Number of days can not exceed 4.", vbOKOnly
                    txtNumberDays.SetFocus
                End If
            Case vbFriday
                If imNumberDays > 3 Then
                    gMsgBox "Number of days can not exceed 3.", vbOKOnly
                    txtNumberDays.SetFocus
                End If
            Case vbSaturday
                If imNumberDays > 2 Then
                    gMsgBox "Number of days can not exceed 2.", vbOKOnly
                    txtNumberDays.SetFocus
                End If
            Case vbSunday
                If imNumberDays > 1 Then
                    gMsgBox "Number of days can not exceed 1.", vbOKOnly
                    txtNumberDays.SetFocus
                End If
        End Select
        If DateValue(gAdjYear(smDate)) <= DateValue(gAdjYear(slNowDate)) Then
            imExporting = False
            Beep
            gMsgBox "Date must be after today's date " & slNowDate, vbCritical
            txtDate.SetFocus
            Exit Sub
        End If
    End If
    If (txtStationInfo.Text = "") And (ckcGenerate(0).Value = vbChecked) Then
        imExporting = False
        gMsgBox "Import file must be specified.", vbOKOnly
        txtStationInfo.SetFocus
        Exit Sub
    End If
    'If (rbcSpots(0).Value = False) And (rbcSpots(1).Value = False) Then
    '    Beep
    '    gMsgBox "Please Specify Export Spots Type.", vbCritical
    '    Exit Sub
    'End If
    Screen.MousePointer = vbHourglass
    If ckcGenerate(1).Value = vbChecked Then
        If Not gPopCopy(smDate, "Export OLA") Then
            imExporting = False
            Exit Sub
        End If
    End If
    imExporting = True
    If ckcGenerate(0).Value = vbChecked Then
        edcMsg.Text = "Reading Station Info...."
        edcMsg.Visible = True
        DoEvents
        ilRet = mReadStationReceiverRecords()
        edcMsg.Visible = False
        DoEvents
        If ilRet <> 0 Then
            imExporting = False
            Screen.MousePointer = vbDefault
            If ilRet = 1 Then
                Exit Sub
            ElseIf ilRet = 2 Then
                Exit Sub
            ElseIf ilRet = 3 Then
                Exit Sub
            Else
                ilRet = gMsgBox("Some Stations Not Defined within the Affiliate system, Continue anyway", vbYesNo + vbQuestion, "Information")
                If ilRet = vbNo Then
                    Exit Sub
                End If
            End If
            imExporting = True
            Screen.MousePointer = vbHourglass
        End If
        edcMsg.Text = "Retrieving Market, Format, State and Time Zone info...."
        edcMsg.Visible = True
        DoEvents
        ilRet = gPopStates()
        ilRet = gPopFormats()
        ilRet = gPopTimeZones()
        ilRet = gPopMarkets()
        ilRet = gPopStations()
        'edcMsg.text = "Updating Station Info...."
        'edcMsg.Visible = True
        'DoEvents
        'ilRet = mUpdateShttUsedForOLA()
        edcMsg.Visible = False
        DoEvents
        'If Not ilRet Then
        '    Screen.MousePointer = vbDefault
        '    ilRet = gMsgBox("Unable to set Used for OLA with all Stations, Continue anyway", vbYesNo + vbQuestion, "Information")
        '    If ilRet = vbNo Then
        '        imExporting = False
        '        Exit Sub
        '    End If
        'End If
    End If
    Screen.MousePointer = vbHourglass
    On Error GoTo 0
    lacResult.Caption = ""
    If ckcGenerate(1).Value = vbChecked Then
        'If rbcSpots(0).Value = True Then
            slExportType = "!! Exporting All Spots, "
        'Else
        '    slExportType = "!! Exporting Regional Spots, "
        'End If
        'gLogMsg slExportType & "Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", "OLAExportLog.Txt", False
        ilRet = 0
        slToFile = sgMsgDirectory & "OLAExportLog.Txt"
        On Error GoTo mFileErr
        'slDateTime = FileDateTime(slToFile)
        ilRet = gFileExist(slToFile)
        If ilRet = 0 Then
            slDateTime = gFileDateTime(slToFile)
            slFileDate = Format$(slDateTime, "m/d/yy")
            If DateValue(gAdjYear(slFileDate)) = DateValue(gAdjYear(slNowDate)) Then  'Append
                gLogMsg slExportType & "Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", "OLAExportLog.Txt", False
            Else
                gLogMsg slExportType & "Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", "OLAExportLog.Txt", True
            End If
        Else
            gLogMsg slExportType & "Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", "OLAExportLog.Txt", False
        End If
        On Error GoTo ErrHand
        ilRet = mExportSpots()
        gCloseRegionSQLRst
        edcMsg.Visible = False
        DoEvents
        If (ilRet = False) Then
            gLogMsg "** Terminated - mExportSpots returned False **", "OLAExportLog.Txt", False
            imExporting = False
            Screen.MousePointer = vbDefault
            cmdCancel.SetFocus
            Exit Sub
        End If
        If imTerminate Then
            gLogMsg "** User Terminated **", "OLAExportLog.Txt", False
            imExporting = False
            Screen.MousePointer = vbDefault
            cmdCancel.SetFocus
            Exit Sub
        End If
        On Error GoTo ErrHand:
        'Print #hmMsg, "** Completed Export of StarGuide: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
        gLogMsg "** Completed Export of OLA **", "OLAExportLog.Txt", False
        'Close #hmMsg
        If slOutputType <> "T" Then
            lacResult.Caption = "Exports placed into: " & smExportPath
        Else
            lacResult.Caption = ""
        End If
    End If
    imExporting = False
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    gLogMsg "", "OLAExportLog.Txt", False
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmExportOLA-cmdExport"
    imExporting = False
    Exit Sub
mFileErr:
    ilRet = Err.Number
    Resume Next
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    txtDate.Text = ""
    Unload FrmExportOLA
End Sub


Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    edcMsg.Move (Me.Width - edcMsg.Width) / 2, txtStationInfo.Top   '(Me.Height - Msg.Height) / 2

End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    Dim slSvIniPathFileName As String
        
    Screen.MousePointer = vbHourglass
    FrmExportOLA.Caption = "Export OLA - " & sgClientName
    smDate = gObtainNextMonday(Format$(gNow(), sgShowDateForm))
    txtDate.Text = smDate
    txtNumberDays.Text = 7
    imAllClick = False
    imAllStationClick = False
    imTerminate = False
    imExporting = False
    'txtStationInfo.text = Left$(sgImportDirectory, Len(sgImportDirectory) - 1)
    mFillVehicle
    chkAll.Value = vbChecked
    mGetAdminInfo
    slSvIniPathFileName = sgIniPathFileName
    'dan m 11/01/10 xml.ini gotten from general procedure; checking different folders
    'sgIniPathFileName = sgStartupDirectory & "\XML.Ini"
    smIniPathFileName = gXmlIniPath()
    sgIniPathFileName = smIniPathFileName
    If Not gLoadOption("OLA", "GrpFilePath", smGrpFileName) Then
        smGrpFileName = ""
    End If
    If Not gLoadOption("OLA", "Export", smExportPath) Then
        smExportPath = sgExportDirectory
    End If
    smExportPath = gSetPathEndSlash(smExportPath, True)
    sgIniPathFileName = slSvIniPathFileName
    ilRet = gPopAvailNames()
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Resize()
    lacFileInfo.FontSize = 7
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    Erase tmRegionBreakSpots
    Erase tmTempRegionBreakSpots
    Erase tmRegionDefinition
    Erase tmSplitCategoryInfo
    Erase tmMergeRegionDefinition
    Erase tmMergeSplitCategoryInfo
    Erase tmCustomGroupNames
    Erase tmCategoryName
    Erase tmUniqueGroupNames
    'Erase imShttCodes
    cprst.Close
    lst_rst.Close
    vff_rst.Close
    rsf_rst.Close
    cpf_rst.Close
    adrst.Close
    err_rst.Close
    Set FrmExportOLA = Nothing
End Sub



Private Sub lbcVehicles_Click()
    Dim iLoop As Integer
    Dim iCount As Integer
    
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = vbChecked Then
        imAllClick = True
        chkAll.Value = vbUnchecked
        imAllClick = False
    End If
    For iLoop = 0 To lbcVehicles.ListCount - 1 Step 1
        If lbcVehicles.Selected(iLoop) Then
            imVefCode = lbcVehicles.ItemData(iLoop)
            iCount = iCount + 1
            If iCount > 1 Then
                Exit For
            End If
        End If
    Next iLoop
End Sub

Private Sub txtDate_Change()
    lbcMsg.Clear
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
End Sub

Private Sub txtDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Function mExportSpots() As Integer
    'Export all spots with its general copy for the specified vehicle and days
    'Each vehicle will create a separate export file.
    'All days will be within the same export file
    'The spots are obtained from LST instead of AST as OLA will create the spots for each station
    'If any spot within a break has region copy, then all spots within that break must be export
    'along with each region definition for the spot.
    'This part of the export must be after all the general copy is exported for the vehicle and days.
    '
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim slSDate As String
    Dim slEDate As String
    Dim slWeekNo As String
    Dim slYearNo As String
    Dim slVehName As String
    Dim slVehGroupName As String
    Dim slVehExportID As String
    Dim llODate As Long
    Dim slDate As String
    Dim llDate As Long
    Dim slXMLFileName As String
    Dim slCSVFileName As String
    Dim slNowDT As String
    Dim slTimeID As String
    Dim llEventID As Long
    Dim slEventID As String
    Dim llBreakNo As Long
    Dim llLogTime As Long
    Dim llLstLogDate As Long
    Dim llLstLogTime As Long
    Dim ilPos As Integer
    Dim slOutputType As String
    Dim ilLoop As Integer
    Dim ilAnyExports As Integer
    Dim ilPositionNo As Integer
    Dim slCSVRecord As String
    Dim llAdf As Long
    Dim slAdvtName As String
    Dim slGrpName As String
    Dim slStdStartDate As String
    Dim slHour As String
    Dim slVefCode As String
    Dim blSpotOk As Boolean
    Dim ilAnf As Integer
    
    On Error GoTo ErrHand
    mExportSpots = True
    imCsfOpenStatus = mOpenCSF()

    ilAnyExports = False
    slSDate = smDate
    slEDate = DateAdd("d", imNumberDays - 1, smDate)
    'slWeekNo = Format(slSDate, "ww")
    'If Len(slWeekNo) = 1 Then
    '    slWeekNo = "0" & slWeekNo
    'End If
    'slYearNo = Format(slSDate, "yy")
    'If Len(slYearNo) = 1 Then
    '    slYearNo = "0" & slYearNo
    'End If
    slStdStartDate = gObtainYearStartDate(slSDate)
    slWeekNo = Trim$(Str$(DateDiff("ww", slStdStartDate, slSDate, vbMonday) + 1))
    If Len(slWeekNo) = 1 Then
        slWeekNo = "0" & slWeekNo
    End If
    slYearNo = Format(gObtainEndStd(slSDate), "yy")
    If Len(slYearNo) = 1 Then
        slYearNo = "0" & slYearNo
    End If
    
    llEventID = 0
    imCustomGroupNo = 0
    For ilVef = 0 To lbcVehicles.ListCount - 1 Step 1
        If lbcVehicles.Selected(ilVef) Then
            DoEvents
            If imTerminate Then
                mAddMsgToList "User Cancelled Export"
                mExportSpots = False
                Exit For
            End If
            slVehName = Trim$(lbcVehicles.List(ilVef))
            imVefCode = lbcVehicles.ItemData(ilVef)
            smCustomGroupName = "CG" & slWeekNo
            slVefCode = Trim$(Str$(imVefCode))
            Do While Len(slVefCode) < 4
                slVefCode = "0" & slVefCode
            Loop
            smCustomGroupName = smCustomGroupName & slVefCode
            SQLQuery = "SELECT * "
            SQLQuery = SQLQuery + " FROM VFF_Vehicle_Features"
            SQLQuery = SQLQuery + " WHERE (vffVefCode = " & imVefCode & ")"
            Set vff_rst = gSQLSelectCall(SQLQuery)
            If Not vff_rst.EOF Then
                imStartHourNumber = -1
                edcMsg.Text = "Generating General Schedule for " & slVehName & "..."
                edcMsg.Visible = True
                DoEvents
                ilAnyExports = True
                slVehGroupName = Trim$(vff_rst!VffGroupName)
                slVehExportID = Trim$(vff_rst!VffOLAExportID)
                smVehicleGroupName = slVehGroupName
                slXMLFileName = slVehExportID & "~SCH~" & slWeekNo & slYearNo & ".XML"
                slCSVFileName = slVehExportID & "~SCH~" & slWeekNo & slYearNo & ".CSV"
                slOutputType = "F"
                '6808
                If Not gDeleteFile(sgExportDirectory & "OLASpot.XML") Then
                    mAddMsgToList "Could not delete file in mExportSpots before writing.  Appended."
                End If
                'User wamts CrLf.  The ulity is adding a Cr so only add the Lf in this code
                'ilRet = csiXMLStart(sgStartupDirectory & "\xml.ini", "OLASched", slOutputType, sgExportDirectory & slXMLFileName, sgCRLF)
                ' Dan M 11/01/10 use smIniPathFileName that is created at formload
               ' ilRet = csiXMLStart(sgStartupDirectory & "\xml.ini", "OLA", slOutputType, sgExportDirectory & "OLASpot.XML", sgCRLF)
                '6807
                'ilRet = csiXMLStart(smIniPathFileName, "OLA", slOutputType, sgExportDirectory & "OLASpot.XML", sgCRLF)
                ilRet = csiXMLStart(smIniPathFileName, "OLA", slOutputType, sgExportDirectory & "OLASpot.XML", sgCRLF, "")
                ilRet = csiXMLSetMethod("", "", "", "trafficOutput")
                DoEvents
                csiXMLData "OT", "fileHeader", ""
                csiXMLData "CD", "filename", slVehExportID & slWeekNo & slYearNo
                csiXMLData "CD", "description", gXMLNameFilter(Trim$(lbcVehicles.List(ilVef))) & " #" & slWeekNo & "-" & slYearNo
                slNowDT = Now
                csiXMLData "OT", "creationDate", "" 'Format$(slNowDT, "yyyy-mm-ddThh:mm:ss")
                csiXMLData "CD", "day", Format$(slNowDT, "d")
                csiXMLData "CD", "month", Format$(slNowDT, "m")
                csiXMLData "CD", "year", Format$(slNowDT, "yyyy")
                slHour = Format$(slNowDT, "hh")
                'slHour = Trim$(Str$(Val(slHour) + 1))
                'If Len(slHour) = 1 Then
                '    slHour = "0" & slHour
                'End If
                csiXMLData "CD", "hour", slHour 'Format$(slNowDT, "hh")
                csiXMLData "CD", "minute", Format$(slNowDT, "nn")
                csiXMLData "CD", "seconds", Format$(slNowDT, "ss")
                csiXMLData "CT", "creationDate", ""
                csiXMLData "CD", "contactName", "Dial Global Affiliate Relations"
                csiXMLData "CD", "contactPhoneNumber", smAdminPhone
                csiXMLData "CD", "contactEmail", smAdminEMail
                csiXMLData "CT", "fileHeader", ""
                
                csiXMLData "OT", "affidavit", ""
                csiXMLData "CD", "vehicleName", gXMLNameFilter(slVehName)
                csiXMLData "CD", "cycleCode", slWeekNo & slYearNo
                csiXMLData "CD", "headComments", ""
                csiXMLData "CD", "footComments", ""
                
                slTimeID = Timer
                ilPos = InStr(1, slTimeID, ".", vbTextCompare)
                If ilPos > 0 Then
                    slTimeID = Left(slTimeID, ilPos - 1) & Mid(slTimeID, ilPos + 1)
                End If
                slTimeID = Left$(slTimeID, 3)
                
                If ckcGenCSV.Value = vbChecked Then
                    gLogMsgWODT "ON", hmCSV, sgExportDirectory & slCSVFileName
                    slCSVRecord = "Vehicle " & gAddQuotes(Trim$(lbcVehicles.List(ilVef)))
                    gLogMsgWODT "W", hmCSV, slCSVRecord
                    slCSVRecord = "Start Date " & slSDate & " End Date " & slEDate
                    gLogMsgWODT "W", hmCSV, slCSVRecord
                    slCSVRecord = "Creation Date and Time " & Format$(slNowDT, "yyyy-mm-ddThh:mm:ss")
                    gLogMsgWODT "W", hmCSV, slCSVRecord
                    slCSVRecord = "Message ID " & slWeekNo & slYearNo & slTimeID
                    gLogMsgWODT "W", hmCSV, slCSVRecord
                    slCSVRecord = "LogDate,LogTime,Break#,Position#,Contract#,Advertiser,Line#,Length,ISCI,Region Names..."
                    gLogMsgWODT "W", hmCSV, slCSVRecord
                End If
                slCSVRecord = ""
                ReDim tmRegionBreakSpots(0 To 0) As REGIONBREAKSPOTS
                llODate = -1
                SQLQuery = "SELECT * FROM lst "
                SQLQuery = SQLQuery + " WHERE (lstLogVefCode = " & imVefCode
                SQLQuery = SQLQuery + " AND lstBkoutLstCode = 0"
                '3/9/16: Fix the filter
                'SQLQuery = SQLQuery + " AND lstStatus < 20" 'Bypass MG/Bonus
                SQLQuery = SQLQuery + " AND Mod(lstStatus, 100) < " & ASTEXTENDED_MG 'Bypass MG/Bonus
                SQLQuery = SQLQuery + " AND (lstLogDate >= '" & Format$(slSDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(slEDate, sgSQLDateForm) & "')" & ")"
                SQLQuery = SQLQuery + " ORDER BY lstLogDate, lstLogTime, lstBreakNo, lstPositionNo"
                Set lst_rst = gSQLSelectCall(SQLQuery)
                Do While Not lst_rst.EOF
                    blSpotOk = True
                    ilAnf = gBinarySearchAnf(lst_rst!lstAnfCode)
                    If ilAnf <> -1 Then
                        If tgAvailNamesInfo(ilAnf).sAudioExport = "N" Then
                            blSpotOk = False
                        End If
                    End If
                    If (blSpotOk) And ((UCase(Trim$(lst_rst!lstZone)) = "EST") Or (Trim$(lst_rst!lstZone) = "")) Then
                        slDate = Format$(lst_rst!lstLogDate, sgShowDateForm)
                        llDate = DateValue(gAdjYear(slDate))
                        If llODate <> DateValue(lst_rst!lstLogDate) Then
                            llODate = llDate
                            llBreakNo = 0
                            ilPositionNo = 0
                            llLogTime = -1
                        End If
                        llLstLogDate = DateValue(gAdjYear(Format$(lst_rst!lstLogDate, sgShowDateForm)))
                        llLstLogTime = gTimeToLong(Format$(lst_rst!lstLogTime, sgShowTimeWSecForm), False)
                        If llLogTime <> llLstLogTime Then
                            If llLogTime <> -1 Then
                                If Format$(gLongToTime(llLogTime), "hh") <> Format$(gLongToTime(llLstLogTime), "hh") Then
                                    llBreakNo = 0
                                End If
                            End If
                            llLogTime = llLstLogTime
                            llEventID = llEventID + 1
                            slEventID = Trim$(Str$(llEventID))
                            Do While Len(slEventID) < 5
                                slEventID = "0" & slEventID
                            Loop
                            llBreakNo = llBreakNo + 1
                            ilPositionNo = 0
                        End If
                        ilPositionNo = ilPositionNo + 1
                        'Output spot
                        mCreateXMLSpot "NAT_000", -1, lst_rst!lstCode, llBreakNo, ilPositionNo
                    End If
                    lst_rst.MoveNext
                Loop
                gLogMsgWODT "C", hmCSV, ""
                'Output Breaks with region spots
                ReDim tmCustomGroupNames(0 To 0) As OLACUSTOMGROUPNAMES
                ReDim tmCategoryName(0 To 0) As OLACATEGORYNAME
                ReDim tmUniqueGroupNames(0 To 0) As OLAUNIQUEGROUPNAMES
                If UBound(tmRegionBreakSpots) > 0 Then
                    'Output Region spots plus all generic spots within the break
                    edcMsg.Text = "Generating Region Schedule for " & slVehName & "..."
                    DoEvents
                    mExportRegionSpot slVehGroupName, slVehName, slVehExportID, slTimeID, llEventID
                End If
                csiXMLData "CT", "affidavit", ""
                ilRet = csiXMLWrite(1)
                ilRet = csiXMLEnd()
                DoEvents
                If imTerminate Then
                    mAddMsgToList "User Cancelled Export"
                    mExportSpots = False
                    Exit For
                End If
                edcMsg.Text = "Generating Custom Groups for " & slVehName & "..."
                DoEvents
                ilRet = mExportGroup(slNowDT, slOutputType)
                'Output Unique names then merge Unique names with spots and region definition
                edcMsg.Text = "Generating Copy Split Information " & slVehName & "..."
                DoEvents
                ilRet = mExportUniqueNames(slNowDT, slOutputType)
                'Merge files
                edcMsg.Text = "Merging Schedule and Copy Split Information for " & slVehName & "..."
                DoEvents
                ilRet = mMergeOLAFiles(slXMLFileName)
            Else
                gLogMsg "Vehicle not found for " & slVehName, "OLAExportLog.Txt", False
                mAddMsgToList "Vehicle not found " & slVehName
            End If
        End If
    Next ilVef
    'If ilAnyExports Then
    '    ilRet = csiXMLEnd()
    'End If
    On Error Resume Next
    ilRet = mCloseCSF()
    lst_rst.Close
    vff_rst.Close
    rsf_rst.Close
    mExportSpots = True
    Exit Function
mExportSpotsErr:
    ilRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmExportOLA-mExportSpots"
    Exit Function

End Function


Function mXMLNameFilter(slInName As String) As String
'    Dim slName As String
'    Dim ilPos As Integer
'    Dim ilStartPos As Integer
'    Dim ilFound As Integer
'
'    slName = slInName
'    'Remove " and '
'    ilStartPos = 1
'    Do
'        ilFound = False
'        ilPos = InStr(ilStartPos, slName, "&", 1)
'        If ilPos > 0 Then
'            slName = Left$(slName, ilPos - 1) & "&amp;" & Mid$(slName, ilPos + 1)
'            ilStartPos = ilPos + Len("&amp;")
'            ilFound = True
'        End If
'    Loop While ilFound
'    Do
'        ilFound = False
'        ilPos = InStr(1, slName, "<", 1)
'        If ilPos > 0 Then
'            slName = Left$(slName, ilPos - 1) & "&lt;" & Mid$(slName, ilPos + 1)
'            ilFound = True
'        End If
'    Loop While ilFound
'    Do
'        ilFound = False
'        ilPos = InStr(1, slName, ">", 1)
'        If ilPos > 0 Then
'            slName = Left$(slName, ilPos - 1) & "&gt;" & Mid$(slName, ilPos + 1)
'            ilFound = True
'        End If
'    Loop While ilFound
'    Do
'        ilFound = False
'        ilPos = InStr(1, slName, "'", 1)
'        If ilPos > 0 Then
'            slName = Left$(slName, ilPos - 1) & "&apos;" & Mid$(slName, ilPos + 1)
'            ilFound = True
'        End If
'    Loop While ilFound
'    Do
'        ilFound = False
'        ilPos = InStr(1, slName, """", 1)
'        If ilPos > 0 Then
'            slName = Left$(slName, ilPos - 1) & "&quot;" & Mid$(slName, ilPos + 1)
'            ilFound = True
'        End If
'    Loop While ilFound
'    mXMLNameFilter = Trim$(slName)
End Function
Function mXMLCommentFilter(ilMaxLen As Integer, slInName As String) As String
    Dim slName As String
    Dim ilPos As Integer
    Dim ilStartPos As Integer
    Dim ilFound As Integer
    
    slName = slInName
    'Remove " and '
    ilStartPos = 1
    Do
        ilFound = False
        ilPos = InStr(ilStartPos, slName, "&", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & "&amp;" & Mid$(slName, ilPos + 1)
            ilStartPos = ilPos + Len("and")
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, "<", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, ">", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, "'", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, """", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    If ilMaxLen > 0 Then
        slName = Left$(slName, ilMaxLen)
    End If
    mXMLCommentFilter = Trim$(slName)
End Function

Private Sub txtNumberDays_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub


Private Sub txtStationInfo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub mExportRegionSpot(slVehGroupName As String, slVehName As String, slVehExportID As String, slTimeID As String, llEventID As Long)
    
    'Export any break that has region copy.  Only copts with region split copy are exported

    Dim ilIndex As Integer
    Dim llLoop As Long
    Dim ilRet As Integer
    Dim slRegionName As String
    Dim slGroup As String
    Dim ilCheck As Integer
    Dim ilFound As Integer
    Dim slGroupType As String
    Dim slGroupName As String
    Dim slRegionError As String
    ReDim tlRegionDefinition(0 To 0) As REGIONDEFINITION
    ReDim tlSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
    
    On Error GoTo ErrHand
    
    For ilIndex = 0 To UBound(tmRegionBreakSpots) - 1 Step 1
        DoEvents
        If imTerminate Then
            mAddMsgToList "User Cancelled Export"
            Exit Sub
        End If
        ilRet = gBuildRegionDefinitions("O", tmRegionBreakSpots(ilIndex).lSdfCode, imVefCode, tlRegionDefinition(), tlSplitCategoryInfo())
        If ilRet Then
            ReDim tmRegionDefinition(0 To 0) As REGIONDEFINITION
            ReDim tmSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
            gSeparateRegions tlRegionDefinition(), tlSplitCategoryInfo(), tmRegionDefinition(), tmSplitCategoryInfo()
            For llLoop = 0 To UBound(tmRegionDefinition) - 1 Step 1
                DoEvents
                If imTerminate Then
                    Exit Sub
                End If
                ReDim tmMergeRegionDefinition(0 To 0) As REGIONDEFINITION
                ReDim tmMergeSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
                ilRet = mMergeCategory(llLoop)
                If ilRet Then
                    DoEvents
                    If imTerminate Then
                        Exit Sub
                    End If
                    ilRet = mAnyStations(slGroup)
                    If ilRet Then
                        slGroupName = mFormRegionAddress(slGroup, slGroupType)
                        mCreateXMLSpot slGroupName, llLoop, tmRegionBreakSpots(ilIndex).lLstCode, tmRegionBreakSpots(ilIndex).lBreakNo, tmRegionBreakSpots(ilIndex).iPositionNo
                        ilFound = False
                        For ilCheck = 0 To UBound(tmUniqueGroupNames) - 1 Step 1
                            If (StrComp(Trim$(tmUniqueGroupNames(ilCheck).sGroupName), slGroupName, vbTextCompare) = 0) And (StrComp(Trim$(tmUniqueGroupNames(ilCheck).sGroupType), slGroupType, vbTextCompare) = 0) Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilCheck
                        If Not ilFound Then
                            tmUniqueGroupNames(UBound(tmUniqueGroupNames)).sGroupType = slGroupType
                            tmUniqueGroupNames(UBound(tmUniqueGroupNames)).sGroupName = slGroupName
                            tmUniqueGroupNames(UBound(tmUniqueGroupNames)).sName = Trim$(tmRegionDefinition(llLoop).sRegionName)
                            ReDim Preserve tmUniqueGroupNames(0 To UBound(tmUniqueGroupNames) + 1) As OLAUNIQUEGROUPNAMES
                        End If
                    Else
                        slRegionError = "Empty Region "
                        slRegionError = slRegionError & Trim$(tmRegionDefinition(llLoop).sRegionName)
                        SQLQuery = "SELECT chfCntrNo, adfName, sdfDate, sdfTime "
                        SQLQuery = SQLQuery + " FROM chf_contract_Header, adf_Advertisers, sdf_Spot_Detail"
                        SQLQuery = SQLQuery + " WHERE (sdfCode = " & tmRegionBreakSpots(ilIndex).lSdfCode
                        SQLQuery = SQLQuery + " AND chfCode = sdfChfCode AND adfCode = sdfadfCode " & ")"
                        Set err_rst = gSQLSelectCall(SQLQuery)
                        If Not err_rst.EOF Then
                            slRegionError = slRegionError & " Advertiser " & Trim$(err_rst!adfName) & " Contract # " & err_rst!chfCntrNo & slVehName & " Date " & Format$(err_rst!sdfDate, sgShowDateForm) & " Time " & Format$(err_rst!sdfTime, sgShowTimeWSecForm)
                        Else
                            slRegionError = slRegionError & " " & slVehName
                        End If
                        mAddMsgToList slRegionError
                        gLogMsg slRegionError, "OLAExportLog.Txt", False
                    End If
                End If
            Next llLoop
        End If
    Next ilIndex
    
    Erase tmRegionDefinition
    Erase tmSplitCategoryInfo
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "ExportOLA-mExportRegionSpot"
    Exit Sub
End Sub

Private Function mAnyStations(slGroupInfo As String) As Integer
    'Add loop thru Wegener stations to determine if any station meets the region definition criteria
    'Region defined as Fmt1 and St1 and Not K111.  Look to see if any station will air within this region
    'This is used to avoid exporting regions that have no stations associated with it as Wegener will reject this region and all
    'other commands after this rejected command
    Dim ilRet As Integer
    Dim ilShtt As Integer
    Dim llFormatIndex As Long
    Dim llOtherIndex As Long
    Dim llExcludeIndex As Long
    Dim ilShttCode As Integer
    Dim ilMktCode As Integer
    Dim ilMSAMktCode As Integer
    Dim slState As String
    Dim ilFmtCode As Integer
    Dim ilTztCode As Integer
    Dim llRegion As Long
    
    
    For ilShtt = 0 To UBound(tgStationInfoByCode) - 1 Step 1
        If tgStationInfoByCode(ilShtt).sUsedForOLA = "Y" Then
            ilShttCode = tgStationInfoByCode(ilShtt).iCode
            ilMktCode = tgStationInfoByCode(ilShtt).iMktCode
            ilMSAMktCode = tgStationInfoByCode(ilShtt).iMSAMktCode
            slState = tgStationInfoByCode(ilShtt).sPostalName
            ilFmtCode = tgStationInfoByCode(ilShtt).iFormatCode
            ilTztCode = tgStationInfoByCode(ilShtt).iTztCode
            ilRet = gRegionTestDefinition(ilShttCode, ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, tmMergeRegionDefinition(), tmMergeSplitCategoryInfo(), llRegion, slGroupInfo)
            If ilRet Then
                mAnyStations = True
                Exit Function
            End If
        End If
    Next ilShtt
    mAnyStations = False
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadStationReceiverRecords     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File to get OLA       *
'*                      stations                       *
'*                                                     *
'*******************************************************
Private Function mReadStationReceiverRecords() As Integer
    Dim ilEof As Integer
    Dim slPath As String
    Dim slLine As String
    Dim slChar As String
    Dim ilRet As Integer
    Dim slFromFile As String
    Dim slCallLetters As String
    Dim slFreq As String
    Dim slLicCity As String
    Dim slDMAName As String
    Dim slDMACode As String
    Dim slFormatName As String
    Dim slFormatCode As String
    Dim slStateName As String
    Dim slPostalName As String
    Dim slStateCode As String
    Dim slZoneCode As String
    Dim slCurDate As String
    Dim slCurTime As String
    Dim ilMktCode As Integer
    Dim ilFmtCode As Integer
    Dim ilSntCode As Integer
    Dim ilTztCode As Integer
    Dim slCSIName As String
    Dim ilCode As Integer
    Dim ilWegener As Integer
    Dim ilOLA As Integer
    Dim ilHeaderFd As Integer
    
    mReadStationReceiverRecords = 0
    ilHeaderFd = False
    On Error GoTo ErrHand
    SQLQuery = "Update SHTT Set shttUsedForOLA = '" & "N" & "'"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "Export OLA-mReadStationReceiverRecords"
        mReadStationReceiverRecords = 1
        Exit Function
    End If
    'ReDim imShttCodes(0 To 0) As Integer
    On Error GoTo mReadStationReceiverRecordsErr:
    slCurDate = Format(gNow(), sgShowDateForm)
    slCurTime = Format(gNow(), sgShowTimeWSecForm)
    'slPath = txtStationInfo.text
    'If Right$(slPath, 1) <> "\" Then
    '    slPath = slPath & "\"
    'End If
    'slFromFile = slPath & "JNS_RecSerialNum.Csv"
    slFromFile = txtStationInfo.Text
    'ilRet = 0
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        mAddMsgToList "Open " & slFromFile & " error#" & Str$(ilRet)
        gLogMsg "Open " & slFromFile & " error#" & Str$(ilRet), "OLAExportLog.Txt", False
        mReadStationReceiverRecords = 1
        Exit Function
    End If
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mReadStationReceiverRecordsErr:
        slLine = ""
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                Exit Do
            ElseIf slChar <> sgCR Then
                slLine = slLine & slChar
            End If
        Loop
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            mAddMsgToList "User Cancelled Export"
            gLogMsg "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **", "OLAExportLog.Txt", False
            mReadStationReceiverRecords = 2
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, True, smFields()
                'slCallLetters = Trim$(smFields(1))
                slCallLetters = Trim$(smFields(0))
                'Bypass title line
                'call_letters,frequency,license_city,dma_description,dma_code,format_description,format_code,state_name,state,state_code,zone_code,wegener,OLA
                If Not ilHeaderFd Then
                    'If (StrComp(slCallLetters, "Call_Letters", vbTextCompare) = 0) And (StrComp(smFields(2), "frequency", vbTextCompare) = 0) And (StrComp(smFields(12), "wegener", vbTextCompare) = 0) Then
                    If (StrComp(slCallLetters, "Call_Letters", vbTextCompare) = 0) And (StrComp(smFields(1), "frequency", vbTextCompare) = 0) And (StrComp(smFields(11), "wegener", vbTextCompare) = 0) Then
                        ilHeaderFd = True
                    End If
                Else
                    'slFreq = Trim$(smFields(2))
                    slFreq = Trim$(smFields(1))
                    'slLicCity = Trim$(smFields(3))
                    slLicCity = Trim$(smFields(2))
                    'slDMAName = Trim$(smFields(4))
                    slDMAName = Trim$(smFields(3))
                    'slDMACode = Trim$(smFields(5))
                    slDMACode = Trim$(smFields(4))
                    ilMktCode = mFindMkt(slDMAName, slDMACode)
                    'slFormatName = Trim$(smFields(6))
                    slFormatName = Trim$(smFields(5))
                    'slFormatCode = Trim$(smFields(7))
                    slFormatCode = Trim$(smFields(6))
                    ilFmtCode = mFindFmt(slFormatName, slFormatCode)
                    'slStateName = Trim$(smFields(8))
                    slStateName = Trim$(smFields(7))
                    'slPostalName = Trim$(smFields(9))
                    slPostalName = Trim$(smFields(8))
                    'slStateCode = Trim$(smFields(10))
                    slStateCode = Trim$(smFields(9))
                    ilSntCode = mFindSnt(slStateName, slPostalName, slStateCode)
                    'slZoneCode = Trim$(smFields(11))
                    slZoneCode = Trim$(smFields(10))
                    ilTztCode = mFindTzt(slZoneCode, slCSIName)
                    'ilWegener = Val(smFields(12))
                    ilWegener = Val(smFields(11))
                    'ilOLA = Val(smFields(13))
                    ilOLA = Val(smFields(12))
                    On Error GoTo ErrHand
                    
                    ilRet = gBinarySearchStation(slCallLetters)
                    If ilRet = -1 Then
                        If (ilWegener <> 0) Or (ilOLA <> 0) Then
                            ilRet = 0
                            Do
                                SQLQuery = "SELECT MAX(shttCode) from shtt"
                                Set rst = gSQLSelectCall(SQLQuery)
                                If Not rst.EOF Then
                                    ilCode = rst(0).Value + 1
                                Else
                                    ilCode = 1
                                End If
                                SQLQuery = "Insert Into shtt ( "
                                SQLQuery = SQLQuery & "shttCode, "
                                SQLQuery = SQLQuery & "shttCallLetters, "
                                SQLQuery = SQLQuery & "shttAddress1, "
                                SQLQuery = SQLQuery & "shttAddress2, "
                                SQLQuery = SQLQuery & "shttCity, "
                                SQLQuery = SQLQuery & "shttState, "
                                SQLQuery = SQLQuery & "shttCountry, "
                                SQLQuery = SQLQuery & "shttZip, "
                                SQLQuery = SQLQuery & "shttSelected, "
                                SQLQuery = SQLQuery & "shttEmail, "
                                SQLQuery = SQLQuery & "shttFax, "
                                SQLQuery = SQLQuery & "shttPhone, "
                                SQLQuery = SQLQuery & "shttTimeZone, "
                                SQLQuery = SQLQuery & "shttHomePage, "
                                SQLQuery = SQLQuery & "shttPDName, "
                                SQLQuery = SQLQuery & "shttPDPhone, "
                                SQLQuery = SQLQuery & "shttTDName, "
                                SQLQuery = SQLQuery & "shttTDPhone, "
                                SQLQuery = SQLQuery & "shttMDName, "
                                'SQLQuery = SQLQuery & "shttMDPhone, "
                                ''SQLQuery = SQLQuery & "shttPC, "
                                ''SQLQuery = SQLQuery & "shttHdDrive, "
                                'SQLQuery = SQLQuery & "shttMonthlyWebPost, "
                                SQLQuery = SQLQuery & "shttFrequency, "
                                SQLQuery = SQLQuery & "shttPermStationID, "
                                SQLQuery = SQLQuery & "shttUnused, "
                                SQLQuery = SQLQuery & "shttACName, "
                                SQLQuery = SQLQuery & "shttACPhone, "
                                SQLQuery = SQLQuery & "shttMntCode, "
                                SQLQuery = SQLQuery & "shttChecked, "
                                SQLQuery = SQLQuery & "shttMarket, "
                                SQLQuery = SQLQuery & "shttRank, "
                                SQLQuery = SQLQuery & "shttUsfCode, "
                                SQLQuery = SQLQuery & "shttEnterDate, "
                                SQLQuery = SQLQuery & "shttEnterTime, "
                                SQLQuery = SQLQuery & "shttType, "
                                SQLQuery = SQLQuery & "shttONAddress1, "
                                SQLQuery = SQLQuery & "shttONAddress2, "
                                SQLQuery = SQLQuery & "shttONCity, "
                                SQLQuery = SQLQuery & "shttONState, "
                                SQLQuery = SQLQuery & "shttONZip, "
                                SQLQuery = SQLQuery & "shttStationID, "
                                SQLQuery = SQLQuery & "shttCityLic, "
                                SQLQuery = SQLQuery & "shttStateLic, "
                                SQLQuery = SQLQuery & "shttAckDaylight, "
                                SQLQuery = SQLQuery & "shttWebEmail, "
                                SQLQuery = SQLQuery & "shttWebPW, "
                                SQLQuery = SQLQuery & "shttOwnerArttCode, "
                                SQLQuery = SQLQuery & "shttMktCode, "
                                SQLQuery = SQLQuery & "shttWebAddress, "
                                SQLQuery = SQLQuery & "shttfmtCode, "
                                SQLQuery = SQLQuery & "shttSerialNo1, "
                                SQLQuery = SQLQuery & "shttSerialNo2, "
                                SQLQuery = SQLQuery & "shttTztCode, "
                                SQLQuery = SQLQuery & "shttWebNumber, "
                                SQLQuery = SQLQuery & "shttUsedForAtt, "
                                SQLQuery = SQLQuery & "shttUsedForXDigital, "
                                SQLQuery = SQLQuery & "shttUsedForWegener, "
                                SQLQuery = SQLQuery & "shttUsedForOLA, "
                                SQLQuery = SQLQuery & "shttPort, "
                                SQLQuery = SQLQuery & "shttUnused "
                                SQLQuery = SQLQuery & ") "
                                SQLQuery = SQLQuery & "Values ( "
                                SQLQuery = SQLQuery & ilCode & ", "
                                SQLQuery = SQLQuery & "'" & gFixQuote(slCallLetters) & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & gFixQuote(slPostalName) & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & -1 & ", "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & gFixQuote(slCSIName) & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                'SQLQuery = SQLQuery & "'" & "" & "', "
                                ''SQLQuery = SQLQuery & 0 & ", "
                                ''SQLQuery = SQLQuery & 0 & ", "
                                'SQLQuery = SQLQuery & "'" & "N" & "', "     'Allow Monthly Web Posting
                                SQLQuery = SQLQuery & "'" & "" & "', "  'Frequency
                                SQLQuery = SQLQuery & 0 & ", "  'Permanent Station ID
                                SQLQuery = SQLQuery & "'" & "" & "', "      'Unused
                                SQLQuery = SQLQuery & "'" & "" & "', "      'AC Name
                                SQLQuery = SQLQuery & "'" & "" & "', "      'AC Phone
                                SQLQuery = SQLQuery & 0 & ", "              'MntCode
                                SQLQuery = SQLQuery & -1 & ", "             'Checked
                                SQLQuery = SQLQuery & "'" & gFixQuote(slDMAName) & "', "    'DMA Market
                                SQLQuery = SQLQuery & 0 & ", "              'Rank
                                SQLQuery = SQLQuery & igUstCode & ", "
                                SQLQuery = SQLQuery & "'" & Format$(slCurDate, sgSQLDateForm) & "', "
                                SQLQuery = SQLQuery & "'" & Format$(slCurTime, sgSQLTimeForm) & "', "
                                SQLQuery = SQLQuery & 0 & ", "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & 0 & ", "
                                SQLQuery = SQLQuery & "'" & gFixQuote(slLicCity) & "', "
                                SQLQuery = SQLQuery & "'" & gFixQuote(slPostalName) & "', "
                                SQLQuery = SQLQuery & 1 & ", "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & 0 & ", "
                                SQLQuery = SQLQuery & ilMktCode & ", "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & ilFmtCode & ", "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & ilTztCode & ", "
                                SQLQuery = SQLQuery & "'" & "1" & "', "
                                If (ilWegener = 0) And (ilOLA = 0) Then
                                    SQLQuery = SQLQuery & "'" & gFixQuote("Y") & "', "
                                Else
                                    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
                                End If
                                SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
                                'Wegener set via Wegener import
                                SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
                                If ilOLA = 0 Then
                                    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
                                Else
                                    SQLQuery = SQLQuery & "'" & gFixQuote("Y") & "', "
                                End If
                                SQLQuery = SQLQuery & "'" & "" & "', "
                                SQLQuery = SQLQuery & "'" & "" & "' "
                                SQLQuery = SQLQuery & ") "
                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    '6/10/16: Replaced GoSub
                                    'GoSub ErrHand:
                                    Screen.MousePointer = vbDefault
                                    If Not gHandleError4994("AffErrorLog.txt", "Export OLA-mReadStationReceiverRecords") Then
                                        mReadStationReceiverRecords = 1
                                        Exit Function
                                    End If
                                    ilRet = 1
                                End If
                            Loop While ilRet <> 0
                            'If ilOLA <> 0 Then
                            '    imShttCodes(UBound(imShttCodes)) = ilCode
                            '    ReDim Preserve imShttCodes(0 To UBound(imShttCodes) + 1) As Integer
                            'End If
                            gLogMsg slCallLetters & " Added", "OLAExportLog.Txt", False
                            mAddMsgToList slCallLetters & " Added"
                            '11/26/17: Set Changed date/time
                            gFileChgdUpdate "shtt.mkd", True
                        End If
                    Else
                        'Update information
                        SQLQuery = "Update shtt Set "
                        SQLQuery = SQLQuery & "shttCallLetters = '" & slCallLetters & "', "
                        SQLQuery = SQLQuery & "shttState = '" & gFixQuote(slPostalName) & "', "
                        SQLQuery = SQLQuery & "shttTimeZone = '" & gFixQuote(slCSIName) & "', "
                        SQLQuery = SQLQuery & "shttMarket = '" & gFixQuote(slDMAName) & "', "
                        SQLQuery = SQLQuery & "shttUsfCode = " & igUstCode & ", "
                        SQLQuery = SQLQuery & "shttEnterDate = '" & Format$(slCurDate, sgSQLDateForm) & "', "
                        SQLQuery = SQLQuery & "shttEnterTime = '" & Format$(slCurTime, sgSQLTimeForm) & "', "
                        SQLQuery = SQLQuery & "shttCityLic = '" & gFixQuote(slLicCity) & "', "
                        SQLQuery = SQLQuery & "shttStateLic = '" & gFixQuote(slPostalName) & "', "
                        SQLQuery = SQLQuery & "shttMktCode = " & ilMktCode & ", "
                        SQLQuery = SQLQuery & "shttfmtCode = " & ilFmtCode & ", "
                        SQLQuery = SQLQuery & "shttTztCode = " & ilTztCode & ", "
                        If (ilWegener = 0) And (ilOLA = 0) Then
                            SQLQuery = SQLQuery & "shttUsedForAtt = '" & gFixQuote("Y") & "', "
                        End If
                        If ilOLA <> 0 Then
                            SQLQuery = SQLQuery & "shttUsedForOLA = '" & gFixQuote("Y") & "', "
                        End If
                        SQLQuery = SQLQuery & "shttUnused = '" & "" & "' "
                        SQLQuery = SQLQuery & " WHERE shttCode = " & tgStationInfo(ilRet).iCode
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "AffErrorLog.txt", "Export OLA-mReadStationReceiverRecords"
                            mReadStationReceiverRecords = 1
                            Exit Function
                        End If
                        '11/26/17: Set Changed date/time
                        gFileChgdUpdate "shtt.mkd", True
                        'If ilOLA <> 0 Then
                        '    imShttCodes(UBound(imShttCodes)) = tgStationInfo(ilRet).iCode
                        '    ReDim Preserve imShttCodes(0 To UBound(imShttCodes) + 1) As Integer
                        'End If
                    End If
                End If
            End If
        End If
        ilRet = 0
    Loop
    If Not ilHeaderFd Then
        mAddMsgToList "Header record not found in " & slFromFile
        gLogMsg "Header record not found in " & slFromFile, "OLAExportLog.Txt", False
        mReadStationReceiverRecords = 3
        Exit Function
    End If
    Close hmFrom
    Exit Function
mReadStationReceiverRecordsErr:
    ilRet = Err.Number
    Resume Next
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frnExportOLA-mReadStationReceiverRecords"
    mReadStationReceiverRecords = 1
End Function


Private Sub mSeparateRegions(tlRegionDefinition() As REGIONDEFINITION, tlSplitCategoryInfo() As SPLITCATEGORYINFO)
    'If a region is defined as:
    '(Fmt1 or Fmt2 or Fmt3) and (St1 or St2) and (Not K1111 and Not K222)
    'Convert to:
    'Region 1: Fmt1 and St1 And Not K111 and Not K222
    'Region 2: Fmt1 and St2 And Not K111 and Not K222
    'Region 3: Fmt2 and St1 And Not K111 and Not K222
    'Region 4: Fmt2 and St2 And Not K111 and Not K222
    'Region 5: Fmt3 and St1 And Not K111 and Not K222
    'Region 6: Fmt3 and St2 And Not K111 and Not K222
    Dim llFormatIndex As Long
    Dim llRegion As Long
    Dim llOtherIndex As Long
    Dim llExcludeIndex As Long
    
    For llRegion = 0 To UBound(tlRegionDefinition) - 1 Step 1
        llFormatIndex = tlRegionDefinition(llRegion).lFormatFirst
            
        If tlRegionDefinition(llRegion).lFormatFirst <> -1 Then
            'Test Format
            llFormatIndex = tlRegionDefinition(llRegion).lFormatFirst
            Do
                tmRegionDefinition(UBound(tmRegionDefinition)) = tlRegionDefinition(llRegion)
                tmRegionDefinition(UBound(tmRegionDefinition)).lFormatFirst = UBound(tmSplitCategoryInfo)
                tmRegionDefinition(UBound(tmRegionDefinition)).lOtherFirst = -1
                tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = -1
                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llFormatIndex)
                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                If tlRegionDefinition(llRegion).lOtherFirst <> -1 Then
                    llOtherIndex = tlRegionDefinition(llRegion).lOtherFirst
                    Do
                        tmRegionDefinition(UBound(tmRegionDefinition)).lOtherFirst = UBound(tmSplitCategoryInfo)
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llFormatIndex)
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                        ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
                        If llExcludeIndex <> -1 Then
                            tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = UBound(tmSplitCategoryInfo)
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                            ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                            llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                            Do While llExcludeIndex <> -1
                                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo) - 1).lNext = UBound(tmSplitCategoryInfo)
                                ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                                llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                            Loop
                        End If
                        ReDim Preserve tmRegionDefinition(0 To UBound(tmRegionDefinition) + 1) As REGIONDEFINITION
                        llOtherIndex = tlSplitCategoryInfo(llOtherIndex).lNext
                        If llOtherIndex <> -1 Then
                            tmRegionDefinition(UBound(tmRegionDefinition)) = tlRegionDefinition(llRegion)
                            tmRegionDefinition(UBound(tmRegionDefinition)).lFormatFirst = UBound(tmSplitCategoryInfo)
                            tmRegionDefinition(UBound(tmRegionDefinition)).lOtherFirst = -1
                            tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = -1
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llFormatIndex)
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                            ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        End If
                    Loop While llOtherIndex <> -1
                Else
                    llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
                    If llExcludeIndex <> -1 Then
                        tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = UBound(tmSplitCategoryInfo)
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                        ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                        Do While llExcludeIndex <> -1
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                            tmSplitCategoryInfo(UBound(tmSplitCategoryInfo) - 1).lNext = UBound(tmSplitCategoryInfo)
                            ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                            llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                        Loop
                        ReDim Preserve tmRegionDefinition(0 To UBound(tmRegionDefinition) + 1) As REGIONDEFINITION
                    End If
                End If
                llFormatIndex = tlSplitCategoryInfo(llFormatIndex).lNext
            Loop While llFormatIndex <> -1
        ElseIf tlRegionDefinition(llRegion).lOtherFirst <> -1 Then
            llOtherIndex = tlRegionDefinition(llRegion).lOtherFirst
            Do
                tmRegionDefinition(UBound(tmRegionDefinition)) = tlRegionDefinition(llRegion)
                tmRegionDefinition(UBound(tmRegionDefinition)).lFormatFirst = -1
                tmRegionDefinition(UBound(tmRegionDefinition)).lOtherFirst = UBound(tmSplitCategoryInfo)
                tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = -1
                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llOtherIndex)
                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
                If llExcludeIndex <> -1 Then
                    tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = UBound(tmSplitCategoryInfo)
                    tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                    tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                    ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                    llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                    Do While llExcludeIndex <> -1
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                        tmSplitCategoryInfo(UBound(tmSplitCategoryInfo) - 1).lNext = UBound(tmSplitCategoryInfo)
                        ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                    Loop
                End If
                ReDim Preserve tmRegionDefinition(0 To UBound(tmRegionDefinition) + 1) As REGIONDEFINITION
                llOtherIndex = tlSplitCategoryInfo(llOtherIndex).lNext
            Loop While llOtherIndex <> -1
        Else
            'Exclude only
            llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
            If llExcludeIndex <> -1 Then
                tmRegionDefinition(UBound(tmRegionDefinition)) = tlRegionDefinition(llRegion)
                tmRegionDefinition(UBound(tmRegionDefinition)).lFormatFirst = -1
                tmRegionDefinition(UBound(tmRegionDefinition)).lOtherFirst = -1
                tmRegionDefinition(UBound(tmRegionDefinition)).lExcludeFirst = UBound(tmSplitCategoryInfo)
                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                Do While llExcludeIndex <> -1
                    tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)) = tlSplitCategoryInfo(llExcludeIndex)
                    tmSplitCategoryInfo(UBound(tmSplitCategoryInfo)).lNext = -1
                    tmSplitCategoryInfo(UBound(tmSplitCategoryInfo) - 1).lNext = UBound(tmSplitCategoryInfo)
                    ReDim Preserve tmSplitCategoryInfo(0 + UBound(tmSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                    llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                Loop
                ReDim Preserve tmRegionDefinition(0 To UBound(tmRegionDefinition) + 1) As REGIONDEFINITION
            End If
        End If
    Next llRegion

End Sub

Private Function mFormRegionAddress(slInGroupInfo As String, slGroupType As String) As String
    'Translate the User defined region into region definition that Wegener understands
    'User enters:
    'Urban and California
    'Wegener wants
    'Fmt_123 ^ St_CA
    'Wegener names are call Group Names and are defined as menu items for each category
    '(Format, State, Time zone and Market.  For stations, the call letters are used)
    'Symbols:  ^ = And; ~ = Not
    Dim ilPos As Integer
    Dim slGroupInfo As String
    Dim slStr As String
    Dim slInclExcl As String
    Dim slCategory As String
    Dim slvalue As String
    Dim ilValue As Integer
    Dim ilSnt As Integer
    Dim slAddress As String
    Dim ilRet As Integer
    Dim slGroupName As String
    Dim ilPreCatNameIndex As Integer
    Dim slCustomNo As String
    Dim slMergedCategoryName As String
    Dim slTestCategoryName As String
    Dim ilGroup As Integer
    Dim ilCatIndex As Integer
    Dim ilCheck As Integer
    Dim ilFound As Integer
    Dim slName As String
    
    On Error GoTo mFormRegionAddressErr:
    ilPreCatNameIndex = -1
    slGroupInfo = slInGroupInfo
    ilPos = InStr(1, slGroupInfo, "|", vbTextCompare)
    If (ilPos > 0) Or (Left$(slGroupInfo, 1) = "E") Then
        slMergedCategoryName = ""
        Do
            ilPos = InStr(1, slGroupInfo, "|", vbTextCompare)
            If ilPos = 0 Then
                If Len(slGroupInfo) = 0 Then
                    Exit Do
                Else
                    ilPos = Len(slGroupInfo) + 1
                End If
            End If
            slStr = Left(slGroupInfo, ilPos - 1)
            slGroupInfo = Mid$(slGroupInfo, ilPos + 1)
            slInclExcl = Left$(slStr, 1)
            slCategory = Mid$(slStr, 2, 1)
            slvalue = Trim$(Mid$(slStr, 3))
            If slCategory <> "N" Then
                ilValue = Val(slvalue)
            End If
            Select Case slCategory
                Case "M"    'Market
                    ilRet = gBinarySearchMkt(CLng(ilValue))
                    If ilRet <> -1 Then
                        slGroupName = Trim$(tgMarketInfo(ilRet).sGroupName)
                        slGroupType = "DMA"
                        slName = Trim$(tgMarketInfo(ilRet).sName)
                    End If
                Case "N"    'State Name
                    For ilSnt = 0 To UBound(tgStateInfo) - 1 Step 1
                        If StrComp(Trim$(tgStateInfo(ilSnt).sPostalName), slvalue, vbTextCompare) = 0 Then
                            slGroupName = Trim$(tgStateInfo(ilSnt).sGroupName)
                            slGroupType = "State"
                            slName = Trim$(tgStateInfo(ilSnt).sName)
                            Exit For
                        End If
                    Next ilSnt
                Case "F"    'Format
                    ilRet = gBinarySearchFmt(CLng(ilValue))
                    If ilRet <> -1 Then
                        slGroupName = Trim$(tgFormatInfo(ilRet).sGroupName)
                        slGroupType = "Format"
                        slName = Trim$(tgFormatInfo(ilRet).sName)
                    End If
                Case "T"    'Time zone
                    ilRet = gBinarySearchTzt(ilValue)
                    If ilRet <> -1 Then
                        slGroupName = Trim$(tgTimeZoneInfo(ilRet).sGroupName)
                        slGroupType = "TimeZone"
                        slName = Trim$(tgTimeZoneInfo(ilRet).sName)
                    End If
                Case "S"    'Station
                    ilRet = gBinarySearchStationInfoByCode(ilValue)
                    If ilRet <> -1 Then
                        slGroupName = Trim$(tgStationInfoByCode(ilRet).sCallLetters)
                        slGroupType = "Affiliate"
                        slName = Trim$(tgStationInfoByCode(ilRet).sCallLetters)
                    End If
            End Select
            '2/3/09: Don't include individual root names that are part of the custom group name in the Unique name area.
            'ilFound = False
            'For ilCheck = 0 To UBound(tmUniqueGroupNames) - 1 Step 1
            '    If (StrComp(Trim$(tmUniqueGroupNames(ilCheck).sGroupName), slGroupName, vbTextCompare) = 0) And (StrComp(Trim$(tmUniqueGroupNames(ilCheck).sGroupType), slGroupType, vbTextCompare) = 0) Then
            '        ilFound = True
            '        Exit For
            '    End If
            'Next ilCheck
            'If Not ilFound Then
            '    tmUniqueGroupNames(UBound(tmUniqueGroupNames)).sGroupType = slGroupType
            '    tmUniqueGroupNames(UBound(tmUniqueGroupNames)).sGroupName = slGroupName
            '    tmUniqueGroupNames(UBound(tmUniqueGroupNames)).sName = slName
            '    ReDim Preserve tmUniqueGroupNames(0 To UBound(tmUniqueGroupNames) + 1) As OLAUNIQUEGROUPNAMES
            'End If
            If slInclExcl = "E" Then
                slGroupName = "~" & slGroupName
            End If
            slMergedCategoryName = slMergedCategoryName & Trim$(slGroupName)
        Loop While Len(slGroupInfo) > 0
        For ilGroup = 0 To UBound(tmCustomGroupNames) - 1 Step 1
            slTestCategoryName = ""
            ilCatIndex = tmCustomGroupNames(ilGroup).iFirst
            Do While ilCatIndex <> -1
                slTestCategoryName = slTestCategoryName & Trim$(tmCategoryName(ilCatIndex).sName)
                ilCatIndex = tmCategoryName(ilCatIndex).iNext
            Loop
            If StrComp(slMergedCategoryName, slTestCategoryName, vbTextCompare) = 0 Then
                slGroupType = "Custom"
                mFormRegionAddress = Trim$(tmCustomGroupNames(ilGroup).sName)
                Exit Function
            End If
        Next ilGroup
        slGroupInfo = slInGroupInfo
        imCustomGroupNo = imCustomGroupNo + 1
        slCustomNo = Trim$(Str$(imCustomGroupNo))
        Do While Len(slCustomNo) < 4
            slCustomNo = "0" & slCustomNo
        Loop
        tmCustomGroupNames(UBound(tmCustomGroupNames)).sName = smCustomGroupName & slCustomNo
        tmCustomGroupNames(UBound(tmCustomGroupNames)).iFirst = UBound(tmCategoryName)
        slAddress = smCustomGroupName & slCustomNo
        slGroupType = "Custom"
        ReDim Preserve tmCustomGroupNames(0 To UBound(tmCustomGroupNames) + 1) As OLACUSTOMGROUPNAMES
    Else
        slAddress = ""
        slGroupType = ""
    End If
    Do
        ilPos = InStr(1, slGroupInfo, "|", vbTextCompare)
        If ilPos = 0 Then
            If Len(slGroupInfo) = 0 Then
                Exit Do
            Else
                ilPos = Len(slGroupInfo) + 1
            End If
        End If
        slStr = Left(slGroupInfo, ilPos - 1)
        slGroupInfo = Mid$(slGroupInfo, ilPos + 1)
        slInclExcl = Left$(slStr, 1)
        slCategory = Mid$(slStr, 2, 1)
        slvalue = Trim$(Mid$(slStr, 3))
        If slCategory <> "N" Then
            ilValue = Val(slvalue)
        End If
        Select Case slCategory
            Case "M"    'Market
                ilRet = gBinarySearchMkt(CLng(ilValue))
                If ilRet <> -1 Then
                    slGroupName = Trim$(tgMarketInfo(ilRet).sGroupName)
                    If slGroupType = "" Then
                        slGroupType = "DMA"
                    End If
                End If
            Case "N"    'State Name
                For ilSnt = 0 To UBound(tgStateInfo) - 1 Step 1
                    If StrComp(Trim$(tgStateInfo(ilSnt).sPostalName), slvalue, vbTextCompare) = 0 Then
                        slGroupName = Trim$(tgStateInfo(ilSnt).sGroupName)
                        If slGroupType = "" Then
                            slGroupType = "State"
                        End If
                        Exit For
                    End If
                Next ilSnt
            Case "F"    'Format
                ilRet = gBinarySearchFmt(CLng(ilValue))
                If ilRet <> -1 Then
                    slGroupName = Trim$(tgFormatInfo(ilRet).sGroupName)
                    If slGroupType = "" Then
                        slGroupType = "Format"
                    End If
                End If
            Case "T"    'Time zone
                ilRet = gBinarySearchTzt(ilValue)
                If ilRet <> -1 Then
                    slGroupName = Trim$(tgTimeZoneInfo(ilRet).sGroupName)
                    If slGroupType = "" Then
                        slGroupType = "TimeZone"
                    End If
                End If
            Case "S"    'Station
                ilRet = gBinarySearchStationInfoByCode(ilValue)
                If ilRet <> -1 Then
                    slGroupName = Trim$(tgStationInfoByCode(ilRet).sCallLetters)
                    If slGroupType = "" Then
                        slGroupType = "Affiliate"
                    End If
                End If
        End Select
        If slInclExcl = "E" Then
            slGroupName = "~" & slGroupName
        End If
        If slAddress = "" Then
            slAddress = slGroupName
        Else
            tmCategoryName(UBound(tmCategoryName)).sName = slGroupName
            tmCategoryName(UBound(tmCategoryName)).sCategory = slCategory
            tmCategoryName(UBound(tmCategoryName)).iNext = -1
            If ilPreCatNameIndex <> -1 Then
                tmCategoryName(ilPreCatNameIndex).iNext = UBound(tmCategoryName)
            End If
            ilPreCatNameIndex = UBound(tmCategoryName)
            ReDim Preserve tmCategoryName(0 To UBound(tmCategoryName) + 1) As OLACATEGORYNAME
        End If
    Loop While Len(slGroupInfo) > 0
    
    mFormRegionAddress = slAddress
    Exit Function
mFormRegionAddressErr:
    Resume Next
End Function




Private Sub mAddMsgToList(slMsg As String)
    'Add horizontal scroll if required and add message to list box
    'The control pbcArial is used to get the approximate width of the text as the list box does not has a TextWidth command
    Dim llValue As Long
    Dim llRg As Long
    Dim llMaxWidth
    Dim llRet As Long
    Dim llRow As Long
    
    llMaxWidth = (pbcArial.TextWidth(slMsg))
    If llMaxWidth > lmMaxWidth Then
        lmMaxWidth = llMaxWidth
    End If
    If lmMaxWidth > lbcMsg.Width Then
        'Divide by 15 to convert units and add 120 for little extra room
        'Scale Mode is in Twips
        llValue = lmMaxWidth / 15 + 120
        llRg = 0
        llRet = SendMessageByNum(lbcMsg.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    End If
    llRow = SendMessageByString(lbcMsg.hwnd, LB_FINDSTRING, -1, slMsg)
    If llRow < 0 Then
        lbcMsg.AddItem slMsg
    End If
End Sub

Private Function mExportGroup(slNowDT As String, slOutputType As String) As Integer
    Dim slGrpName As String
    Dim ilGroup As Integer
    Dim slStr As String
    Dim ilCatIndex As Integer
    Dim ilRet As Integer
    Dim slFileName As String
    Dim ilNotSymbol As Integer
    Dim slHour As String
    
    mExportGroup = True
    For ilGroup = 0 To UBound(tmCustomGroupNames) - 1 Step 1
        slStr = Format$(slNowDT, "yy") & Format$(slNowDT, "mm") & Format$(slNowDT, "dd") & Format$(slNowDT, "hh") & Format$(slNowDT, "nn") & Format$(slNowDT, "ss")
        slFileName = "CSCGD~" & Trim$(tmCustomGroupNames(ilGroup).sName) & "~" & slStr & ".XML"
         '6808
        If Not gDeleteFile(smExportPath & slFileName) Then
            mAddMsgToList "Could not delete file in mExportGroup before writing.  Appended."
        End If
        ' Dan M 11/01/10 use smIniPathFileName that is created at formload
        'ilRet = csiXMLStart(sgStartupDirectory & "\xml.ini", "OLA", slOutputType, smExportPath & slFileName, sgCRLF)
        '6807
        'ilRet = csiXMLStart(smIniPathFileName, "OLA", slOutputType, smExportPath & slFileName, sgCRLF)
        ilRet = csiXMLStart(smIniPathFileName, "OLA", slOutputType, smExportPath & slFileName, sgCRLF, "")
        ilRet = csiXMLSetMethod("", "", "", "CopySplitCustomGroupDefinition")
        DoEvents
        csiXMLData "OT", "fileHeader", ""
        csiXMLData "CD", "filename", smGrpFileName & slFileName
        
        csiXMLData "CD", "description", "Copy Split Custom Groups Export"
        
        csiXMLData "OT", "creationDate", ""
        csiXMLData "CD", "day", Format$(slNowDT, "d")
        csiXMLData "CD", "month", Format$(slNowDT, "m")
        csiXMLData "CD", "year", Format$(slNowDT, "yyyy")
        slHour = Format$(slNowDT, "hh")
        'slHour = Trim$(Str$(Val(slHour) + 1))
        'If Len(slHour) = 1 Then
        '    slHour = "0" & slHour
        'End If
        csiXMLData "CD", "hour", slHour 'Format$(slNowDT, "hh")
        csiXMLData "CD", "minute", Format$(slNowDT, "nn")
        csiXMLData "CD", "seconds", Format$(slNowDT, "ss")
        csiXMLData "CT", "creationDate", ""
        csiXMLData "CD", "contactName", "Dial Global Affiliate Relations"
        csiXMLData "CD", "contactPhoneNumber", smAdminPhone
        csiXMLData "CD", "contactEmail", gXMLNameFilter(smAdminEMail)
        csiXMLData "CT", "fileHeader", ""
        
        
        csiXMLData "OT", "newCustomGroup", ""
        csiXMLData "CD", "csCustomGroupCode", Trim$(tmCustomGroupNames(ilGroup).sName)
        csiXMLData "CD", "description", ""
        ilCatIndex = tmCustomGroupNames(ilGroup).iFirst
        Do While ilCatIndex <> -1
            csiXMLData "OT", "csCustomGroupCriteria", ""
            slStr = Trim$(tmCategoryName(ilCatIndex).sName)
            If Left$(slStr, 1) = "~" Then
                slStr = Mid(slStr, 2)
                ilNotSymbol = True
            Else
                ilNotSymbol = False
            End If
            csiXMLData "CD", "csRootGroupCode", slStr
            Select Case tmCategoryName(ilCatIndex).sCategory
                Case "M"    'Market
                    csiXMLData "CD", "csRootGroupType", "DMA"
                Case "N"    'State Name
                    csiXMLData "CD", "csRootGroupType", "State"
                Case "F"    'Format
                    csiXMLData "CD", "csRootGroupType", "Format"
                Case "T"    'Time zone
                    csiXMLData "CD", "csRootGroupType", "TimeZone"
                Case "S"    'Station
                    csiXMLData "CD", "csRootGroupType", "Affiliate"
            End Select
            If ilNotSymbol Then
                'csiXMLData "CT", "csrgNot", ""
                csiXMLData "OT", "csrgNot/", ""
            End If
            csiXMLData "CT", "csCustomGroupCriteria", ""
            ilCatIndex = tmCategoryName(ilCatIndex).iNext
        Loop
        csiXMLData "CT", "newCustomGroup", ""
        ilRet = csiXMLWrite(1)
        ilRet = csiXMLEnd()
    Next ilGroup
End Function



Private Sub mGetAdminInfo()
    On Error GoTo ErrHand
    
    smAdminPhone = ""
    smAdminEMail = ""
    SQLQuery = "SELECT * FROM Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If rst!siteAdminArttCode > 0 Then
            SQLQuery = "SELECT * FROM ARTT Where arttCode = " & rst!siteAdminArttCode
            Set adrst = gSQLSelectCall(SQLQuery)
            If Not adrst.EOF Then
                smAdminPhone = Trim$(adrst!arttPhone)
                smAdminEMail = Trim$(adrst!arttEmail)
            End If
        End If
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "ExportOLA-mGetAdminInfo"
    Exit Sub
End Sub

Private Sub mCreateXMLSpot(slRegionName As String, llRegionIndex As Long, llLstCode As Long, llBreakNo As Long, ilPositionNo As Integer)
    'slType: "S" from ExportSpots, "R" from ExportRegionSpot
    Dim llAdf As Long
    Dim slAdvtName As String
    Dim llCntrNo As Long
    Dim slProd As String
    Dim llLstLogDate As Long
    Dim llLstLogTime As Long
    Dim slLogDate As String
    Dim slLogTime As String
    Dim ilLen As Integer
    Dim llCpfCode As Long
    Dim slISCI As String
    Dim slCart As String
    Dim slCreativeTitle As String
    Dim llCrfCsfCode As Long
    Dim slCSVRecord As String
    Dim ilRet As Integer
    Dim slRCartNo As String
    Dim slRProduct As String
    Dim slRISCI As String
    Dim slRCreativeTitle As String
    Dim llRCrfCsfCode As Long
    Dim llRCpfCode As Long
    Dim slComment As String
    Dim ilCifAdfCode As Integer
    Dim slHour As String
    Dim ilAnf As Integer
    
    On Error GoTo ErrHand
    If slRegionName <> "NAT_000" Then
        SQLQuery = "SELECT * FROM lst WHERE lstCode = " & llLstCode
        Set lst_rst = gSQLSelectCall(SQLQuery)
        If lst_rst.EOF Then
            'Error message
            Exit Sub
        End If
    End If

    ilAnf = gBinarySearchAnf(lst_rst!lstAnfCode)
    If ilAnf <> -1 Then
        If tgAvailNamesInfo(ilAnf).sAudioExport = "N" Then
            Exit Sub
        End If
    End If

    llAdf = gBinarySearchAdf(lst_rst!lstAdfCode)
    If llAdf <> -1 Then
        slAdvtName = Trim$(tgAdvtInfo(llAdf).sAdvtName)
    Else
        slAdvtName = "Advertiser Name Missing"
    End If
    
    llCntrNo = lst_rst!lstCntrNo
    slProd = Trim$(lst_rst!lstProd)
    If slProd = "" Then
        slProd = "N/A"
    End If
    llLstLogDate = DateValue(gAdjYear(Format$(lst_rst!lstLogDate, sgShowDateForm)))
    llLstLogTime = gTimeToLong(Format$(lst_rst!lstLogTime, sgShowTimeWSecForm), False)
    slLogDate = Format$(lst_rst!lstLogDate, "mm/dd/yyyy")
    slLogTime = Format$(lst_rst!lstLogTime, sgShowTimeWSecForm)
    ilLen = lst_rst!lstLen
    
    If slRegionName = "NAT_000" Then
        slCart = lst_rst!lstCart
        slISCI = lst_rst!lstISCI
        llCrfCsfCode = lst_rst!lstCrfCsfCode
        slCSVRecord = slLogDate & "," & slLogTime & "," & llBreakNo & "," & ilPositionNo & "," & llCntrNo & "," & gAddQuotes(slAdvtName) & "," & lst_rst!lstLineNo & "," & ilLen & "," & gAddQuotes(Trim$(lst_rst!lstISCI))
        SQLQuery = "Select rsfCode, rstPtType, rsfCopyCode, rsfCrfCode, rafName from RSF_Region_Schd_Copy, RAF_Region_Area"
        SQLQuery = SQLQuery & " Where (rsfSdfCode = " & lst_rst!lstSdfCode
        SQLQuery = SQLQuery & " AND rsfType <> 'B'"     'Blackout
        SQLQuery = SQLQuery & " AND rsfType <> 'A'"     'Airing vehicle copy
        SQLQuery = SQLQuery & " AND rafType = 'C'"     'Split copy
        SQLQuery = SQLQuery & " AND rafCode = rsfRafCode" & ")"
        Set rsf_rst = gSQLSelectCall(SQLQuery)
        If Not rsf_rst.EOF Then
            tmRegionBreakSpots(UBound(tmRegionBreakSpots)).lBreakNo = llBreakNo
            tmRegionBreakSpots(UBound(tmRegionBreakSpots)).iPositionNo = ilPositionNo
            tmRegionBreakSpots(UBound(tmRegionBreakSpots)).lLstCode = lst_rst!lstCode
            tmRegionBreakSpots(UBound(tmRegionBreakSpots)).lLogDate = llLstLogDate
            tmRegionBreakSpots(UBound(tmRegionBreakSpots)).lLogTime = llLstLogTime
            tmRegionBreakSpots(UBound(tmRegionBreakSpots)).lSdfCode = lst_rst!lstSdfCode
            tmRegionBreakSpots(UBound(tmRegionBreakSpots)).sISCI = UCase$(lst_rst!lstISCI)
            ReDim Preserve tmRegionBreakSpots(0 To UBound(tmRegionBreakSpots) + 1) As REGIONBREAKSPOTS
            If ckcGenCSV.Value = vbChecked Then
                Do
                    ilRet = gGetCopy(rsf_rst!rstPtType, rsf_rst!rsfCopyCode, rsf_rst!rsfCrfCode, True, slRCartNo, slRProduct, slRISCI, slRCreativeTitle, llRCrfCsfCode, llRCpfCode, ilCifAdfCode, lst_rst!lstLogVefCode)
                    slCSVRecord = slCSVRecord & "," & gAddQuotes(slRISCI & " (" & Trim$(rsf_rst!rafName) & ")")
                    rsf_rst.MoveNext
                Loop While Not rsf_rst.EOF
            End If
        End If
        If ckcGenCSV.Value = vbChecked Then
            gLogMsgWODT "W", hmCSV, slCSVRecord
        End If
        SQLQuery = "Select * from cpf_Copy_Prodct_ISCI WHERE cpfCode = " & lst_rst!lstCpfCode
        Set cpf_rst = gSQLSelectCall(SQLQuery)
        If Not cpf_rst.EOF Then
            slCreativeTitle = cpf_rst!cpfCreative
        Else
            slCreativeTitle = ""
        End If
    Else
        ilRet = gGetCopy(tmRegionDefinition(llRegionIndex).sPtType, tmRegionDefinition(llRegionIndex).lCopyCode, tmRegionDefinition(llRegionIndex).lCrfCode, True, slCart, slProd, slISCI, slCreativeTitle, llCrfCsfCode, llCpfCode, ilCifAdfCode, lst_rst!lstLogVefCode)
        If ilCifAdfCode <> lst_rst!lstAdfCode Then
            llAdf = gBinarySearchAdf(CLng(ilCifAdfCode))
            If llAdf <> -1 Then
                slAdvtName = tgAdvtInfo(llAdf).sAdvtName
            Else
                slAdvtName = "Blackout Advertiser Missing"
            End If
        End If
    End If
    csiXMLData "OT", "commercial", ""
    csiXMLData "CD", "csGroupCode", gXMLNameFilter(Trim$(slRegionName))
    csiXMLData "CD", "advertiserName", gXMLNameFilter(Trim$(slAdvtName))
    slISCI = gXMLNameFilter(Trim$(slISCI))
    If slISCI = "" Then
        slISCI = "N/A"
        gLogMsg "ISCI missing for " & slAdvtName & " Contract " & llCntrNo & " on " & slLogDate & " at " & slLogTime, "OLAExportLog.Txt", False
        mAddMsgToList "ISCI missing for " & slAdvtName & " Contract " & llCntrNo & " on " & slLogDate & " at " & slLogTime
    End If
    csiXMLData "CD", "ISCI", slISCI
    csiXMLData "CD", "orderNumber", llCntrNo
    slProd = gXMLNameFilter(Trim$(slProd))
    If slProd = "" Then
        slProd = "N/A"
    End If
    csiXMLData "CD", "product", slProd
    'csiXMLData "CD", "hourNumber", Format$(slLogTime, "h")
    slHour = Format$(slLogTime, "h")
    If imStartHourNumber = -1 Then
        imStartHourNumber = Val(slHour)
    End If
    slHour = Trim$(Str$(Val(slHour) - imStartHourNumber + 1))
    csiXMLData "CD", "hourNumber", slHour
    csiXMLData "CD", "breakNumber", llBreakNo
    csiXMLData "CD", "position", ilPositionNo
    csiXMLData "CD", "length", ilLen
    slCreativeTitle = gXMLNameFilter(Trim$(slCreativeTitle))
    If slCreativeTitle = "" Then
        slCreativeTitle = "N/A"
    End If
    csiXMLData "CD", "creativeTitle", slCreativeTitle
    csiXMLData "CD", "cartNumber", Trim$(slCart)
    csiXMLData "OT", "scheduledDate", ""
    csiXMLData "CD", "day", Format$(slLogDate, "d")
    csiXMLData "CD", "month", Format$(slLogDate, "m")
    csiXMLData "CD", "year", Format$(slLogDate, "yyyy")
    slHour = Format$(slLogTime, "h")
    'slHour = Trim$(Str$(Val(slHour) + 1))
    csiXMLData "CD", "hour", slHour 'Format$(slLogTime, "h")
    csiXMLData "CD", "minute", Format$(slLogTime, "nn")
    csiXMLData "CD", "seconds", Format$(slLogTime, "ss")
    csiXMLData "CT", "scheduledDate", ""
    'Add
    slComment = mGetCSFComment(llCrfCsfCode)
    csiXMLData "CD", "rotationComments", mXMLCommentFilter(255, slComment)
    
    
    csiXMLData "CT", "commercial", ""
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "Export OLA-mCreativeXMLSpot"
    Exit Sub
End Sub

Private Function mMergeCategory(llRegionIndex As Long) As Integer
    'Combine Region category definition togather.
    'This is used to form region from each spot with a break (The intersection of regions acrodd spots)
    'The result is used to determine if any station will receive the region copy defined by the merge.
    Dim llFormatIndex As Long
    Dim llOtherIndex As Long
    Dim llExcludeIndex As Long
    Dim llMergeOtherIndex As Long
    Dim llLastMergeOtherIndex As Long
    Dim llMergeExcludeIndex As Long
    Dim llLastMergeExcludeIndex As Long
    Dim ilShtt As Integer
    Dim ilAllowMerge As Integer
    Dim ilShttCode As Integer
    
    If UBound(tmMergeRegionDefinition) = 0 Then
        tmMergeRegionDefinition(UBound(tmMergeRegionDefinition)).lFormatFirst = -1
        tmMergeRegionDefinition(UBound(tmMergeRegionDefinition)).lOtherFirst = -1
        tmMergeRegionDefinition(UBound(tmMergeRegionDefinition)).lExcludeFirst = -1
        ReDim Preserve tmMergeRegionDefinition(0 To UBound(tmMergeRegionDefinition) + 1) As REGIONDEFINITION
    End If
    
    If tmRegionDefinition(llRegionIndex).lOtherFirst <> -1 Then
        llOtherIndex = tmRegionDefinition(llRegionIndex).lOtherFirst
        If tmSplitCategoryInfo(llOtherIndex).sCategory = "S" Then
            ilAllowMerge = False
            For ilShtt = 0 To UBound(tgStationInfoByCode) - 1 Step 1
                If tgStationInfoByCode(ilShtt).sUsedForOLA = "Y" Then
                    ilShttCode = tgStationInfoByCode(ilShtt).iCode
                    If tmSplitCategoryInfo(llOtherIndex).iIntCode = ilShttCode Then
                        ilAllowMerge = True
                        Exit For
                    End If
                End If
            Next ilShtt
            If ilAllowMerge = False Then
                mMergeCategory = False
                Exit Function
            End If
        End If
    End If

    If tmRegionDefinition(llRegionIndex).lFormatFirst <> -1 Then
        If tmMergeRegionDefinition(0).lFormatFirst <> -1 Then
            mMergeCategory = False
            Exit Function
        End If
        llFormatIndex = tmRegionDefinition(llRegionIndex).lFormatFirst
        tmMergeRegionDefinition(0).lFormatFirst = UBound(tmMergeSplitCategoryInfo)
        tmMergeSplitCategoryInfo(UBound(tmMergeSplitCategoryInfo)) = tmSplitCategoryInfo(llFormatIndex)
        ReDim Preserve tmMergeSplitCategoryInfo(0 To UBound(tmMergeSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
    End If
    If tmRegionDefinition(llRegionIndex).lOtherFirst <> -1 Then
        llOtherIndex = tmRegionDefinition(llRegionIndex).lOtherFirst
        If tmMergeRegionDefinition(0).lOtherFirst <> -1 Then
            llMergeOtherIndex = tmMergeRegionDefinition(0).lOtherFirst
            Do While llMergeOtherIndex <> -1
                If tmSplitCategoryInfo(llOtherIndex).sCategory = tmMergeSplitCategoryInfo(llMergeOtherIndex).sCategory Then
                    mMergeCategory = False
                    Exit Function
                End If
                llLastMergeOtherIndex = llMergeOtherIndex
                llMergeOtherIndex = tmMergeSplitCategoryInfo(llMergeOtherIndex).lNext
            Loop
            tmMergeSplitCategoryInfo(llLastMergeOtherIndex).lNext = UBound(tmMergeSplitCategoryInfo)
        Else
            tmMergeRegionDefinition(0).lOtherFirst = UBound(tmMergeSplitCategoryInfo)
        End If
        tmMergeSplitCategoryInfo(UBound(tmMergeSplitCategoryInfo)) = tmSplitCategoryInfo(llOtherIndex)
        tmMergeSplitCategoryInfo(UBound(tmMergeSplitCategoryInfo)).lNext = -1
        ReDim Preserve tmMergeSplitCategoryInfo(0 To UBound(tmMergeSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
    End If
    If tmRegionDefinition(llRegionIndex).lExcludeFirst <> -1 Then
        llExcludeIndex = tmRegionDefinition(llRegionIndex).lExcludeFirst
        Do While llExcludeIndex <> -1
            If tmMergeRegionDefinition(0).lExcludeFirst <> -1 Then
                llMergeExcludeIndex = tmMergeRegionDefinition(0).lExcludeFirst
                Do While llMergeExcludeIndex <> -1
                    llLastMergeExcludeIndex = llMergeExcludeIndex
                    llMergeExcludeIndex = tmMergeSplitCategoryInfo(llMergeExcludeIndex).lNext
                Loop
                tmMergeSplitCategoryInfo(llLastMergeExcludeIndex).lNext = UBound(tmMergeSplitCategoryInfo)
            Else
                tmMergeRegionDefinition(0).lExcludeFirst = UBound(tmMergeSplitCategoryInfo)
            End If
            tmMergeSplitCategoryInfo(UBound(tmMergeSplitCategoryInfo)) = tmSplitCategoryInfo(llExcludeIndex)
            tmMergeSplitCategoryInfo(UBound(tmMergeSplitCategoryInfo)).lNext = -1
            ReDim Preserve tmMergeSplitCategoryInfo(0 To UBound(tmMergeSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
            llExcludeIndex = tmSplitCategoryInfo(llExcludeIndex).lNext
        Loop
    End If
    mMergeCategory = True
End Function

Private Function mOpenCSF() As Integer

    Dim ilRet As Integer
    
    hmCsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCsf, "", sgDBPath & "CSF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gMsgBox "btrOpen Failed on CSF.BTR"
        ilRet = btrClose(hmCsf)
        btrDestroy hmCsf
        mOpenCSF = False
        Exit Function
    End If
    
    mOpenCSF = True

End Function
Private Function mCloseCSF()

    Dim ilRet As Integer
    
    ilRet = btrClose(hmCsf)
    If ilRet <> BTRV_ERR_NONE Then
        gMsgBox "btrClose Failed on CSF.BTR"
        btrDestroy hmCsf
        mCloseCSF = False
        Exit Function
    End If
    
    btrDestroy hmCsf
    mCloseCSF = True
    Exit Function

End Function

Private Function mGetCSFComment(lCSFCode As Long) As String

    Dim ilRet, i, ilLen, ilActualLen As Integer
    Dim ilRecLen As Integer
    Dim tlCSF As CSF
    Dim tlCsfSrchKey As LONGKEY0
    Dim slComment As String
    Dim slTemp As String
    Dim blOneChar As Byte
    
    On Error GoTo ErrHand
    
    mGetCSFComment = ""
    If (lCSFCode <= 0) Or (imCsfOpenStatus = False) Then
        Exit Function
    End If
    tlCsfSrchKey.lCode = lCSFCode
    tlCSF.sComment = ""
    ilRecLen = Len(tlCSF) '5011
    ilRet = btrGetEqual(hmCsf, tlCSF, ilRecLen, tlCsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Function
    End If

    slComment = gStripChr0(tlCSF.sComment)
    'If tlCSF.iStrLen > 0 Then
    If slComment <> "" Then
        'slComment = Trim$(Left$(tlCSF.sComment, tlCSF.iStrLen))
        ' Strip off any trailing non ascii characters.
        ilLen = Len(slComment)
        ' Find the first valid ascii character from the end and assume the rest of the string is good.
        For i = ilLen To 1 Step -1
            blOneChar = Asc(Mid(slComment, i, 1))
            If blOneChar >= 32 Then
                ' The first valid ASCII character has been found.
                slTemp = Left(slComment, i)
                Exit For
            End If
        Next i
        ilActualLen = i
        ' Scan through and remove any non ASCII characters. This was causing a problem for the web site.
        slComment = ""
        For i = 1 To ilActualLen
            blOneChar = Asc(Mid(slTemp, i, 1))
            If blOneChar >= 32 Then
                slComment = slComment + Mid(slTemp, i, 1)
            Else
                slComment = slComment + " "
            End If
        Next i
        mGetCSFComment = slComment
    End If

    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occurred in modPervasive-mGetCSFComment: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "WebExportLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Function

Private Function mExportUniqueNames(slNowDT As String, slOutputType As String) As Integer
    Dim ilUnique As Integer
    Dim ilRet As Integer
    
    mExportUniqueNames = True
         '6808
        If Not gDeleteFile(sgExportDirectory & "OLACopySplit.xml") Then
            mAddMsgToList "Could not delete file in mExportUniqueNames before writing.  Appended."
        End If
    ' Dan M 11/01/10 use smIniPathFileName that is created at formload
   ' ilRet = csiXMLStart(sgStartupDirectory & "\xml.ini", "OLA", slOutputType, sgExportDirectory & "OLACopySplit.xml", sgCRLF)
    '6807
    'ilRet = csiXMLStart(smIniPathFileName, "OLA", slOutputType, sgExportDirectory & "OLACopySplit.xml", sgCRLF)
    ilRet = csiXMLStart(smIniPathFileName, "OLA", slOutputType, sgExportDirectory & "OLACopySplit.xml", sgCRLF, "")
    ilRet = csiXMLSetMethod("", "", "", "CopySplitCustomGroupDefinition")
    For ilUnique = 0 To UBound(tmUniqueGroupNames) - 1 Step 1
        DoEvents
        If Trim$(tmUniqueGroupNames(ilUnique).sGroupType) = "Custom" Then
            csiXMLData "OT", "copySplit", ""
            csiXMLData "CD", "csGroupCode", Trim$(tmUniqueGroupNames(ilUnique).sGroupName)
            csiXMLData "CD", "csGroupType", Trim$(tmUniqueGroupNames(ilUnique).sGroupType)
            csiXMLData "CD", "description", gXMLNameFilter(Trim$(tmUniqueGroupNames(ilUnique).sName))
            csiXMLData "CT", "copySplit", ""
        End If
    Next ilUnique
    For ilUnique = 0 To UBound(tmUniqueGroupNames) - 1 Step 1
        DoEvents
        If Trim$(tmUniqueGroupNames(ilUnique).sGroupType) <> "Custom" Then
            csiXMLData "OT", "copySplit", ""
            csiXMLData "CD", "csGroupCode", Trim$(tmUniqueGroupNames(ilUnique).sGroupName)
            csiXMLData "CD", "csGroupType", Trim$(tmUniqueGroupNames(ilUnique).sGroupType)
            csiXMLData "CD", "description", gXMLNameFilter(Trim$(tmUniqueGroupNames(ilUnique).sName))
            csiXMLData "CT", "copySplit", ""
        End If
    Next ilUnique
    
    ilRet = csiXMLWrite(1)
    ilRet = csiXMLEnd()
End Function


Private Function mMergeOLAFiles(slXMLFileName As String) As Integer
    Dim ilRet As Integer
    Dim hlOLASpot As Integer
    Dim hlOLACopy As Integer
    Dim hlOLAMerge As Integer
    Dim slRecord As String
    Dim ilWriteCopy As Integer
    Dim slTemp As String

    'On Error GoTo mMergeOLAFilesErr:
    'ilRet = 0
    'hlOLASpot = FreeFile
    'Open sgExportDirectory & "OLASpot.XML" For Input Access Read As hlOLASpot
    slTemp = sgExportDirectory & "OLASpot.XML"
    ilRet = gFileOpen(slTemp, "Input Access Read", hlOLASpot)
    If ilRet <> 0 Then
        mAddMsgToList "Open " & sgExportDirectory & "OLACopySplit.xml" & " error#" & Str$(ilRet)
        gLogMsg "Open " & sgExportDirectory & "OLACopySplit.xml" & " error#" & Str$(ilRet), "OLAExportLog.Txt", False
        mMergeOLAFiles = False
        Exit Function
    End If
    'ilRet = 0
    'hlOLACopy = FreeFile
    'Open sgExportDirectory & "OLACopySplit.xml" For Input Access Read As hlOLACopy
    slTemp = sgExportDirectory & "OLACopySplit.xml"
    ilRet = gFileOpen(slTemp, "Input Access Read", hlOLACopy)
    If ilRet <> 0 Then
        mAddMsgToList "Open " & sgExportDirectory & "OLASpot.XML" & " error#" & Str$(ilRet)
        gLogMsg "Open " & sgExportDirectory & "OLASpot.XML" & " error#" & Str$(ilRet), "OLAExportLog.Txt", False
        mMergeOLAFiles = False
        Exit Function
    End If
    gLogMsgWODT "ON", hlOLAMerge, smExportPath & slXMLFileName
    Do While Not EOF(hlOLASpot)
        ilRet = 0
        On Error GoTo mMergeOLAFilesErr:
        Line Input #hlOLASpot, slRecord
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            mAddMsgToList "User Terminated"
            gLogMsg "User Terminated Export " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm), "OLAExportLog.Txt", False
            Close hlOLASpot
            Close hlOLACopy
            mMergeOLAFiles = False
            Exit Function
        End If
        slRecord = Trim$(slRecord)
        If Len(slRecord) > 0 Then
            If (Asc(slRecord) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                gLogMsgWODT "W", hlOLAMerge, slRecord
            End If
            If StrComp(slRecord, "</fileHeader>", vbTextCompare) = 0 Then
                Exit Do
            End If
        End If
        ilRet = 0
    Loop
    
    ilWriteCopy = False
    Do While Not EOF(hlOLACopy)
        ilRet = 0
        On Error GoTo mMergeOLAFilesErr:
        Line Input #hlOLACopy, slRecord
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            mAddMsgToList "User Terminated"
            gLogMsg "User Terminated Export " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm), "OLAExportLog.Txt", False
            Close hlOLASpot
            Close hlOLACopy
            mMergeOLAFiles = False
            Exit Function
        End If
        slRecord = Trim$(slRecord)
        If Len(slRecord) > 0 Then
            If (Asc(slRecord) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                If StrComp(slRecord, "<copySplit>", vbTextCompare) = 0 Then
                    ilWriteCopy = True
                End If
                If ilWriteCopy Then
                    gLogMsgWODT "W", hlOLAMerge, slRecord
                End If
                If StrComp(slRecord, "</copySplit>", vbTextCompare) = 0 Then
                    ilWriteCopy = False
                End If
            End If
        End If
        ilRet = 0
    Loop
    
    Do While Not EOF(hlOLASpot)
        ilRet = 0
        On Error GoTo mMergeOLAFilesErr:
        Line Input #hlOLASpot, slRecord
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            mAddMsgToList "User Terminated"
            gLogMsg "User Terminated Export " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm), "OLAExportLog.Txt", False
            Close hlOLASpot
            Close hlOLACopy
            mMergeOLAFiles = False
            Exit Function
        End If
        slRecord = Trim$(slRecord)
        If Len(slRecord) > 0 Then
            If (Asc(slRecord) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                gLogMsgWODT "W", hlOLAMerge, slRecord
            End If
        End If
        ilRet = 0
    Loop
    
    gLogMsgWODT "C", hlOLAMerge, ""
    Close hlOLASpot
    Close hlOLACopy
    mMergeOLAFiles = True
    Exit Function
mMergeOLAFilesErr:
    ilRet = Err.Number
    Resume Next

End Function
Private Function mUpdateShttUsedForOLA() As Integer
    Dim ilShtt As Integer
    Dim ilRet As Integer
    
    'On Error GoTo ErrHand
    'SQLQuery = "Update SHTT Set shttUsedForOLA = '" & "N" & "'"
    'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
    '    GoSub ErrHand:
    'End If
    'For ilShtt = 0 To UBound(imShttCodes) - 1 Step 1
    '    DoEvents
    '    SQLQuery = "Update SHTT Set shttUsedForOLA = '" & "Y" & "'" & " Where shttCode = " & imShttCodes(ilShtt)
    '    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
    '        GoSub ErrHand:
    '    End If
    'Next ilShtt
    'ilRet = gPopStations()
    mUpdateShttUsedForOLA = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmExporOla-mUpdateShttUsedForOLA"
    mUpdateShttUsedForOLA = False
    Exit Function
End Function


Private Function mFindMkt(slName As String, slGroupName As String) As Integer
    Dim ilLoop As Integer
    Dim ilCode As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    mFindMkt = 0
    If slName = "" Then
        Exit Function
    End If
    For ilLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
        If StrComp(Trim$(tgMarketInfo(ilLoop).sName), slName, vbTextCompare) = 0 Then
            mFindMkt = tgMarketInfo(ilLoop).lCode
            If StrComp(Trim$(tgMarketInfo(ilLoop).sGroupName), slGroupName, vbBinaryCompare) <> 0 Then
                'Update name
                SQLQuery = "UPDATE mkt"
                SQLQuery = SQLQuery & " SET mktUsfCode = " & igUstCode & ","
                SQLQuery = SQLQuery & "mktGroupName = '" & slGroupName & "'"
                SQLQuery = SQLQuery & " WHERE mktCode = " & tgMarketInfo(ilLoop).lCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "Export OLA-mFindMkt"
                    mFindMkt = 0
                    Exit Function
                End If
                tgMarketInfo(ilLoop).sGroupName = slGroupName
            End If
            Exit Function
        End If
    Next ilLoop
    'Add Market
    Do
        SQLQuery = "SELECT MAX(mktCode) from mkt"
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            ilCode = rst(0).Value + 1
        Else
            ilCode = 1
        End If
        ilRet = 0
        SQLQuery = "Insert into mkt "
        SQLQuery = SQLQuery & "(mktCode, mktName, mktUSFCode, mktGroupName, mktUnused) "
        SQLQuery = SQLQuery & " VALUES (" & ilCode & ",'" & gFixQuote(slName) & "'," & igUstCode & ",'" & slGroupName & "'," & "''" & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            If Not gHandleError4994("AffErrorLog.txt", "Export OLA-mFindMkt") Then
                mFindMkt = 0
                Exit Function
            End If
            ilRet = 1
        End If
    Loop While ilRet <> 0
    tgMarketInfo(UBound(tgMarketInfo)).lCode = ilCode
    tgMarketInfo(UBound(tgMarketInfo)).sName = slName
    tgMarketInfo(UBound(tgMarketInfo)).sGroupName = slGroupName
    ReDim Preserve tgMarketInfo(0 To UBound(tgMarketInfo) + 1) As MARKETINFO
    mFindMkt = ilCode
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frnExportOLA-mFindMkt"
End Function

Private Function mFindFmt(slName As String, slGroupName As String) As Integer
    Dim ilLoop As Integer
    Dim ilCode As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    mFindFmt = 0
    If slName = "" Then
        Exit Function
    End If
    For ilLoop = 0 To UBound(tgFormatInfo) - 1 Step 1
        If StrComp(Trim$(tgFormatInfo(ilLoop).sName), slName, vbTextCompare) = 0 Then
            mFindFmt = tgFormatInfo(ilLoop).lCode
            If StrComp(Trim$(tgFormatInfo(ilLoop).sGroupName), slGroupName, vbBinaryCompare) <> 0 Then
                'Update name
                SQLQuery = "UPDATE Fmt_Station_Format"
                SQLQuery = SQLQuery & " SET fmtUstCode = " & igUstCode & ","
                SQLQuery = SQLQuery & "fmtGroupName = '" & slGroupName & "'"
                SQLQuery = SQLQuery & " WHERE FmtCode = " & tgFormatInfo(ilLoop).lCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "Export OLA-mFinfFmt"
                    mFindFmt = 0
                    Exit Function
                End If
                tgFormatInfo(ilLoop).sGroupName = slGroupName
            End If
            Exit Function
        End If
    Next ilLoop
    'Add Format
    Do
        SQLQuery = "SELECT MAX(fmtCode) from Fmt_Station_Format"
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            ilCode = rst(0).Value + 1
        Else
            ilCode = 1
        End If
        ilRet = 0
        SQLQuery = "Insert into Fmt_Station_Format "
        SQLQuery = SQLQuery & "(fmtCode, fmtName, fmtUstCode, fmtGroupName, fmtDftCode, fmtUnused) "
        SQLQuery = SQLQuery & " VALUES (" & ilCode & ",'" & gFixQuote(slName) & "'," & igUstCode & ",'" & slGroupName & "'," & 0 & ",''" & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            If Not gHandleError4994("AffErrorLog.txt", "Export OLA-mFindFmt") Then
                mFindFmt = 0
                Exit Function
            End If
            ilRet = 1
        End If
    Loop While ilRet <> 0
    tgFormatInfo(UBound(tgFormatInfo)).lCode = ilCode
    tgFormatInfo(UBound(tgFormatInfo)).sName = slName
    tgFormatInfo(UBound(tgFormatInfo)).sGroupName = slGroupName
    ReDim Preserve tgFormatInfo(0 To UBound(tgFormatInfo) + 1) As FORMATINFO
    mFindFmt = ilCode
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError4994 "AffErorLog.txt", "frmExportOLA-mFindFmt"
End Function

Private Function mFindSnt(slName As String, slPostalName As String, slGroupName As String) As Integer
    Dim ilLoop As Integer
    Dim ilCode As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    mFindSnt = 0
    If slName = "" Then
        Exit Function
    End If
    For ilLoop = 0 To UBound(tgStateInfo) - 1 Step 1
        If StrComp(Trim$(tgStateInfo(ilLoop).sPostalName), slPostalName, vbTextCompare) = 0 Then
            mFindSnt = tgStateInfo(ilLoop).iCode
            If StrComp(Trim$(tgStateInfo(ilLoop).sGroupName), slGroupName, vbBinaryCompare) <> 0 Then
                'Update name
                SQLQuery = "UPDATE SNT"
                SQLQuery = SQLQuery & " SET sntGroupName = '" & slGroupName & "',"
                SQLQuery = SQLQuery & "sntUstCode = " & igUstCode
                SQLQuery = SQLQuery & " WHERE SntCode = " & tgStateInfo(ilLoop).iCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "Export OLA-mFindSnt"
                    mFindSnt = 0
                    Exit Function
                End If
                tgStateInfo(ilLoop).sGroupName = slGroupName
            End If
            Exit Function
        End If
    Next ilLoop
    'Add State
    Do
        SQLQuery = "SELECT MAX(sntCode) from snt"
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            ilCode = rst(0).Value + 1
        Else
            ilCode = 1
        End If
        ilRet = 0
        SQLQuery = "Insert into SNT "
        SQLQuery = SQLQuery & "(sntCode, sntName, sntPostalName, sntGroupName, sntUstCode, sntUnused) "
        SQLQuery = SQLQuery & " VALUES (" & ilCode & ",'" & gFixQuote(slName) & "','" & slPostalName & "','" & slGroupName & "'," & igUstCode & "," & "''" & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            If Not gHandleError4994("AffErrorLog.txt", "Export OLA-mFindSnt") Then
                mFindSnt = 0
                Exit Function
            End If
            ilRet = 1
        End If
    Loop While ilRet <> 0
    tgStateInfo(UBound(tgStateInfo)).iCode = ilCode
    tgStateInfo(UBound(tgStateInfo)).sName = slName
    tgStateInfo(UBound(tgStateInfo)).sPostalName = slPostalName
    tgStateInfo(UBound(tgStateInfo)).sGroupName = slGroupName
    ReDim Preserve tgStateInfo(0 To UBound(tgStateInfo) + 1) As STATEINFO
    mFindSnt = ilCode
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError4994 "AffErorLog.txt", "frmExportOLA-mFindSnt"
End Function

Private Function mFindTzt(slGroupName As String, slCSIName As String) As Integer
    Dim ilLoop As Integer
    Dim ilCode As Integer
    Dim ilRet As Integer
    Dim slName As String
    
    On Error GoTo ErrHand
    
    mFindTzt = 0
    If slGroupName = "" Then
        Exit Function
    End If
    Select Case slGroupName
        Case "TZ_EAS"
            slName = "Eastern"
            slCSIName = "EST"
        Case "TZ_CEN"
            slName = "Central"
            slCSIName = "CST"
        Case "TZ_MTN"
            slName = "Mountain"
            slCSIName = "MST"
        Case "TZ_PAC"
            slName = "Pacific"
            slCSIName = "PST"
        Case Else
            slName = ""
    End Select
    If slName <> "" Then
        For ilLoop = 0 To UBound(tgTimeZoneInfo) - 1 Step 1
            If StrComp(Trim$(tgTimeZoneInfo(ilLoop).sName), slName, vbTextCompare) = 0 Then
                mFindTzt = tgTimeZoneInfo(ilLoop).iCode
                slCSIName = Trim$(tgTimeZoneInfo(ilLoop).sCSIName)
                If StrComp(Trim$(tgTimeZoneInfo(ilLoop).sGroupName), slGroupName, vbBinaryCompare) <> 0 Then
                    'Update name
                    SQLQuery = "UPDATE TZT"
                    SQLQuery = SQLQuery & " SET TztGroupName = '" & slGroupName & "',"
                    SQLQuery = SQLQuery & "tztUstCode =" & igUstCode
                    SQLQuery = SQLQuery & " WHERE TztCode = " & tgTimeZoneInfo(ilLoop).iCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "Export OLA-mFindTzt"
                        mFindTzt = 0
                        Exit Function
                    End If
                    tgTimeZoneInfo(ilLoop).sGroupName = slGroupName
                End If
                Exit Function
            End If
        Next ilLoop
    End If
    For ilLoop = 0 To UBound(tgTimeZoneInfo) - 1 Step 1
        If StrComp(Trim$(tgTimeZoneInfo(ilLoop).sGroupName), slGroupName, vbTextCompare) = 0 Then
            mFindTzt = tgTimeZoneInfo(ilLoop).iCode
            slCSIName = Trim$(tgTimeZoneInfo(ilLoop).sCSIName)
            Exit Function
        End If
    Next ilLoop
    If slName <> "" Then
        'Add Time zone
        Do
            SQLQuery = "SELECT MAX(tztCode) from tzt"
            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                ilCode = rst(0).Value + 1
            Else
                ilCode = 1
            End If
            ilRet = 0
            SQLQuery = "Insert into TZT "
            SQLQuery = SQLQuery & "(tztCode, tztName, tztGroupName, tztCSIName, tztUstCode, TztUnused) "
            SQLQuery = SQLQuery & " VALUES (" & ilCode & ",'" & slName & "','" & slGroupName & "','" & slCSIName & "'," & igUstCode & "," & "''" & ")"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                If Not gHandleError4994("AffErrorLog.txt", "Export OLA-mFinfTzt") Then
                    mFindTzt = 0
                    Exit Function
                End If
                ilRet = 1
            End If
        Loop While ilRet <> 0
        tgTimeZoneInfo(UBound(tgTimeZoneInfo)).iCode = ilCode
        tgTimeZoneInfo(UBound(tgTimeZoneInfo)).sName = slName
        tgTimeZoneInfo(UBound(tgTimeZoneInfo)).sCSIName = slCSIName
        tgTimeZoneInfo(UBound(tgTimeZoneInfo)).sGroupName = slGroupName
        ReDim Preserve tgTimeZoneInfo(0 To UBound(tgTimeZoneInfo) + 1) As TIMEZONEINFO
    Else
        gLogMsg "Not a standard Time Zone Group Name found " & slGroupName, "OLAExportLog.Txt", False
        mAddMsgToList "Time Zone Group Missing " & slGroupName
    End If
    mFindTzt = ilCode
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError4994 "AffErorLog.txt", "frmExportOLA-mFindTzt"
End Function
