VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmExportSchdSpot 
   Caption         =   "Export Scheduled Station Spots"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "AffExportSchdSpot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   9615
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9405
      Top             =   3765
   End
   Begin V81Affiliate.CSI_Calendar edcDate 
      Height          =   285
      Left            =   1485
      TabIndex        =   1
      Top             =   180
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      Text            =   "5/7/2020"
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   3
   End
   Begin VB.TextBox edcTitle2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6570
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Results"
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox edcTitle3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4185
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Stations"
      Top             =   1530
      Width           =   1635
   End
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   1530
      Width           =   3150
   End
   Begin VB.CheckBox chkAllStation 
      Caption         =   "All"
      Height          =   195
      Left            =   4215
      TabIndex        =   8
      Top             =   3930
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.ListBox lbcStation 
      Height          =   2010
      ItemData        =   "AffExportSchdSpot.frx":08CA
      Left            =   4200
      List            =   "AffExportSchdSpot.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1770
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtNumberDays 
      Height          =   285
      Left            =   4605
      TabIndex        =   3
      Text            =   "7"
      Top             =   180
      Width           =   405
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3930
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   3375
      ItemData        =   "AffExportSchdSpot.frx":08CE
      Left            =   6585
      List            =   "AffExportSchdSpot.frx":08D0
      TabIndex        =   9
      Top             =   450
      Width           =   2820
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   2010
      ItemData        =   "AffExportSchdSpot.frx":08D2
      Left            =   135
      List            =   "AffExportSchdSpot.frx":08D4
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1770
      Width           =   3855
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   615
      Top             =   4305
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4755
      FormDesignWidth =   9615
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   5910
      TabIndex        =   10
      Top             =   4290
      Width           =   1665
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7755
      TabIndex        =   11
      Top             =   4290
      Width           =   1665
   End
   Begin V81Affiliate.AffExportCriteria udcCriteria 
      Height          =   810
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1429
   End
   Begin VB.Label Label2 
      Caption         =   "Number of Days"
      Height          =   225
      Left            =   3195
      TabIndex        =   2
      Top             =   210
      Width           =   1335
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   120
      TabIndex        =   12
      Top             =   4230
      Width           =   5580
   End
   Begin VB.Label Label1 
      Caption         =   "Export Start Date"
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   1395
   End
End
Attribute VB_Name = "frmExportSchdSpot"
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
Private imNumberDays As Integer
Private smEndDate As String
Private imVefCode As Integer
Private imAdfCode As Integer
Private smVefName As String
Private imAllClick As Integer
Private imAllStationClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
'Private hmMsg As Integer
Private hmTo As Integer
Private hmFrom As Integer
Private hmAst As Integer
Private cprst As ADODB.Recordset
Private smMessage As String
Private smWarnFlag As Integer
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private tmAetInfo() As AETINFO
Private tmAet() As AETINFO
Private lmEqtCode As Long
Private aet_rst As ADODB.Recordset
Private lst_rst As ADODB.Recordset
Private shtt_rst As ADODB.Recordset
Private vef_rst As ADODB.Recordset


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
'    'Print #hmMsg, "** Export Scheduled Station Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
'    'Print #hmMsg, ""
'    'sMsgFileName = slToFile
'    'mOpenMsgFile = True
'    Exit Function
'mOpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function

Private Sub mFillVehicle()
    Dim ilLoop As Integer
    Dim llVef As Long
    Dim slNowDate As String
    
    lbcVehicles.Clear
    lbcMsg.Clear
    chkAll.Value = vbUnchecked
    'For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
    '    'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
    '        lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
    '        lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
    '    'End If
    'Next iLoop
    slNowDate = Format(gNow(), sgSQLDateForm)
    SQLQuery = "SELECT DISTINCT attVefCode FROM att WHERE attDropDate > '" & slNowDate & "' AND attOffAir > '" & slNowDate & "' AND attExportType <> 0" & " AND attExportToUnivision = 'Y'"
    Set vef_rst = gSQLSelectCall(SQLQuery)
    Do While Not vef_rst.EOF
        llVef = gBinarySearchVef(CLng(vef_rst!attvefCode))
        If llVef <> -1 Then
            lbcVehicles.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
            lbcVehicles.ItemData(lbcVehicles.NewIndex) = vef_rst!attvefCode
        End If
        vef_rst.MoveNext
    Loop
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
        If lbcVehicles.ListCount > 1 Then
            edcTitle3.Visible = False
            chkAllStation.Visible = False
            lbcStation.Visible = False
            lbcStation.Clear
        Else
            edcTitle3.Visible = True
            chkAllStation.Visible = True
            lbcStation.Visible = True
        End If
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
    If lbcStation.ListCount > 0 Then
        imAllStationClick = True
        lRg = CLng(lbcStation.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStation.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllStationClick = False
    End If

End Sub


Private Sub cmdExport_Click()
    Dim iLoop As Integer
    Dim sFileName As String
    Dim ilRet As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim sToFile As String
    Dim sDateTime As String
    Dim sMsgFileName As String
    Dim sMoDate As String
    Dim sNowDate As String
    Dim llSDate As Long
    Dim llEDate As Long
    Dim llDate As Long
    Dim slDate As String
    Dim ilLoop As Integer
    Dim slExportType As String

    On Error GoTo ErrHand
    
    If imExporting Then
        Exit Sub
    End If
    imExporting = True
    
    lbcMsg.Clear
    If lbcVehicles.ListIndex < 0 Then
        igExportReturn = 2
        imExporting = False
        Exit Sub
    End If
    If edcDate.Text = "" Then
        imExporting = False
        gMsgBox "Date must be specified.", vbOKOnly
        edcDate.SetFocus
        Exit Sub
    End If
    If gIsDate(edcDate.Text) = False Then
        imExporting = False
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        edcDate.SetFocus
        Exit Sub
    Else
        smDate = Format(edcDate.Text, sgShowDateForm)
    End If
    sNowDate = Format$(gNow(), "m/d/yy")
    If DateValue(gAdjYear(smDate)) <= DateValue(gAdjYear(sNowDate)) Then
        imExporting = False
        Beep
        gMsgBox "Date must be after today's date " & sNowDate, vbCritical
        edcDate.SetFocus
        Exit Sub
    End If
    sMoDate = gObtainPrevMonday(smDate)
    llSDate = DateValue(gAdjYear(smDate))
    imNumberDays = Val(txtNumberDays.Text)
    If imNumberDays <= 0 Then
        imExporting = False
        gMsgBox "Number of days must be specified.", vbOKOnly
        txtNumberDays.SetFocus
        Exit Sub
    End If
    llEDate = DateValue(gAdjYear(Format$(DateAdd("d", imNumberDays - 1, smDate), "mm/dd/yy")))
    If (udcCriteria.rbcUSpot(0) = False) And (udcCriteria.rbcUSpot(1) = False) Then
        imExporting = False
        Beep
        gMsgBox "Please Specify Export Spots Type.", vbCritical
        Exit Sub
    End If
    If udcCriteria.rbcUSpot(0) = True Then
        slExportType = "!! Exporting All Spots for: "
    Else
        slExportType = "!! Exporting Spot Changes for: "
    End If


    Screen.MousePointer = vbHourglass
    mSaveCustomValues
    If igExportSource = 2 Then DoEvents
    If Not gPopCopy(sMoDate, "Export Schedule Spot") Then
        igExportReturn = 2
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        Exit Sub
    End If
    
    smWarnFlag = False
    If Not mCheckSelection() Then
        igExportReturn = 2
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        imExporting = False
        Screen.MousePointer = vbDefault
        'If rbcSpots(0).Value Then
        '    rbcSpots(1).SetFocus
        'Else
        '    rbcSpots(0).SetFocus
        'End If
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
'    If Not mOpenMsgFile(sMsgFileName) Then
'        Screen.MousePointer = vbDefault
'        cmdCancel.SetFocus
'        Exit Sub
'    End If
    imExporting = True
    ilRet = 0
    'On Error GoTo cmdExportErr:
    'sToFile = txtFile.Text
    sToFile = udcCriteria.edcUFile()
    'sDateTime = FileDateTime(sToFile)
    ilRet = gFileExist(sToFile)
    If ilRet = 0 Then
        sDateTime = gFileDateTime(sToFile)
        Screen.MousePointer = vbDefault
        ilRet = gMsgBox("Export Previously Created " & sDateTime & " Continue with Export by Replacing File?", vbOKCancel, "File Exist")
        If ilRet = vbCancel Then
            gLogMsg "** Terminated Because Export File Existed **", "UnivisionExportLog.Txt", False
            'Print #hmMsg, "** Terminated Because Export File Existed **"
            'Close #hmMsg
            Close #hmTo
            ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
            imExporting = False
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        Kill sToFile
    End If
    On Error GoTo 0
    'ilRet = 0
    'On Error GoTo cmdExportErr:
    'hmTo = FreeFile
    'Open sToFile For Output As hmTo
    ilRet = gFileOpen(sToFile, "Output", hmTo)
    If ilRet <> 0 Then
        gLogMsg "** Terminated because " & sToFile & " failed to open. **", "UnivisionExportLog.Txt", False
        'Print #hmMsg, "** Terminated **"
        'Close #hmMsg
        Close #hmTo
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        imExporting = False
        Screen.MousePointer = vbDefault
        gMsgBox "Open Error #" & Str$(Err.Numner) & sToFile, vbOKOnly, "Open Error"
        Exit Sub
    End If
    gLogMsg "** Storing Output into " & sToFile & " **", "UnivisionExportLog.Txt", False
    'Print #hmMsg, "** Storing Output into " & sToFile & " **"
    On Error GoTo 0
    lacResult.Caption = ""
    'D.S. 11/21/05
'    ilRet = gGetMaxAstCode()
'    If Not ilRet Then
'        Exit Sub
'    End If
    For iLoop = 0 To lbcVehicles.ListCount - 1
        If igExportSource = 2 Then DoEvents
        If lbcVehicles.Selected(iLoop) Then
            'Get hmTo handle
            imVefCode = lbcVehicles.ItemData(iLoop)
            smVefName = Trim$(lbcVehicles.List(iLoop))
            If sgShowByVehType = "Y" Then
                smVefName = Mid$(smVefName, 3)
            End If
            Screen.MousePointer = vbHourglass
            gLogMsg slExportType & smVefName & " With a Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", "UnivisionExportLog.Txt", False
            If smWarnFlag Then
                gLogMsg "Warning: Notified user that All Spots were previously exported for " & smVefName & " for the week of " & sMoDate & ", but the user chose to continue.", "UnivisionExportLog.Txt", False
                smWarnFlag = False
            End If
            ilRet = mExportSpots()
            If (ilRet = False) Then
                gCloseRegionSQLRst
                gLogMsg "** Terminated - mExportSpots returned False **", "UnivisionExportLog.Txt", False
                'Print #hmMsg, "** Terminated **"
                'Close #hmMsg
                Close #hmTo
                ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
                imExporting = False
                Screen.MousePointer = vbDefault
                cmdCancel.SetFocus
                Exit Sub
            End If
            If imTerminate Then
                gCloseRegionSQLRst
                gLogMsg "** User Terminated **", "UnivisionExportLog.Txt", False
                'Print #hmMsg, "** User Terminated **"
                'Close #hmMsg
                Close #hmTo
                ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
                imExporting = False
                Screen.MousePointer = vbDefault
                cmdCancel.SetFocus
                Exit Sub
            End If
            ilRet = gUpdateLastExportDate(imVefCode, smEndDate)
       End If
    Next iLoop
    Close #hmTo
    gCloseRegionSQLRst
    'Clear old aet records out
    On Error GoTo ErrHand:
    
    For ilLoop = 0 To lbcVehicles.ListCount - 1
        If igExportSource = 2 Then DoEvents
        If lbcVehicles.Selected(ilLoop) Then
            imVefCode = lbcVehicles.ItemData(ilLoop)
            For llDate = llSDate To llEDate Step 7
                If igExportSource = 2 Then DoEvents
                slDate = Format$(llDate, "m/d/yy")
                ilRet = gAlertClear("A", "F", "S", imVefCode, slDate)
                ilRet = gAlertClear("A", "R", "S", imVefCode, slDate)
            Next llDate
        End If
    Next ilLoop
    ilRet = gAlertForceCheck()
    
    sNowDate = Format$(gNow(), "m/d/yy")
    sNowDate = gObtainPrevMonday(sNowDate)
    If DateValue(gAdjYear(sNowDate)) < DateValue(gAdjYear(sMoDate)) Then
        sMoDate = sNowDate
    End If
    cnn.BeginTrans
    SQLQuery = "DELETE FROM Aet WHERE (aetStatus = 'I' And aetFeedDate <= '" & Format$(DateAdd("d", -28, sMoDate), sgSQLDateForm) & "')"
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "UnivisionExportLog.txt", "frmExportSchdSpot-cmdExport_Click"
        cnn.RollbackTrans
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        imExporting = False
        Exit Sub
    End If
    cnn.CommitTrans
    ilRet = gCustomEndStatus(lmEqtCode, 1, "")
    imExporting = False
    'Print #hmMsg, "** Completed Export Scheduled Station Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    gLogMsg "** Completed Export Scheduled Station Spots **", "UnivisionExportLog.Txt", False
    'Close #hmMsg
    lacResult.Caption = "Results: " & sMsgFileName
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    gLogMsg "", "UnivisionExportLog.Txt", False
    Exit Sub
'cmdExportErr:
'    ilRet = Err
'    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "UnivisionExportLog.txt", "frmExportSchdSpot-mcmdExport_Click"
    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
    imExporting = False
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    edcDate.Text = ""
    Unload frmExportSchdSpot
End Sub


Private Sub Form_Activate()
    Dim llVef As Long
    Dim ilLoop As Integer
    Dim hlResult As Integer
    Dim slNowStart As String
    Dim slNowEnd As String
    
    If imFirstTime Then
        udcCriteria.Left = Label1.Left
        udcCriteria.Height = (7 * Me.Height) / 10
        udcCriteria.Width = (7 * Me.Width) / 10
        'udcCriteria.Top = txtDate.Top + (3 * txtDate.Height) / 4
        udcCriteria.Top = Label1.Top + Label1.Height
        udcCriteria.Action 6
        If UBound(tgEvtInfo) > 0 Then
            chkAll.Value = vbUnchecked
            lbcStation.Clear
            lbcVehicles.Clear
            For ilLoop = 0 To UBound(tgEvtInfo) - 1 Step 1
                llVef = gBinarySearchVef(CLng(tgEvtInfo(ilLoop).iVefCode))
                If llVef <> -1 Then
                    lbcVehicles.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
                    lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgEvtInfo(ilLoop).iVefCode
                End If
            Next ilLoop
            chkAll.Value = vbChecked
            If lbcVehicles.ListCount = 1 Then
                imVefCode = lbcVehicles.ItemData(0)
                edcTitle3.Visible = True
                chkAllStation.Visible = True
                chkAllStation.Value = vbUnchecked
                lbcStation.Visible = True
                mFillStations
                chkAllStation.Value = vbChecked
            End If
        End If
        If igExportSource = 2 Then
            slNowStart = gNow()
            edcDate.Text = sgExporStartDate
            txtNumberDays.Text = igExportDays
            igExportReturn = 1
            '6394 move before 'click'
            sgExportResultName = "UnivisionResultList.Txt"
            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
            gLogMsgWODT "W", hlResult, "Univision Result List, Started: " & slNowStart
            hgExportResult = hlResult
            cmdExport_Click
            slNowEnd = gNow()
            'Output result list box
'            sgExportResultName = "UnivisionResultList.Txt"
'            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
'            gLogMsgWODT "W", hlResult, "Univision Result List, Started: " & slNowStart
            If lbcMsg.ListCount > 0 Then
                For ilLoop = 0 To lbcMsg.ListCount - 1 Step 1
                    gLogMsgWODT "W", hlResult, Trim$(lbcMsg.List(ilLoop))
                Next ilLoop
            End If
            gLogMsgWODT "W", hlResult, "Univision Result List, Completed: " & slNowEnd
            gLogMsgWODT "C", hlResult, ""
            '6394 clear values
            hgExportResult = 0
            imTerminate = True
            tmcTerminate.Enabled = True
        End If
        imFirstTime = False
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts Me
    If igExportSource = 2 Then
        Me.Top = -(2 * Me.Top + Screen.Height)
    End If
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    frmExportSchdSpot.Caption = "Scheduled Station Spots - " & sgClientName
    smDate = gObtainNextMonday(Format$(gNow(), sgShowDateForm))
    edcDate.Text = smDate
    imNumberDays = 7
    txtNumberDays.Text = Trim$(Str$(imNumberDays))
    imAllClick = False
    imAllStationClick = False
    imTerminate = False
    imExporting = False
    imFirstTime = True
    ReDim Preserve tgEvtInfo(0 To 0) As EVTINFO
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    lbcStation.Clear
    mFillVehicle
    'txtFile.Text = sgExportDirectory & "MktSpots.txt"
    chkAll.Value = vbChecked
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    If imExporting Then
        imTerminate = True
        Cancel = True
        Exit Sub
    End If
    
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    
    Erase tmCPDat
    Erase tmAstInfo
    Erase tmAetInfo
    Erase tmAet
    cprst.Close
    aet_rst.Close
    shtt_rst.Close
    lst_rst.Close
    vef_rst.Close
    Set frmExportSchdSpot = Nothing
End Sub


Private Sub lbcStation_Click()
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

Private Sub lbcVehicles_Click()
    Dim iLoop As Integer
    Dim iCount As Integer
    
    lbcStation.Clear
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
    If chkAllStation.Value = vbChecked Then
        chkAllStation.Value = vbUnchecked
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
    If iCount = 1 Then
        edcTitle3.Visible = True
        chkAllStation.Visible = True
        lbcStation.Visible = True
        mFillStations
    Else
        edcTitle3.Visible = False
        chkAllStation.Visible = False
        lbcStation.Visible = False
    End If
End Sub

Private Sub edcDate_Change()
    lbcMsg.Clear
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
End Sub

Private Function mExportSpots()
    Dim sDate As String
    Dim iNoWeeks As Integer
    Dim iLoop As Integer
    Dim iRet As Integer
    Dim sMoDate As String
    Dim sEndDate As String
    Dim sAdvt As String
    Dim sProd As String
    Dim sPledgeStartDate As String
    Dim sPledgeEndDate As String
    Dim iIndex As Integer
    ReDim iDays(0 To 6) As Integer
    Dim sPledgeStartTime As String
    Dim sPledgeEndTime As String
    Dim sLen As String
    Dim sCart As String
    Dim sISCI As String
    Dim sCreative As String
    Dim iDay As Integer
    Dim iAddDelete As Integer
    Dim iUpper As Integer
    Dim slStr As String
    Dim iAet As Integer
    Dim iFound As Integer
    Dim ilOkStation As Integer
    Dim slTemp As String
    Dim slSDate As String
    Dim slEDate As String
    Dim ilAddRecs As Integer
    Dim ilDeleteRecs As Integer
    Dim ilWriteHeader As Boolean
    Dim iExport As Integer  '0=Don't export as it did not change
                            '1=Export and create aet record
                            '2=Export and don't create aet reord (nothing changed but generating all spot export)
    
    Dim sRCart As String
    Dim sRISCI As String
    Dim sRCreative As String
    Dim sRProd As String
    Dim lRCrfCsfCode As Long
    Dim lRCrfCode As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    sMoDate = gObtainPrevMonday(smDate)
    sEndDate = DateAdd("d", imNumberDays - 1, smDate)
    slSDate = smDate
    slEDate = gObtainNextSunday(slSDate)
    If DateValue(gAdjYear(sEndDate)) < DateValue(gAdjYear(slEDate)) Then
        slEDate = sEndDate
    End If
    smEndDate = slEDate
    Do
        If igExportSource = 2 Then DoEvents
        ''Get CPTT so that Stations requiring CP can be obtained
        'SQLQuery = "SELECT shttCallLetters, shttMarket, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate"
        'SQLQuery = SQLQuery & " FROM shtt, cptt, att"
        'SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate, mktName"
        'SQLQuery = SQLQuery & " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode, cptt, att"
        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate"
        SQLQuery = SQLQuery & " FROM shtt, cptt, att"
        SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
        SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
        '10/29/14: Bypass Service agreements
        SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"
        SQLQuery = SQLQuery & " AND attExportToUnivision = 'Y'"
        SQLQuery = SQLQuery & " AND cpttVefCode = " & imVefCode
        SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sMoDate, sgSQLDateForm) & "')"
        Set cprst = gSQLSelectCall(SQLQuery)
        While Not cprst.EOF
            If igExportSource = 2 Then DoEvents
            If lbcStation.ListCount > 0 Then
                ilOkStation = False
                For iLoop = 0 To lbcStation.ListCount - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    If lbcStation.Selected(iLoop) Then
                        If lbcStation.ItemData(iLoop) = cprst!shttCode Then
                            ilOkStation = True
                            Exit For
                        End If
                    End If
                Next iLoop
            Else
                ilOkStation = True
            End If
            If ilOkStation Then
                ReDim tgCPPosting(0 To 1) As CPPOSTING
                tgCPPosting(0).lCpttCode = cprst!cpttCode
                tgCPPosting(0).iStatus = cprst!cpttStatus
                tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                tgCPPosting(0).lAttCode = cprst!cpttatfCode
                tgCPPosting(0).iAttTimeType = cprst!attTimeType
                tgCPPosting(0).iVefCode = imVefCode
                tgCPPosting(0).iShttCode = cprst!shttCode
                tgCPPosting(0).sZone = cprst!shttTimeZone
                tgCPPosting(0).sDate = Format$(sMoDate, sgShowDateForm)
                tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                ilWriteHeader = True
'                Print #hmTo, "A," & """" & "Counterpoint Software" & """"       'Network Provider Name
'                Print #hmTo, "B," & """" & "Marketron" & """"                   'Web Provider
'                Print #hmTo, "C," & """" & "Marketron" & """"                   'Station Provider
'                Print #hmTo, "D," & """" & "HBC" & """"                         'Station Provider
'                Print #hmTo, "E," & """" & Trim$(smVefName) & """" 'Vehicle name
'                Print #hmTo, "F," & """" & Trim$(cprst!shttCallLetters) & """"
                'Create AST records
                igTimes = 1 'By Week
                imAdfCode = -1
                If igExportSource = 2 Then DoEvents
                iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, True, True)
                gFilterAstExtendedTypes tmAstInfo
                ReDim tmAet(0 To 0) As AETINFO
                ReDim tmAetInfo(0 To 0) As AETINFO
                'Obtain past image
                SQLQuery = "SELECT aetCode, aetSdfCode, aetFeedDate, aetFeedTime, aetPledgeStartDate, aetPledgeEndDate, aetPledgeStartTime, aetPledgeEndTime, aetAdvt, aetProd, aetCart, aetISCI, aetCreative, aetAstCode, aetLen, aetCntrNo"
                SQLQuery = SQLQuery & " FROM aet"
                SQLQuery = SQLQuery & " WHERE (aetShfCode = " & cprst!shttCode
                'D.S. 10/25/04 commented out line below
                'SQLQuery = SQLQuery & " AND aetAtfCode = " & cprst!cpttatfCode
                SQLQuery = SQLQuery & " AND aetVefCode = " & imVefCode
                SQLQuery = SQLQuery & " AND aetStatus <> 'D' " & ")"
                'SQLQuery = SQLQuery & " AND (aetFeedDate >= '" & Format$(smDate, sgSQLDateForm) & "' AND aetFeedDate <= '" & Format$(sEndDate, sgSQLDateForm) & "')"
                SQLQuery = SQLQuery & " AND (aetFeedDate >= '" & Format$(slSDate, sgSQLDateForm) & "' AND aetFeedDate <= '" & Format$(slEDate, sgSQLDateForm) & "')"
                Set aet_rst = gSQLSelectCall(SQLQuery)
                While Not aet_rst.EOF
                    If igExportSource = 2 Then DoEvents
                    iUpper = UBound(tmAetInfo)
                    tmAetInfo(iUpper).lCode = aet_rst!aetCode
                    tmAetInfo(iUpper).lSdfCode = aet_rst!aetSdfCode
                    tmAetInfo(iUpper).sFeedDate = Format$(aet_rst!aetFeedDate, "mm/dd/yyyy")    'sgShowDateForm)
                    If Second(aet_rst!aetFeedTime) <> 0 Then
                        tmAetInfo(iUpper).sFeedTime = Format$(aet_rst!aetFeedTime, sgShowTimeWSecForm)
                    Else
                        tmAetInfo(iUpper).sFeedTime = Format$(aet_rst!aetFeedTime, sgShowTimeWOSecForm)
                    End If
                    tmAetInfo(iUpper).sPledgeStartDate = Format$(aet_rst!aetPledgeStartDate, "mm/dd/yyyy")    'sgShowDateForm)
                    tmAetInfo(iUpper).sPledgeEndDate = Format$(aet_rst!aetPledgeEndDate, "mm/dd/yyyy")    'sgShowDateForm)
                    If Second(aet_rst!aetPledgeStartTime) <> 0 Then
                        tmAetInfo(iUpper).sPledgeStartTime = Format$(aet_rst!aetPledgeStartTime, sgShowTimeWSecForm)
                    Else
                        tmAetInfo(iUpper).sPledgeStartTime = Format$(aet_rst!aetPledgeStartTime, sgShowTimeWOSecForm)
                    End If
                    If Not IsNull(aet_rst!aetPledgeEndTime) Then
                        If Second(aet_rst!aetPledgeEndTime) <> 0 Then
                            tmAetInfo(iUpper).sPledgeEndTime = Format$(aet_rst!aetPledgeEndTime, sgShowTimeWSecForm)
                        Else
                            tmAetInfo(iUpper).sPledgeEndTime = Format$(aet_rst!aetPledgeEndTime, sgShowTimeWOSecForm)
                        End If
                    Else
                        tmAetInfo(iUpper).sPledgeEndTime = ""
                    End If
                    tmAetInfo(iUpper).sAdvt = aet_rst!aetAdvt
                    tmAetInfo(iUpper).sProd = aet_rst!aetProd
                    tmAetInfo(iUpper).sCart = aet_rst!aetCart
                    tmAetInfo(iUpper).sISCI = aet_rst!aetISCI
                    tmAetInfo(iUpper).sCreative = aet_rst!aetCreative
                    tmAetInfo(iUpper).lAstCode = aet_rst!aetAstCode
                    tmAetInfo(iUpper).iLen = aet_rst!aetLen
                    tmAetInfo(iUpper).lCntrNo = aet_rst!aetCntrNo
                    tmAetInfo(iUpper).iProcessed = False
                    ReDim Preserve tmAetInfo(0 To iUpper + 1) As AETINFO
                    aet_rst.MoveNext
                Wend
                'Output AST
                ilAddRecs = 0
                For iLoop = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    'If (DateValue(tmAstInfo(iLoop).sFeedDate) >= DateValue(smDate)) And (DateValue(tmAstInfo(iLoop).sFeedDate) <= DateValue(sEndDate)) And (tgStatusTypes(tmAstInfo(iLoop).iPledgeStatus).iPledged <> 2) Then
                    If (DateValue(gAdjYear(tmAstInfo(iLoop).sFeedDate)) >= DateValue(gAdjYear(smDate))) And (DateValue(gAdjYear(tmAstInfo(iLoop).sFeedDate)) <= DateValue(gAdjYear(sEndDate))) And (tgStatusTypes(gGetAirStatus(tmAstInfo(iLoop).iStatus)).iPledged <> 2) Then
                        iAddDelete = 0
                        sAdvt = "Missing"
                        sCart = ""
                        sISCI = ""
                        sCreative = ""
                        SQLQuery = "SELECT lstProd, lstCart, lstISCI, adfName, cpfCreative"
                        SQLQuery = SQLQuery & " FROM (LST LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on lstCpfCode = cpfCode) LEFT OUTER JOIN ADF_Advertisers on lstadfCode = adfCode"
                        SQLQuery = SQLQuery & " WHERE lstCode =" & Str(tmAstInfo(iLoop).lLstCode)
                        Set lst_rst = gSQLSelectCall(SQLQuery)
                        If Not lst_rst.EOF Then
                            If igExportSource = 2 Then DoEvents
                            If IsNull(lst_rst!adfName) = True Then
                                sAdvt = "Missing"
                            Else
                                sAdvt = Trim$(lst_rst!adfName)
                            End If
                            If IsNull(lst_rst!lstProd) = True Then
                                sProd = ""
                            Else
                                sProd = Trim$(lst_rst!lstProd)
                            End If
                            If IsNull(lst_rst!lstCart) Or Left$(lst_rst!lstCart, 1) = Chr$(0) Then
                                sCart = ""
                            Else
                                sCart = Trim$(lst_rst!lstCart)
                            End If
                            If IsNull(lst_rst!lstISCI) = True Then
                                sISCI = ""
                            Else
                                sISCI = Trim$(lst_rst!lstISCI)
                            End If
                            If IsNull(lst_rst!cpfCreative) = True Then
                                sCreative = ""
                            Else
                                sCreative = Trim$(lst_rst!cpfCreative)
                            End If
                        End If
                        ''6/12/06- Check if any region copy defined for the spots
                        ''ilRet = gGetRegionCopy(tmAstInfo(iLoop).iShttCode, tmAstInfo(iLoop).lSdfCode, tmAstInfo(iLoop).iVefCode, sRCart, sRProd, sRISCI, sRCreative, lRCrfCsfCode, lRCrfCode)
                        'ilRet = gGetRegionCopy(tmAstInfo(iLoop), sRCart, sRProd, sRISCI, sRCreative, lRCrfCsfCode, lRCrfCode)
                        'If ilRet Then
                        If tmAstInfo(iLoop).iRegionType > 0 Then
                            sCart = Trim$(tmAstInfo(iLoop).sRCart) 'sRCart
                            sProd = Trim$(tmAstInfo(iLoop).sRProduct)  'sRProd
                            sISCI = Trim$(tmAstInfo(iLoop).sRISCI) 'sRISCI
                            sCreative = Trim$(tmAstInfo(iLoop).sRCreativeTitle)    'sRCreative
                        End If
                        sPledgeStartDate = Format$(tmAstInfo(iLoop).sPledgeDate, "m/d/yyyy")
                        If tgStatusTypes(gGetAirStatus(tmAstInfo(iLoop).iPledgeStatus)).iPledged = 0 Then
                            sPledgeEndDate = sPledgeStartDate
                        Else
                            gUnMapDays tmAstInfo(iLoop).sPdDays, iDays()
                            iDay = Weekday(sPledgeStartDate, vbMonday) - 1
                            sPledgeEndDate = sPledgeStartDate
                            For iIndex = iDay + 1 To 6 Step 1
                                If igExportSource = 2 Then DoEvents
                                If iDays(iIndex) Then
                                    sPledgeEndDate = DateAdd("d", 1, sPledgeEndDate)
                                Else
                                    Exit For
                                End If
                            Next iIndex
                        End If
                        If Second(tmAstInfo(iLoop).sPledgeStartTime) <> 0 Then
                            sPledgeStartTime = Format$(tmAstInfo(iLoop).sPledgeStartTime, "h:mm:ssa/p")
                        Else
                            sPledgeStartTime = Format$(tmAstInfo(iLoop).sPledgeStartTime, "h:mma/p")
                        End If
                        If Len(Trim$(tmAstInfo(iLoop).sPledgeEndTime)) <= 0 Then
                            sPledgeEndTime = sPledgeStartTime
                        Else
                            If Second(tmAstInfo(iLoop).sPledgeEndTime) <> 0 Then
                                sPledgeEndTime = Format$(tmAstInfo(iLoop).sPledgeEndTime, "h:mm:ssa/p")
                            Else
                                sPledgeEndTime = Format$(tmAstInfo(iLoop).sPledgeEndTime, "h:mma/p")
                            End If
                        End If
                        sLen = Trim$(Str$(tmAstInfo(iLoop).iLen))
                        iFound = False
                        iExport = 1
                        For iAet = 0 To UBound(tmAetInfo) - 1 Step 1
                            If igExportSource = 2 Then DoEvents
                            'Add astCode test to handle Load factor other then one
                            'If (tmAetInfo(iAet).lSdfCode = tmAstInfo(iLoop).lSdfCode) And (tmAetInfo(iAet).iProcessed = False) Then
                            If (tmAetInfo(iAet).lSdfCode = tmAstInfo(iLoop).lSdfCode) And (tmAetInfo(iAet).lAstCode = tmAstInfo(iLoop).lCode) And (tmAetInfo(iAet).iProcessed = False) Then
                                iFound = True
                                tmAetInfo(iAet).iProcessed = True
                                'Compare to see if different
                                iExport = 0
                                If DateValue(gAdjYear(sPledgeStartDate)) <> DateValue(gAdjYear(tmAetInfo(iAet).sPledgeStartDate)) Then
                                    iExport = 1
                                End If
                                If DateValue(gAdjYear(sPledgeEndDate)) <> DateValue(gAdjYear(tmAetInfo(iAet).sPledgeEndDate)) Then
                                    iExport = 1
                                End If
                                If TimeValue(sPledgeStartTime) <> TimeValue(tmAetInfo(iAet).sPledgeStartTime) Then
                                    iExport = 1
                                End If
                                If TimeValue(sPledgeEndTime) <> TimeValue(tmAetInfo(iAet).sPledgeEndTime) Then
                                    iExport = 1
                                End If
                                If StrComp(sAdvt, Trim$(tmAetInfo(iAet).sAdvt), vbTextCompare) <> 0 Then
                                    iExport = 1
                                End If
                                If StrComp(sProd, Trim$(tmAetInfo(iAet).sProd), vbTextCompare) <> 0 Then
                                    iExport = 1
                                End If
                                If StrComp(sCart, Trim$(tmAetInfo(iAet).sCart), vbTextCompare) <> 0 Then
                                    iExport = 1
                                End If
                                If StrComp(sISCI, Trim$(tmAetInfo(iAet).sISCI), vbTextCompare) <> 0 Then
                                    iExport = 1
                                End If
                                If StrComp(sCreative, Trim$(tmAetInfo(iAet).sCreative), vbTextCompare) <> 0 Then
                                    iExport = 1
                                End If
                                If tmAstInfo(iLoop).lCode <> tmAetInfo(iAet).lAstCode Then
                                    iExport = 1
                                End If
                                If tmAstInfo(iLoop).iLen <> tmAetInfo(iAet).iLen Then
                                    iExport = 1
                                End If
                                'If ((rbcSpots(1).Value) And (iExport = 1)) Or (rbcSpots(0).Value) Then
                                If ((udcCriteria.rbcUSpot(1)) And (iExport = 1)) Or (udcCriteria.rbcUSpot(0)) Then
                                    iExport = 1
                                    If ilWriteHeader Then
                                        If igExportSource = 2 Then DoEvents
                                        Print #hmTo, "A," & """" & "Counterpoint Software" & """"       'Network Provider Name
                                        Print #hmTo, "B," & """" & "Marketron" & """"                   'Web Provider
                                        Print #hmTo, "C," & """" & "Marketron" & """"                   'Station Provider
                                        Print #hmTo, "D," & """" & "HBC" & """"                         'Station Provider
                                        Print #hmTo, "E," & """" & Trim$(smVefName) & """" 'Vehicle name
                                        Print #hmTo, "F," & """" & Trim$(cprst!shttCallLetters) & """"
                                        ilWriteHeader = False
                                        If igExportSource = 2 Then DoEvents
                                    End If
                                    ''Print #hmTo, "S," & """" & sAdvtProd & """" & "," & sPledgeStartDate & "-" & sPledgeEndDate & "," & sPledgeStartTime & "," & sPledgeEndTime & "," & sLen & "," & """" & sCart & """" & "," & """" & sISCI & """" & "," & """" & sCreative & """" & "," & Trim$(Str$(tmAstInfo(iLoop).lCode))
                                    'slStr = "S," & """" & Trim$(tmAetInfo(iAet).sAdvt) & """" & "," & """" & Trim$(tmAetInfo(iAet).sProd) & """" & ",D," & Trim$(tmAetInfo(iAet).sPledgeStartDate) & "," & Trim$(tmAetInfo(iAet).sPledgeEndDate) & "," & Trim$(tmAetInfo(iAet).sPledgeStartTime) & "," & Trim$(tmAetInfo(iAet).sPledgeEndTime) & "," & Trim$(Str$(tmAetInfo(iAet).iLen)) & ","
                                    slStr = "S," & """" & Trim$(tmAetInfo(iAet).sAdvt) & """" & ","
                                    If Trim$(tmAetInfo(iAet).sProd) <> "" Then
                                        slStr = slStr & """" & Trim$(tmAetInfo(iAet).sProd) & """" & ","
                                    Else
                                        slStr = slStr & ","
                                    End If
                                    slStr = slStr & "D," & Trim$(tmAetInfo(iAet).sPledgeStartDate) & "," & Trim$(tmAetInfo(iAet).sPledgeEndDate) & "," & Trim$(tmAetInfo(iAet).sPledgeStartTime) & "," & Trim$(tmAetInfo(iAet).sPledgeEndTime) & "," & Trim$(Str$(tmAetInfo(iAet).iLen)) & ","
                                    If Trim$(tmAetInfo(iAet).sCart) <> "" Then
                                        slStr = slStr & """" & Trim$(tmAetInfo(iAet).sCart) & """" & ","
                                    Else
                                        slStr = slStr & ","
                                    End If
                                    If Trim$(tmAetInfo(iAet).sISCI) <> "" Then
                                        slStr = slStr & """" & Trim$(tmAetInfo(iAet).sISCI) & """" & ","
                                    Else
                                        slStr = slStr & ","
                                    End If
                                    If Trim$(tmAetInfo(iAet).sCreative) <> "" Then
                                        slStr = slStr & """" & Trim$(tmAetInfo(iAet).sCreative) & """" & ","
                                    Else
                                        slStr = slStr & ","
                                    End If
                                    slStr = slStr & Trim$(Str$(tmAetInfo(iAet).lAstCode))
                                    Print #hmTo, slStr
                                    ilAddRecs = ilAddRecs + 1
                                Else
                                    tmAetInfo(iAet).lCode = 0
                                End If
                                Exit For
                            End If
                        Next iAet
                        If igExportSource = 2 Then DoEvents
                        If InStr(1, sAdvt, "Missing", vbTextCompare) = 1 Then
                            gLogMsg Trim$(smVefName) & ": Advertiser Missing on " & Format$(tmAstInfo(iLoop).sAirDate, "m/d/yy") & " at " & Format$(tmAstInfo(iLoop).sAirTime, "h:mm:ssAM/PM"), "UnivisionExportLog.Txt", False
                            'Print #hmMsg, Trim$(smVefName) & ": Advertiser Missing on " & Format$(tmAstInfo(iLoop).sAirDate, "m/d/yy") & " at " & Format$(tmAstInfo(iLoop).sAirTime, "h:mm:ssAM/PM")
                            lbcMsg.AddItem Trim$(smVefName) & ": Advertiser Missing on " & Format$(tmAstInfo(iLoop).sAirDate, "m/d/yy") & " at " & Format$(tmAstInfo(iLoop).sAirTime, "h:mm:ssAM/PM")
                        Else
                            If iExport <> 0 Then
                                If ilWriteHeader Then
                                    If igExportSource = 2 Then DoEvents
                                    Print #hmTo, "A," & """" & "Counterpoint Software" & """"       'Network Provider Name
                                    Print #hmTo, "B," & """" & "Marketron" & """"                   'Web Provider
                                    Print #hmTo, "C," & """" & "Marketron" & """"                   'Station Provider
                                    Print #hmTo, "D," & """" & "HBC" & """"                         'Station Provider
                                    Print #hmTo, "E," & """" & Trim$(smVefName) & """" 'Vehicle name
                                    Print #hmTo, "F," & """" & Trim$(cprst!shttCallLetters) & """"
                                    ilWriteHeader = False
                                    If igExportSource = 2 Then DoEvents
                                End If
                                ''Print #hmTo, "S," & """" & sAdvtProd & """" & "," & sPledgeStartDate & "-" & sPledgeEndDate & "," & sPledgeStartTime & "," & sPledgeEndTime & "," & sLen & "," & """" & sCart & """" & "," & """" & sISCI & """" & "," & """" & sCreative & """" & "," & Trim$(Str$(tmAstInfo(iLoop).lCode))
                                'slStr = "S," & """" & sAdvt & """" & "," & """" & sProd & """" & ",A," & sPledgeStartDate & "," & sPledgeEndDate & "," & sPledgeStartTime & "," & sPledgeEndTime & "," & sLen & ","
                                slStr = "S," & """" & sAdvt & """" & ","
                                If sProd <> "" Then
                                    slStr = slStr & """" & sProd & """" & ","
                                Else
                                    slStr = slStr & ","
                                End If
                                slStr = slStr & "A," & sPledgeStartDate & "," & sPledgeEndDate & "," & sPledgeStartTime & "," & sPledgeEndTime & "," & sLen & ","
                                If sCart <> "" Then
                                    slStr = slStr & """" & sCart & """" & ","
                                Else
                                    slStr = slStr & ","
                                End If
                                If sISCI <> "" Then
                                    slStr = slStr & """" & sISCI & """" & ","
                                Else
                                    slStr = slStr & ","
                                End If
                                If sCreative <> "" Then
                                    slStr = slStr & """" & sCreative & """" & ","
                                Else
                                    slStr = slStr & ","
                                End If
                                slStr = slStr & Trim$(Str$(tmAstInfo(iLoop).lCode))
                                Print #hmTo, slStr
                                ilAddRecs = ilAddRecs + 1
                                If iExport = 1 Then
                                    If igExportSource = 2 Then DoEvents
                                    iUpper = UBound(tmAet)
                                    tmAet(iUpper).lCode = 0
                                    tmAet(iUpper).lAtfCode = tmAstInfo(iLoop).lAttCode
                                    tmAet(iUpper).iShfCode = tmAstInfo(iLoop).iShttCode
                                    tmAet(iUpper).iVefCode = tmAstInfo(iLoop).iVefCode
                                    tmAet(iUpper).lSdfCode = tmAstInfo(iLoop).lSdfCode
                                    tmAet(iUpper).sFeedDate = tmAstInfo(iLoop).sFeedDate
                                    tmAet(iUpper).sFeedTime = tmAstInfo(iLoop).sFeedTime
                                    tmAet(iUpper).sPledgeStartDate = sPledgeStartDate
                                    tmAet(iUpper).sPledgeEndDate = sPledgeEndDate
                                    tmAet(iUpper).sPledgeStartTime = sPledgeStartTime
                                    tmAet(iUpper).sPledgeEndTime = sPledgeEndTime
                                    tmAet(iUpper).sAdvt = sAdvt
                                    tmAet(iUpper).sProd = sProd
                                    tmAet(iUpper).sCart = sCart
                                    tmAet(iUpper).sISCI = sISCI
                                    tmAet(iUpper).sCreative = sCreative
                                    tmAet(iUpper).lAstCode = tmAstInfo(iLoop).lCode
                                    tmAet(iUpper).iLen = Val(sLen)
                                    tmAet(iUpper).lCntrNo = tmAstInfo(iLoop).lCntrNo
                                    ReDim Preserve tmAet(0 To iUpper + 1) As AETINFO
                                    If igExportSource = 2 Then DoEvents
                                End If
                            End If
                        End If
                    End If
                Next iLoop
                ilDeleteRecs = 0
                For iAet = 0 To UBound(tmAetInfo) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    If tmAetInfo(iAet).iProcessed = False Then
                        tmAetInfo(iAet).iProcessed = True
                        If ilWriteHeader Then
                            If igExportSource = 2 Then DoEvents
                            Print #hmTo, "A," & """" & "Counterpoint Software" & """"       'Network Provider Name
                            Print #hmTo, "B," & """" & "Marketron" & """"                   'Web Provider
                            Print #hmTo, "C," & """" & "Marketron" & """"                   'Station Provider
                            Print #hmTo, "D," & """" & "HBC" & """"                         'Station Provider
                            Print #hmTo, "E," & """" & Trim$(smVefName) & """" 'Vehicle name
                            Print #hmTo, "F," & """" & Trim$(cprst!shttCallLetters) & """"
                            ilWriteHeader = False
                            If igExportSource = 2 Then DoEvents
                        End If
                        ''Print #hmTo, "S," & """" & sAdvtProd & """" & "," & sPledgeStartDate & "-" & sPledgeEndDate & "," & sPledgeStartTime & "," & sPledgeEndTime & "," & sLen & "," & """" & sCart & """" & "," & """" & sISCI & """" & "," & """" & sCreative & """" & "," & Trim$(Str$(tmAstInfo(iLoop).lCode))
                        'slStr = "S," & """" & Trim$(tmAetInfo(iAet).sAdvt) & """" & "," & """" & Trim$(tmAetInfo(iAet).sProd) & """" & ",D," & Trim$(tmAetInfo(iAet).sPledgeStartDate) & "," & Trim$(tmAetInfo(iAet).sPledgeEndDate) & "," & Trim$(tmAetInfo(iAet).sPledgeStartTime) & "," & Trim$(tmAetInfo(iAet).sPledgeEndTime) & "," & Trim$(Str$(tmAetInfo(iAet).iLen)) & ","
                        slStr = "S," & """" & Trim$(tmAetInfo(iAet).sAdvt) & """" & ","
                        If Trim$(tmAetInfo(iAet).sProd) <> "" Then
                            slStr = slStr & """" & Trim$(tmAetInfo(iAet).sProd) & """" & ","
                        Else
                            slStr = slStr & ","
                        End If
                        slStr = slStr & "D," & Trim$(tmAetInfo(iAet).sPledgeStartDate) & "," & Trim$(tmAetInfo(iAet).sPledgeEndDate) & "," & Trim$(tmAetInfo(iAet).sPledgeStartTime) & "," & Trim$(tmAetInfo(iAet).sPledgeEndTime) & "," & Trim$(Str$(tmAetInfo(iAet).iLen)) & ","
                        If Trim$(tmAetInfo(iAet).sCart) <> "" Then
                            slStr = slStr & """" & Trim$(tmAetInfo(iAet).sCart) & """" & ","
                        Else
                            slStr = slStr & ","
                        End If
                        If tmAetInfo(iAet).sISCI <> "" Then
                            slStr = slStr & """" & Trim$(tmAetInfo(iAet).sISCI) & """" & ","
                        Else
                            slStr = slStr & ","
                        End If
                        If Trim$(tmAetInfo(iAet).sCreative) <> "" Then
                            slStr = slStr & """" & Trim$(tmAetInfo(iAet).sCreative) & """" & ","
                        Else
                            slStr = slStr & ","
                        End If
                        slStr = slStr & Trim$(Str$(tmAetInfo(iAet).lAstCode))
                        Print #hmTo, slStr
                        'Update status only
                        If igExportSource = 2 Then DoEvents
                        SQLQuery = "UPDATE aet SET "
                        SQLQuery = SQLQuery & "aetStatus = 'D'"
                        SQLQuery = SQLQuery & " WHERE aetCode = " & tmAetInfo(iAet).lCode & ""
                        cnn.BeginTrans
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "UnivisionExportLog.txt", "frmExportSchdSpot-mExportSpots"
                            cnn.RollbackTrans
                            mExportSpots = False
                            Exit Function
                        End If
                        cnn.CommitTrans
                        tmAetInfo(iAet).lCode = 0
                        ilDeleteRecs = ilDeleteRecs + 1
                    End If
                Next iAet
                
                'Remove previously created aet
                If igExportSource = 2 Then DoEvents
                For iAet = 0 To UBound(tmAetInfo) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    If tmAetInfo(iAet).lCode > 0 Then
                        cnn.BeginTrans
                        SQLQuery = "DELETE FROM Aet WHERE (aetCode = " & tmAetInfo(iAet).lCode & ")"
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "UnivisionExportLog.txt", "frmExportSchdSpot-mExportSpots"
                            cnn.RollbackTrans
                            mExportSpots = False
                            Exit Function
                        End If
                        cnn.CommitTrans
                    End If
                Next iAet
                'Add ast as aet
                For iAet = 0 To UBound(tmAet) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    SQLQuery = "INSERT INTO aet"
                    SQLQuery = SQLQuery & "(aetAtfCode, aetShfCode, aetVefCode, "
                    SQLQuery = SQLQuery & "aetSdfCode, aetFeedDate, aetFeedTime, "
                    SQLQuery = SQLQuery & " aetPledgeStartDate, aetPledgeEndDate, "
                    SQLQuery = SQLQuery & "aetPledgeStartTime, aetPledgeEndTime, aetAdvt, aetProd, "
                    SQLQuery = SQLQuery & "aetCart, aetISCI, aetCreative, "
                    SQLQuery = SQLQuery & "aetAstCode, aetLen, aetCntrNo, aetStatus)"
                    SQLQuery = SQLQuery & " VALUES "
                    SQLQuery = SQLQuery & "(" & tmAet(iAet).lAtfCode & ", " & tmAet(iAet).iShfCode & ", "
                    SQLQuery = SQLQuery & tmAet(iAet).iVefCode & ", " & tmAet(iAet).lSdfCode & ", "
                    SQLQuery = SQLQuery & "'" & Format$(tmAet(iAet).sFeedDate, sgSQLDateForm) & "', '" & Format$(tmAet(iAet).sFeedTime, sgSQLTimeForm) & "', "
                    SQLQuery = SQLQuery & "'" & Format$(tmAet(iAet).sPledgeStartDate, sgSQLDateForm) & "', '" & Format$(tmAet(iAet).sPledgeEndDate, sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery & "'" & Format$(tmAet(iAet).sPledgeStartTime, sgSQLTimeForm) & "', '" & Format$(tmAet(iAet).sPledgeEndTime, sgSQLTimeForm) & "', "
                    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmAet(iAet).sAdvt)) & "', '" & gFixQuote(Trim$(tmAet(iAet).sProd)) & "', '" & gFixQuote(Trim$(tmAet(iAet).sCart)) & "', '" & gFixQuote(Trim$(tmAet(iAet).sISCI)) & "', "
                    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmAet(iAet).sCreative)) & "', " & tmAet(iAet).lAstCode & ", " & tmAet(iAet).iLen & ", "
                    SQLQuery = SQLQuery & tmAet(iAet).lCntrNo & ", " & "'A'" & ")"
                    cnn.BeginTrans
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "UnivisionExportLog.txt", "frmExportSchdSpot-mExportSpots"
                        cnn.RollbackTrans
                        mExportSpots = False
                        Exit Function
                    End If
                    cnn.CommitTrans
                    If igExportSource = 2 Then DoEvents
                    SQLQuery = "Select MAX(aetCode) from aet"
                    Set aet_rst = gSQLSelectCall(SQLQuery)
                    tmAet(iAet).lCode = aet_rst(0).Value
                Next iAet
            End If
            
            ' JD 04/25/05
            'Delete Cptt records if necessary
            iRet = mAdjCpttRecs(cprst)
            
            gLogMsg Trim$(cprst!shttCallLetters) & " Exported " & CStr(ilAddRecs) & " Add records and " & CStr(ilDeleteRecs) & " Delete Records", "UnivisionExportLog.Txt", False
            cprst.MoveNext
        Wend
        If (lbcStation.ListCount = 0) Or (chkAllStation.Value = vbChecked) Or (lbcStation.ListCount = lbcStation.SelCount) Then
            gClearASTInfo True
            'gClearAbf imVefCode, 0, sMoDate, gObtainNextSunday(sMoDate)
        Else
            gClearASTInfo False
        End If
        sMoDate = DateAdd("d", 7, sMoDate)
        slSDate = sMoDate
        slEDate = gObtainNextSunday(slSDate)
        If DateValue(gAdjYear(sEndDate)) < DateValue(gAdjYear(slEDate)) Then
            slEDate = sEndDate
        End If
    Loop While DateValue(gAdjYear(sMoDate)) < DateValue(gAdjYear(sEndDate))

    mExportSpots = True
    Exit Function
mExportSpotsErr:
    iRet = Err

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "UnivisionExportLog.txt", "frmExportSchdSpot-mExportSpots"
    mExportSpots = False
    Exit Function
    
End Function

Private Sub mFillStations()
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
    SQLQuery = SQLQuery & " FROM shtt, att"
    SQLQuery = SQLQuery & " WHERE (attVefCode = " & imVefCode
    SQLQuery = SQLQuery & " AND attExportType <> 0 "
    SQLQuery = SQLQuery & " AND attExportToUnivision = 'Y' "
    SQLQuery = SQLQuery & " AND shttCode = attShfCode)"
    SQLQuery = SQLQuery & " ORDER BY shttCallLetters"
    Set shtt_rst = gSQLSelectCall(SQLQuery)
    While Not shtt_rst.EOF
        lbcStation.AddItem Trim$(shtt_rst!shttCallLetters)
        lbcStation.ItemData(lbcStation.NewIndex) = shtt_rst!shttCode
        shtt_rst.MoveNext
    Wend
    chkAllStation.Value = vbChecked
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "UnivisionExportLog.txt", "frmExportSchdSpot-mFillStations"

End Sub

Private Function mCheckLastExportDate() As Integer
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim slDate As String
    Dim ilLoop As Integer
    'Dim slFields(1 To 15) As String
    Dim slFields(0 To 14) As String
    
    'slFromFile = txtFile.Text
    slFromFile = udcCriteria.edcUFile()
    'ilRet = 0
    'On Error GoTo mCheckLastExportDateErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Output", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        If udcCriteria.rbcUSpot(0) Then
            mCheckLastExportDate = True
        Else
        End If
        Exit Function
    End If
    slDate = ""
    Do While Not EOF(hmFrom)
        If igExportSource = 2 Then DoEvents
        ilRet = 0
        On Error GoTo mCheckLastExportDateErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, False, slFields()
                For ilLoop = LBound(slFields) To UBound(slFields) Step 1
                    If igExportSource = 2 Then DoEvents
                    slFields(ilLoop) = Trim$(slFields(ilLoop))
                Next ilLoop
                'If slFields(1) = "S" Then
                If slFields(0) = "S" Then
                    On Error GoTo ErrHand
                    If slDate = "" Then
                        'slDate = slFields(5)
                        slDate = slFields(4)
                    Else
                        'If DateValue(gAdjYear(slFields(5))) > DateValue(gAdjYear(slDate)) Then
                        If DateValue(gAdjYear(slFields(4))) > DateValue(gAdjYear(slDate)) Then
                            'slDate = slFields(5)
                            slDate = slFields(4)
                        End If
                    End If
                End If
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    If slDate = "" Then
    Else
    End If
    Exit Function
mCheckLastExportDateErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "UnivisionExportLog.txt", "frmExportSchdSpot-mCheckLastExportDate"
    ilRet = 1
End Function

Private Function mCheckSelection() As Integer

    Dim ilRet As Integer
    Dim slMsg As String
    Dim slSQLQuery As String
    
    If igExportSource = 2 Then DoEvents
    slSQLQuery = "SELECT DISTINCT aetFeedDate"
    slSQLQuery = slSQLQuery & " FROM aet"
    slSQLQuery = slSQLQuery & " WHERE (aetFeedDate >= '" & Format$(smDate, sgSQLDateForm) & "' AND aetFeedDate <= '" & Format$(DateAdd("d", imNumberDays - 1, smDate), sgSQLDateForm) & "')"
    Set aet_rst = gSQLSelectCall(slSQLQuery)
    If aet_rst.EOF Then
        If udcCriteria.rbcUSpot(1) Then
            Screen.MousePointer = vbDefault
            gMsgBox "You must select 'all spots' for the specified dates before selecting 'spot changes'.", vbOKOnly
            gLogMsg "You must select 'all spots' for the specified dates before selecting 'spot changes'.", "UnivisionExportLog.Txt", False
            mCheckSelection = False
            Exit Function
        End If
    Else
        If udcCriteria.rbcUSpot(0) Then
            Screen.MousePointer = vbDefault
            ilRet = gMsgBox("Warning: You have already generated 'all spots' for this period." & Chr$(13) & Chr$(10) & "Do not proceed if stations have already received 'all spots' export." & Chr$(13) & Chr$(10) & "Continue with 'All Spot' Export?", vbYesNo)
            If ilRet = vbNo Then
                mCheckSelection = False
                Exit Function
            Else
                smWarnFlag = True
            End If
        End If
    End If
    mCheckSelection = True
    Exit Function
mCheckSelectionErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "UnivisionExportLog.txt", "frmExportSchdSpot-mCheckSelection"
    ilRet = 1
    Resume Next

End Function

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload frmExportSchdSpot
End Sub

Private Sub txtNumberDays_Change()
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
End Sub

Private Function mAdjCpttRecs(cprst As ADODB.Recordset) As Integer
    Dim slTemp As String

    'D.S. 10/25/04
    mAdjCpttRecs = True
    If DateValue(gAdjYear(Trim$(cprst!attOffAir))) <= DateValue(gAdjYear(Trim$(cprst!attDropDate))) Then
        slTemp = Trim$(cprst!attOffAir)
    Else
        slTemp = Trim$(cprst!attDropDate)
    End If

    ' Markettron delete the aet records to.

'    If DateValue(gAdjYear(cprst!cpttStartDate)) < DateValue(gAdjYear(cprst!attOnAir)) - 1 Then
'        SQLQuery = "DELETE FROM Ast"
'        SQLQuery = SQLQuery & " WHERE astAtfCode = " & cprst!cpttatfCode
'        SQLQuery = SQLQuery & " And astPledgeDate < '" & Format$(DateValue(gAdjYear(cprst!attOnAir)), sgSQLDateForm) & "'"
'        cnn.Execute SQLQuery, rdExecDirect
'
'        SQLQuery = "DELETE FROM AET"
'        SQLQuery = SQLQuery & " WHERE aetAtfCode = " & cprst!cpttatfCode
'        SQLQuery = SQLQuery & " And aetPledgeStartDate < '" & Format$(DateValue(gAdjYear(cprst!attOnAir)), sgSQLDateForm) & "'"
'        cnn.Execute SQLQuery, rdExecDirect
'    End If
'    If DateValue(gAdjYear(cprst!cpttStartDate)) > DateValue(gAdjYear(slTemp)) + 1 Then
'        SQLQuery = "DELETE FROM Ast"
'        SQLQuery = SQLQuery & " WHERE astAtfCode = " & cprst!cpttatfCode
'        SQLQuery = SQLQuery & " And astPledgeDate > '" & Format$(DateValue(gAdjYear(slTemp)), sgSQLDateForm) & "'"
'        cnn.Execute SQLQuery, rdExecDirect
'
'        SQLQuery = "DELETE FROM AET"
'        SQLQuery = SQLQuery & " WHERE aetAtfCode = " & cprst!cpttatfCode
'        SQLQuery = SQLQuery & " And aetPledgeStartDate > '" & Format$(DateValue(gAdjYear(slTemp)), sgSQLDateForm) & "'"
'        cnn.Execute SQLQuery, rdExecDirect
'    End If

    ' Delete ast records using Station, Vehicle, Dates
    If DateValue(gAdjYear(cprst!CpttStartDate)) < DateValue(gAdjYear(cprst!attOnAir)) Or DateValue(gAdjYear(cprst!CpttStartDate)) > DateValue(gAdjYear(slTemp)) Then
        SQLQuery = "DELETE FROM Ast"
        SQLQuery = SQLQuery & " WHERE astAtfCode = " & cprst!cpttatfCode
        '12/13/13: Replace Pledge with Feed
        'SQLQuery = SQLQuery & " And astPledgeDate >= '" & Format$(DateValue(gAdjYear(cprst!CpttStartDate)), sgSQLDateForm) & "'"
        'SQLQuery = SQLQuery & " And astPledgeDate <= '" & Format$(DateValue(gAdjYear(cprst!CpttStartDate)) + 6, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " And astFeedDate >= '" & Format$(DateValue(gAdjYear(cprst!CpttStartDate)), sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " And astFeedDate <= '" & Format$(DateValue(gAdjYear(cprst!CpttStartDate)) + 6, sgSQLDateForm) & "'"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "UnivisionExportLog.txt", "frmExportSchdSpot-mAdjCpttRecs"
            mAdjCpttRecs = False
            Exit Function
        End If
        
        SQLQuery = "DELETE FROM AET"
        SQLQuery = SQLQuery & " WHERE aetAtfCode = " & cprst!cpttatfCode
        '12/13/13: Change to use Feed
        'SQLQuery = SQLQuery & " And aetPledgeStartDate >= '" & Format$(DateValue(gAdjYear(cprst!CpttStartDate)), sgSQLDateForm) & "'"
        'SQLQuery = SQLQuery & " And aetPledgeStartDate <= '" & Format$(DateValue(gAdjYear(cprst!CpttStartDate)) + 6, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " And aetFeedDate >= '" & Format$(DateValue(gAdjYear(cprst!CpttStartDate)), sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " And aetFeedDate <= '" & Format$(DateValue(gAdjYear(cprst!CpttStartDate)) + 6, sgSQLDateForm) & "'"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "UnivisionExportLog.txt", "frmExportSchdSpot-mAdjCpttRecs"
            mAdjCpttRecs = False
            Exit Function
        End If

    End If


    'D.S. 10/25/04 Determine if this is a valid week date for this agreement.
    'If not then delete the cptt record.
    ' If the start date is < On Air date then
    If DateValue(gAdjYear(cprst!CpttStartDate)) < DateValue(gAdjYear(cprst!attOnAir)) Or DateValue(gAdjYear(cprst!CpttStartDate)) > DateValue(gAdjYear(slTemp)) Then
        SQLQuery = "DELETE FROM Cptt"
        SQLQuery = SQLQuery & " WHERE CpttCode = " & cprst!cpttCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "UnivisionExportLog.txt", "frmExportSchdSpot-mAdjCpttRecs"
            mAdjCpttRecs = False
            Exit Function
        End If
    End If

Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "UnivisionExportLog.txt", "frmExportSchdSpot-mAdjCpttRecs"
    Exit Function
End Function
Private Sub mSaveCustomValues()
    Dim ilLoop As Integer
    ReDim ilVefCode(0 To 0) As Integer
    ReDim ilShttCode(0 To 0) As Integer
    If igExportSource <> 2 Then
        ReDim tgEhtInfo(0 To 1) As EHTINFO
        ReDim tgEvtInfo(0 To 0) As EVTINFO
        ReDim tgEctInfo(0 To 0) As ECTINFO
        lgExportEhtInfoIndex = 0
        tgEhtInfo(lgExportEhtInfoIndex).lFirstEct = -1
        For ilLoop = 0 To lbcVehicles.ListCount - 1
            If lbcVehicles.Selected(ilLoop) Then
                ilVefCode(UBound(ilVefCode)) = lbcVehicles.ItemData(ilLoop)
                ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
            End If
        Next ilLoop
        For ilLoop = 0 To lbcStation.ListCount - 1
            If lbcStation.Selected(ilLoop) Then
                ilShttCode(UBound(ilShttCode)) = lbcStation.ItemData(ilLoop)
                ReDim Preserve ilShttCode(0 To UBound(ilShttCode) + 1) As Integer
            End If
        Next ilLoop
        udcCriteria.Action 5
        lmEqtCode = gCustomStartStatus("A", "Univision", "2", Trim$(edcDate.Text), Trim$(txtNumberDays.Text), ilVefCode(), ilShttCode())
    End If
End Sub

