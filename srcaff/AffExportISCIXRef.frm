VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmExportISCIXRef 
   Caption         =   "Export ISCI Cross Reference"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "AffExportISCIXRef.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   9615
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8400
      Top             =   4410
   End
   Begin V81Affiliate.CSI_Calendar edcDate 
      Height          =   285
      Left            =   1500
      TabIndex        =   1
      Top             =   165
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      Text            =   "8/4/2022"
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
      Height          =   240
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Results"
      Top             =   1110
      Width           =   960
   End
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   1560
      Width           =   3825
   End
   Begin VB.TextBox txtNumberDays 
      Height          =   285
      Left            =   3915
      TabIndex        =   3
      Text            =   "1"
      Top             =   165
      Width           =   405
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   2985
      ItemData        =   "AffExportISCIXRef.frx":08CA
      Left            =   4800
      List            =   "AffExportISCIXRef.frx":08CC
      TabIndex        =   10
      Top             =   1395
      Width           =   4605
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   2595
      ItemData        =   "AffExportISCIXRef.frx":08CE
      Left            =   120
      List            =   "AffExportISCIXRef.frx":08D0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1800
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
      FormDesignHeight=   5610
      FormDesignWidth =   9615
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   5910
      TabIndex        =   7
      Top             =   5070
      Width           =   1665
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7755
      TabIndex        =   8
      Top             =   5070
      Width           =   1665
   End
   Begin V81Affiliate.AffExportCriteria udcCriteria 
      Height          =   810
      Left            =   120
      TabIndex        =   4
      Top             =   690
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1429
   End
   Begin VB.Label Label2 
      Caption         =   "# of Days"
      Height          =   255
      Left            =   3030
      TabIndex        =   2
      Top             =   210
      Width           =   795
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   120
      TabIndex        =   11
      Top             =   5010
      Width           =   5580
   End
   Begin VB.Label Label1 
      Caption         =   "Export Start Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   1395
   End
End
Attribute VB_Name = "frmExportISCIXRef"
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

Private smStartDate As String     'Export Date
Private smEndDate As String
Private imNumberDays As Integer
Private imVefCode As Integer
Private imAdfCode As Integer
Private smVefName As String
Private smXDXMLForm As String
Private smISCIPrefix As String
Private imAllClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
'Private hmMsg As Integer
Private smExportDirectory As String
Private hmTo As Integer
Private hmFrom As Integer
Private lstrst As ADODB.Recordset
Private rsfrst As ADODB.Recordset
Private rafrst As ADODB.Recordset
Private cpfrst As ADODB.Recordset
Private smMessage As String
Private smWarnFlag As Integer
Private lmSdfCode() As Long
Private tmXFerSplitSdfCode() As XFRESPLITSDFCODE
Private lmEqtCode As Long
Private smAddAdvtToISCI As String
'7459 0=none 1=isci 2=break 3=both
Private imISCIPrefix As Integer





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

Private Sub mFillVehicle()
    Dim iLoop As Integer
    Dim ilRet As Integer
    
    ilRet = gPopVff()
    lbcVehicles.Clear
    lbcMsg.Clear
    chkAll.Value = 0
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
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





Private Sub cmdExport_Click()
    Dim sNowDate As String
    Dim ilRet As Integer
    Dim slExportType As String
    Dim ilVef As Integer
    Dim llIndex As Long
    Dim slToFile As String
    Dim slDateTime As String
    Dim ilVff As Integer

    On Error GoTo ErrHand
    
    lbcMsg.Clear
    If lbcVehicles.ListIndex < 0 Then
        igExportReturn = 2
        Exit Sub
    End If
    If edcDate.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        edcDate.SetFocus
        Exit Sub
    End If
    If gIsDate(edcDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        edcDate.SetFocus
        Exit Sub
    Else
        smStartDate = Format(edcDate.Text, sgShowDateForm)
    End If
    imNumberDays = Val(txtNumberDays.Text)
    If imNumberDays <= 0 Then
        gMsgBox "Number of days must be specified.", vbOKOnly
        txtNumberDays.SetFocus
        Exit Sub
    End If
    Select Case Weekday(gAdjYear(smStartDate))
        Case vbMonday
            If imNumberDays > 7 Then
                gMsgBox "Number of days can not exceed 7.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbTuesday
            If imNumberDays > 6 Then
                gMsgBox "Number of days can not exceed 6.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbWednesday
            If imNumberDays > 5 Then
                gMsgBox "Number of days can not exceed 5.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbThursday
            If imNumberDays > 4 Then
                gMsgBox "Number of days can not exceed 4.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbFriday
            If imNumberDays > 3 Then
                gMsgBox "Number of days can not exceed 3.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbSaturday
            If imNumberDays > 2 Then
                gMsgBox "Number of days can not exceed 2.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
        Case vbSunday
            If imNumberDays > 1 Then
                gMsgBox "Number of days can not exceed 1.", vbOKOnly
                txtNumberDays.SetFocus
                Exit Sub
            End If
    End Select
    smEndDate = DateAdd("d", imNumberDays - 1, smStartDate)
    sNowDate = Format$(gNow(), "m/d/yy")
    If DateValue(gAdjYear(smStartDate)) <= DateValue(gAdjYear(sNowDate)) Then
        Beep
        gMsgBox "Date must be after today's date " & sNowDate, vbCritical
        edcDate.SetFocus
        Exit Sub
    End If
    smExportDirectory = udcCriteria.RExportToPath
    If smExportDirectory = "" Then
        smExportDirectory = sgExportDirectory
    End If
    smExportDirectory = gSetPathEndSlash(smExportDirectory, False)
    
    
    Screen.MousePointer = vbHourglass
    smWarnFlag = False
    imExporting = True
    mSaveCustomValues
    If Not gPopCopy(smStartDate, "Export ISCI Cross Reference") Then
        igExportReturn = 2
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        imExporting = False
        Exit Sub
    End If
    lacResult.Caption = ""
    slExportType = "!! Exporting ISCI Cross Reference, "
    gLogMsg slExportType & "Start Date of: " & smStartDate & " For: " & CStr(imNumberDays) & " Days.", "ISCIXRefExportLog.Txt", False
    For ilVef = 0 To lbcVehicles.ListCount - 1
        If igExportSource = 2 Then DoEvents
        If lbcVehicles.Selected(ilVef) Then
            imVefCode = lbcVehicles.ItemData(ilVef)
            llIndex = gBinarySearchVef(CLng(imVefCode))
            If llIndex <> -1 Then
                If imNumberDays > 1 Then
                    slToFile = smExportDirectory & Trim$(tgVehicleInfo(llIndex).sCodeStn) & Format(smStartDate, "YYYYMMDD") & "-" & Format(smEndDate, "YYYYMMDD") & ".xis"
                Else
                    slToFile = smExportDirectory & Trim$(tgVehicleInfo(llIndex).sCodeStn) & Format(smStartDate, "YYYYMMDD") & ".xis"
                End If
                ilRet = 0
                On Error GoTo mExportErr:
                'slDateTime = FileDateTime(slToFile)
                ilRet = gFileExist(slToFile)
                If ilRet = 0 Then
                    Kill slToFile
                End If
                'ilRet = 0
                'hmTo = FreeFile
                'Open slToFile For Output As hmTo
                ilRet = gFileOpen(slToFile, "Output", hmTo)
                If ilRet <> 0 Then
                    Close #hmTo
                    ilRet = gCustomEndStatus(lmEqtCode, 2, "")
                    imExporting = False
                    Screen.MousePointer = vbDefault
                    gMsgBox "Open Error #" & Str$(Err.Numner) & slToFile, vbOKOnly, "Open Error"
                    Exit Sub
                End If
                On Error GoTo ErrHand
                ilRet = mGatherLST(llIndex)
                If imTerminate Then
                    Close #hmTo
                    gLogMsg "** User Terminated **", "ISCIXRefExportLog.Txt", False
                    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
                    imExporting = False
                    Screen.MousePointer = vbDefault
                    cmdCancel.SetFocus
                    Exit Sub
                End If
                If Not ilRet Then
                    Close #hmTo
                    gLogMsg "** Error, ISCI Export Stopped **", "ISCIXRefExportLog.Txt", False
                    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
                    imExporting = False
                    Screen.MousePointer = vbDefault
                    cmdCancel.SetFocus
                    Exit Sub
                End If
                smISCIPrefix = ""
                smXDXMLForm = "P"
                ilVff = gBinarySearchVff(imVefCode)
                If ilVff <> -1 Then
                    '7459
'                    '7456  -1 = error 0=none 1=isci 2=break 3=both
'                    ilRet = mSiteAllowXDS()
'                    If ilRet = 1 Then
'                        smISCIPrefix = Trim$(tgVffInfo(ilVff).sXDSISCIPrefix)
'                    ElseIf ilRet < 3 Then
'                        smISCIPrefix = Trim$(tgVffInfo(ilVff).sXDISCIPrefix)
'                    'both: use radio buttons 7459
'                    Else
'                        smISCIPrefix = Trim$(tgVffInfo(ilVff).sXDISCIPrefix)
'                    End If
'                    smXDXMLForm = Trim$(tgVffInfo(ilVff).sXDXMLForm)
                    Select Case imISCIPrefix
                        Case 0
                            smISCIPrefix = ""
                        Case 1
                            smISCIPrefix = Trim$(tgVffInfo(ilVff).sXDSISCIPrefix)
                        Case 2
                            smISCIPrefix = Trim$(tgVffInfo(ilVff).sXDISCIPrefix)
                        Case 3
                            If udcCriteria.RPrefix(1) Then
                                smISCIPrefix = Trim$(tgVffInfo(ilVff).sXDSISCIPrefix)
                            Else
                                smISCIPrefix = Trim$(tgVffInfo(ilVff).sXDISCIPrefix)
                            End If
                    End Select
                    smXDXMLForm = Trim$(tgVffInfo(ilVff).sXDXMLForm)
                End If
                ilRet = mGenerateXRef()
                If imTerminate Then
                    Close #hmTo
                    gLogMsg "** User Terminated **", "ISCIXRefExportLog.Txt", False
                    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
                    imExporting = False
                    Screen.MousePointer = vbDefault
                    cmdCancel.SetFocus
                    Exit Sub
                End If
                Close #hmTo
            End If
        End If
    Next ilVef
    
    ilRet = gCustomEndStatus(lmEqtCode, 1, "")
    imExporting = False
    'Print #hmMsg, "** Completed Export of StarGuide: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    gLogMsg "** Completed ISCI Export **", "ISCIXRefExportLog.Txt", False
    'Close #hmMsg
    lacResult.Caption = "Exports placed into: " & smExportDirectory
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    gLogMsg "", "ISCIXRefExportLog.Txt", False
    Exit Sub
mExportErr:
    ilRet = 1
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "ISCIXRefExportLog.txt", "Export ISCI XRef-cmdExport_Click"
    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    edcDate.Text = ""
    Unload frmExportISCIXRef
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
        udcCriteria.Top = txtNumberDays.Top + txtNumberDays.Height
        udcCriteria.Action 6
        If UBound(tgEvtInfo) > 0 Then
            chkAll.Value = vbUnchecked
            lbcVehicles.Clear
            For ilLoop = 0 To UBound(tgEvtInfo) - 1 Step 1
                llVef = gBinarySearchVef(CLng(tgEvtInfo(ilLoop).iVefCode))
                If llVef <> -1 Then
                    lbcVehicles.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
                    lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgEvtInfo(ilLoop).iVefCode
                End If
            Next ilLoop
            chkAll.Value = vbChecked
        End If
        If igExportSource = 2 Then
            slNowStart = gNow()
            edcDate.Text = sgExporStartDate
            txtNumberDays.Text = igExportDays
            igExportReturn = 1
            '6394 move before 'click'
            sgExportResultName = "ISCIXrefResultList.Txt"
            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
            gLogMsgWODT "W", hlResult, "ISCI Xref Result List, Started: " & slNowStart
            hgExportResult = hlResult
            cmdExport_Click
            slNowEnd = gNow()
            'Output result list box
'            sgExportResultName = "ISCIXrefResultList.Txt"
'            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
'            gLogMsgWODT "W", hlResult, "ISCI Xref Result List, Started: " & slNowStart
            If lbcMsg.ListCount > 0 Then
                For ilLoop = 0 To lbcMsg.ListCount - 1 Step 1
                    gLogMsgWODT "W", hlResult, Trim$(lbcMsg.List(ilLoop))
                Next ilLoop
            End If
            gLogMsgWODT "W", hlResult, "ISCI Xref Result List, Completed: " & slNowEnd
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
    Me.Height = Screen.Height / 1.2 '1.6
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
    Dim ilValue10 As Integer
    Dim ilValue7 As Integer
    Dim ilValue8 As Integer
    
    Screen.MousePointer = vbHourglass
    frmExportISCIXRef.Caption = "Export ISCI Cross Reference - " & sgClientName
    ilRet = 0
    smStartDate = gObtainNextMonday(Format$(gNow(), sgShowDateForm))
    edcDate.Text = smStartDate
    txtNumberDays.Text = 1
    imAllClick = False
    imTerminate = False
    imExporting = False
    imFirstTime = True
    
    mFillVehicle
    
    smAddAdvtToISCI = "N"
    'SQLQuery = "Select spfBSlspBack From SPF_Site_Options"
    'Set rst = gSQLSelectCall(SQLQuery)
    'If Not rst.EOF Then
    '    smAddAdvtToISCI = Trim$(rst!spfBSlspBack)
    'End If
    '7459
   ' SQLQuery = "Select spfUsingFeatures10 From SPF_Site_Options"
    SQLQuery = "Select spfUsingFeatures7,spfUsingFeatures8,spfUsingFeatures10 From SPF_Site_Options"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        ilValue10 = Asc(rst!spfUsingFeatures10)
        If (ilValue10 And ADDADVTTOISCI) = ADDADVTTOISCI Then
            smAddAdvtToISCI = "Y"
        End If
        ilValue7 = Asc(rst!spfUsingFeatures7)
        If (ilValue7 And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT Then
            ilRet = 1
        End If
        ilValue8 = Asc(rst!spfUsingFeatures8)
        If (ilValue8 And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT Then
            ilRet = ilRet + 2
        End If
    End If
    imISCIPrefix = ilRet
    If imISCIPrefix = 3 Then
        udcCriteria.RPrefixVisible = True
    Else
        udcCriteria.RPrefixVisible = False
    End If
    ilRet = gPopAvailNames()
    'txtFile.Text = sgExportDirectory & "MktSpots.txt"
    chkAll.Value = vbChecked
    ilRet = gPopAvailNames()
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
    lstrst.Close
    rsfrst.Close
    rafrst.Close
    cpfrst.Close

    Erase lmSdfCode
    Erase tmXFerSplitSdfCode
    Set frmExportISCIXRef = Nothing
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
End Sub

Private Sub edcDate_Change()
    lbcMsg.Clear
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
End Sub

Private Function mGatherLST(llVef As Long) As Integer
    Dim llLoop As Long
    Dim llTest As Long
    Dim ilFound As Integer

    On Error GoTo ErrHand
    ReDim lmSdfCode(0 To 0) As Long
    If tgVehicleInfo(llVef).sVehType = "L" Then
        For llLoop = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
            If igExportSource = 2 Then DoEvents
            If tgVehicleInfo(llLoop).iVefCode = imVefCode Then
                SQLQuery = "SELECT DISTINCT lstSdfCode FROM lst WHERE (lstLogVefCode = " & tgVehicleInfo(llLoop).iCode
                SQLQuery = SQLQuery + " AND Mod(lstStatus, 100) < " & ASTEXTENDED_MG 'Bypass MG/Bonus
                SQLQuery = SQLQuery & " AND lstLogDate >= '" & Format$(smStartDate, sgSQLDateForm) & "'"
                SQLQuery = SQLQuery & " AND lstLogDate <= '" & Format$(smEndDate, sgSQLDateForm) & "')"
                Set lstrst = gSQLSelectCall(SQLQuery)
                While Not lstrst.EOF
                    If igExportSource = 2 Then DoEvents
                    If imTerminate Then
                        mGatherLST = False
                        Exit Function
                    End If
                    ilFound = False
                    If tgVehicleInfo(llLoop).sVehType = "A" Then
                        ilFound = False
                        For llTest = 0 To UBound(lmSdfCode) - 1 Step 1
                            If lmSdfCode(llTest) = lstrst!lstSdfCode Then
                                ilFound = True
                                Exit For
                            End If
                        Next llTest
                    End If
                    If Not ilFound Then
                        lmSdfCode(UBound(lmSdfCode)) = lstrst!lstSdfCode
                        ReDim Preserve lmSdfCode(0 To UBound(lmSdfCode) + 1) As Long
                    End If
                    lstrst.MoveNext
                Wend
            End If
        Next llLoop
    Else
        SQLQuery = "SELECT DISTINCT lstSdfCode FROM lst WHERE (lstLogVefCode = " & imVefCode
        SQLQuery = SQLQuery + " AND Mod(lstStatus, 100) < " & ASTEXTENDED_MG 'Bypass MG/Bonus
        SQLQuery = SQLQuery & " AND lstLogDate >= '" & Format$(smStartDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND lstLogDate <= '" & Format$(smEndDate, sgSQLDateForm) & "')"
        Set lstrst = gSQLSelectCall(SQLQuery)
        While Not lstrst.EOF
            If igExportSource = 2 Then DoEvents
            If imTerminate Then
                mGatherLST = False
                Exit Function
            End If
            lmSdfCode(UBound(lmSdfCode)) = lstrst!lstSdfCode
            ReDim Preserve lmSdfCode(0 To UBound(lmSdfCode) + 1) As Long
            lstrst.MoveNext
        Wend
    End If
    mGatherLST = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "ISCIXRefExportLog.txt", "Export ISCI XRef-mGatherLst"
    mGatherLST = False
    Exit Function
End Function

Private Function mGenerateXRef() As Integer
    'Dan M 10/24/14 changed mFileNameFilter to gFileNameFilter
    Dim llSection As Long
    Dim llTest As Long
    Dim ilFound As Integer
    Dim llAdf As Long
    Dim ilRet As Integer
    Dim llSpot As Long
    Dim slCreativeTitle As String
    Dim llTime As Long
    Dim llDate As Long

    Dim slRCartNo As String
    Dim slRProduct As String
    Dim slRISCI As String
    Dim slRCreativeTitle As String
    Dim llRCrfCsfCode As Long
    Dim llRCpfCode As Long
    Dim ilCifAdfCode As Integer
    Dim llRegion As Long
    Dim slShortTitle As String
    Dim llSplits As Long
    Dim ilRegionSpotFound As Integer
    Dim blSpotOk As Boolean
    Dim ilAnf As Integer

    On Error GoTo ErrHand
    ReDim tmXFerSplitSdfCode(0 To 0) As XFRESPLITSDFCODE
    llSection = 0
    For llSpot = 0 To UBound(lmSdfCode) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        SQLQuery = "SELECT * FROM lst WHERE (lstSdfCode = " & lmSdfCode(llSpot) & ")"
        Set lstrst = gSQLSelectCall(SQLQuery)
        If Not lstrst.EOF Then
            'Output Generic Copy
            'llSection = llSection + 1
            'Print #hmTo, "[Copy" & llSection & "]"
            'Print #hmTo, "GenericCartNumber=" & Trim$(lstrst!lstCart)
            'Print #hmTo, "GenericIsci=" & mFileNameFilter(Trim$(smISCIPrefix & lstrst!lstISCI))
            'Creative title
            blSpotOk = True
            ilAnf = gBinarySearchAnf(lstrst!lstAnfCode)
            If ilAnf <> -1 Then
                If tgAvailNamesInfo(ilAnf).sISCIExport = "N" Then
                    blSpotOk = False
                End If
            End If
            If igExportSource = 2 Then DoEvents
            If Not blSpotOk Then
                lmSdfCode(llSpot) = -lmSdfCode(llSpot)
            Else
                slCreativeTitle = ""
                If lstrst!lstCpfCode > 0 Then
                    SQLQuery = "Select * From CPF_Copy_Prodct_ISCI"
                    SQLQuery = SQLQuery & " Where (cpfCode = " & lstrst!lstCpfCode & ")"
                    Set cpfrst = gSQLSelectCall(SQLQuery)
                    If Not cpfrst.EOF Then
                        'slCreativeTitle = Trim$(cpf_rst!cpfCreative)
                        If IsNull(cpfrst!cpfCreative) = False Then
                            If Asc(cpfrst!cpfCreative) <> 0 Then
                                slCreativeTitle = Trim$(cpfrst!cpfCreative)
                            End If
                        End If
                    End If
                End If
                'If slCreativeTitle <> "" Then
                '    Print #hmTo, "GenericCreativeTitle=" & slCreativeTitle
                'Else
                '    llAdf = gBinarySearchAdf(CLng(lstrst!lstAdfCode))
                '    If llAdf <> -1 Then
                '        Print #hmTo, "GenericCreativeTitle=" & Trim$(tgAdvtInfo(llAdf).sAdvtName)
                '    End If
                'End If
                'Output Region Copy
                '4/10/13: Bypass Airing copy as it is assigned to lst
                llRegion = 0
                'SQLQuery = "SELECT * FROM rsf_Region_Schd_Copy WHERE (rsfsdfCode = " & lmSdfCode(llSpot) & ")"
                SQLQuery = "SELECT * FROM rsf_Region_Schd_Copy "
                SQLQuery = SQLQuery & " Where (rsfSdfCode = " & lmSdfCode(llSpot)
                SQLQuery = SQLQuery & " AND rsfType <> 'B'"     'Blackout
                SQLQuery = SQLQuery & " AND rsfType <> 'A'" & ")"     'Airing vehicle copy
                Set rsfrst = gSQLSelectCall(SQLQuery)
                If (udcCriteria.RIncludeGeneric = vbChecked) Or (Not rsfrst.EOF) Then
                    llSection = llSection + 1
                    Print #hmTo, "[Copy" & llSection & "]"
                    Print #hmTo, "GenericCartNumber=" & Trim$(lstrst!lstCart)
                    If smAddAdvtToISCI = "Y" Then
                        '2/7/13: Replace short title with Advertiser abbreviation
                        'slShortTitle = Trim$(gGetShortTitle(lmSdfCode(llSpot)))
                        'slShortTitle = Left$(UCase$(mFileNameFilter(slShortTitle)), 15)
                        'Print #hmTo, "GenericIsci=" & slShortTitle & "(" & UCase$(mFileNameFilter(Trim$(smISCIPrefix & lstrst!lstISCI))) & ")" & ".mp2"
                        'see 7219 below
'                        llAdf = gBinarySearchAdf(CLng(lstrst!lstAdfCode))
'                        If llAdf <> -1 Then
'                            slShortTitle = Trim$(Left$(UCase$(mFileNameFilter(Trim$(tgAdvtInfo(llAdf).sAdvtAbbr))), 6))
'                            If slShortTitle = "" Then
'                                slShortTitle = Trim$(Left$(UCase$(mFileNameFilter(Trim$(tgAdvtInfo(llAdf).sAdvtName))), 6))
'                            End If
'                            Print #hmTo, "GenericIsci=" & slShortTitle & "(" & UCase$(mFileNameFilter(Trim$(smISCIPrefix & lstrst!lstISCI))) & ")" & ".mp2"
'                        Else
'                            Print #hmTo, "GenericIsci=" & mFileNameFilter(Trim$(smISCIPrefix & lstrst!lstISCI))
'                        End If
'                        '7219
                        slShortTitle = gXDSShortTitle(CLng(lstrst!lstAdfCode), "", False, False)
                        ' couldn't find advertiser
                        If slShortTitle = "" Then
                            Print #hmTo, "GenericIsci=" & gFileNameFilter(Trim$(smISCIPrefix & lstrst!lstISCI))
                        Else
                            '7496
                            'Print #hmTo, "GenericIsci=" & slShortTitle & "(" & UCase$(gFileNameFilter(Trim$(smISCIPrefix & lstrst!lstISCI))) & ")" & ".mp2"
                            Print #hmTo, "GenericIsci=" & slShortTitle & "(" & UCase$(gFileNameFilter(Trim$(smISCIPrefix & lstrst!lstISCI))) & ")" & sgAudioExtension
                        End If
                    Else
                        Print #hmTo, "GenericIsci=" & gFileNameFilter(Trim$(smISCIPrefix & lstrst!lstISCI))
                    End If
                    If slCreativeTitle <> "" Then
                        Print #hmTo, "GenericCreativeTitle=" & slCreativeTitle
                    Else
                        llAdf = gBinarySearchAdf(CLng(lstrst!lstAdfCode))
                        If llAdf <> -1 Then
                            Print #hmTo, "GenericCreativeTitle=" & Trim$(tgAdvtInfo(llAdf).sAdvtName)
                        End If
                    End If
                End If
                ilRegionSpotFound = False
                While Not rsfrst.EOF
                    ilRegionSpotFound = True
                    If igExportSource = 2 Then DoEvents
                    If imTerminate Then
                        mGenerateXRef = False
                        Exit Function
                    End If
                    ilRet = gGetCopy(rsfrst!rstPtType, rsfrst!rsfCopyCode, rsfrst!rsfCrfCode, False, slRCartNo, slRProduct, slRISCI, slRCreativeTitle, llRCrfCsfCode, llRCpfCode, ilCifAdfCode)
                    llRegion = llRegion + 1
                    Print #hmTo, "RegionIsci" & llRegion & "=" & gFileNameFilter(Trim$(smISCIPrefix & slRISCI))
                    'Region Name
                    SQLQuery = "SELECT rafName, rafAdfCode FROM raf_Region_Area WHERE (rafCode = " & rsfrst!rsfRafCode & ")"
                    Set rafrst = gSQLSelectCall(SQLQuery)
                    If Not rafrst.EOF Then
                        Print #hmTo, "RegionName" & llRegion & "=" & Trim$(rafrst!rafName)
                    Else
                        Print #hmTo, "RegionName" & llRegion & "=" & Trim$(slRCreativeTitle)
                    End If
                    'ShortTitle(isci).mp2 or isci.mp2
                    If smXDXMLForm <> "P" Then
                        If smAddAdvtToISCI = "Y" Then
                            '2/7/13: Replace shorttitle with Advertiser abbreviation
                            'slShortTitle = Trim$(gGetShortTitle(lmSdfCode(llSpot)))
                            'slShortTitle = Left$(UCase$(mFileNameFilter(slShortTitle)), 15)
                            'Print #hmTo, "RegionFilename" & llRegion & "=" & slShortTitle & "(" & mFileNameFilter(Trim$(smISCIPrefix & slRISCI)) & ")" & ".mp2"
'                            'see 7219 below
'                            llAdf = gBinarySearchAdf(CLng(rafrst!rafAdfCode))
'                            If llAdf <> -1 Then
'                                slShortTitle = Trim$(Left$(UCase$(mFileNameFilter(Trim$(tgAdvtInfo(llAdf).sAdvtAbbr))), 6))
'                                If slShortTitle = "" Then
'                                    slShortTitle = Trim$(Left$(UCase$(mFileNameFilter(Trim$(tgAdvtInfo(llAdf).sAdvtName))), 6))
'                                End If
'                                Print #hmTo, "RegionFilename" & llRegion & "=" & slShortTitle & "(" & mFileNameFilter(Trim$(smISCIPrefix & slRISCI)) & ")" & ".mp2"
'                            Else
'                                Print #hmTo, "RegionFilename" & llRegion & "=" & slShortTitle & "(" & mFileNameFilter(Trim$(smISCIPrefix & slRISCI)) & ")" & ".mp2"
'                            End If
                            '7219
                            slShortTitle = gXDSShortTitle(CLng(rafrst!rafAdfCode), "", False, False)
                            ' couldn't find advertiser
                            If slShortTitle = "" Then
                                '7496
                               ' Print #hmTo, "RegionFilename" & llRegion & "=" & slShortTitle & "(" & gFileNameFilter(Trim$(smISCIPrefix & slRISCI)) & ")" & ".mp2"
                                Print #hmTo, "RegionFilename" & llRegion & "=" & slShortTitle & "(" & gFileNameFilter(Trim$(smISCIPrefix & slRISCI)) & ")" & sgAudioExtension
                            Else
                                '7496
                               ' Print #hmTo, "RegionFilename" & llRegion & "=" & slShortTitle & "(" & gFileNameFilter(Trim$(smISCIPrefix & slRISCI)) & ")" & ".mp2"
                                Print #hmTo, "RegionFilename" & llRegion & "=" & slShortTitle & "(" & gFileNameFilter(Trim$(smISCIPrefix & slRISCI)) & ")" & sgAudioExtension
                            End If
                        Else
                            '7496
                            'Print #hmTo, "RegionFilename" & llRegion & "=" & gFileNameFilter(Trim$(smISCIPrefix & slRISCI)) & ".mp2"
                            Print #hmTo, "RegionFilename" & llRegion & "=" & gFileNameFilter(Trim$(smISCIPrefix & slRISCI)) & sgAudioExtension
                        End If
                    Else
                        slShortTitle = Trim$(gGetShortTitle(lmSdfCode(llSpot), 6))
                        slShortTitle = UCase$(gFileNameFilter(slShortTitle))
                        '7496
                        'Print #hmTo, "RegionFilename" & llRegion & "=" & slShortTitle & "(" & gFileNameFilter(Trim$(smISCIPrefix & slRISCI)) & ")" & ".mp2"
                        Print #hmTo, "RegionFilename" & llRegion & "=" & slShortTitle & "(" & gFileNameFilter(Trim$(smISCIPrefix & slRISCI)) & ")" & sgAudioExtension
                    End If
                    rsfrst.MoveNext
                Wend
                If ilRegionSpotFound Then
                    If (udcCriteria.RIncludeGeneric = vbUnchecked) And (smXDXMLForm = "S") Then
                        tmXFerSplitSdfCode(UBound(tmXFerSplitSdfCode)).lSdfCode = lmSdfCode(llSpot)
                        tmXFerSplitSdfCode(UBound(tmXFerSplitSdfCode)).lLogDate = DateValue(gAdjYear(lstrst!lstLogDate))
                        tmXFerSplitSdfCode(UBound(tmXFerSplitSdfCode)).lLogTime = gTimeToLong(lstrst!lstLogTime, False)
                        ReDim Preserve tmXFerSplitSdfCode(0 To UBound(tmXFerSplitSdfCode) + 1) As XFRESPLITSDFCODE
                        lmSdfCode(llSpot) = 0
                    End If
                End If
            End If
        End If
    Next llSpot
    If (udcCriteria.RIncludeGeneric = vbUnchecked) And (smXDXMLForm = "S") Then
        For llSplits = 0 To UBound(tmXFerSplitSdfCode) - 1 Step 1
            If igExportSource = 2 Then DoEvents
            For llSpot = 0 To UBound(lmSdfCode) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If (lmSdfCode(llSpot) > 0) And (tmXFerSplitSdfCode(llSplits).lSdfCode <> lmSdfCode(llSpot)) And (lmSdfCode(llSpot) <> 0) Then
                    SQLQuery = "SELECT * FROM lst WHERE (lstSdfCode = " & lmSdfCode(llSpot) & ")"
                    Set lstrst = gSQLSelectCall(SQLQuery)
                    If Not lstrst.EOF Then
                        llDate = DateValue(gAdjYear(lstrst!lstLogDate))
                        llTime = gTimeToLong(lstrst!lstLogTime, False)
                        If (llDate = tmXFerSplitSdfCode(llSplits).lLogDate) And (llTime = tmXFerSplitSdfCode(llSplits).lLogTime) Then
                            slCreativeTitle = ""
                            If lstrst!lstCpfCode > 0 Then
                                SQLQuery = "Select * From CPF_Copy_Prodct_ISCI"
                                SQLQuery = SQLQuery & " Where (cpfCode = " & lstrst!lstCpfCode & ")"
                                Set cpfrst = gSQLSelectCall(SQLQuery)
                                If Not cpfrst.EOF Then
                                    'slCreativeTitle = Trim$(cpf_rst!cpfCreative)
                                    If IsNull(cpfrst!cpfCreative) = False Then
                                        If Asc(cpfrst!cpfCreative) <> 0 Then
                                            slCreativeTitle = Trim$(cpfrst!cpfCreative)
                                        End If
                                    End If
                                End If
                            End If
                            If igExportSource = 2 Then DoEvents
                            llSection = llSection + 1
                            Print #hmTo, "[Copy" & llSection & "]"
                            Print #hmTo, "GenericCartNumber=" & Trim$(lstrst!lstCart)
                            Print #hmTo, "GenericIsci=" & gFileNameFilter(Trim$(smISCIPrefix & lstrst!lstISCI))
                            If slCreativeTitle <> "" Then
                                Print #hmTo, "GenericCreativeTitle=" & slCreativeTitle
                            Else
                                llAdf = gBinarySearchAdf(CLng(lstrst!lstAdfCode))
                                If llAdf <> -1 Then
                                    Print #hmTo, "GenericCreativeTitle=" & Trim$(tgAdvtInfo(llAdf).sAdvtName)
                                End If
                            End If
                            If igExportSource = 2 Then DoEvents
                        End If
                    End If
                End If
            Next llSpot
        Next llSplits
    End If
    mGenerateXRef = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "ISCIXRefExportLog.txt", "Export ISCI XRef-mGenerateXRef"
    mGenerateXRef = False
    Exit Function
End Function

'Private Function mFileNameFilter(slInName As String) As String
'    'Same as in ExportXDigital
'    Dim slName As String
'    Dim ilPos As Integer
'    Dim ilFound As Integer
'    slName = slInName
'    'Remove " and '
'    Do
'        If igExportSource = 2 Then DoEvents
'        ilFound = False
'        ilPos = InStr(1, slName, "'", 1)
'        If ilPos > 0 Then
'            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
'            ilFound = True
'        End If
'    Loop While ilFound
'    Do
'        If igExportSource = 2 Then DoEvents
'        ilFound = False
'        ilPos = InStr(1, slName, """", 1)
'        If ilPos > 0 Then
'            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
'            ilFound = True
'        End If
'    Loop While ilFound
'    Do
'        If igExportSource = 2 Then DoEvents
'        ilFound = False
'        ilPos = InStr(1, slName, "&", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "/", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "\", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "*", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, ":", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "?", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "%", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        'ilPos = InStr(1, slName, """", 1)
'        'If ilPos > 0 Then
'        '    Mid$(slName, ilPos, 1) = "'"
'        '    ilFound = True
'        'End If
'        ilPos = InStr(1, slName, "=", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "+", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "<", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, ">", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "|", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, ";", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "@", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "[", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "]", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "{", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "}", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, "^", 1)
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "-"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, ".", 1)    'If period, use underscore
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "_"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, ",", 1)    'If comma, use underscore
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "_"
'            ilFound = True
'        End If
'        ilPos = InStr(1, slName, " ", 1)    'If space, use underscore
'        If ilPos > 0 Then
'            Mid$(slName, ilPos, 1) = "_"
'            ilFound = True
'        End If
'    Loop While ilFound
'    mFileNameFilter = slName
'End Function

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload frmExportISCIXRef

End Sub
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
        udcCriteria.Action 5
        lmEqtCode = gCustomStartStatus("R", "ISCI Cross Reference", "R", Trim$(edcDate.Text), Trim$(txtNumberDays.Text), ilVefCode(), ilShttCode())
    End If
End Sub
'Private Function mSiteAllowXDS() As Integer
''O: -1 error 0 none 1 isci only 2 break only 3 both
'    Dim ilValue7 As Integer
'    Dim ilValue8 As Integer
'    Dim ilRet As Integer
'
' On Error GoTo ERRORBOX
'    ilRet = 0
'    SQLQuery = "Select spfUsingFeatures7,spfUsingFeatures8 From SPF_Site_Options"
'    Set rst = gSQLSelectCall(SQLQuery)
'    If Not rst.EOF Then
'        ilValue7 = Asc(rst!spfusingfeatures7)
'        If (ilValue7 And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT Then
'            ilRet = 1
'        End If
'        ilValue8 = Asc(rst!spfUsingFeatures8)
'        If (ilValue8 And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT Then
'            ilRet = ilRet + 2
'        End If
'    End If
'    mSiteAllowXDS = ilRet
'   Exit Function
'ERRORBOX:
'    mSiteAllowXDS = -1
'    gHandleError "ISCIXRefExportLog.txt", "Export ISCI XRef-mSiteAllowXds"
'End Function
