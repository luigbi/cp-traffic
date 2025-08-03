VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmExportStarGuide 
   Caption         =   "Export StarGuide"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "AffExportStarGuide.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   9615
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9360
      Top             =   3825
   End
   Begin V81Affiliate.CSI_Calendar edcDate 
      Height          =   285
      Left            =   1515
      TabIndex        =   1
      Top             =   180
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      Text            =   "11/8/2010"
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
   Begin VB.TextBox edcTitle3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Stations"
      Top             =   1500
      Width           =   1635
   End
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   1500
      Width           =   3825
   End
   Begin VB.TextBox txtNumberDays 
      Height          =   285
      Left            =   4005
      TabIndex        =   3
      Text            =   "1"
      Top             =   180
      Width           =   405
   End
   Begin VB.CheckBox chkAllStation 
      Caption         =   "All"
      Height          =   195
      Left            =   4215
      TabIndex        =   10
      Top             =   3900
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.ListBox lbcStation 
      Height          =   2010
      ItemData        =   "AffExportStarGuide.frx":08CA
      Left            =   4200
      List            =   "AffExportStarGuide.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1770
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3900
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   3375
      ItemData        =   "AffExportStarGuide.frx":08CE
      Left            =   6585
      List            =   "AffExportStarGuide.frx":08D0
      TabIndex        =   14
      Top             =   450
      Width           =   2820
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   2010
      ItemData        =   "AffExportStarGuide.frx":08D2
      Left            =   120
      List            =   "AffExportStarGuide.frx":08D4
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   6
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
      FormDesignHeight=   4785
      FormDesignWidth =   9615
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   5910
      TabIndex        =   11
      Top             =   4290
      Width           =   1665
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7755
      TabIndex        =   12
      Top             =   4290
      Width           =   1665
   End
   Begin V81Affiliate.AffExportCriteria udcCriteria 
      Height          =   810
      Left            =   120
      TabIndex        =   4
      Top             =   570
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1429
   End
   Begin VB.Label Label2 
      Caption         =   "# of Days"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   210
      Width           =   795
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   120
      TabIndex        =   15
      Top             =   4230
      Width           =   5580
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   7035
      TabIndex        =   13
      Top             =   120
      Width           =   1965
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
Attribute VB_Name = "frmExportStarGuide"
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
Private imVefCode As Integer
Private imAdfCode As Integer
Private smVefName As String
Private imAllClick As Integer
Private imAllStationClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
'Private hmMsg As Integer
Private smExportDirectory As String
Private hmTo As Integer
Private hmFrom As Integer
Private hmAst As Integer
Private cprst As ADODB.Recordset
Private smMessage As String
Private smWarnFlag As Integer
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private tmStarGuideAst() As STARGUIDEAST
Private tmStarGuideXRef() As STARGUIDEXREF
Private lmEqtCode As Long




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
    Dim sNowDate As String
    Dim ilRet As Integer
    Dim slExportType As String

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
        smDate = Format(edcDate.Text, sgShowDateForm)
    End If
    imNumberDays = Val(txtNumberDays.Text)
    If imNumberDays <= 0 Then
        gMsgBox "Number of days must be specified.", vbOKOnly
        txtNumberDays.SetFocus
        Exit Sub
    End If
    Select Case Weekday(gAdjYear(smDate))
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
    sNowDate = Format$(gNow(), "m/d/yy")
    If DateValue(gAdjYear(smDate)) <= DateValue(gAdjYear(sNowDate)) Then
        Beep
        gMsgBox "Date must be after today's date " & sNowDate, vbCritical
        edcDate.SetFocus
        Exit Sub
    End If
    If (udcCriteria.SSpots(0) = False) And (udcCriteria.SSpots(1) = False) Then
        Beep
        gMsgBox "Please Specify Export Spots Type.", vbCritical
        Exit Sub
    End If
    smExportDirectory = udcCriteria.SExportToPath
    If smExportDirectory = "" Then
        smExportDirectory = sgExportDirectory
    End If
    smExportDirectory = gSetPathEndSlash(smExportDirectory, False)
    Screen.MousePointer = vbHourglass
    smWarnFlag = False
    imExporting = True
    mSaveCustomValues
    On Error GoTo 0
    lacResult.Caption = ""
    If udcCriteria.SSpots(0) = True Then
        slExportType = "!! Exporting All Spots, "
    Else
        slExportType = "!! Exporting Regional Spots, "
    End If
    gLogMsg slExportType & "Start Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days.", "StarGuideExportLog.Txt", False
    
    bgTaskBlocked = False
    sgTaskBlockedName = "Star Guide Export"
    ilRet = mExportSpots()
    gCloseRegionSQLRst
    
    If (ilRet = False) Then
        bgTaskBlocked = False
        sgTaskBlockedName = ""
        gLogMsg "** Terminated - mExportSpots returned False **", "StarGuideExportLog.Txt", False
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        imExporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    If imTerminate Then
        bgTaskBlocked = False
        sgTaskBlockedName = ""
        gLogMsg "** User Terminated **", "StarGuideExportLog.Txt", False
        ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
        imExporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    On Error GoTo ErrHand:
    ilRet = gCustomEndStatus(lmEqtCode, 1, "")
    If bgTaskBlocked And igExportSource <> 2 Then
         gMsgBox "Some spots were blocked during the Export generation." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
    End If
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    imExporting = False
    'Print #hmMsg, "** Completed Export of StarGuide: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    gLogMsg "** Completed Export of StarGuide **", "StarGuideExportLog.Txt", False
    'Close #hmMsg
    lacResult.Caption = "Exports placed into: " & smExportDirectory
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    gLogMsg "", "StarGuideExportLog.Txt", False
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "StarGuideExportLog.txt", "Export StarGuide-cmdExport_Click"
    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    edcDate.Text = ""
    Unload frmExportStarGuide
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
            sgExportResultName = "StarGuideResultList.Txt"
            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
            gLogMsgWODT "W", hlResult, "StarGuide Result List, Started: " & slNowStart
            hgExportResult = hlResult
            cmdExport_Click
            slNowEnd = gNow()
            'Output result list box
'            sgExportResultName = "StarGuideResultList.Txt"
'            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
'            gLogMsgWODT "W", hlResult, "StarGuide Result List, Started: " & slNowStart
            If lbcMsg.ListCount > 0 Then
                For ilLoop = 0 To lbcMsg.ListCount - 1 Step 1
                    gLogMsgWODT "W", hlResult, Trim$(lbcMsg.List(ilLoop))
                Next ilLoop
            End If
            gLogMsgWODT "W", hlResult, "StarGuide Result List, Completed: " & slNowEnd
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
    frmExportStarGuide.Caption = "Export StarGuide - " & sgClientName
    smDate = gObtainNextMonday(Format$(gNow(), sgShowDateForm))
    edcDate.Text = smDate
    txtNumberDays.Text = 1
    imAllClick = False
    imAllStationClick = False
    imTerminate = False
    imExporting = False
    imFirstTime = True
    
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
    Erase tmStarGuideAst
    Erase tmStarGuideXRef
    Set frmExportStarGuide = Nothing
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
    Dim iLoop As Integer
    Dim iRet As Integer
    Dim sMoDate As String
    Dim sEndDate As String
    Dim slStr As String
    Dim ilOkStation As Integer
    Dim ilOkVehicle As Integer
    Dim ilVef As Integer
    Dim slSDate As String
    Dim slEDate As String
    Dim ilRet As Integer
    Dim ilPrevShttCode As Integer
    Dim ilPos As Integer
    Dim slCallLetters As String
    Dim slBand As String
    Dim slToFile As String
    Dim slFeedNo As String
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim slRunLetter As String
    Dim slDateTime As String
    Dim llTime As Long
    Dim llSpotTime As Long
    Dim slTime As String
    Dim llDate As Long
    Dim slDate As String
    Dim ilIncludeSpot As Integer
    Dim ilIndex As Integer
    Dim ilSIndex As Integer
    Dim ilEIndex As Integer
    Dim ilLoop As Integer
    Dim slSeqNo As String
    Dim ilBreakLen As Integer
    Dim slRCart As String
    Dim slRISCI As String
    Dim slRCreative As String
    Dim slRProd As String
    Dim llRCrfCsfCode As Long
    Dim llRCrfCode As Long
    
    On Error GoTo ErrHand
    sMoDate = gObtainPrevMonday(smDate)
    sEndDate = DateAdd("d", imNumberDays - 1, smDate)
    slSDate = smDate
    slEDate = gObtainNextSunday(slSDate)
    If DateValue(gAdjYear(sEndDate)) < DateValue(gAdjYear(slEDate)) Then
        slEDate = sEndDate
    End If
    slStr = smDate
    slStr = gAdjYear(slStr)
    gObtainYearMonthDayStr slStr, True, slFYear, slFMonth, slFDay
    slFeedNo = slFDay
    slRunLetter = Trim$(udcCriteria.SRunLetter)
    imVefCode = 0
    For ilVef = 0 To lbcVehicles.ListCount - 1
        If igExportSource = 2 Then DoEvents
        If lbcVehicles.Selected(ilVef) Then
            If imVefCode = 0 Then
                imVefCode = lbcVehicles.ItemData(ilVef)
            Else
                imVefCode = -1
                Exit For
            End If
        End If
    Next ilVef
    Do
        ReDim tmStarGuideXRef(0 To 0) As STARGUIDEXREF
        ilPrevShttCode = -1
        If igExportSource = 2 Then DoEvents
        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate"
        SQLQuery = SQLQuery & " FROM shtt, cptt, att"
        SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
        SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
        '10/29/14: Bypass Service agreements
        SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"
        If imVefCode > 0 Then
            SQLQuery = SQLQuery & " AND cpttVefCode = " & imVefCode
        End If
        SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sMoDate, sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " ORDER BY shttCallLetters, shttCode"
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
                ilOkVehicle = False
                For ilVef = 0 To lbcVehicles.ListCount - 1
                    If igExportSource = 2 Then DoEvents
                    If lbcVehicles.Selected(ilVef) Then
                        If lbcVehicles.ItemData(ilVef) = cprst!cpttvefcode Then
                            imVefCode = lbcVehicles.ItemData(ilVef)
                            ilOkVehicle = True
                            Exit For
                        End If
                    End If
                Next ilVef
            End If
            If ilOkStation And ilOkVehicle Then
                If ilPrevShttCode <> cprst!shttCode Then
                    If ilPrevShttCode <> -1 Then
                        'Output record
                        ilRet = mOutputSch(ilPrevShttCode)
                        'Close File
                        Close #hmTo
                    End If
                    'Open File
                    slCallLetters = cprst!shttCallLetters
                    slBand = ""
                    ilPos = InStr(1, slCallLetters, "-", vbTextCompare)
                    If ilPos > 0 Then
                        slBand = Trim$(Mid$(slCallLetters, ilPos + 1))
                        slBand = Left$(slBand, 1)
                        slCallLetters = Left$(slCallLetters, ilPos - 1)
                    End If
                    slCallLetters = gFileNameFilter(slCallLetters)
                    slToFile = smExportDirectory & Trim$(slCallLetters) & slBand & slFeedNo & slRunLetter & ".sch"
                    ilRet = 0
                    On Error GoTo mExportSpotsErr:
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
                        imExporting = False
                        Screen.MousePointer = vbDefault
                        gMsgBox "Open Error #" & Str$(Err.Numner) & slToFile, vbOKOnly, "Open Error"
                        mExportSpots = False
                        Exit Function
                    End If
                    ilPrevShttCode = cprst!shttCode
                    ReDim tmStarGuideAst(0 To 0) As STARGUIDEAST
                    Print #hmTo, "SCHEDULE"
                End If
                On Error GoTo ErrHand
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
                'Create AST records
                igTimes = 1 'By Week
                imAdfCode = -1
                If igExportSource = 2 Then DoEvents
                'Dan M 9/26/13  6442
                iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, False, True)
                gFilterAstExtendedTypes tmAstInfo
               ' iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, False, False, True)
                ilIndex = LBound(tmAstInfo)
                Do While ilIndex < UBound(tmAstInfo)
                    If igExportSource = 2 Then DoEvents
                    'If (DateValue(tmAstInfo(iLoop).sFeedDate) >= DateValue(smDate)) And (DateValue(tmAstInfo(iLoop).sFeedDate) <= DateValue(sEndDate)) And (tgStatusTypes(tmAstInfo(iLoop).iPledgeStatus).iPledged <> 2) Then
                    If (DateValue(gAdjYear(tmAstInfo(ilIndex).sFeedDate)) >= DateValue(gAdjYear(smDate))) And (DateValue(gAdjYear(tmAstInfo(ilIndex).sFeedDate)) <= DateValue(gAdjYear(sEndDate))) And (tgStatusTypes(gGetAirStatus(tmAstInfo(ilIndex).iStatus)).iPledged <> 2) Then
                        ilBreakLen = tmAstInfo(ilIndex).iLen
                        ilIncludeSpot = False
                        ilSIndex = ilIndex
                        ilEIndex = ilSIndex + 1
                        Do While ilEIndex < UBound(tmAstInfo)
                            If igExportSource = 2 Then DoEvents
                            If gTimeToLong(tmAstInfo(ilSIndex).sFeedTime, False) = gTimeToLong(tmAstInfo(ilEIndex).sFeedTime, False) Then
                                ilBreakLen = ilBreakLen + tmAstInfo(ilEIndex).iLen
                                ilEIndex = ilEIndex + 1
                            Else
                                Exit Do
                            End If
                        Loop
                        ilEIndex = ilEIndex - 1
                        If udcCriteria.SSpots(1) Then
                            'Check if regional copy defined with spots within same avail
                            For ilLoop = ilSIndex To ilEIndex Step 1
                                If igExportSource = 2 Then DoEvents
                                ''ilRet = gGetRegionCopy(tmAstInfo(ilLoop).iShttCode, tmAstInfo(ilLoop).lSdfCode, tmAstInfo(ilLoop).iVefCode, slRCart, slRProd, slRISCI, slRCreative, llRCrfCsfCode, llRCrfCode)
                                'ilRet = gGetRegionCopy(tmAstInfo(ilLoop), slRCart, slRProd, slRISCI, slRCreative, llRCrfCsfCode, llRCrfCode)
                                'If ilRet Then
                                If tmAstInfo(ilLoop).iRegionType > 0 Then
                                    ilIncludeSpot = True
                                    Exit For
                                End If
                            Next ilLoop
                        Else
                            ilIncludeSpot = True
                        End If
                        If ilIncludeSpot Then
                            For ilLoop = ilSIndex To ilEIndex Step 1
                                If igExportSource = 2 Then DoEvents
                                llDate = DateValue(tmAstInfo(ilLoop).sFeedDate)
                                llTime = gTimeToLong(tmAstInfo(ilLoop).sFeedTime, False)
                                'Translate time based on zone
                                Select Case UCase$(Trim$(cprst!shttTimeZone))
                                    Case "EST"
                                        llSpotTime = llTime
                                    Case "CST"
                                        llSpotTime = llTime + 3600
                                    Case "MST"
                                        llSpotTime = llTime + 2 * 3600
                                    Case "PST"
                                        llSpotTime = llTime + 3 * 3600
                                    Case Else
                                        llSpotTime = llTime
                                End Select
                                If (llSpotTime >= 24 * CLng(3600)) Then
                                    'Adjust date
                                    llDate = llDate + 1
                                    llSpotTime = llSpotTime - 24 * CLng(3600)
                                End If
                                tmAstInfo(ilLoop).sFeedDate = Format$(llDate, "m/d/yy")
                                tmAstInfo(ilLoop).sFeedTime = gLongToTime(llSpotTime)
                                slDate = Trim$(Str$(llDate))
                                Do While Len(slDate) < 6
                                    slDate = "0" & slDate
                                Loop
                                slTime = Trim$(Str$(llSpotTime))
                                Do While Len(slTime) < 6
                                    slTime = "0" & slTime
                                Loop
                                slSeqNo = Trim$(Str$(UBound(tmStarGuideAst)))
                                Do While Len(slSeqNo) < 6
                                    slSeqNo = "0" & slSeqNo
                                Loop
                                tmStarGuideAst(UBound(tmStarGuideAst)).sKey = slDate & "|" & slTime & "|" & slSeqNo
                                tmStarGuideAst(UBound(tmStarGuideAst)).tAstInfo = tmAstInfo(ilLoop)
                                tmStarGuideAst(UBound(tmStarGuideAst)).iBreakLen = ilBreakLen
                                tmStarGuideAst(UBound(tmStarGuideAst)).iVefCode = imVefCode
                                ReDim Preserve tmStarGuideAst(0 To UBound(tmStarGuideAst) + 1) As STARGUIDEAST
                            Next ilLoop
                        End If
                        ilIndex = ilEIndex
                    End If
                    ilIndex = ilIndex + 1
                Loop
                If igExportSource = 2 Then DoEvents
            End If
            
            'gLogMsg Trim$(cprst!shttCallLetters) & " Exported " & CStr(ilAddRecs) & " Add records and " & CStr(ilDeleteRecs) & " Delete Records", "StarGuideExportLog.Txt", False
            cprst.MoveNext
        Wend
        If ilPrevShttCode <> -1 Then
            'Output record
            ilRet = mOutputSch(ilPrevShttCode)
            'Close File
            Close #hmTo
            ilRet = mOutputXRef(slFeedNo)
            ilPrevShttCode = -1
        End If
        If (lbcStation.ListCount = 0) Or (chkAllStation.Value = vbChecked) Or (lbcStation.ListCount = lbcStation.SelCount) Then
            gClearASTInfo True
        Else
            gClearASTInfo False
        End If
        sMoDate = DateAdd("d", 7, sMoDate)
        slSDate = sMoDate
        slEDate = gObtainNextSunday(slSDate)
        If DateValue(gAdjYear(sEndDate)) < DateValue(gAdjYear(slEDate)) Then
            slEDate = sEndDate
        End If
        slStr = sMoDate
        slStr = gAdjYear(slStr)
        gObtainYearMonthDayStr slStr, True, slFYear, slFMonth, slFDay
        slFeedNo = slFDay
    Loop While DateValue(gAdjYear(sMoDate)) < DateValue(gAdjYear(sEndDate))

    mExportSpots = True
    Exit Function
mExportSpotsErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "StarGuideExportLog.txt", "Export StarGuide-mExportSpots"
    mExportSpots = False
    Exit Function
    
End Function

Private Sub mFillStations()
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
    SQLQuery = SQLQuery & " FROM shtt, att"
    SQLQuery = SQLQuery & " WHERE (attVefCode = " & imVefCode
    SQLQuery = SQLQuery & " AND shttCode = attShfCode)"
    SQLQuery = SQLQuery & " ORDER BY shttCallLetters"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        lbcStation.AddItem Trim$(rst!shttCallLetters)
        lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
        rst.MoveNext
    Wend
    chkAllStation.Value = vbChecked
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "StarGuideExportLog.txt", "Export StarGuide-mFillStations"

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
    Dim llRCrfCode As Long
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
    Dim llAdf As Long
    
    On Error GoTo ErrHand:
    If UBound(tmStarGuideAst) > 1 Then
        ArraySortTyp fnAV(tmStarGuideAst(), 0), UBound(tmStarGuideAst), 0, LenB(tmStarGuideAst(0)), 0, LenB(tmStarGuideAst(0).sKey), 0
    End If
    ilLoop = 0
    Do While ilLoop <= UBound(tmStarGuideAst) - 1
        If igExportSource = 2 Then DoEvents
        tlAstInfo = tmStarGuideAst(ilLoop).tAstInfo
        slAdvt = "Missing"
        slCart = ""
        slISCI = ""
        slCreative = ""
        SQLQuery = "SELECT lstProd, lstCart, lstISCI, lstLen, adfName, cpfCreative"
        SQLQuery = SQLQuery & " FROM (LST LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on lstCpfCode = cpfCode) LEFT OUTER JOIN ADF_Advertisers on lstadfCode = adfCode"
        SQLQuery = SQLQuery & " WHERE lstCode =" & Str(tlAstInfo.lLstCode)
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            If igExportSource = 2 Then DoEvents
            If IsNull(rst!adfName) = True Then
                slAdvt = "Missing"
            Else
                slAdvt = Trim$(rst!adfName)
            End If
            If IsNull(rst!lstProd) = True Then
                slProd = ""
            Else
                slProd = Trim$(rst!lstProd)
            End If
            If IsNull(rst!lstCart) Or Left$(rst!lstCart, 1) = Chr$(0) Then
                slCart = ""
            Else
                slCart = Trim$(rst!lstCart)
            End If
            If IsNull(rst!lstISCI) = True Then
                slISCI = ""
            Else
                slISCI = Trim$(rst!lstISCI)
            End If
            If IsNull(rst!cpfCreative) = True Then
                slCreative = ""
            Else
                slCreative = Trim$(rst!cpfCreative)
            End If
            slLen = Trim$(Str$(rst!lstLen))
        End If
        ''6/12/06- Check if any region copy defined for the spots
        ''ilRet = gGetRegionCopy(tlAstInfo.iShttCode, tlAstInfo.lSdfCode, tlAstInfo.iVefCode, slRCart, slRProd, slRISCI, slRCreative, llRCrfCsfCode, llRCrfCode)
        'ilRet = gGetRegionCopy(tlAstInfo, slRCart, slRProd, slRISCI, slRCreative, llRCrfCsfCode, llRCrfCode)
        'If ilRet Then
        If tlAstInfo.iRegionType > 0 Then
            slCart = Trim$(tlAstInfo.sRCart)   'slRCart
            slProd = Trim$(tlAstInfo.sRProduct)    'slRProd
            slISCI = Trim$(tlAstInfo.sRISCI)   'slRISCI
            slCreative = Trim$(tlAstInfo.sRCreativeTitle)  'slRCreative
        End If
        'Get Short Title
        slShortTitle = gGetShortTitle(tlAstInfo.lSdfCode)
        If tlAstInfo.iRegionType = 2 Then
            If igExportSource = 2 Then DoEvents
            llAdf = gBinarySearchAdf(CLng(tlAstInfo.iAdfCode))
            If llAdf <> -1 Then
                slShortTitle = Trim$(tgAdvtInfo(llAdf).sAdvtAbbr) & ", " & Trim$(tlAstInfo.sRProduct)
            Else
                slShortTitle = Trim$(tlAstInfo.sRProduct)
            End If
            slShortTitle = Left$(slShortTitle, 15)
        End If
        slVehicle = ""
        llVeh = gBinarySearchVef(CLng(tlAstInfo.iVefCode))
        If llVeh <> -1 Then
            slVehicle = Trim$(tgVehicleInfo(llVeh).sVehicle)
        End If
        If igExportSource = 2 Then DoEvents
        ilFound = False
        slKey = slShortTitle & "|" & slCart & "|" & slISCI & "|" & slCreative & "|" & slVehicle
        For ilXRef = 0 To UBound(tmStarGuideXRef) - 1 Step 1
            If StrComp(Trim$(tmStarGuideXRef(ilXRef).sKey), slKey, vbTextCompare) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilXRef
        If Not ilFound Then
            tmStarGuideXRef(UBound(tmStarGuideXRef)).sKey = slKey
            tmStarGuideXRef(UBound(tmStarGuideXRef)).iVefCode = tlAstInfo.iVefCode
            ReDim Preserve tmStarGuideXRef(0 To UBound(tmStarGuideXRef) + 1) As STARGUIDEXREF
        End If
        slISCI = gFileNameFilter(slISCI)
        If igExportSource = 2 Then DoEvents
        'Correct Date and Time
        slDate = gAdjYear(tlAstInfo.sFeedDate)
        slTime = tlAstInfo.sFeedTime
        slDate = Format$(slDate, "yyyy-mm-dd")
        slTime = Format$(slTime, "hh:mm:ss")
        If slTime = "12M" Then
            slTime = "00:00:00"
        End If
        slRecord = "EVENT: " & slDate & " " & slTime
        'Window
        llVpf = gBinarySearchVpf(CLng(tmStarGuideAst(ilLoop).iVefCode))
        If llVpf <> -1 Then
            ilWindow = tgVpfOptions(llVpf).lEDASWindow
        Else
            ilWindow = 400
        End If
        slRecord = slRecord & "," & Trim$(Str$(ilWindow))
        If tmStarGuideAst(ilLoop).iBreakLen = 30 Then
            slRecord = slRecord & ",0004"
        ElseIf tmStarGuideAst(ilLoop).iBreakLen = 60 Then
            slRecord = slRecord & ",0002"
        Else
            slRecord = slRecord & ",0001"
        End If
        slShortTitle = gFileNameFilter(slShortTitle)
        '7496
       ' slRecord = slRecord & "," & """" & slShortTitle & "(" & slISCI & ")" & ".mp2" & """"
        slRecord = slRecord & "," & """" & slShortTitle & "(" & slISCI & ")" & sgAudioExtension & """"
        ilLoop1 = ilLoop + 1
        Do While ilLoop1 <= UBound(tmStarGuideAst) - 1
            If igExportSource = 2 Then DoEvents
            tlAstInfo = tmStarGuideAst(ilLoop1).tAstInfo
            If DateValue(slDate) = DateValue(gAdjYear(tlAstInfo.sFeedDate)) Then
                If gTimeToLong(slTime, False) = gTimeToLong(tlAstInfo.sFeedTime, False) Then
                    If igExportSource = 2 Then DoEvents
                    slISCI = ""
                    SQLQuery = "SELECT lstProd, lstCart, lstISCI, lstLen, adfName, cpfCreative"
                    SQLQuery = SQLQuery & " FROM (LST LEFT OUTER JOIN CPF_Copy_Prodct_ISCI on lstCpfCode = cpfCode) LEFT OUTER JOIN ADF_Advertisers on lstadfCode = adfCode"
                    SQLQuery = SQLQuery & " WHERE lstCode =" & Str(tlAstInfo.lLstCode)
                    Set rst = gSQLSelectCall(SQLQuery)
                    If Not rst.EOF Then
                        If igExportSource = 2 Then DoEvents
                        If IsNull(rst!lstCart) Or Left$(rst!lstCart, 1) = Chr$(0) Then
                            slCart = ""
                        Else
                            slCart = Trim$(rst!lstCart)
                        End If
                        If IsNull(rst!lstISCI) = True Then
                            slISCI = ""
                        Else
                            slISCI = Trim$(rst!lstISCI)
                        End If
                        If IsNull(rst!cpfCreative) = True Then
                            slCreative = ""
                        Else
                            slCreative = Trim$(rst!cpfCreative)
                        End If
                    End If
                    ''6/12/06- Check if any region copy defined for the spots
                    ''ilRet = gGetRegionCopy(tlAstInfo.iShttCode, tlAstInfo.lSdfCode, tlAstInfo.iVefCode, slRCart, slRProd, slRISCI, slRCreative, llRCrfCsfCode, llRCrfCode)
                    'ilRet = gGetRegionCopy(tlAstInfo, slRCart, slRProd, slRISCI, slRCreative, llRCrfCsfCode, llRCrfCode)
                    'If ilRet Then
                    If tlAstInfo.iRegionType > 0 Then
                        slCart = Trim$(tlAstInfo.sRCart)   'slRCart
                        slISCI = Trim$(tlAstInfo.sRISCI)   'slRISCI
                        slCreative = Trim$(tlAstInfo.sRCreativeTitle)  'slRCreative
                    End If
                    If tlAstInfo.iRegionType = 2 Then
                        llAdf = gBinarySearchAdf(CLng(tlAstInfo.iAdfCode))
                        If llAdf <> -1 Then
                            slShortTitle = Trim$(tgAdvtInfo(llAdf).sAdvtAbbr) & ", " & Trim$(tlAstInfo.sRProduct)
                        Else
                            slShortTitle = Trim$(tlAstInfo.sRProduct)
                        End If
                        slShortTitle = Left$(slShortTitle, 15)
                    Else
                        slShortTitle = gGetShortTitle(tlAstInfo.lSdfCode)
                    End If
                    slVehicle = ""
                    llVeh = gBinarySearchVef(CLng(tlAstInfo.iVefCode))
                    If llVeh <> -1 Then
                        slVehicle = Trim$(tgVehicleInfo(llVeh).sVehicle)
                    End If
                    If igExportSource = 2 Then DoEvents
                    ilFound = False
                    slKey = slShortTitle & "|" & slCart & "|" & slISCI & "|" & slCreative & "|" & slVehicle
                    For ilXRef = 0 To UBound(tmStarGuideXRef) - 1 Step 1
                        If igExportSource = 2 Then DoEvents
                        If StrComp(Trim$(tmStarGuideXRef(ilXRef).sKey), slKey, vbTextCompare) = 0 Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilXRef
                    If Not ilFound Then
                        tmStarGuideXRef(UBound(tmStarGuideXRef)).sKey = slKey
                        tmStarGuideXRef(UBound(tmStarGuideXRef)).iVefCode = tlAstInfo.iVefCode
                        ReDim Preserve tmStarGuideXRef(0 To UBound(tmStarGuideXRef) + 1) As STARGUIDEXREF
                    End If
                    slISCI = gFileNameFilter(slISCI)
                    slShortTitle = gFileNameFilter(slShortTitle)
                    '7496
                    'slRecord = slRecord & "," & """" & slShortTitle & "(" & slISCI & ")" & ".mp2" & """"
                    slRecord = slRecord & "," & """" & slShortTitle & "(" & slISCI & ")" & sgAudioExtension & """"
                    ilLoop = ilLoop + 1
                    ilLoop1 = ilLoop1 + 1
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
        Print #hmTo, slRecord
        If igExportSource = 2 Then DoEvents
        ilLoop = ilLoop + 1
    Loop
    'Output EDAS
    If igExportSource = 2 Then DoEvents
    SQLQuery = "SELECT shttSerialNo1, shttSerialNo2"
    SQLQuery = SQLQuery & " FROM SHTT"
    SQLQuery = SQLQuery & " WHERE shttCode = " & Str(ilShttCode)
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If Trim$(rst!shttSerialNo1) <> "" Then
            Print #hmTo, "ADDR: " & Trim$(rst!shttSerialNo1)
        End If
        If Trim$(rst!shttSerialNo2) <> "" Then
            Print #hmTo, "ADDR: " & Trim$(rst!shttSerialNo2)
        End If
    End If
    mOutputSch = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "StarGuideExportLog.txt", "Export StarGuide-mOutputSch"
    mOutputSch = False
End Function

Private Function mOutputXRef(slFeedNo As String) As Integer
    Dim slStnCode As String
    Dim slXRefLetter As String
    Dim slExportFile As String
    Dim slTimeStamp As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilShowName As Integer
    Dim ilVpf As Integer
    Dim slRecord As String
    Dim slProdISCITitle As String
    Dim slPrevProdISCITitle As String
    Dim slProduct As String
    Dim slCart As String
    Dim slISCI As String
    Dim slCreative As String
    Dim slVehicle As String
    Dim slKey As String
    
    slStnCode = "X"
    slXRefLetter = Trim$(udcCriteria.SRunLetter)
    Do
        If igExportSource = 2 Then DoEvents
        slExportFile = smExportDirectory & slStnCode & slFeedNo & slXRefLetter & ".xrf"
        ilRet = 0
        'On Error GoTo mOutputXRefErr:
        'slTimeStamp = FileDateTime(slExportFile)
        ilRet = gFileExist(slExportFile)
        If ilRet = 0 Then
            slXRefLetter = Chr$(Asc(slXRefLetter) + 1)
        End If
    Loop While ilRet = 0    'equal zero if file exist

    'ilRet = 0
    'On Error GoTo mOutputXRefErr:
    'hmTo = FreeFile
    'Open slExportFile For Output As hmTo
    ilRet = gFileOpen(slExportFile, "Output", hmTo)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        gMsgBox "Open " & slExportFile & ", Error #" & Str$(ilRet), vbOKOnly, "Open Error"
        mOutputXRef = False
        Exit Function
    End If
    If igExportSource = 2 Then DoEvents
    If UBound(tmStarGuideXRef) > 1 Then
        ArraySortTyp fnAV(tmStarGuideXRef(), 0), UBound(tmStarGuideXRef), 0, LenB(tmStarGuideXRef(0)), 0, LenB(tmStarGuideXRef(0).sKey), 0
    End If
    slRecord = " "
    Do While Len(slRecord) < 35
        slRecord = slRecord & " "
    Loop
    slRecord = slRecord & Trim$(sgClientName)
    Print #hmTo, slRecord
    slRecord = " "
    Do While Len(slRecord) < 35
        slRecord = slRecord & " "
    Loop
    slRecord = slRecord & "Cross Reference"
    Print #hmTo, slRecord
    slRecord = " "
    Do While Len(slRecord) < 35
        slRecord = slRecord & " "
    Loop
    slRecord = slRecord & smDate
    Print #hmTo, slRecord
    Print #hmTo, ""
    Print #hmTo, ""
    
    slRecord = "Short Title"
    Do While Len(slRecord) < 20
        slRecord = slRecord & " "
    Loop
    slRecord = slRecord & " Cart"
    Do While Len(slRecord) < 31
        slRecord = slRecord & " "
    Loop
    slRecord = slRecord & " ISCI"
    Do While Len(slRecord) < 52
        slRecord = slRecord & " "
    Loop
    slRecord = slRecord & " Creative Title"
    Do While Len(slRecord) < 83
        slRecord = slRecord & " "
    Loop
    slRecord = slRecord & " Vehicle"
    Print #hmTo, slRecord
    slPrevProdISCITitle = ""
    For ilLoop = 0 To UBound(tmStarGuideXRef) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        slKey = Trim$(tmStarGuideXRef(ilLoop).sKey)
        ilRet = gParseItem(slKey, 1, "|", slProduct)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 2, "|", slCart)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 3, "|", slISCI)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 4, "|", slCreative)  'Obtain Index and code number
        ilRet = gParseItem(slKey, 5, "|", slVehicle)  'Obtain Index and code number
        ilShowName = False
        ilVpf = gBinarySearchVpf(CLng(tmStarGuideXRef(ilLoop).iVefCode))
        If ilVpf <> -1 Then
            If tgVpfOptions(ilVpf).sStnFdXRef = "Y" Then
                ilShowName = True
            End If
        End If
        If igExportSource = 2 Then DoEvents
        If ilShowName Then
            Do While Len(slProduct) < 20
                slProduct = slProduct & " "
            Loop
            If Len(slProduct) > 20 Then
                slProduct = Left$(slProduct, 20)
            End If
            Do While Len(slCart) < 10
                slCart = slCart & " "
            Loop
            Do While Len(slISCI) < 20
                slISCI = slISCI & " "
            Loop
            Do While Len(slCreative) < 30
                slCreative = slCreative & " "
            Loop
            slProdISCITitle = slProduct & " " & slCart & " " & slISCI & " " & slCreative
            If slPrevProdISCITitle <> slProdISCITitle Then
                Print #hmTo, ""
                slPrevProdISCITitle = slProdISCITitle
                slRecord = slProdISCITitle
                slRecord = slRecord & " " & slVehicle
            Else
                slRecord = " "
                Do While Len(slRecord) < Len(slProdISCITitle)
                    slRecord = slRecord & " "
                Loop
                slRecord = slRecord & " " & slVehicle
            End If
            Print #hmTo, slRecord
        End If
    Next ilLoop
    Close #hmTo
    mOutputXRef = True
    Exit Function
'mOutputXRefErr:
'    ilRet = Err.Number
'    Resume Next
End Function

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload frmExportStarGuide
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
        For ilLoop = 0 To lbcStation.ListCount - 1
            If lbcStation.Selected(ilLoop) Then
                ilShttCode(UBound(ilShttCode)) = lbcStation.ItemData(ilLoop)
                ReDim Preserve ilShttCode(0 To UBound(ilShttCode) + 1) As Integer
            End If
        Next ilLoop
        udcCriteria.Action 5
        lmEqtCode = gCustomStartStatus("S", "StarGuide", "S", Trim$(edcDate.Text), Trim$(txtNumberDays.Text), ilVefCode(), ilShttCode())
    End If
End Sub
