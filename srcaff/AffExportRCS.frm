VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmExportRCS 
   Caption         =   "Export RCS"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   Icon            =   "AffExportRCS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   9885
   Begin V81Affiliate.CSI_Calendar edcDate 
      Height          =   285
      Left            =   1275
      TabIndex        =   1
      Top             =   60
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
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   1755
      Width           =   3825
   End
   Begin VB.TextBox edcTitle2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5235
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Results"
      Top             =   1755
      Width           =   3825
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9675
      Top             =   3300
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   4605
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   2400
      ItemData        =   "AffExportRCS.frx":08CA
      Left            =   5070
      List            =   "AffExportRCS.frx":08CC
      TabIndex        =   7
      Top             =   2010
      Width           =   4455
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   2400
      ItemData        =   "AffExportRCS.frx":08CE
      Left            =   135
      List            =   "AffExportRCS.frx":08D0
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   2010
      Width           =   3855
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9480
      Top             =   4740
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5610
      FormDesignWidth =   9885
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   5820
      TabIndex        =   5
      Top             =   5115
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7860
      TabIndex        =   6
      Top             =   5100
      Width           =   1575
   End
   Begin V81Affiliate.AffExportCriteria udcCriteria 
      Height          =   810
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1429
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   120
      TabIndex        =   8
      Top             =   5055
      Width           =   5490
   End
   Begin VB.Label Label1 
      Caption         =   "Export Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1395
   End
End
Attribute VB_Name = "frmExportRCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmContact - allows for selection of station/vehicle/advertiser for contact information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private smDate As String     'Export Date
Private imVefCode As Integer
Private smVefName As String
Private smZone As String
Private imAllClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
Private smGetZone As String
Private imLocalAdj As Integer
Private smExportDirectory As String
Private hmMsg As Integer
Private hmTo As Integer
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
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer

    'On Error GoTo mOpenMsgFileErr:
    slToFile = smExportDirectory & "ExptRCS.Txt"
    slNowDate = Format$(gNow(), sgShowDateForm)
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, sgShowDateForm)
        If DateValue(gAdjYear(slFileDate)) = DateValue(gAdjYear(slNowDate)) Then  'Append
            On Error GoTo 0
            'ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Close hmMsg
                hmMsg = -1
                gMsgBox "Open File " & slToFile & " error #" & Str$(Err.Number), vbOKOnly
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            'ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Close hmMsg
                hmMsg = -1
                gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        'ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, "** Export Vehicle RCS Info: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    sMsgFileName = slToFile
    mOpenMsgFile = True
    Exit Function
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
    If chkAll.Value = 1 Then
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
    Dim iLoop As Integer
    Dim sYear As String
    Dim sMonth As String
    Dim sDay As String
    Dim sFileName As String
    Dim sLetter As String
    Dim iRet As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim sToFile As String
    Dim sDateTime As String
    Dim sMsgFileName As String
    Dim ilRet As Integer

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
    Else
        smDate = Format(edcDate.Text, sgShowDateForm)
    End If
    smExportDirectory = udcCriteria.RCSExportToPath
    If smExportDirectory = "" Then
        smExportDirectory = sgExportDirectory
    End If
    smExportDirectory = gSetPathEndSlash(smExportDirectory, False)
    Screen.MousePointer = vbHourglass
    
    If Not mOpenMsgFile(sMsgFileName) Then
        igExportReturn = 2
        cmdCancel.SetFocus
        Exit Sub
    End If
    imExporting = True
    mSaveCustomValues
    gObtainYearMonthDayStr smDate, True, sYear, sMonth, sDay
    sFileName = sMonth & sDay & right$(sYear, 2)
    lacResult.Caption = ""
    For iLoop = 0 To lbcVehicles.ListCount - 1
        If igExportSource = 2 Then DoEvents
        If lbcVehicles.Selected(iLoop) Then
            'Get hmTo handle
            imVefCode = lbcVehicles.ItemData(iLoop)
            
            For iVef = 0 To UBound(tgVehicleInfo) - 1 Step 1
                If igExportSource = 2 Then DoEvents
                If tgVehicleInfo(iVef).iCode = imVefCode Then
                    'For iZone = LBound(tgVehicleInfo(iVef).sZone) To UBound(tgVehicleInfo(iVef).sZone) Step 1
                    '    If tgVehicleInfo(iVef).sFed(iZone) = "*" Then
                    For iZone = 0 To 3 Step 1
                        If igExportSource = 2 Then DoEvents
                        If udcCriteria.RCSZone(iZone) = vbChecked Then
                            smVefName = Trim$(tgVehicleInfo(iVef).sVehicle)
                            'smZone = Trim$(tgVehicleInfo(iVef).sZone(iZone))
                            'Print #hmMsg, "** Generating Data for " & smVefName & "-" & Trim$(tgVehicleInfo(iVef).sZone(iZone)) & " **"
                            Select Case iZone
                                Case 0
                                    smZone = "EST"
                                Case 1
                                    smZone = "CST"
                                Case 2
                                    smZone = "MST"
                                Case 3
                                    smZone = "PST"
                            End Select
                            imLocalAdj = tgVehicleInfo(iVef).iLocalAdj(iZone)
                            smGetZone = ""
                            If tgVehicleInfo(iVef).sFed(iZone) = "*" Then
                                smGetZone = Trim$(tgVehicleInfo(iVef).sZone(iZone))
                            Else
                                smGetZone = Trim$(tgVehicleInfo(iVef).sFed(iZone)) & "ST"
                            End If
                            Print #hmMsg, "** Generating Data for " & smVefName & "-" & Trim$(smZone) & " **"
                            sLetter = Trim$(Left$(tgVehicleInfo(iVef).sCodeStn, 3))
                            iRet = 0
                            'On Error GoTo cmdExportErr:
                            'sToFile = sFileName & Left$(tgVehicleInfo(iVef).sZone(iZone), 1) & sLetter & ".Txt"
                            sToFile = sFileName & Left$(smZone, 1) & sLetter & ".Log"
                            'sDateTime = FileDateTime(smExportDirectory & sToFile)
                            iRet = gFileExist(smExportDirectory & sToFile)
                            If iRet = 0 Then
                                Kill smExportDirectory & sToFile
                            End If
                            On Error GoTo 0
                            'sToFile = sgExportDirectory & sFileName & sLetter & ".RCS"
                            sToFile = smExportDirectory & sFileName & Left$(smZone, 1) & sLetter & ".Log"
                            On Error GoTo 0
                            'iRet = 0
                            'On Error GoTo cmdExportErr:
                            'hmTo = FreeFile
                            'Open sToFile For Output As hmTo
                            iRet = gFileOpen(sToFile, "Output", hmTo)
                            If iRet <> 0 Then
                                Print #hmMsg, "** Terminated **"
                                Close #hmMsg
                                Close #hmTo
                                ilRet = gCustomEndStatus(lmEqtCode, 2, "")
                                imExporting = False
                                Screen.MousePointer = vbDefault
                                gMsgBox "Open Error #" & Str$(Err.Numner) & sToFile, vbOKOnly, "Open Error"
                                Exit Sub
                            End If
                            Print #hmMsg, "** Storing Output into " & sToFile & " **"
                            If igExportSource = 2 Then DoEvents
                            iRet = mExportRCS()
                            If (iRet = False) Then
                                Print #hmMsg, "** Terminated **"
                                Close #hmMsg
                                Close #hmTo
                                ilRet = gCustomEndStatus(lmEqtCode, 2, "")
                                imExporting = False
                                Screen.MousePointer = vbDefault
                                cmdCancel.SetFocus
                                Exit Sub
                            End If
                            If imTerminate Then
                                Print #hmMsg, "** User Terminated **"
                                Close #hmMsg
                                Close #hmTo
                                ilRet = gCustomEndStatus(lmEqtCode, 2, "")
                                imExporting = False
                                Screen.MousePointer = vbDefault
                                cmdCancel.SetFocus
                                Exit Sub
                            End If
                            Print #hmMsg, "** Completed " & smVefName & " **"
                            Close #hmTo
                        End If
                    Next iZone
                    Exit For
                End If
            Next iVef
        End If
    Next iLoop
    ilRet = gCustomEndStatus(lmEqtCode, 1, "")
    imExporting = False
    Print #hmMsg, "** Completed Export RCS Linker: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Close #hmMsg
    lacResult.Caption = "See: " & sMsgFileName & " for Result Summary"
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Exit Sub
'cmdExportErr:
'    iRet = Err
'    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmExportRCS-mcmdExport_Click"
    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    edcDate.Text = ""
    Unload frmExportRCS
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
        udcCriteria.Top = Label1.Top + (2 * Label1.Height)
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
            igExportReturn = 1
            '6394 move before 'click'
            sgExportResultName = "RCSResultList.Txt"
            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
            gLogMsgWODT "W", hlResult, "RCS Result List, Started: " & slNowStart
            hgExportResult = hlResult
            cmdExport_Click
            slNowEnd = gNow()
            'Output result list box
'            sgExportResultName = "RCSResultList.Txt"
'            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
'            gLogMsgWODT "W", hlResult, "RCS Result List, Started: " & slNowStart
            If lbcMsg.ListCount > 0 Then
                For ilLoop = 0 To lbcMsg.ListCount - 1 Step 1
                    gLogMsgWODT "W", hlResult, Trim$(lbcMsg.List(ilLoop))
                Next ilLoop
            End If
            gLogMsgWODT "W", hlResult, "RCS Result List, Completed: " & slNowEnd
            gLogMsgWODT "C", hlResult, ""
            '6394 clear values
            hgExportResult = 0
            tmcTerminate.Enabled = True
        End If
        imFirstTime = False
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.7
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmExportRCS
    If igExportSource = 2 Then
        Me.Top = -(2 * Me.Top + Screen.Height)
    End If
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    frmExportRCS.Caption = "Export RCS - " & sgClientName
    'Me.Width = Screen.Width / 1.5
    'Me.Height = Screen.Height / 1.7
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    smDate = ""
    imAllClick = False
    imTerminate = False
    imExporting = False
    imFirstTime = True
    
    mFillVehicle
'    chkZone(0).Enabled = True
'    chkZone(1).Enabled = True
'    chkZone(2).Enabled = True
'    chkZone(3).Enabled = True
'    chkZone(0).Value = 0
'    chkZone(1).Value = 0
'    chkZone(2).Value = 0
'    chkZone(3).Value = 0
'    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
'        For iZone = LBound(tgVehicleInfo(iLoop).sZone) To UBound(tgVehicleInfo(iLoop).sZone) Step 1
'            Select Case Left$(tgVehicleInfo(iLoop).sZone(iZone), 1)
'                Case "E"
'                    If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
'                        'chkZone(0).Enabled = True
'                        chkZone(0).Value = 1
'                    End If
'                Case "C"
'                    If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
'                        'chkZone(1).Enabled = True
'                        chkZone(1).Value = 1
'                    End If
'                Case "M"
'                    If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
'                        'chkZone(2).Enabled = True
'                        chkZone(2).Value = 1
'                    End If
'                Case "P"
'                    If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
'                        'chkZone(3).Enabled = True
'                        chkZone(3).Value = 1
'                    End If
'            End Select
'        Next iZone
'    Next iLoop
    ilRet = gPopAvailNames()
    Screen.MousePointer = vbDefault
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If imExporting Then
        imTerminate = True
        Cancel = True
        Exit Sub
    End If
    Set frmExportRCS = Nothing
End Sub


Private Sub lbcVehicles_Click()
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        imAllClick = True
        chkAll.Value = 0
        imAllClick = False
    End If
End Sub

Private Sub edcDate_Change()
    lbcMsg.Clear
End Sub


Private Function mExportRCS()
    Dim sHour As String
    Dim sMin As String
    Dim sSec As String
    Dim sRecord As String
    Dim sTime As String * 7
    Dim sCart As String * 4
    Dim sAdvtProd As String * 24
    Dim slTime As String * 8
    Dim slCart As String * 5
    Dim slAdvtName As String
    Dim slProd As String
    Dim sUnused12 As String * 12
    Dim sUnused3 As String * 3
    Dim sUnused4 As String * 4
    Dim sLen As String
    Dim sMsg As String
    Dim iLoop As Integer
    Dim iLen As Integer
    Dim sCode As String
    Dim iRet As Integer
    Dim sStr As String
    Dim lSpotTime As Long
    Dim lAvailTime As Long
    Dim lRunTime As Long
    Dim lDate As Long
    Dim sAirDate As String
    Dim lTime As Long
    Dim rstDat As ADODB.Recordset
    Dim blSpotOk As Boolean
    Dim ilAnf As Integer
       
    If igExportSource = 2 Then DoEvents
    
    sUnused12 = "            "
    sUnused3 = "   "
    sUnused4 = "    "
    
    On Error GoTo ErrHand
    lDate = DateValue(gAdjYear(smDate))
    lAvailTime = -1
    lRunTime = 0
    SQLQuery = "SELECT * FROM lst "
    SQLQuery = SQLQuery + " WHERE (lstLogVefCode = " & imVefCode
    SQLQuery = SQLQuery + " AND lstType = " & 0
    SQLQuery = SQLQuery & " AND lstZone = '" & smGetZone & "'"
    SQLQuery = SQLQuery + " AND lstBkoutLstCode = 0"
    SQLQuery = SQLQuery + " AND Mod(lstStatus, 100) < " & ASTEXTENDED_MG 'Bypass MG/Bonus
    If imLocalAdj = 0 Then
        SQLQuery = SQLQuery + " AND lstLogDate = '" & Format$(smDate, sgSQLDateForm) & "')"
    Else
        If imLocalAdj < 0 Then
            SQLQuery = SQLQuery + " AND (lstLogDate >= '" & Format$(lDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(lDate + 1, sgSQLDateForm) & "')" & ")"
        Else
            SQLQuery = SQLQuery + " AND (lstLogDate >= '" & Format$(lDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(lDate, sgSQLDateForm) & "')" & ")"
        End If
    End If
    If igExportSource = 2 Then DoEvents
    SQLQuery = SQLQuery + " ORDER BY lstLogDate, lstLogTime, lstBreakNo, lstPositionNo"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        If igExportSource = 2 Then DoEvents
        blSpotOk = True
        ilAnf = gBinarySearchAnf(rst!lstAnfCode)
        If ilAnf <> -1 Then
            If tgAvailNamesInfo(ilAnf).sAutomationExport = "N" Then
                blSpotOk = False
            End If
        End If
        If (blSpotOk) And (tgStatusTypes(rst!lstStatus).iPledged <> 2) Then
            'Record type
            If igRCSExportBy <> 5 Then
                sRecord = "C"
            Else
                sRecord = ""
            End If
            'Time
            lSpotTime = gTimeToLong(Format$(rst!lstLogTime, "hh:mm:ssam/pm"), False)
            If lSpotTime = lAvailTime Then
                lSpotTime = lRunTime
                'sTime = Format$(gLongToTime(lRunTime), "HHMM:SS")
                lRunTime = lRunTime + rst!lstLen
            Else
                lAvailTime = lSpotTime
                'sTime = Format$(rst!lstLogTime, "HHMM:SS")
                lRunTime = lAvailTime + rst!lstLen
            End If
            sAirDate = Format$(rst!lstLogDate, sgShowDateForm)
            lTime = lSpotTime + 3600 * imLocalAdj
            If lTime < 0 Then
                lTime = lTime + 86400
                sAirDate = Format$(DateValue(gAdjYear(sAirDate)) - 1, sgShowDateForm)
            ElseIf lTime > 86400 Then
                lTime = lTime - 86400
                sAirDate = Format$(DateValue(gAdjYear(sAirDate)) + 1, sgShowDateForm)
            End If
            If igRCSExportBy <> 5 Then
                sTime = Format$(gLongToTime(lTime), "HHMM:SS")
            Else
                slTime = Format$(gLongToTime(lTime), "HH:MM:SS")
            End If
            If (DateValue(gAdjYear(sAirDate)) = lDate) Then
                'Cart (4 characters)
                If IsNull(rst!lstCart) Or Left$(rst!lstCart, 1) = Chr$(0) Then
                    sCart = ""
                    slCart = ""
                Else
                    sStr = UCase$(rst!lstCart)
                    If igRCSExportBy <> 5 Then
                        If (Left$(sStr, 1) >= "A") And (Left$(sStr, 1) <= "Z") Then
                            sCart = Mid$(rst!lstCart, 2, 4)
                        Else
                            sCart = Mid$(rst!lstCart, 1, 4)
                        End If
                    Else
                        If (Left$(sStr, 1) >= "A") And (Left$(sStr, 1) <= "Z") Then
                            slCart = Mid$(rst!lstCart, 2, 5)
                        Else
                            slCart = Mid$(rst!lstCart, 1, 5)
                        End If
                    End If
                End If
                If igRCSExportBy <> 5 Then
                    If Len(Trim$(sCart)) = 0 Then
                        sMsg = "Cart Missing: " & smVefName & " " & smDate & " " & Format$(gLongToTime(lAvailTime), sgShowTimeWSecForm)
                        Print #hmMsg, sMsg
                        lbcMsg.AddItem sMsg
                    End If
                Else
                    If Len(Trim$(slCart)) = 0 Then
                        sMsg = "Cart Missing: " & smVefName & " " & smDate & " " & Format$(gLongToTime(lAvailTime), sgShowTimeWSecForm)
                        Print #hmMsg, sMsg
                        lbcMsg.AddItem sMsg
                    End If
                End If
                sAdvtProd = ""
                slAdvtName = ""
                slProd = ""
                For iLoop = 0 To UBound(tgAdvtInfo) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    If tgAdvtInfo(iLoop).iCode = rst!lstAdfCode Then
                        If igRCSExportBy <> 5 Then
                            If IsNull(rst!lstProd) Then
                                sAdvtProd = Trim$(tgAdvtInfo(iLoop).sAdvtAbbr)
                            Else
                                sAdvtProd = Trim$(tgAdvtInfo(iLoop).sAdvtAbbr) & "/" & rst!lstProd
                            End If
                        Else
                            If IsNull(rst!lstProd) Then
                                sAdvtProd = Trim$(tgAdvtInfo(iLoop).sAdvtAbbr)
                                slAdvtName = Trim$(tgAdvtInfo(iLoop).sAdvtName)
                            Else
                                sAdvtProd = Trim$(tgAdvtInfo(iLoop).sAdvtAbbr) & "/" & rst!lstProd
                                slAdvtName = Trim$(tgAdvtInfo(iLoop).sAdvtName)
                                slProd = Trim$(rst!lstProd)
                            End If
                        End If
                        Exit For
                    End If
                Next iLoop
                If sAdvtProd <> "" Then
                    iLen = rst!lstLen
                    If iLen < 60 Then
                        If igRCSExportBy <> 5 Then
                            sLen = Trim$(Str$(iLen))
                        Else
                            sLen = Trim$(Str$(iLen))
                            If Len(sLen) = 1 Then
                                sLen = ":0" & sLen
                            Else
                                sLen = ":" & sLen
                            End If
                        End If
                    Else
                        sLen = Trim$(Str$(iLen - 60))
                        Do While Len(sLen) < 2
                            sLen = "0" & sLen
                        Loop
                        If igRCSExportBy <> 5 Then
                            sLen = "01" & sLen
                        Else
                            sLen = "01:" & sLen
                        End If
                    End If
                    If igRCSExportBy <> 5 Then
                        Do While Len(sLen) < 4
                            sLen = "0" & sLen
                        Loop
                    Else
                        Do While Len(sLen) < 5
                            sLen = "0" & sLen
                        Loop
                    End If
                    sCode = Trim$(Str$(rst!lstCode))
                    If igRCSExportBy <> 5 Then
                        Do While Len(sCode) < 8
                            sCode = " " & sCode
                        Loop
                    End If
                    If igRCSExportBy <> 5 Then
                        sRecord = sRecord & sTime & sCart & sAdvtProd & sUnused3 & sLen & sUnused4 & sCode & sUnused12
                    Else
                        sRecord = slTime & Chr(254) & sLen & Chr(254) & "COM" & Chr(254) & slCart & Chr(254) & slAdvtName & Chr(254) & slProd & Chr(254) & Chr(254) & sCode & Chr(254)
                        For iLoop = 1 To 11 Step 1
                            sRecord = sRecord & Chr(254)
                        Next iLoop
                    End If
                    iRet = 0
                    On Error GoTo mExportRCSErr:
                    Print #hmTo, sRecord
                    If iRet <> 0 Then
                        mExportRCS = False
                        Exit Function
                    End If
                Else
                    sMsg = "Advertiser Missing: " & smVefName & " " & smDate & " " & sTime
                    Print #hmMsg, sMsg
                    lbcMsg.AddItem sMsg
                End If
                If igExportSource = 2 Then DoEvents
            End If
        End If
        On Error GoTo ErrHand
        rst.MoveNext
    Wend
    mExportRCS = True
    Exit Function
mExportRCSErr:
    iRet = Err
    Resume Next

    'Record format:
    'Column  Length  Field
    '  1        1    Commercial Indicator always C
    '  2        7    Start Time HRMN:SC  Military Hours
    '  9        4    Cart Number
    ' 13       24    Advertiser Abbr/Product
    ' 37        3    Unused (Priority number)
    ' 40        4    Length  MNSC (60=> 0100; 90=> 0130; 30=>0030)
    ' 44        4    Unused (Commercial Type)
    ' 48        6    First Part of Sdf.lCode (was Unused (Customer ID))
    ' 54        2    Second Part of Sdf.lCode (was Unused (Internal Code))
    '                use 8 bytes to make up sdd.lcode
    ' 56        4    Unused (Product Code)
    ' 60        8    Unused (Ordered Time)
    ' 68        1    Carriage Return <cr>
    ' 69        1    Line Feed <lf>
    '
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmExportRCS-mExportRCS"
    mExportRCS = False
    Exit Function

End Function

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload frmExportRCS
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
        If igRCSExportBy <> 5 Then
            lmEqtCode = gCustomStartStatus("4", "RCS 4 Digit Cart #'s", "4", Trim$(edcDate.Text), "1", ilVefCode(), ilShttCode())
        Else
            lmEqtCode = gCustomStartStatus("5", "RCS 5 Digit Cart #'s", "5", Trim$(edcDate.Text), "1", ilVefCode(), ilShttCode())
        End If
    End If
End Sub

