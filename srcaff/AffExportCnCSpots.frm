VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmExportCnCSpots 
   Caption         =   "Export Clearance n Compensation Spots"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   Icon            =   "AffExportCnCSpots.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   9645
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9360
      Top             =   3840
   End
   Begin V81Affiliate.CSI_Calendar edcDate 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   150
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
      Left            =   4275
      Locked          =   -1  'True
      TabIndex        =   8
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
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   1530
      Width           =   3810
   End
   Begin VB.CheckBox chkAllStation 
      Caption         =   "All"
      Height          =   195
      Left            =   4215
      TabIndex        =   10
      Top             =   3915
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.ListBox lbcStation 
      Height          =   2010
      ItemData        =   "AffExportCnCSpots.frx":08CA
      Left            =   4200
      List            =   "AffExportCnCSpots.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1770
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtNumberDays 
      Height          =   285
      Left            =   4605
      TabIndex        =   3
      Text            =   "7"
      Top             =   150
      Width           =   405
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3915
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   3375
      ItemData        =   "AffExportCnCSpots.frx":08CE
      Left            =   6585
      List            =   "AffExportCnCSpots.frx":08D0
      TabIndex        =   14
      Top             =   450
      Width           =   2820
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   2010
      ItemData        =   "AffExportCnCSpots.frx":08D2
      Left            =   120
      List            =   "AffExportCnCSpots.frx":08D4
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
      FormDesignWidth =   9645
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
      Top             =   750
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1429
   End
   Begin VB.Label Label2 
      Caption         =   "Number of Days"
      Height          =   255
      Left            =   3195
      TabIndex        =   2
      Top             =   195
      Width           =   1335
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
      Top             =   195
      Width           =   1395
   End
End
Attribute VB_Name = "frmExportCnCSpots"
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
Private imNumberDays As Integer
Private imVefCode As Integer
Private imAdfCode As Integer
Private smVefName As String
Private imAllClick As Integer
Private imAllStationClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
Private hmMsg As Integer
Private hmTo As Integer
Private hmFrom As Integer
Private hmAst As Integer
Private lmEqtCode As Long
Private cprst As ADODB.Recordset
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private tmAetInfo() As AETINFO
Private tmAet() As AETINFO






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
    ilRet = 0
    slNowDate = Format$(gNow(), sgShowDateForm)
    slToFile = sgMsgDirectory & "ExptCnCSpots.Txt"
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
    Print #hmMsg, "** Export CnC Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
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
    Dim iRet As Integer
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
        Exit Sub
    Else
        smDate = Format(edcDate.Text, sgShowDateForm)
    End If
    sNowDate = Format$(gNow(), "m/d/yy")
    If DateValue(gAdjYear(smDate)) > DateValue(gAdjYear(sNowDate)) Then
        Beep
        gMsgBox "Date must be prior to today's date " & sNowDate, vbCritical
        edcDate.SetFocus
        Exit Sub
    End If
    sMoDate = gObtainPrevMonday(smDate)
    llSDate = DateValue(gAdjYear(smDate))
    imNumberDays = Val(txtNumberDays.Text)
    If imNumberDays <= 0 Then
        gMsgBox "Number of days must be specified.", vbOKOnly
        txtNumberDays.SetFocus
        Exit Sub
    End If
    llEDate = DateValue(gAdjYear(Format$(DateAdd("d", imNumberDays - 1, smDate), "mm/dd/yy")))
    Screen.MousePointer = vbHourglass
    
    'Open the CEF file with an API call
    If Not mOpenCEFFile Then
        igExportReturn = 2
        'Stop the Pervasive API engine
        imExporting = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If Not mOpenMsgFile(sMsgFileName) Then
        igExportReturn = 2
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    imExporting = True
    iRet = 0
    On Error GoTo cmdExportErr:
    sToFile = udcCriteria.edcCFile
    'sDateTime = FileDateTime(sToFile)
    iRet = gFileExist(sToFile)
    If iRet = 0 Then
        Screen.MousePointer = vbDefault
        sDateTime = gFileDateTime(sToFile)
        iRet = gMsgBox("Export Previously Created " & sDateTime & " Continue with Export by Replacing File?", vbOKCancel, "File Exist")
        If (iRet = vbCancel) Or (igExportSource = 2) Then
            'Close the CEF File with an API call
            Call mCloseCEFFile
            'Stop the Pervasive API engine
            Print #hmMsg, "** Terminated Because Export File Existed **"
            Close #hmMsg
            Close #hmTo
            imExporting = False
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        Kill sToFile
    End If
    On Error GoTo 0
    'iRet = 0
    'On Error GoTo cmdExportErr:
    'hmTo = FreeFile
    'Open sToFile For Output As hmTo
    iRet = gFileOpen(sToFile, "Output", hmTo)
    If iRet <> 0 Then
        'Close the CEF File with an API call
        Call mCloseCEFFile
        'Stop the Pervasive API engine
        Print #hmMsg, "** Terminated **"
        Close #hmMsg
        Close #hmTo
        imExporting = False
        Screen.MousePointer = vbDefault
        gMsgBox "Open Error #" & Str$(Err.Numner) & sToFile, vbOKOnly, "Open Error"
        Exit Sub
    End If
    Print #hmMsg, "** Storing Output into " & sToFile & " **"
    On Error GoTo 0
    
    mSaveCustomValues
    bgTaskBlocked = False
    sgTaskBlockedName = "CnC Export"
    lacResult.Caption = ""
    For iLoop = 0 To lbcVehicles.ListCount - 1
        If igExportSource = 2 Then DoEvents
        If lbcVehicles.Selected(iLoop) Then
            'Get hmTo handle
            imVefCode = lbcVehicles.ItemData(iLoop)
            smVefName = Trim$(lbcVehicles.List(iLoop))
            Screen.MousePointer = vbHourglass
            iRet = mExportSpots()
            If (iRet = False) Then
                igExportReturn = 2
                gCloseRegionSQLRst
                bgTaskBlocked = False
                sgTaskBlockedName = ""
                'Close the CEF File with an API call
                Call mCloseCEFFile
                'Stop the Pervasive API engine
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
                igExportReturn = 2
                gCloseRegionSQLRst
                bgTaskBlocked = False
                sgTaskBlockedName = ""
                'Close the CEF File with an API call
                Call mCloseCEFFile
                'Stop the Pervasive API engine
                Print #hmMsg, "** User Terminated **"
                Close #hmMsg
                Close #hmTo
                ilRet = gCustomEndStatus(lmEqtCode, 2, "")
                imExporting = False
                Screen.MousePointer = vbDefault
                cmdCancel.SetFocus
                Exit Sub
            End If
       End If
    Next iLoop
    gCloseRegionSQLRst
    If bgTaskBlocked And igExportSource <> 2 Then
         gMsgBox "Some spots were blocked during the Export generation." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
    End If
    bgTaskBlocked = False
    sgTaskBlockedName = ""

    'Close the CEF File with an API call
    Call mCloseCEFFile
    Close #hmTo
    ilRet = gCustomEndStatus(lmEqtCode, 1, "")
    imExporting = False
    Print #hmMsg, "** Completed Export CnC Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Close #hmMsg
    lacResult.Caption = "Results: " & sMsgFileName
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Exit Sub
cmdExportErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Export CnC-cmdExport"
    ilRet = gCustomEndStatus(lmEqtCode, igExportReturn, "")
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    edcDate.Text = ""
    Unload frmExportCnCSpots
End Sub


Private Sub edcDate_Change()
    lbcMsg.Clear
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
End Sub

Private Sub Form_Activate()
    Dim llVef As Long
    Dim ilLoop As Integer
    Dim hlResult As Integer
    Dim slNowStart As String
    Dim slNowEnd As String
    Dim llEqtCode As Long
    
    If imFirstTime Then
        udcCriteria.Left = Label1.Left
        udcCriteria.Height = (7 * Me.Height) / 10
        udcCriteria.Width = (7 * Me.Width) / 10
        udcCriteria.Top = txtNumberDays.Top + txtNumberDays.Height  '(3 * edcDate.Height) / 4
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
            sgExportResultName = "CnCResultList.Txt"
            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
            gLogMsgWODT "W", hlResult, "CnC Result List, Started: " & slNowStart
            hgExportResult = hlResult
            cmdExport_Click
            slNowEnd = gNow()
            'Output result list box
'            sgExportResultName = "CnCResultList.Txt"
'            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
'            gLogMsgWODT "W", hlResult, "CnC Result List, Started: " & slNowStart
            If lbcMsg.ListCount > 0 Then
                For ilLoop = 0 To lbcMsg.ListCount - 1 Step 1
                    gLogMsgWODT "W", hlResult, Trim$(lbcMsg.List(ilLoop))
                Next ilLoop
            End If
            gLogMsgWODT "W", hlResult, "CnC Result List, Completed: " & slNowEnd
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
    frmExportCnCSpots.Caption = "Clearance n Compensation Spots - " & sgClientName
    smDate = gObtainPrevMonday(Format$(gNow(), sgShowDateForm))
    smDate = DateAdd("d", -7, smDate)
    edcDate.Text = smDate
    imNumberDays = 7
    txtNumberDays.Text = Trim$(Str$(imNumberDays))
    imAllClick = False
    imAllStationClick = False
    imTerminate = False
    imExporting = False
    imFirstTime = True
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    
    lbcStation.Clear
    mFillVehicle
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
    Set frmExportCnCSpots = Nothing
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

Private Function mExportSpots()
    Dim sDate As String
    Dim iNoWeeks As Integer
    Dim iLoop As Integer
    Dim iRet As Integer
    Dim sMoDate As String
    Dim sEndDate As String
    Dim slLogVefCode As String
    Dim slStationID As String
    Dim slLogDate As String
    Dim slSchDate As String
    Dim slAirTime As String
    Dim slDate As String
    Dim slTimeZone As String
    Dim slEventID As String
    Dim slBreakID As String
    Dim slPositionID As String
    Dim slStr As String
    Dim ilOkStation As Integer
    On Error GoTo ErrHand
    sMoDate = gObtainPrevMonday(smDate)
    sEndDate = DateAdd("d", imNumberDays - 1, smDate)
    
    'D.S. 11/21/05
'    iRet = gGetMaxAstCode()
'    If Not iRet Then
'        Exit Function
'    End If
    
    Do
        If igExportSource = 2 Then DoEvents
        'Get CPTT so that Stations requiring CP can be obtained
        SQLQuery = "SELECT shttTimeZone, shttStationID, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP"
        SQLQuery = SQLQuery + " FROM shtt, cptt, att"
        SQLQuery = SQLQuery + " WHERE (ShttCode = cpttShfCode"
        SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
        '10/29/14: Bypass Service agreements
        SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"
        SQLQuery = SQLQuery + " AND cpttVefCode = " & imVefCode
        SQLQuery = SQLQuery + " AND cpttPostingStatus = 2"
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
                If igExportSource = 2 Then DoEvents
                slStationID = Trim$(Str$(cprst!shttStationId))
                Do While Len(slStationID) < 8
                    slStationID = "0" & slStationID
                Loop
                slTimeZone = cprst!shttTimeZone
                Do While Len(slTimeZone) < 3
                    slTimeZone = slTimeZone & " "
                Loop
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
                iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, True, True)
                gFilterAstExtendedTypes tmAstInfo
                'Output AST
                For iLoop = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                    If igExportSource = 2 Then DoEvents
                    'If (DateValue(tmAstInfo(iLoop).sFeedDate) >= DateValue(smDate)) And (DateValue(tmAstInfo(iLoop).sFeedDate) <= DateValue(sEndDate)) And (tgStatusTypes(tmAstInfo(iLoop).iPledgeStatus).iPledged <> 2) Then
                    If (DateValue(gAdjYear(tmAstInfo(iLoop).sFeedDate)) >= DateValue(gAdjYear(smDate))) And (DateValue(gAdjYear(tmAstInfo(iLoop).sFeedDate)) <= DateValue(gAdjYear(sEndDate))) And (tgStatusTypes(gGetAirStatus(tmAstInfo(iLoop).iStatus)).iPledged <> 2) Then
                        SQLQuery = "SELECT lstLogVefCode, lstEvtIDCefCode, lstBreakNo, lstPositionNo, lstLogDate"
                        SQLQuery = SQLQuery & " FROM LST "
                        SQLQuery = SQLQuery + " WHERE lstCode =" & Str(tmAstInfo(iLoop).lLstCode)
                        Set rst = gSQLSelectCall(SQLQuery)
                        If Not rst.EOF Then
                            If igExportSource = 2 Then DoEvents
                            slLogVefCode = Trim$(Str$(rst!lstLogVefCode))
                            Do While Len(slLogVefCode) < 10
                                slLogVefCode = "0" & slLogVefCode
                            Loop
                            slLogDate = Format$(rst!lstLogDate, "mmddyyyy")
                            slSchDate = Format$(rst!lstLogDate, "mm/dd/yyyy")
                            If rst!lstEvtIDCefCode > 0 Then
                                iRet = mGetCefComment(rst!lstEvtIDCefCode, slEventID)
                            Else
                                slEventID = ""
                            End If
                            If Len(slEventID) > 12 Then
                                slEventID = Left$(slEventID, 12)
                            End If
                            Do While Len(slEventID) < 12
                                slEventID = slEventID & " "
                            Loop
                            slPositionID = Trim$(Str$(rst!lstPositionNo))
                            Do While Len(slPositionID) < 2
                                slPositionID = "0" & slPositionID
                            Loop
                            If igExportSource = 2 Then DoEvents
                            'net_nbr (network number)
                            slStr = slLogVefCode
                            'sta_nbr (station number)
                            slStr = slStr & slStationID
                            'Sched_start_dt (Week Start date- always monday)
                            slStr = slStr & Format(sMoDate, "mmddyyyy")
                            'sched_dt (scheduled date)
                            slStr = slStr & slLogDate 'Format$(tmAstInfo(iLoop).sAirDate, "mmddyyyy")
                            'feed_time-zn_cd (station local feed time zone code)
                            slStr = slStr & slTimeZone
                            'event_id
                            slStr = slStr & slEventID
                            'break_id
                            slStr = slStr & slBreakID
                            'position_id
                            slStr = slStr & slPositionID
                            'airplay_nbr
                            'for now use 1
                            slStr = slStr & "1"
                            'clrnc_cd (airing status)
                            If tgStatusTypes(gGetAirStatus(tmAstInfo(iLoop).iStatus)).iPledged = 0 Then
                                slStr = slStr & "01"
                            ElseIf tgStatusTypes(gGetAirStatus(tmAstInfo(iLoop).iStatus)).iPledged = 1 Then
                                slStr = slStr & "02"
                            ElseIf tgStatusTypes(gGetAirStatus(tmAstInfo(iLoop).iStatus)).iPledged = 2 Then
                                slStr = slStr & "04"
                            ElseIf tgStatusTypes(gGetAirStatus(tmAstInfo(iLoop).iStatus)).iPledged = 3 Then
                                slStr = slStr & "03"
                            End If
                            'cause_cd
                            If tgStatusTypes(gGetAirStatus(tmAstInfo(iLoop).iStatus)).iPledged <> 2 Then
                                slStr = slStr & "01"
                            Else
                                slStr = slStr & "06"
                            End If
                            'time_aired
                            slAirTime = Format$(tmAstInfo(iLoop).sAirTime, "hhmma/p")
                            If InStr(1, slAirTime, "a", vbTextCompare) > 0 Then
                                slAirTime = Left$(slAirTime, 4) & "1"
                            Else
                                slAirTime = Left$(slAirTime, 4) & "2"
                            End If
                            slStr = slStr & slAirTime
                            'date_aired if different then scheduled date
                            slDate = Format$(tmAstInfo(iLoop).sAirDate, "mm/dd/yyyy")
                            If DateValue(gAdjYear(slDate)) <> DateValue(gAdjYear(slSchDate)) Then
                                slDate = Format$(tmAstInfo(iLoop).sAirDate, "mmddyyyy")
                            Else
                                slDate = " "
                                Do While Len(slDate) < 8
                                    slDate = " " & slDate
                                Loop
                            End If
                            slStr = slStr & slDate
                            'line_aired_pledged
                            slStr = slStr & "N"
                            'aff_aired_pledged
                            slStr = slStr & " "
                            Print #hmTo, slStr
                            If igExportSource = 2 Then DoEvents
                        End If
                    End If
                Next iLoop
            End If
            cprst.MoveNext
        Wend
        If (lbcStation.ListCount = 0) Or (chkAllStation.Value = vbChecked) Or (lbcStation.ListCount = lbcStation.SelCount) Then
            gClearASTInfo True
        Else
            gClearASTInfo False
        End If
        sMoDate = DateAdd("d", 7, sMoDate)
    Loop While DateValue(gAdjYear(sMoDate)) < DateValue(gAdjYear(sEndDate))

    mExportSpots = True
    Exit Function
mExportSpotsErr:
    iRet = Err
    Resume Next

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Export CnC-mExportSpot"
    mExportSpots = False
    Exit Function
    
End Function

Private Sub mFillStations()
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
    SQLQuery = SQLQuery + " FROM shtt, att"
    SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode
    'SQLQuery = SQLQuery + " AND attExportType = 2 "
    SQLQuery = SQLQuery + " AND shttCode = attShfCode)"
    SQLQuery = SQLQuery + " ORDER BY shttCallLetters"
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
    gHandleError "AffErrorLog.txt", "Export CnC-mFileStations"
End Sub





Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload frmExportCnCSpots
End Sub

Private Sub txtNumberDays_Change()
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
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
        lmEqtCode = gCustomStartStatus("C", "Clearance and Compensation", "C", Trim$(edcDate.Text), Trim$(txtNumberDays.Text), ilVefCode(), ilShttCode())
    End If
End Sub
