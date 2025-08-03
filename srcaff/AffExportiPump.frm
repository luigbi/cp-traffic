VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form FrmExportiPump 
   Caption         =   "Export iPump"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   Icon            =   "AffExportiPump.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9975
   Begin VB.PictureBox pbcTextWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9105
      ScaleHeight     =   225
      ScaleWidth      =   1005
      TabIndex        =   17
      Top             =   4185
      Visible         =   0   'False
      Width           =   1035
   End
   Begin V81Affiliate.CSI_Calendar edcStartDate 
      Height          =   285
      Left            =   1515
      TabIndex        =   1
      Top             =   150
      Width           =   1380
      _ExtentX        =   2434
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
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin VB.TextBox edcTitle1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Vehicles"
      Top             =   1695
      Width           =   3810
   End
   Begin VB.TextBox edcTitle3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4275
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Stations"
      Top             =   1695
      Width           =   1635
   End
   Begin VB.ListBox lbcSort 
      Height          =   255
      Left            =   8205
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   5145
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7710
      Top             =   5040
   End
   Begin VB.TextBox txtNumberDays 
      Height          =   285
      Left            =   3990
      TabIndex        =   3
      Text            =   "1"
      Top             =   165
      Width           =   405
   End
   Begin VB.CheckBox chkAllStation 
      Caption         =   "All"
      Height          =   195
      Left            =   4215
      TabIndex        =   8
      Top             =   4695
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.ListBox lbcStation 
      Height          =   2595
      ItemData        =   "AffExportiPump.frx":08CA
      Left            =   4200
      List            =   "AffExportiPump.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1950
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4695
      Width           =   900
   End
   Begin VB.ListBox lbcMsg 
      Height          =   3960
      ItemData        =   "AffExportiPump.frx":08CE
      Left            =   6585
      List            =   "AffExportiPump.frx":08D0
      TabIndex        =   10
      Top             =   570
      Width           =   2820
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   2595
      ItemData        =   "AffExportiPump.frx":08D2
      Left            =   90
      List            =   "AffExportiPump.frx":08D4
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1950
      Width           =   3855
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   885
      Top             =   5610
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5940
      FormDesignWidth =   9975
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   3015
      TabIndex        =   11
      Top             =   5460
      Width           =   1665
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4980
      TabIndex        =   12
      Top             =   5460
      Width           =   1665
   End
   Begin V81Affiliate.AffExportCriteria udcCriteria 
      Height          =   375
      Left            =   -30
      TabIndex        =   4
      Top             =   1095
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1429
   End
   Begin VB.Label Label2 
      Caption         =   "# of Days"
      Height          =   255
      Left            =   3105
      TabIndex        =   2
      Top             =   210
      Width           =   795
   End
   Begin VB.Label lacResult 
      Height          =   405
      Left            =   120
      TabIndex        =   13
      Top             =   5055
      Width           =   9240
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   7035
      TabIndex        =   9
      Top             =   240
      Width           =   1965
   End
   Begin VB.Label lacStartDate 
      Caption         =   "Export Start Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   1395
   End
End
Attribute VB_Name = "FrmExportiPump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  Created April 12 by Dan Michaelson
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text
Private imGenerating As Integer
Private smDate As String      'Export Date
Private imNumberDays As Integer
Private imVefCode As Integer
Private imAdfCode As Integer
Private smVefName As String
Private imAllClick As Integer
Private imAllStationClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private imFirstTime As Integer
Private smExportDirectory As String
Private hmAst As Integer
Private cprst As ADODB.Recordset
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
'Private bmMgsPrevExisted As Boolean
Private hmCsf As Integer
'Dan M for writing messages in list box
Private lmMaxWidth As Long
Private Const myForm As String = "iPump"
Private Const FORMNAME As String = "FrmExportiPump"
Private Const FILEFACTS As String = "iPumpFacts"
Private Const FILEERROR As String = "iPumpExport"
Private Const MESSAGEBLACK As Long = 0
Private Const MESSAGERED As Long = 255
Private Const MESSAGEGREEN As Long = 39680
Private Const XMLDATE As String = "yyyy-mm-dd"
Private Const XMLTIME As String = "hh:mm:ss"
Private Const SORTORDER As String = "AirDate asc, Hour asc,Minute asc,Index asc"
'time zone subtract station from 'site' to get time difference. no saving is only for station
Private Const NOTIMEDIFFERENCE As Integer = 0
Private Const EASTERN As Integer = 8
Private Const CENTRAL As Integer = 7
Private Const MOUNTAIN As Integer = 6
Private Const PACIFIC As Integer = 5
Private Const ALASKAN As Integer = 4
Private Const HAWAIIAN As Integer = 3
Private Const TIMENONE  As String = "1/1/1970"
Private Const ZONENODAYLIGHTCHANGE As Integer = 0
Private Const ZONEDAYLIGHTATSTATION As Integer = 1
Private Const ZONEDAYLIGHTATTIME As Integer = 2
Private lmEqtCode As Long
'logging:
'Private myFacts As CLogger
Private smPathForgLogMsg As String
Private myErrors As CLogger
'To help debugging
Private bmWriteFacts As Boolean
Private rsTimeAdjust As Recordset
'6166
Private rsDeletePrevious As Recordset
'7/30/13
Private rsBreakNumbers As Recordset
'7458
Dim myEnt As CENThelper

' Template functions
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
Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    edcStartDate.Text = ""
    Unload Me
End Sub
Private Sub edcStartDate_Change()
    lbcMsg.Clear
    If cmdExport.Enabled = False Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
End Sub
Private Sub edcStartDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Load()
    mInit
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
Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.2
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts Me
    gCenterForm Me
    mAdjustForm

    If igExportSource = 2 Then
        Me.Top = -(2 * Me.Top + Screen.Height)
    End If
End Sub
Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload Me
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
        lmEqtCode = gCustomStartStatus("P", myForm, "P", Trim$(edcStartDate.Text), Trim$(txtNumberDays.Text), ilVefCode(), ilShttCode())
    End If
End Sub
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
Private Function mCloseCSF() As Integer
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
End Function

Private Sub txtNumberDays_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub mClearAlerts()
    Dim ilVef As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim llStartDate As Long
    Dim ilRet As Integer
    
    Dim slEDate As String
    Dim llSDate As Long
    Dim llEDate As Long
    
    
    slEDate = mEarliestEndDate(smDate, imNumberDays)
    llSDate = gDateValue(smDate)
    llEDate = gDateValue(slEDate)
    
    slDate = gObtainPrevMonday(Format(llSDate, "m/d/yy"))
    llStartDate = gDateValue(slDate)
    For ilVef = 0 To lbcVehicles.ListCount - 1
        If lbcVehicles.Selected(ilVef) Then
            imVefCode = lbcVehicles.ItemData(ilVef)
            For llDate = llStartDate To llEDate Step 7
                slDate = Format$(llDate, "m/d/yy")
                ilRet = gAlertClear("A", "F", "S", imVefCode, slDate)
                ilRet = gAlertClear("A", "R", "S", imVefCode, slDate)
            Next llDate
        End If
    Next ilVef
    ilRet = gAlertForceCheck()
End Sub
Private Sub cmdExport_Click()
    mCleanFolders
    mExport
End Sub
Private Sub mCleanFolders()
    If Not myErrors Is Nothing Then
        With myErrors
            .CleanThisFolder = messages
            .CleanFolder myForm
            If Len(.ErrorMessage) > 0 Then
                .WriteWarning "Couldn't delete old files from 'messages': " & .ErrorMessage
            End If
        End With
    End If
End Sub
Private Sub Form_Activate()
    Dim llVef As Long
    Dim ilLoop As Integer
    Dim hlResult As Integer
    Dim slNowStart As String
    Dim slNowEnd As String
    
    If imFirstTime Then
        udcCriteria.Left = lacStartDate.Left
        udcCriteria.Height = (7 * Me.Height) / 10
        udcCriteria.Width = (7 * Me.Width) / 10
        udcCriteria.Top = txtNumberDays.Top + (3 * txtNumberDays.Height / 2)
        udcCriteria.Action 6
        If UBound(tgEvtInfo) > 0 Then
            chkAll.Value = vbUnchecked
            lbcStation.Clear
            lbcVehicles.Clear
            For ilLoop = 0 To UBound(tgEvtInfo) - 1 Step 1
                llVef = gBinarySearchVef(CLng(tgEvtInfo(ilLoop).iVefCode))
                'added merge, which doesn't allow 'override' vehicles to appear
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
            edcStartDate.Text = sgExporStartDate
            txtNumberDays.Text = igExportDays
            igExportReturn = 1
            '6394 move before 'click'
             'Output result list box
            sgExportResultName = myForm & "ResultList.Txt"
            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
            gLogMsgWODT "W", hlResult, myForm & " Result List, Started: " & slNowStart
            ' pass global so glogMsg will write messages to sgExportResultName
            hgExportResult = hlResult
            cmdExport_Click
            slNowEnd = gNow()
'            'Output result list box
'            sgExportResultName = myForm & "ResultList.Txt"
'            gLogMsgWODT "O", hlResult, sgMsgDirectory & sgExportResultName
'            gLogMsgWODT "W", hlResult, myForm & " Result List, Started: " & slNowStart
            If lbcMsg.ListCount > 0 Then
                For ilLoop = 0 To lbcMsg.ListCount - 1 Step 1
                    gLogMsgWODT "W", hlResult, Trim$(lbcMsg.List(ilLoop))
                Next ilLoop
            End If
            gLogMsgWODT "W", hlResult, myForm & " Result List, Completed: " & slNowEnd
            gLogMsgWODT "C", hlResult, ""
            '6394 clear values
            hgExportResult = 0
            imTerminate = True
            tmcTerminate.Enabled = True
        End If
        imFirstTime = False
    End If
End Sub
Private Sub mSetResults(slMsg As String, llFGC As Long)
    Dim llLoop As Long
    gAddMsgToListBox Me, lmMaxWidth, slMsg, lbcMsg
    lbcMsg.ListIndex = lbcMsg.ListCount - 1
    If lbcMsg.ForeColor <> MESSAGERED Then
        lbcMsg.ForeColor = llFGC
    End If
    If igExportSource = 2 Then DoEvents
End Sub
Private Function mVehicleName(iVefCode As Integer) As String
    Dim llLoop As Integer
    
    mVehicleName = ""
    For llLoop = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        If tgVehicleInfo(llLoop).iCode = iVefCode Then
            mVehicleName = Trim(tgVehicleInfo(llLoop).sVehicle)
            Exit For
        End If
    Next
End Function
Private Function mStationName(iShttCode As Integer) As String
    Dim llLoop As Integer
    
    mStationName = ""
    For llLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
        If igExportSource = 2 Then DoEvents
        If tgStationInfo(llLoop).iCode = iShttCode Then
            mStationName = Trim(tgStationInfo(llLoop).sCallLetters)
            Exit For
        End If
    Next
End Function
Private Sub mLogEndDate()
    Dim ilVef As Integer
    Dim ilVefCode As Integer
    Dim slEDate As String

    For ilVef = 0 To lbcVehicles.ListCount - 1
        ilVefCode = lbcVehicles.ItemData(ilVef)
        slEDate = mEarliestEndDate(smDate, imNumberDays)
        gUpdateLastExportDate ilVefCode, slEDate
    Next ilVef

End Sub
Private Function mEarliestEndDate(ByVal slStartDate As String, ilNumberDays As Integer) As String
    Dim slEndChosen As String
    Dim slEndOfWeek As String
    
    slEndChosen = DateAdd("d", ilNumberDays - 1, slStartDate)
    slEndOfWeek = gObtainNextSunday(slStartDate)
    If gDateValue(gAdjYear(slEndChosen)) < gDateValue(gAdjYear(slEndOfWeek)) Then
         mEarliestEndDate = slEndChosen
    Else
        mEarliestEndDate = slEndOfWeek
    End If
End Function
Private Sub mAdjustForm()
    Dim llLeft As Long
    
    llLeft = lbcVehicles.Left
    lacStartDate.Left = llLeft
    chkAll.Left = llLeft
    udcCriteria.Left = llLeft
End Sub
'End Template functions
'Template function that need minor changes

Private Sub mFillStations()
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
    SQLQuery = SQLQuery & " FROM shtt, att"
    SQLQuery = SQLQuery & " WHERE (attVefCode = " & imVefCode
    SQLQuery = SQLQuery & " AND shttCode = attShfCode) AND RTrim(shttIPumpID) <> ''"
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
    gHandleError smPathForgLogMsg, FORMNAME & "-mFillStations"

End Sub
Private Sub mFillVehicle()
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim llVef As Long
    Dim ilVff As Integer
    Dim ilVehicleCode As Integer
    
    On Error GoTo ErrHand
    lbcVehicles.Clear
    chkAll.Value = vbUnchecked
    slNowDate = Format(gNow(), sgSQLDateForm)
    'all vehicles
    SQLQuery = "SELECT DISTINCT attVefCode FROM att WHERE attDropDate > '" & slNowDate & "' AND attOffAir > '" & slNowDate & "'"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        llVef = gBinarySearchVef(CLng(rst!attvefCode))
        If llVef <> -1 Then
            ilVehicleCode = tgVehicleInfo(llVef).iCode
            ilVff = gBinarySearchVff(ilVehicleCode)
            If ilVff <> -1 Then
                'block non ipump and overridden vehicles
                If Trim$(tgVffInfo(ilVff).sExportIPump) = "Y" And Len(Trim$(tgVffInfo(ilVff).sIPumpEventTypeOV)) = 0 Then
                    lbcVehicles.AddItem Trim$(tgVehicleInfo(llVef).sVehicle)
                    lbcVehicles.ItemData(lbcVehicles.NewIndex) = ilVehicleCode
                End If
            End If
        End If
        rst.MoveNext
    Loop
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, FORMNAME & "-mFillVehicle"
    Resume Next
    Exit Sub
IndexErr:
    ilRet = 1
    Resume Next
End Sub
Private Sub mInit()
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    lmMaxWidth = lbcMsg.Width
    imTerminate = False
    imFirstTime = True
    Me.Caption = "Export Wegener " & myForm & " - " & sgClientName
    Set myErrors = New CLogger
    myErrors.LogPath = myErrors.CreateLogName(sgMsgDirectory & FILEERROR)
    smPathForgLogMsg = FILEERROR & "Log_" & Format(gNow(), "mm-dd-yy") & ".txt"
    smDate = gObtainNextMonday(Format$(gNow(), sgShowDateForm))
    edcStartDate.Text = smDate
    txtNumberDays.Text = 7
    imAllClick = False
    imAllStationClick = False
    imTerminate = False
    imExporting = False
    smExportDirectory = mExportDirectory()
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    If Not ilRet Then
        imTerminate = True
    End If
    ilRet = mOpenCSF()
    If Not ilRet Then
        imTerminate = True
    End If
    lbcStation.Clear
    mFillVehicle
    chkAll.Value = vbChecked
    If lbcVehicles.ListCount = 1 Then
        imVefCode = lbcVehicles.ItemData(0)
        edcTitle3.Visible = True
        chkAllStation.Visible = True
        lbcStation.Visible = True
        mFillStations
    End If
    ilRet = gPopAvailNames()
    If Not ilRet Then
        imTerminate = True
    End If
    If imTerminate Then
        tmcTerminate.Enabled = True
    End If
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
    ilRet = mCloseCSF()
    Erase tmCPDat
    Erase tmAstInfo
    If Not cprst Is Nothing Then
        If (cprst.State And adStateOpen) <> 0 Then
            cprst.Close
        End If
        Set cprst = Nothing
    End If
    rsTimeAdjust.Close
    rsDeletePrevious.Close
    rsBreakNumbers.Close
   ' Set myFacts = Nothing
    Set myErrors = Nothing
    Set FrmExportiPump = Nothing
End Sub
'End Minor Changes
'Common Functions
Private Function mPrepRecordset() As ADODB.Recordset
    Dim myRs As ADODB.Recordset

    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "EventType", adChar, 2
            .Append "NetworkID", adChar, 2
            .Append "AirDate", adDate
            .Append "Minute", adChar, 2
            .Append "Hour", adChar, 2
            .Append "NameSpace", adChar, 60
            .Append "FileName", adChar, 14
            .Append "StationID", adChar, 10
            .Append "AstCode", adInteger
            .Append "Found", adBoolean
            .Append "Index", adInteger
            'for writing debug file
            .Append "Adv", adInteger
            .Append "ISCI", adChar, 40
            .Append "Vehicle", adChar, 40
            'for deletions
            .Append "ShttCode", adInteger
            .Append "AttCode", adInteger
            .Append "VefCode", adInteger
            .Append "ZoneAdjust", adInteger
            .Append "DaylightAdjust", adInteger
        End With
    myRs.Open
    myRs!Hour.Properties("optimize") = True
    myRs.Sort = SORTORDER
    Set mPrepRecordset = myRs
End Function
Private Function mFileNameFilter(slInName As String) As String
    Dim slName As String
    Dim ilPos As Integer
    Dim ilFound As Integer
    
    slName = Trim$(slInName)
    'Remove " and '
    Do
        If igExportSource = 2 Then DoEvents
        ilFound = False
        ilPos = InStr(1, slName, "'", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        If igExportSource = 2 Then DoEvents
        ilFound = False
        ilPos = InStr(1, slName, """", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        If igExportSource = 2 Then DoEvents
        ilFound = False
        ilPos = InStr(1, slName, "&", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "/", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "\", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "*", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ":", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "?", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "%", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "=", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "+", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "<", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ">", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "|", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ";", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "@", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "[", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "]", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "{", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "}", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, "^", 1)
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "-"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ".", 1)    'If period, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, ",", 1)    'If comma, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
        ilPos = InStr(1, slName, " ", 1)    'If space, use underscore
        If ilPos > 0 Then
            Mid$(slName, ilPos, 1) = "_"
            ilFound = True
        End If
    Loop While ilFound
    mFileNameFilter = slName
End Function
Private Function mStartFresh() As Boolean
    Dim blRet As Boolean
    Dim slNowDate As String
    Dim ilVef As Integer
    Dim blVehicleSelected As Boolean
    
    blRet = True
    imTerminate = False
    lbcMsg.Clear
    lbcMsg.ForeColor = RGB(0, 0, 0)
    lmMaxWidth = 0
    gClearListScrollBar lbcMsg
    lacResult.Caption = ""
    If Not mNumberDays() Then
        blRet = False
        GoTo Cleanup
    End If
    slNowDate = Format$(gNow(), "m/d/yy")
    'If gDateValue(gAdjYear(smDate)) <= gDateValue(gAdjYear(slNowDate)) Then
    If gDateValue(gAdjYear(smDate)) < gDateValue(gAdjYear(slNowDate)) Then
        Beep
        'gMsgBox "Date must be after today's date " & slNowDate, vbCritical
        gMsgBox "Date must be today's date or greater: " & slNowDate, vbCritical
        'Dan 8/30/13 Dick said remove these as not good for auto exporting
       ' edcStartDate.SetFocus
        blRet = False
        GoTo Cleanup
    End If
    blVehicleSelected = False
    For ilVef = 0 To lbcVehicles.ListCount - 1 Step 1
        If lbcVehicles.Selected(ilVef) Then
            blVehicleSelected = True
            Exit For
        End If
    Next ilVef
    If (Not blVehicleSelected) Then
        Beep
        gMsgBox "Vehicle must be selected.", vbCritical
        blRet = False
        GoTo Cleanup
    End If
    Screen.MousePointer = vbHourglass
    imExporting = True
    mSaveCustomValues
Cleanup:
    mStartFresh = blRet
End Function
Private Function mNumberDays() As Boolean
    mNumberDays = False
    If edcStartDate.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        'Dan 8/30/13 Dick said remove these as not good for auto exporting
        'edcStartDate.SetFocus
        Exit Function
    End If
    If gIsDate(edcStartDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        'Dan 8/30/13 Dick said remove these as not good for auto exporting
        'edcStartDate.SetFocus
        Exit Function
    Else
        smDate = Format(edcStartDate.Text, sgShowDateForm)
    End If
    imNumberDays = Val(txtNumberDays.Text)
    If imNumberDays <= 0 Then
        gMsgBox "Number of days must be specified.", vbOKOnly
        'Dan 8/30/13 Dick said remove these as not good for auto exporting
        'txtNumberDays.SetFocus
        Exit Function
    End If
    Select Case Weekday(gAdjYear(smDate))
        Case vbMonday
            If imNumberDays > 7 Then
                gMsgBox "Number of days can not exceed 7.", vbOKOnly
            'Dan 8/30/13 Dick said remove these as not good for auto exporting
                'txtNumberDays.SetFocus
                Exit Function
            End If
        Case vbTuesday
            If imNumberDays > 6 Then
                gMsgBox "Number of days can not exceed 6.", vbOKOnly
                'Dan 8/30/13 Dick said remove these as not good for auto exporting
                'txtNumberDays.SetFocus
                Exit Function
            End If
        Case vbWednesday
            If imNumberDays > 5 Then
                gMsgBox "Number of days can not exceed 5.", vbOKOnly
                'Dan 8/30/13 Dick said remove these as not good for auto exporting
                'txtNumberDays.SetFocus
                Exit Function
            End If
        Case vbThursday
            If imNumberDays > 4 Then
                gMsgBox "Number of days can not exceed 4.", vbOKOnly
                'Dan 8/30/13 Dick said remove these as not good for auto exporting
                'txtNumberDays.SetFocus
                Exit Function
            End If
        Case vbFriday
            If imNumberDays > 3 Then
                gMsgBox "Number of days can not exceed 3.", vbOKOnly
                'Dan 8/30/13 Dick said remove these as not good for auto exporting
                'txtNumberDays.SetFocus
                Exit Function
           End If
        Case vbSaturday
            If imNumberDays > 2 Then
                gMsgBox "Number of days can not exceed 2.", vbOKOnly
                'Dan 8/30/13 Dick said remove these as not good for auto exporting
                'txtNumberDays.SetFocus
                Exit Function
            End If
        Case vbSunday
            If imNumberDays > 1 Then
                gMsgBox "Number of days can not exceed 1.", vbOKOnly
                'Dan 8/30/13 Dick said remove these as not good for auto exporting
                'txtNumberDays.SetFocus
                Exit Function
            End If
    End Select
    mNumberDays = True
End Function
Private Function mLoseLastLetter(slInput As String) As String
    Dim llLength As Long
    Dim slNewString As String

    llLength = Len(slInput)
    If llLength > 0 Then
        slNewString = Mid(slInput, 1, llLength - 1)
    End If
    mLoseLastLetter = slNewString
End Function
Private Function mExportDirectory() As String
    Dim slExportDir As String
    
    slExportDir = sgExportDirectory
    slExportDir = gSetPathEndSlash(slExportDir, False)
    mExportDirectory = slExportDir
End Function
'End Common functions
Private Sub mExport()
    Dim rsIPump As ADODB.Recordset
    Dim blAtLeastOneExport As Boolean
    Dim slLogInfo As String
    
On Error GoTo ErrHand
    If Not mStartFresh() Then
        Exit Sub
    End If
    '6393
    If udcCriteria.iPOutput(0) = vbChecked Then
    'If ckcFacts.Value = vbChecked Then
        bmWriteFacts = True
    Else
        bmWriteFacts = False
    End If
    mSaveCustomValues
    If Not gPopCopy(smDate, "Export " & myForm) Then
        igExportReturn = 2
        gCustomEndStatus lmEqtCode, igExportReturn, ""
        imExporting = False
        Exit Sub
    End If
    Set rsIPump = mPrepRecordset()
    If rsIPump Is Nothing Then
        Beep
        gMsgBox "Error creating recordset in mExport", vbCritical
        myErrors.WriteError "Error creating recordset in mExport", False, False
        mSetResults "Error creating recordest-export halted", MESSAGERED
        Exit Sub
    End If
    slLogInfo = "Date of: " & smDate & " For: " & CStr(imNumberDays) & " Days."
    myErrors.WriteFacts slLogInfo, True
    '7458
    Set myEnt = New CENThelper
    With myEnt
        .TypeEnt = Exportunposted3rdparty
        .ThirdParty = Vendors.Wegener_IPump
        .ErrorLog = smPathForgLogMsg
        .User = igUstCode
        .ReturnEntCode = True
    End With
    '8689
    bgTaskBlocked = False
    sgTaskBlockedName = "iPump Export"
    If mGatherSpots(rsIPump) Then
        lacResult.Caption = "Writing info to file"
        If Not mWriteSpots(rsIPump, blAtLeastOneExport) Then
            mSetResults "Export Failed", MESSAGERED
            myErrors.WriteWarning "Export Failed", True
            gCustomEndStatus lmEqtCode, igExportReturn, ""
            GoTo Cleanup
        Else
             '6166 write out files even if no spots!
            mDeleteUnwrittenSpots
            If bmWriteFacts Then
                mFactFile rsIPump
            End If
        End If
    Else
        lacResult.Caption = ""
        mSetResults "Export Failed", MESSAGERED
        myErrors.WriteWarning "Export Failed", True
        gCustomEndStatus lmEqtCode, igExportReturn, ""
        GoTo Cleanup
    End If
    gCloseRegionSQLRst
    mFinishUp
    If imTerminate Then
        mSetResults "** User Terminated **", MESSAGERED
        myErrors.WriteFacts "*** User Terminated **", True
        gCustomEndStatus lmEqtCode, igExportReturn, ""
        GoTo Cleanup
    End If
    '8/1/13: Moved to after alerts cleared
    'cmdCancel.Caption = "&Done"
    If blAtLeastOneExport Then
        If lbcMsg.ForeColor = MESSAGERED Then
            myErrors.WriteWarning "Some exports were not successful."
            mSetResults "Some exports were not successful. See " & myErrors.LogPath, MESSAGERED
            gCustomEndStatus lmEqtCode, 2, ""
        Else
            myErrors.WriteFacts "Export completed successfully.", False
            mSetResults "Export Completed Successfully", MESSAGEGREEN
            mLogEndDate
            mClearAlerts
            gCustomEndStatus lmEqtCode, 1, ""
        End If
        lacResult.Caption = "Exports placed into: " & sgExportDirectory
    Else
        myErrors.WriteFacts "No spots to export."
        mSetResults "No spots to Export.", MESSAGEBLACK
        lacResult.Caption = ""
        gCustomEndStatus lmEqtCode, 1, ""
    End If
    '8689
    If bgTaskBlocked And igExportSource <> 2 Then
         mSetResults "Some spots were blocked during export.", MESSAGERED
         gMsgBox "Some spots were blocked during the export." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
         myErrors.WriteWarning "Some spots were blocked during export.", True
         lacResult.Caption = "Please refer to the Messages folder for file TaskBlocked_" & sgTaskBlockedDate & ".txt."
    End If
    cmdCancel.Caption = "&Done"
Cleanup:
    sgTaskBlockedName = ""
    bgTaskBlocked = False
    If Not rsIPump Is Nothing Then
        If (rsIPump.State And adStateOpen) <> 0 Then
            rsIPump.Close
        End If
        Set rsIPump = Nothing
    End If
    If Not cprst Is Nothing Then
        If (cprst.State And adStateOpen) <> 0 Then
            cprst.Close
        End If
    End If
    If Not rsTimeAdjust Is Nothing Then
        If (rsTimeAdjust.State And adStateOpen) <> 0 Then
            rsTimeAdjust.Close
        End If
        Set rsTimeAdjust = Nothing
    End If
    If Not rsDeletePrevious Is Nothing Then
        If (rsDeletePrevious.State And adStateOpen) <> 0 Then
            rsDeletePrevious.Close
        End If
        Set rsDeletePrevious = Nothing
    End If
    '7458
    Set myEnt = Nothing
    imExporting = False
    cmdExport.Enabled = False
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, FORMNAME & "-mExport"
    gCustomEndStatus lmEqtCode, igExportReturn, ""
    GoTo Cleanup
End Sub
Private Sub mCpttInfo()
    Dim ilVef As Integer
    Dim slDate As String
    
    'just 1 vehicle chosen?
    slDate = gObtainPrevMonday(smDate)
    slDate = Format$(slDate, sgSQLDateForm)
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
    If igExportSource = 2 Then DoEvents
    SQLQuery = "SELECT  shttTimeZone, shttCode,  shttIpumpid as iPumpID, ShttackDaylight as Daylight, shttTztCode as TimeZone, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attTimeType"
    '7701
    SQLQuery = SQLQuery & " FROM shtt, cptt,vef_Vehicles, att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode"
    SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
    SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
    '10/29/14: Bypass Service agreements
    SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"
    SQLQuery = SQLQuery & " AND vefCode = cpttVefCode"
    SQLQuery = SQLQuery & " AND RTrim(shttIpumpId) <> ''"
    '7701
    SQLQuery = SQLQuery & " AND vatWvtVendorId  = " & Vendors.Wegener_IPump
    'SQLQuery = SQLQuery & " AND attAudioDelivery = " & "'P'"
    If imVefCode > 0 Then
        SQLQuery = SQLQuery & " AND cpttVefCode = " & imVefCode
    End If
    SQLQuery = SQLQuery & " AND cpttStartDate = '" & slDate & "') AND "
    'merge stuff. exclude non ipump and overrides
    SQLQuery = SQLQuery & " vefcode in (select vffvefcode from vff_vehicle_Features where length(vffIpumpEventTypeOv) = 0 and vffExportIPump = 'Y') "
    SQLQuery = SQLQuery & " ORDER BY vefName, shttCallLetters, shttCode"
    Set cprst = gSQLSelectCall(SQLQuery)

End Sub
Private Sub mAstInfo()
    'not typical!  for 'time zone', get the day before and after what is requested. Change igTimes
    Dim ilVpf As Integer
    Dim slNewDate As String
    
    slNewDate = DateAdd("d", -1, smDate)
    ilVpf = gBinarySearchVpf(CLng(imVefCode))
    If ilVpf <> -1 Then
        ReDim tgCPPosting(0 To 1) As CPPOSTING
        tgCPPosting(0).lCpttCode = cprst!cpttCode
        tgCPPosting(0).iStatus = cprst!cpttStatus
        tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
        tgCPPosting(0).lAttCode = cprst!cpttatfCode
        tgCPPosting(0).iAttTimeType = cprst!attTimeType
        tgCPPosting(0).iVefCode = imVefCode
        tgCPPosting(0).iShttCode = cprst!shttCode
        tgCPPosting(0).sZone = cprst!shttTimeZone
        tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
        tgCPPosting(0).sDate = Format$(slNewDate, sgShowDateForm)
        tgCPPosting(0).iNumberDays = imNumberDays + 2
        'Create AST records
        igTimes = 3 'not By Week
        imAdfCode = -1
        If igExportSource = 2 Then DoEvents
        gGetAstInfo hmAst, tmCPDat(), tmAstInfo(), imAdfCode, True, True, True
        gFilterAstExtendedTypes tmAstInfo
    End If

End Sub
Private Function mStationOk() As Boolean
    Dim blOkStation As Boolean
    Dim ilLoop As Integer
    
    If lbcStation.ListCount > 0 Then
        blOkStation = False
        For ilLoop = 0 To lbcStation.ListCount - 1 Step 1
            If igExportSource = 2 Then DoEvents
            If lbcStation.Selected(ilLoop) Then
                If lbcStation.ItemData(ilLoop) = cprst!shttCode Then
                    blOkStation = True
                    Exit For
                End If
            End If
        Next ilLoop
    Else
        blOkStation = True
    End If
    mStationOk = blOkStation
End Function
Private Function mVehicleOk() As Boolean
    Dim ilVef As Integer
    Dim blOkVehicle As Boolean
    
    blOkVehicle = False
    For ilVef = 0 To lbcVehicles.ListCount - 1
        If igExportSource = 2 Then DoEvents
        If lbcVehicles.Selected(ilVef) Then
            If lbcVehicles.ItemData(ilVef) = cprst!cpttvefcode Then
                imVefCode = lbcVehicles.ItemData(ilVef)
                blOkVehicle = True
                Exit For
            End If
        End If
    Next ilVef
    mVehicleOk = blOkVehicle
End Function
Private Function mGatherSpots(rsIPump As ADODB.Recordset) As Boolean
    Dim blRet As Boolean
    Dim ilLoop As Integer
    Dim blOkStation As Boolean
    Dim blOkVehicle As Boolean
    Dim ilIndex As Integer
    Dim ilVpf As Integer
    Dim slVehicleName As String
    Dim slStationName As String
    Dim slExportEnd As String
    Dim slFileName As String
    Dim slEventType As String
    Dim slNetworkId As String
    Dim slNameSpace As String
    'merging
    Dim blOnMerge As Boolean
    Dim slstations() As String
    'time zone
    Dim ilSiteZone As Integer
    Dim ilZoneAdjust As Integer
    Dim dlStartSavingTime As Date
    Dim dlEndSavingTime As Date
    Dim ilDaylightAdjust As Integer
    Dim slAdjustedDate As String
    Dim llAdjustedFeedDate As Long
    Dim slAdjustedHour As String
    Dim slAdjustedMinute As String
    Dim slEventTypeOverride As String
    Dim blErrorShown As Boolean
    Dim ilPass As Integer
    '6693
    Dim ilDaylightAdjustToPass As Integer
    
    On Error GoTo ErrHand
    blRet = True
    blOnMerge = False
    imExporting = True
    blErrorShown = False
    imVefCode = 0
    If bmWriteFacts Then
        Set rsTimeAdjust = mPrepRecordsetTime()
    End If
    Set rsDeletePrevious = mPrepRecordsetDelete()
    ReDim slstations(0 To 0)
    slExportEnd = mEarliestEndDate(smDate, imNumberDays)
    ilSiteZone = mZoneGetSite()
    If ilSiteZone > 0 Then
        ' get daylight savings times, and see if it matters for this export
        ilDaylightAdjust = mZoneForSavings(dlStartSavingTime, dlEndSavingTime, slExportEnd)
    Else
        ilDaylightAdjust = ZONENODAYLIGHTCHANGE
    End If
    'fills cprst
    mCpttInfo
    '6329, 6319.  Need to get override spots, even if no spots in playlist.
    'first pass, normal vehicles.  2nd pass, overrides.
    For ilPass = 1 To 2 Step 1
        While Not cprst.EOF
            If igExportSource = 2 Then DoEvents
            If blOnMerge Then
                blOkStation = True
                blOkVehicle = True
                imVefCode = cprst!cpttvefcode
            Else
                blOkStation = mStationOk()
                If blOkStation Then
                    'also sets imVefCode
                    blOkVehicle = mVehicleOk()
                End If
            End If
            If blOkStation And blOkVehicle Then
                On Error GoTo ErrHand
                DoEvents
                '6693  daylight only matters if station doesn't honor daylight. If it honors daylight, set to NoDaylightChange
                ilDaylightAdjustToPass = ilDaylightAdjust
                If cprst!DAYLIGHT <> 1 Then
                    ilDaylightAdjustToPass = ZONENODAYLIGHTCHANGE
                End If
                If Not rsDeletePrevious Is Nothing Then
                    rsDeletePrevious.AddNew Array("StationId", "VehicleID", "Override", "Delete"), Array(cprst!shttCode, imVefCode, "", False)
                End If
                mAstInfo
                ilIndex = LBound(tmAstInfo)
                slVehicleName = mVehicleName(cprst!cpttvefcode)
                slStationName = mStationName(cprst!shttCode)
                '6320 tmastInfo empty?  not an error! Continue with next contract
                If ilIndex = UBound(tmAstInfo) Then
                    '7/28/13: Only show the message once in result area
                    If Not blErrorShown Then
                        mSetResults "Spots Missing, see iPumpErrorLog", RGB(255, 0, 0)
                        blErrorShown = True
                    End If
                    myErrors.WriteWarning "Spots missing for: " & slStationName & " " & slVehicleName
                Else
                    ilZoneAdjust = NOTIMEDIFFERENCE
                    If ilSiteZone > NOTIMEDIFFERENCE Then
                        '6693
                        ilZoneAdjust = mZoneGetAdjustForStation(ilSiteZone, ilDaylightAdjustToPass)
                        'ilZoneAdjust = mZoneGetAdjustForStation(ilSiteZone, ilDaylightAdjust)
                    End If
                    lacResult.Caption = "Exporting " & slStationName & ", " & slVehicleName
                    'get vehicle info here to pass to mCifInfo
                    slEventTypeOverride = mGetEventOverride(tmAstInfo(ilIndex).iVefCode)
                    'for 'merge' vehicles. What stations are going to allow?
                    mAddStation slstations, tmAstInfo(ilIndex).iShttCode
                    If bmWriteFacts Then
                        If Not rsTimeAdjust Is Nothing Then
                            '6693 changed
                            rsTimeAdjust.AddNew Array("StationID", "Vehicle", "StationName", "AdjustTime", "daylightAdjust"), Array(cprst!IPUMPID, slVehicleName, slStationName, ilZoneAdjust, ilDaylightAdjustToPass)
                          '  rsTimeAdjust.AddNew Array("StationID", "Vehicle", "StationName", "AdjustTime", "daylightAdjust"), Array(cprst!IPUMPID, slVehicleName, slStationName, ilZoneAdjust, ilDaylightAdjust)
                        End If
                    End If
                End If
                '7458
                With myEnt
                    .Vehicle = imVefCode
                    .Station = cprst!shttCode
                    .Agreement = cprst!cpttatfCode
                    'copied from mCreateFile
                    .fileName = mFileNameFilter(cprst!IPUMPID) & "_" & Format(smDate, "yyyymmdd") & ".weg"
                    .ProcessStart
                End With
                'loop all spots
                Do While ilIndex < UBound(tmAstInfo)
                    DoEvents
                    With tmAstInfo(ilIndex)
                        If igExportSource = 2 Then DoEvents
                        '7458
                        If Not myEnt.Add(tmAstInfo(ilIndex).sFeedDate, tmAstInfo(ilIndex).lgsfCode, Asts) Then
                            myErrors.WriteWarning myEnt.ErrorMessage
                        End If
                        ' 2 means don't air!
                        If (tgStatusTypes(gGetAirStatus(.iStatus)).iPledged <> 2) Then
                            If Len(.sFeedDate) <> 0 And Len(.sFeedTime) <> 0 Then
                                '6693
                                slAdjustedDate = mZoneAdjustTime(.sFeedDate, .sFeedTime, ilZoneAdjust, ilDaylightAdjustToPass, dlStartSavingTime, dlEndSavingTime)
                                'slAdjustedDate = mZoneAdjustTime(.sFeedDate, .sFeedTime, ilZoneAdjust, ilDaylightAdjust, dlStartSavingTime, dlEndSavingTime)
                                llAdjustedFeedDate = gDateValue(gAdjYear(slAdjustedDate))
                                If llAdjustedFeedDate >= gDateValue(gAdjYear(smDate)) And llAdjustedFeedDate <= gDateValue(gAdjYear(slExportEnd)) Then
                                    If mCifInfo(ilIndex, slFileName, slEventType, slNetworkId, slNameSpace) Then
                                        If Len(slEventTypeOverride) > 0 Then
                                            slEventType = slEventTypeOverride
                                        End If
                                        slAdjustedHour = Format(slAdjustedDate, "hh")
                                        slAdjustedMinute = Format(slAdjustedDate, "nn")
                                        slAdjustedDate = Format(slAdjustedDate, sgShowDateForm)
                                        '6325 for break fix, dick now passes time zone adjustment here.  May not need to, but was quickest at time.
                                        '6693
                                        rsIPump.AddNew Array("astCode", "NetworkID", "NameSpace", "StationID", "EventType", "FileName", "AirDate", "Hour", "Minute", "Index", "Adv", "ISCI", "Vehicle", "ShttCode", "AttCode", "VefCode", "ZoneAdjust", "DaylightAdjust"), Array(tmAstInfo(ilIndex).lCode, slNetworkId, slNameSpace, cprst!IPUMPID, slEventType, slFileName, slAdjustedDate, slAdjustedHour, slAdjustedMinute, ilIndex, tmAstInfo(ilIndex).iAdfCode, tmAstInfo(ilIndex).sISCI, slVehicleName, cprst!shttCode, cprst!cpttatfCode, cprst!cpttvefcode, ilZoneAdjust, ilDaylightAdjustToPass)
                                        'rsIPump.AddNew Array("astCode", "NetworkID", "NameSpace", "StationID", "EventType", "FileName", "AirDate", "Hour", "Minute", "Index", "Adv", "ISCI", "Vehicle", "ShttCode", "AttCode", "VefCode", "ZoneAdjust", "DaylightAdjust"), Array(tmAstInfo(ilIndex).lCode, slNetworkId, slNameSpace, cprst!IPUMPID, slEventType, slFileName, slAdjustedDate, slAdjustedHour, slAdjustedMinute, ilIndex, tmAstInfo(ilIndex).iAdfCode, tmAstInfo(ilIndex).sISCI, slVehicleName, cprst!shttCode, cprst!cpttatfCode, cprst!cpttVefCode, ilZoneAdjust, ilDaylightAdjust)
                                        '7458
                                        If Not myEnt.Add(tmAstInfo(ilIndex).sFeedDate, tmAstInfo(ilIndex).lgsfCode, SentOrReceived, , True) Then
                                            myErrors.WriteWarning myEnt.ErrorMessage
                                        End If
                                    Else
                                        myErrors.WriteWarning "Could not create spot info for astcode #" & tmAstInfo(ilIndex).lCode & " --File name = " & slFileName _
                                        & " NetworkID = " & slNetworkId & " EventType = " & slEventType & " NameSpace = " & slNameSpace
                                        mSetResults "Not all spots could be exported. See warnings.", MESSAGERED
                                    End If
                                End If 'date range ok
                            Else
                                myErrors.WriteWarning "missing spot info for astcode #" & tmAstInfo(ilIndex).lCode & " --AirTime = " & .sFeedTime _
                                & " AirDate = " & .sFeedDate
                                mSetResults "Not all spots could be exported. See warnings.", MESSAGERED
                            End If  'date time exist
                        End If  'pledge status
                    End With
                    ilIndex = ilIndex + 1
                    If imTerminate Then
                        imExporting = False
                        Exit Function
                    End If
                Loop 'each spot
                If igExportSource = 2 Then DoEvents
                '7458
                If Not myEnt.CreateEnts(EntError) Then
                    myErrors.WriteWarning myEnt.ErrorMessage
                End If
            End If  'station and vehicle
            cprst.MoveNext
        Wend
        If cprst.EOF And Not blOnMerge Then
            blOnMerge = True
            mAddMissingStations slstations
            mAddMergeVehicles slstations
        End If
    Next ilPass
    If imTerminate Then
        imExporting = False
        blRet = False
        GoTo Cleanup
    End If
    blRet = True
Cleanup:
    mGatherSpots = blRet
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, FORMNAME & "-mGatherSpots"
    blRet = False
    GoTo Cleanup
End Function
Private Function mAddMissingStations(slstations() As String)
    Dim ilVef As Integer
    Dim slDate As String
    Dim blAddStation As Boolean
    
    'just 1 vehicle chosen?
    slDate = gObtainPrevMonday(smDate)
    slDate = Format$(slDate, sgSQLDateForm)
    imVefCode = 0
    '8/1/13: If only one vehicle and no spots exist for that vehicle, then the SQL call will not find stations with ID only
    'For ilVef = 0 To lbcVehicles.ListCount - 1
    '    If igExportSource = 2 Then DoEvents
    '    If lbcVehicles.Selected(ilVef) Then
    '        If imVefCode = 0 Then
    '            imVefCode = lbcVehicles.ItemData(ilVef)
    '        Else
    '            imVefCode = -1
    '            Exit For
    '        End If
    '    End If
    'Next ilVef
    If igExportSource = 2 Then DoEvents
    SQLQuery = "SELECT  shttTimeZone, shttCode,  shttIpumpid as iPumpID, ShttackDaylight as Daylight, shttTztCode as TimeZone, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, cpttVefCode, attTimeType"
    SQLQuery = SQLQuery & " FROM shtt, cptt, vef_Vehicles, att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode"
    SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
    SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
    '10/29/14: Bypass Service agreements
    SQLQuery = SQLQuery + " AND attServiceAgreement <> 'Y'"
    SQLQuery = SQLQuery & " AND vefCode = cpttVefCode"
    SQLQuery = SQLQuery & " AND RTrim(shttIpumpId) <> ''"
     '7701
    SQLQuery = SQLQuery & " AND vatWvtVendorId  = " & Vendors.Wegener_IPump
    'SQLQuery = SQLQuery & " AND attAudioDelivery = " & "'P'"
    If imVefCode > 0 Then
        SQLQuery = SQLQuery & " AND cpttVefCode = " & imVefCode
    End If
    SQLQuery = SQLQuery & " AND cpttStartDate = '" & slDate & "') AND "
    'merge stuff. exclude non ipump and overrides
    SQLQuery = SQLQuery & " vefcode in (select vffvefcode from vff_vehicle_Features where length(vffIpumpEventTypeOv) > 0 and vffExportIPump = 'Y') "
    SQLQuery = SQLQuery & " ORDER BY vefName, shttCallLetters, shttCode"
    Set cprst = gSQLSelectCall(SQLQuery)
    'get vehicles of active agreements.
    Do While Not cprst.EOF
        SQLQuery = "SELECT attVefCode FROM att WHERE attShfCode = " & cprst!shttCode
        SQLQuery = SQLQuery & " AND attOnAir <= '" & Format(smDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND attOffAir >= '" & Format(smDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND attDropDate >= '" & Format(smDate, sgSQLDateForm) & "'"
        Set rst = gSQLSelectCall(SQLQuery)
        Do While Not rst.EOF
            For ilVef = 0 To lbcVehicles.ListCount - 1
                If igExportSource = 2 Then DoEvents
                If lbcVehicles.Selected(ilVef) Then
                    If lbcVehicles.ItemData(ilVef) = rst!attvefCode Then
                        mAddStation slstations, cprst!shttCode
                        Exit Do
                    End If
                End If
            Next ilVef
            rst.MoveNext
        Loop
        cprst.MoveNext
    Loop
End Function
Private Function mGetEventOverride(ilVehicleCode As Integer) As String
    Dim ilVff As Integer
    Dim slEventType As String
    
    slEventType = ""
    If ilVehicleCode > 0 Then
        ilVff = gBinarySearchVff(ilVehicleCode)
        If ilVff <> -1 Then
            'block non ipump and overridden vehicles
            If Len(Trim$(tgVffInfo(ilVff).sIPumpEventTypeOV)) > 0 Then
                slEventType = Trim$(tgVffInfo(ilVff).sIPumpEventTypeOV)
            End If
        End If
    End If
    mGetEventOverride = slEventType
End Function
Private Sub mAddMergeVehicles(slstations() As String)
    Dim slStationList As String
    Dim slVehicles As String
    Dim slSql As String
    Dim myRs As ADODB.Recordset
    Dim slDate As String
   
 On Error GoTo ERRORBOX
    slStationList = Join(slstations, ",")
    If Len(slStationList) > 1 Then
        slStationList = mLoseLastLetter(slStationList)
        slDate = "'" & Format$(gNow(), sgSQLDateForm) & "'"
        slSql = "Select distinct vffvefcode as vef from vff_Vehicle_Features where length(vffIPumpEventTypeOv) > 0 AND vffExportIPump = 'Y' "
        slSql = slSql & "and vffvefcode in (SELECT DISTINCT attVefCode FROM att WHERE attDropDate > " & slDate
        '10/29/14: Bypass Service agreements
        slSql = slSql + " AND attServiceAgreement <> 'Y'"
        slSql = slSql & " AND attOffAir > " & slDate & " and attshfcode in(" & slStationList & " ) )"
        Set myRs = gSQLSelectCall(slSql)
        Do While Not myRs.EOF
            slVehicles = slVehicles & myRs!VEF & ","
            myRs.MoveNext
        Loop
        If Len(slVehicles) > 1 Then
            slVehicles = mLoseLastLetter(slVehicles)
            slDate = gObtainPrevMonday(smDate)
            slDate = "'" & Format$(slDate, sgSQLDateForm) & "'"
            slSql = " SELECT  shttTimeZone, shttCode,  shttIpumpid as iPumpID, ShttackDaylight as Daylight, "
            slSql = slSql & " shttTztCode as TimeZone, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, "
            slSql = slSql & " cpttVefCode, attTimeType FROM shtt,cptt, vef_Vehicles, att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode WHERE (ShttCode = cpttShfCode "
            slSql = slSql & " AND attCode = cpttAtfCode AND vefCode = cpttVefCode AND RTrim(shttIpumpId) <> '' "
            '10/29/14: Bypass Service agreements
            slSql = slSql + " AND attServiceAgreement <> 'Y'"
            '7701
            slSql = slSql & " AND vatWvtVendorId = " & Vendors.Wegener_IPump
            'slSql = slSql & " AND attAudioDelivery = " & "'P'"
            slSql = slSql & " AND cpttStartDate = " & slDate & ") and vefcode in (" & slVehicles & ")"
            slSql = slSql & " AND shttcode in (" & slStationList & ")"
            slSql = slSql & " ORDER BY vefName, shttCallLetters, shttCode"
            Set cprst = gSQLSelectCall(slSql)
        End If
    End If
Cleanup:
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
            Set myRs = Nothing
        End If
    End If
    Exit Sub
ERRORBOX:
    gHandleError smPathForgLogMsg, FORMNAME & "-mAddMergeVehicles"
    GoTo Cleanup
End Sub
Private Sub mAddStation(slstations() As String, ilStation As Integer)
    Dim ilUpper As Integer
    Dim c As Integer
    Dim blFound As Boolean
    
    blFound = False
    ilUpper = UBound(slstations)
    For c = 0 To ilUpper - 1
        If ilStation = slstations(c) Then
            blFound = True
            Exit For
        End If
    Next c
    If Not blFound Then
        slstations(ilUpper) = ilStation
        ReDim Preserve slstations(ilUpper + 1)
    End If
End Sub
Private Function mZoneForSavings(dlStartSavingTime As Date, dlEndSavingTime As Date, slExportEnd As String) As Integer
    'ZONENODAYLIGHTCHANGE ZONEDAYLIGHTATSTATION ZONEDAYLIGHTATTIME
    Dim ilRet As Integer
    
    ilRet = ZONENODAYLIGHTCHANGE
    dlStartSavingTime = mDaylightSavings(True)
    dlEndSavingTime = mDaylightSavings(False)
'    ' this week's export is within daylight savings.  The station's choice to follow dst matters
'    If DateDiff("d", dlStartSavingTime, smDate) > 0 And DateDiff("d", slExportEnd, dlEndSavingTime) > 0 Then
'        ilRet = ZONEDAYLIGHTATSTATION
'    ' this export crosses into daylight savings time! must check at date time for when this happens
'    ElseIf (DateDiff("d", smDate, dlStartSavingTime) > 0 And DateDiff("d", dlStartSavingTime, slExportEnd) > 0) Or (DateDiff("d", slExportEnd, dlEndSavingTime) > 0 And DateDiff("d", dlEndSavingTime, smDate) > 0) Then
'        ilRet = ZONEDAYLIGHTATTIME
'    End If
    '6693
    'export dates entirely within daylight savings time
    If DateDiff("d", dlStartSavingTime, smDate) > 0 And DateDiff("d", slExportEnd, dlEndSavingTime) > 0 Then
        ilRet = ZONEDAYLIGHTATSTATION
    ' this export crosses into daylight savings time! must check at date time for when this happens
    'start date after daylight starts and before daylight ends...or end date is
    ElseIf (DateDiff("d", dlStartSavingTime, smDate) >= 0 And DateDiff("d", smDate, dlEndSavingTime) >= 0) Or (DateDiff("d", slExportEnd, dlEndSavingTime) >= 0 And DateDiff("d", dlStartSavingTime, slExportEnd) >= 0) Then
        ilRet = ZONEDAYLIGHTATTIME
    End If
    mZoneForSavings = ilRet
End Function
Private Function mDaylightSavings(blStart As Boolean) As Date
    ' start: 2nd sunday of march end: 1st sunday in November  both 2:00 am
    Dim slDate As String
    Dim ilDay As Integer
  
 On Error GoTo ERRORBOX
    If blStart Then
        slDate = "3/01/" & DatePart("yyyy", smDate)
        ilDay = DatePart("w", slDate)
        If ilDay <> 1 Then
            slDate = gObtainNextSunday(slDate)
        End If
        slDate = DateAdd("ww", 1, slDate)
        'test only!
       ' slDate = "12/21/1999"
    Else
        slDate = "11/01/" & DatePart("yyyy", smDate)
        ilDay = DatePart("w", slDate)
        If ilDay <> 1 Then
            slDate = gObtainNextSunday(slDate)
        End If
        'test only!
       ' slDate = "12/23/1999"
    End If
    mDaylightSavings = CDate(slDate & " 2:00:00 am")
    Exit Function
ERRORBOX:
    myErrors.WriteError Err.Description, True, True
    mSetResults "Error in mDaylightSavings", MESSAGERED
End Function
Private Function mZoneAdjustTime(slDate As String, slTime As String, ilAdjust As Integer, ilDaylightAdjust As Integer, dlStartSavingTime As Date, dlEndSavingTime As Date) As String
    ' return adjusted date/time if site says to always send as specific zone
    Dim slRet As String
    
    slRet = DateAdd("h", ilAdjust, slDate & " " & slTime)
    '6693 added if.  only adjust if NOT honoring daylight (=1).
    If ilDaylightAdjust = ZONEDAYLIGHTATTIME Then
        ' date is on daylight savings time, is the time after 2:00 am?
        If DateDiff("d", slRet, dlStartSavingTime) = 0 Then
            If DateDiff("h", dlStartSavingTime, slRet) > 0 Then
                 slRet = DateAdd("h", 1, slRet)
            End If
        ' date is when dst ends. Is time before 2:00 am?
        ElseIf DateDiff("d", slRet, dlEndSavingTime) = 0 Then
            If DateDiff("h", slRet, dlEndSavingTime) > 0 Then
                 slRet = DateAdd("h", 1, slRet)
            End If
        ' date is during daylight savings time, but not the start or end date of dst?
        ElseIf DateDiff("d", dlStartSavingTime, slRet) > 0 And DateDiff("d", slRet, dlEndSavingTime) > 0 Then
            slRet = DateAdd("h", 1, slRet)
        End If
    End If
    mZoneAdjustTime = slRet
End Function
'Private Function mZoneAdjustTime(slDate As String, slTime As String, slExportEnd As String, ilAdjust As Integer, ilDaylightAdjust As Integer, dlStartSavingTime As Date, dlEndSavingTime As Date, blOut As Boolean) As String
'    ' return adjusted time zone if site says to always send as specific zone -- formatted as hour only
'    ' return blOut = true if date is outside export range. Don't send this spot
'    Dim slRet As String
'
'    slRet = DateAdd("h", ilAdjust, slDate & " " & slTime)
'    If ilDaylightAdjust = ZONEDAYLIGHTATTIME Then
'        ' date is on daylight savings time, is the time after 2:00 am?
'        If DateDiff("d", slRet, dlStartSavingTime) = 0 Then
'            If DateDiff("h", dlStartSavingTime, slRet) > 0 Then
'                 slRet = DateAdd("h", 1, slRet)
'            End If
'        ' date is when dst ends. Is time before 2:00 am?
'        ElseIf DateDiff("d", slRet, dlEndSavingTime) = 0 Then
'            If DateDiff("h", slRet, dlEndSavingTime) > 0 Then
'                 slRet = DateAdd("h", 1, slRet)
'            End If
'        ' date is during daylight savings time, but not the start or end date of dst?
'        ElseIf DateDiff("d", dlStartSavingTime, slRet) > 0 And DateDiff("d", slRet, dlEndSavingTime) > 0 Then
'            slRet = DateAdd("h", 1, slRet)
'        End If
'    End If
'    'could I skip this for all the spots that aren't adjusted?
'    If DateDiff("d", slExportEnd, slRet) > 0 Or DateDiff("d", smDate, slRet) < 0 Then
'        blOut = True
'    Else
'        blOut = False
'    End If
'    slRet = Format(slRet, "hh")
'    mZoneAdjustTime = slRet
'End Function
Private Function mZoneGetAdjustForStation(ilSiteZone As Integer, ilDaylightAdjust As Integer) As Integer
    'ZONENODAYLIGHTCHANGE ZONEDAYLIGHTATSTATION ZONEDAYLIGHTATTIME
    'out if number to add to adjust for difference from station's zone to 'master zone'.  could be negative
    'an Eastern station to Pacific?  Get a minus number.  When used later with dateAdd, will subtract 3 hours (not incl. daylight issues)
    Dim ilRet As Integer
    Dim ilStation As Integer
    Dim ilZone As Integer
    Dim blNeedDaylightAdjust As Boolean
    Dim slTimeZone As String
    Dim ilIndex As Integer
    
    ilStation = NOTIMEDIFFERENCE
    ilZone = cprst!TimeZone
    For ilIndex = 0 To UBound(tgTimeZoneInfo) - 1 Step 1
        If tgTimeZoneInfo(ilIndex).iCode = ilZone Then
            slTimeZone = tgTimeZoneInfo(ilIndex).sCSIName
            Exit For
        End If
    Next ilIndex
    If Len(slTimeZone) > 0 Then
        slTimeZone = Mid(slTimeZone, 1, 1)
        Select Case slTimeZone
            Case "E"
                ilStation = EASTERN
            Case "C"
                ilStation = CENTRAL
            Case "M"
                ilStation = MOUNTAIN
            Case "P"
                ilStation = PACIFIC
            Case "A"
                ilStation = ALASKAN
            Case Else
                ilStation = HAWAIIAN
        End Select
        ilRet = ilSiteZone - ilStation
    Else
        ilRet = 0
    End If
    '6693
    If ilDaylightAdjust = ZONEDAYLIGHTATSTATION Then
        ilRet = ilRet + 1
    End If
'    ' 0 yes or 1 no aknowledge daylight
'    If cprst!DAYLIGHT = 1 Then
'        blNeedDaylightAdjust = True
'    Else
'        blNeedDaylightAdjust = False
'    End If
'    'station doesn't aknowledge daylight, and this export is during daylight (but not across daylight change!)
'    If blNeedDaylightAdjust And ilDaylightAdjust = ZONEDAYLIGHTATSTATION Then
'        ilRet = ilRet + 1
'    End If
     mZoneGetAdjustForStation = ilRet
End Function
Private Function mZoneGetSite() As Integer
    Dim ilRet As Integer
    Dim slZone As String
    Dim rsZone As ADODB.Recordset
    Dim slSql As String
    
On Error GoTo ERRORBOX
    ilRet = NOTIMEDIFFERENCE
    slSql = "SELECT safIPumpZone as SiteZone From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set rsZone = gSQLSelectCall(slSql)
    If Not rsZone.EOF Then
        slZone = UCase(rsZone!SiteZone)
        Select Case slZone
            Case "E"
                ilRet = EASTERN
            Case "C"
                ilRet = CENTRAL
            Case "M"
                ilRet = MOUNTAIN
            Case "P"
                ilRet = PACIFIC
'            Case "A"
'                ilRet = ALASKAN
'            Case "H"
'                ilRet = HAWAIIAN
        End Select
    End If
Cleanup:
    mZoneGetSite = ilRet
    If Not rsZone Is Nothing Then
        If (rsZone.State And adStateOpen) <> 0 Then
            rsZone.Close
        End If
        Set rsZone = Nothing
    End If
    Exit Function
ERRORBOX:
    gHandleError smPathForgLogMsg, FORMNAME & "-mZoneGetSite"
    ilRet = NOTIMEDIFFERENCE
    GoTo Cleanup
End Function
Private Function mWriteSpots(rsIPump As ADODB.Recordset, blAtLeastOneExport As Boolean) As Boolean
    'Return blAtLeastOneExport
    Dim blRet As Boolean
    Dim myFacts As CLogger
    Dim myClone As ADODB.Recordset
    Dim rscprst As ADODB.Recordset
    Dim slAst As String
    Dim ilSpots As Integer
    Dim slField1 As String
    Dim slEventType As String
    Dim slFileLocation As String
    Dim ilBreaks As Integer
    Dim slFileName As String
    Dim slPreviousDay As String
    Dim slPreviousHour As String
    Dim ilPrevVefCode As Integer
    '6166
    Dim slDeleteCommand As String
    Dim blGatherAvails As Boolean
    Dim blGatherID As Boolean
    '7458
    Dim slIncludeStations As String
    
    blRet = True
    blAtLeastOneExport = False
    slIncludeStations = ""
    slFileName = ""
    slDeleteCommand = ""
    blGatherAvails = True
On Error GoTo ERRORBOX
    If rsIPump Is Nothing Then
    ElseIf (rsIPump.State And adStateOpen) = 0 Then
    ElseIf rsIPump.RecordCount > 0 Then
        Set myFacts = New CLogger
        '7548
        myFacts.BlockUserName = True
        Set myClone = rsIPump.Clone()
        Set rscprst = rsIPump.Clone()
        rscprst.Sort = SORTORDER
        rscprst.MoveFirst
        ilPrevVefCode = -1
        '3 loops.  Get each station id. this is the file (rscprst)
        ' then get each unique time and date for this station id (rsIpump)
        ' use that info to build the # of spots with that station id, date, and time (myclone)
        Do While Not rscprst.EOF
            rscprst.Filter = "Found = false"
             'last step: make .txt into .weg
            If Len(slFileName) > 0 Then
                If myFacts.myFile.FILEEXISTS(slFileName & ".txt") Then
                    myFacts.myFile.MoveFile slFileName & ".txt", slFileName & ".weg"
                End If
            End If
            slFileName = mCreateFile(rscprst!stationid, myFacts)
            If myFacts.isLog Then
                ilBreaks = 0
                rsIPump.Filter = "Found = false AND StationId = '" & rscprst!stationid & "'"
                '6166 write all deletion at start of file for ALL days.
                If Not rsIPump.EOF Then
                    '6306 added network id to know which one needs to be deleted
                    If Not mDeleteAllSpots(rsIPump!stationid, rsIPump!shttCode, rsIPump!NetworkId, myFacts) Then
                        blRet = False
                        mSetResults "Problem deleting spots for " & rsIPump!stationid & ".  See Export log.", MESSAGERED
                    End If
                    slField1 = mFileSafe(rsIPump!stationid) & " STORAGE PLAYLISTDEF "
                    myClone.Sort = SORTORDER
                End If
                '6227 6241
                blGatherID = True
                myClone.Filter = "StationId = '" & rscprst!stationid & "' AND EventType = 'ID'"
                slAst = ""
                slEventType = ""
                slFileLocation = ""
                If Not myClone.EOF Then
                    slEventType = mFileSafe(myClone!EventType) & mFileSafe(myClone!NetworkId)
                    '6302
                    ilSpots = 1
                    Do While Not myClone.EOF
                        slAst = myClone!astCode
                        slFileLocation = mFileSafe(myClone!NameSpace) & mFileSafe(myClone!fileName) & ";"
                        myFacts.WriteFacts slField1 & slEventType & "," & slAst & ",0," & ilSpots & "," & slFileLocation
                        blAtLeastOneExport = True
                        myClone!found = True
                        myClone.MoveNext
                    Loop
                End If
                Do While Not rsIPump.EOF
                    myClone.Filter = "Found = false AND StationId = '" & rsIPump!stationid & "' AND AirDate = '" & rsIPump!airDate & "' AND Hour = '" & rsIPump!Hour & "' AND minute = '" & rsIPump!Minute & "'"
                    ilBreaks = ilBreaks + 1
                    ilSpots = myClone.RecordCount
                    slAst = ""
                    slEventType = ""
                    If Not myClone.EOF Then
                        slEventType = mFileSafe(myClone!EventType) & mFileSafe(myClone!NetworkId) & Trim$(Format(myClone!airDate, "yymmdd")) & UCase(Mid(Format$(myClone!airDate, "ddd"), 1, 2)) & Trim$(myClone!Hour)
                        slFileLocation = ""
                        Do While Not myClone.EOF
                            'a new day restarts the breaks
                            '7/30/13: Added Vehicle test
                            If (slPreviousDay <> myClone!airDate) Or (ilPrevVefCode <> rsIPump!vefCode) Then
                                '7/30/13: Get avails so that correct break number can be determined
                                mGetBreakNumbers rsIPump
                                slPreviousDay = myClone!airDate
                                ilBreaks = 1
                                ilPrevVefCode = rsIPump!vefCode
    '                            '6166
    '                            myFacts.WriteFacts slDeleteCommand
                            End If
                            'a new hour restarts the breaks
                            If slPreviousHour <> myClone!Hour Then
                                slPreviousHour = myClone!Hour
                                ilBreaks = 1
                            End If
                            myClone!found = True
                            slAst = slAst & myClone!astCode & "|"
                            slFileLocation = slFileLocation & mFileSafe(myClone!NameSpace) & mFileSafe(myClone!fileName) & ";"
                            myClone.MoveNext
                        Loop 'day and time match
                            slAst = mLoseLastLetter(slAst)
                            '7/30/13: Get break number
                            rsBreakNumbers.Filter = "Date = " & gDateValue(rsIPump!airDate) & " AND Hour = '" & rsIPump!Hour & "' AND Min = '" & rsIPump!Minute & "'"
                            If Not rsBreakNumbers.EOF Then
                                slEventType = slEventType & "1" & rsBreakNumbers!BreakNumber
                            Else
                                '7/30/13: Retained old break number count if break number code fails
                                slEventType = slEventType & "1" & ilBreaks
                            End If
                            myFacts.WriteFacts slField1 & slEventType & "," & slAst & ",0," & ilSpots & "," & slFileLocation
                            blAtLeastOneExport = True
                    End If
                   rsIPump.Filter = "Found = false AND StationId = '" & rscprst!stationid & "'"
                Loop 'asts that match station id
                '7458
                slIncludeStations = slIncludeStations & CStr(rscprst!shttCode) & ","
                rscprst.Filter = "Found = false"
            Else
                'quit if can't create a file!
                myErrors.WriteError "Could not create log for IPumpStationId: " & rscprst!stationid
                mSetResults "Error in creating file for IPumpStationId:" & rscprst!stationid, MESSAGERED
                blRet = False
                Exit Do
            End If
        Loop 'main loop: get unique stations
        If Len(slFileName) > 0 Then
            If myFacts.myFile.FILEEXISTS(slFileName & ".txt") Then
                myFacts.myFile.MoveFile slFileName & ".txt", slFileName & ".weg"
            End If
        End If
        '7458
        slIncludeStations = mLoseLastLetter(slIncludeStations)
        If Not myEnt.UpdateAsSuccessful(False, slIncludeStations, StationCodes) Then
            myErrors.WriteWarning myEnt.ErrorMessage
        End If
    End If
Cleanup:
    If Not myFacts Is Nothing Then
        Set myFacts = Nothing
    End If
    If Not myClone Is Nothing Then
        If (myClone.State And adStateOpen) <> 0 Then
            myClone.Close
            Set myClone = Nothing
        End If
    End If
    mWriteSpots = blRet
    Exit Function
ERRORBOX:
    myErrors.WriteError "mWriteSpots-" & Err.Description, True, True
    mSetResults "Error in mWriteSpots", MESSAGERED
    blRet = False
End Function

Private Function mFileSafe(ByVal slInput As String) As String
    slInput = Trim$(slInput)
    slInput = Replace(slInput, ",", "_")
    slInput = Replace(slInput, ";", "_")
    mFileSafe = slInput
End Function
Private Sub mFinishUp()
    lacResult.Caption = ""
    If (lbcStation.ListCount = 0) Or (chkAllStation.Value = vbChecked) Or (lbcStation.ListCount = lbcStation.SelCount) Then
        gClearASTInfo True
    Else
        gClearASTInfo False
    End If
    mClearAlerts
End Sub
Private Function mCifInfo(ilIndex As Integer, slFileName As String, slEventType As String, slNetworkId As String, slNameSpace As String) As Boolean
'get cpf(creative title), mef(event type, network id,NameSpace)
' event type comes in with override!  Only fill if blank.
'return false if error OR a field is empty.
    Dim llCif As Long
    Dim blRet As Boolean
    Dim myRs As ADODB.Recordset
    Dim Sql As String
    Dim llMcf As Long
    Dim ilVff As Integer
    Dim ilVehicleCode As Integer
    
    blRet = True
    slEventType = vbNullString
    slFileName = vbNullString
    slNetworkId = vbNullString
    slNameSpace = vbNullString
    llMcf = 0
On Error GoTo ErrHandler
    If tmAstInfo(ilIndex).iRegionType = 0 Then
        llCif = tmAstInfo(ilIndex).lCifCode
    Else
        llCif = tmAstInfo(ilIndex).lRCifCode
    End If
    If llCif > 0 Then
        If igExportSource = 2 Then DoEvents
        Sql = "Select cifName, cifmcfCode FROM cif_Copy_Inventory WHERE cifCode = " & llCif
        Set myRs = gSQLSelectCall(Sql)
        If Not (myRs.EOF Or myRs.BOF) Then
            '6195 removed "format()" from below
            slFileName = myRs!cifName
            llMcf = Format(myRs!cifMcfCode)
        End If
        If igExportSource = 2 Then DoEvents
        myRs.Close
    End If
    If llMcf > 0 And Len(slFileName) > 0 Then
        If igExportSource = 2 Then DoEvents
        Sql = "Select mefEventType,mefNetworkId,mefNameSpace,mefPrefix,mefSuffix FROM mef_Media_Extra WHERE mefmcfCode = " & llMcf
        Set myRs = gSQLSelectCall(Sql)
        If Not (myRs.EOF Or myRs.BOF) Then
            slEventType = Trim$(Format(myRs!mefEventType))
            slNetworkId = Trim$(Format(myRs!mefNetworkId))
            slNameSpace = Trim$(Format(myRs!mefNameSpace))
            slFileName = Trim$(Format(myRs!mefPrefix)) & Trim(slFileName) & Trim$(Format(myRs!mefSuffix))
        End If
    End If
    'required fields  slEventType no longer required...could not exist and be overridden later.
    If Len(slNetworkId) = 0 Or Len(slNameSpace) = 0 Or Len(slFileName) = 0 Then
        blRet = False
    End If
Cleanup:
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
            Set myRs = Nothing
        End If
    End If
    mCifInfo = blRet
    Exit Function
ErrHandler:
    blRet = False
    gHandleError smPathForgLogMsg, FORMNAME & "-mCifInfo"
    GoTo Cleanup
End Function
Private Sub mFactFile(myRs As ADODB.Recordset)
    Dim myFacts As CLogger
    Dim rsWrite As Recordset
    Dim slAdv As String
    Dim blDay As Boolean
    Dim slPreviousStation As String
    Dim slPreviousVehicle As String
    Dim slTime As String
    Dim llAdf As Long
    '6693
    Dim ilTimeAdjust As Integer
    
On Error GoTo ERRORBOX
    slPreviousStation = ""
    Set rsWrite = mPrepRecordsetDebug()
    If Not myRs Is Nothing And Not rsTimeAdjust Is Nothing And Not rsWrite Is Nothing Then
        myRs.Filter = adFilterNone
        If Not myRs.EOF And Not rsTimeAdjust.EOF Then
            myRs.MoveFirst
            With rsTimeAdjust
                .Filter = adFilterNone
                .MoveFirst
                Do While Not .EOF
                    If slPreviousStation <> .Fields("stationid") Then
                        slPreviousStation = .Fields("stationid")
                        slPreviousVehicle = .Fields("Vehicle")
                        myRs.Filter = "StationId = '" & slPreviousStation & "' and Vehicle = '" & slPreviousVehicle & "'"
                    ElseIf slPreviousVehicle <> .Fields("Vehicle") Then
                         slPreviousVehicle = .Fields("Vehicle")
                        myRs.Filter = "StationId = '" & slPreviousStation & "' and Vehicle = '" & slPreviousVehicle & "'"
                    End If
                    '6693
                    ilTimeAdjust = .Fields("AdjustTime")
                    'station or each spot must check to see if date in range of daylight savings because this station doesn't use it.
                    If .Fields("DaylightAdjust") > 0 Then
                        blDay = True
                        '6693
                        If .Fields("DaylightAdjust") = ZONEDAYLIGHTATTIME Then
                            ilTimeAdjust = ilTimeAdjust + 1
                        End If
                      Else
                        blDay = False
                    End If
                    Do While Not myRs.EOF
                        llAdf = gBinarySearchAdf(myRs!adv)
                        If llAdf <> -1 Then
                            slAdv = Trim$(tgAdvtInfo(llAdf).sAdvtName)
                        Else
                            slAdv = "Advertiser Name Missing"
                        End If
                        slTime = myRs!Hour & ":" & myRs!Minute
                        '6693
                        rsWrite.AddNew Array("ISCI", "ADVERTISER", "date", "Time", "TimeAdjust", "DaylightAdjust", "StationName", "vehicle", "AstCode"), Array(myRs!ISCI, slAdv, myRs!airDate, slTime, ilTimeAdjust, blDay, .Fields("StationName"), .Fields("Vehicle"), myRs!astCode)
                        'rsWrite.AddNew Array("ISCI", "ADVERTISER", "date", "Time", "TimeAdjust", "DaylightAdjust", "StationName", "vehicle", "AstCode"), Array(myRs!ISCI, slAdv, myRs!airDate, slTime, .Fields("AdjustTime"), blDay, .Fields("StationName"), .Fields("Vehicle"), myRs!astCode)
                        myRs.MoveNext
                    Loop
                    .MoveNext
                Loop
            End With
        End If
    End If
    slPreviousStation = ""
    slPreviousVehicle = ""
    Set myFacts = New CLogger
    With myFacts
        '7548
        .BlockUserName = True
        .LogPath = .CreateLogName(smExportDirectory & "iPumpFacts")
        .WriteFacts "iPump export for " & smDate & ", " & imNumberDays & " days.", True
        rsWrite.Filter = adFilterNone
        Do While Not rsWrite.EOF
            If slPreviousStation <> rsWrite!StationName Then
                slPreviousStation = rsWrite!StationName
                .WriteFacts "STATION " & slPreviousStation
                slPreviousVehicle = ""
            End If
            If slPreviousVehicle <> rsWrite!Vehicle Then
                slPreviousVehicle = rsWrite!Vehicle
                .WriteFacts " VECHICLE " & slPreviousVehicle
            End If
            .WriteFacts "       " & Trim$(rsWrite!Advertiser) & " " & Trim$(rsWrite!ISCI) & " " & rsWrite!Date & " " & Trim$(rsWrite!TIME) & " Time adjusted:" & rsWrite!TimeAdjust & " Daylight a factor?:" & rsWrite!DaylightAdjust & " astcode: " & rsWrite!astCode
            rsWrite.MoveNext
        Loop
    End With
Cleanup:
    If Not rsWrite Is Nothing Then
        If (rsWrite.State And adStateOpen) <> 0 Then
            rsWrite.Close
        End If
        Set rsWrite = Nothing
    End If
    Set myFacts = Nothing
    Exit Sub
ERRORBOX:
    myErrors.WriteError " mFactFile: " & Err.Description
    GoTo Cleanup
End Sub
Private Function mPrepRecordsetTime() As Recordset
    Dim myRs As ADODB.Recordset

    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "StationID", adChar, 10
            .Append "AdjustTime", adInteger
            .Append "DaylightAdjust", adInteger
            .Append "StationName", adChar, 40
            .Append "Vehicle", adChar, 40
          '  .Append "VehicleId", adInteger
        End With
    myRs.Open
    Set mPrepRecordsetTime = myRs
End Function
Private Function mPrepRecordsetDebug() As Recordset
    Dim myRs As ADODB.Recordset

    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "ISCI", adChar, 40
            .Append "Advertiser", adChar, 30
            .Append "Date", adDate
            .Append "Time", adChar, 8
            .Append "DaylightAdjust", adBoolean
            .Append "StationName", adChar, 40
            .Append "Vehicle", adChar, 40
            .Append "TimeAdjust", adInteger
            .Append "AstCode", adInteger
        End With
    myRs.Open
    myRs!StationName.Properties("optimize") = True
    myRs.Sort = "StationName,Vehicle,Date,Time"
    Set mPrepRecordsetDebug = myRs
End Function
Private Function mPrepRecordsetDelete() As Recordset
    Dim myRs As ADODB.Recordset

    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "StationID", adInteger
            .Append "VehicleID", adInteger
            .Append "Override", adChar, 2
            .Append "Delete", adBoolean
        End With
    myRs.Open
    Set mPrepRecordsetDelete = myRs
End Function
Private Function mDeleteAllSpots(slStationID As String, llShttCode As Long, slNetwork As String, myFacts As CLogger) As Boolean
    Dim blRet As Boolean
    Dim myRs As ADODB.Recordset
    Dim slSql As String
    Dim c As Integer
    Dim slCurrentDate As String
    Dim slDeleteCommand As String
    Dim slDeletePartial As String
    Dim slDeleteAlmostComplete As String
    Dim llVef As Long
    Dim ilVff As Integer
    Dim ilVehicleCode As Integer
    
On Error GoTo ERRORBOX:
    blRet = True
    slNetwork = Trim$(slNetwork)
    'first pass, be safe and write out all form mef.  Don't worry about overrides yet.
    '6235. added where clause to skip 'id'
    '7772 if writing just delete file, I don't know the network id.
    If Len(slNetwork) > 0 Then
        slSql = "select distinct mefnetworkId as NetworkId,mefEventType as EventType from MEF_Media_Extra where ucase(mefEventType) <> 'ID' AND networkid = '" & slNetwork & "'"
    Else
        slSql = "select distinct mefnetworkId as NetworkId,mefEventType as EventType from MEF_Media_Extra where ucase(mefEventType) <> 'ID'"
    End If
    Set myRs = gSQLSelectCall(slSql)
    If Not myRs.EOF Then
        slDeletePartial = mFileSafe(slStationID) & " STORAGE DELPLAYLIST "
    Else
        blRet = False
        GoTo Cleanup
    End If
    Do While Not myRs.EOF
        If Len(Trim(myRs!EventType)) > 0 And Len(Trim(myRs!NetworkId)) > 0 Then
            slDeleteAlmostComplete = slDeletePartial & mFileSafe(myRs!EventType) & mFileSafe(myRs!NetworkId)
            For c = 0 To imNumberDays - 1
                slCurrentDate = DateAdd("d", c, smDate)
                slDeleteCommand = slDeleteAlmostComplete & Trim$(Format(slCurrentDate, "yymmdd")) & UCase(Mid(Format$(slCurrentDate, "ddd"), 1, 2)) & "????"
                myFacts.WriteFacts slDeleteCommand
            Next c
        End If
        myRs.MoveNext
    Loop
    'ok,now worry about overrrides in vff
    If Not rsDeletePrevious Is Nothing Then
        rsDeletePrevious.Filter = "Delete = False and StationID = " & llShttCode
        Do While Not rsDeletePrevious.EOF
            'this vehicle belongs with this station
            llVef = gBinarySearchVef(rsDeletePrevious!vehicleid)
            If llVef <> -1 Then
                ilVehicleCode = tgVehicleInfo(llVef).iCode
                ilVff = gBinarySearchVff(ilVehicleCode)
                If ilVff <> -1 Then
                    'block non ipump and overridden vehicles
                    '6235.  skip 'id'
                    If Trim$(tgVffInfo(ilVff).sExportIPump) = "Y" And Len(Trim$(tgVffInfo(ilVff).sIPumpEventTypeOV)) <> 0 And UCase(Trim$(tgVffInfo(ilVff).sIPumpEventTypeOV)) <> "ID" Then
                        slSql = "select distinct mefnetworkId as NetworkId from MEF_Media_Extra"
                        Set myRs = gSQLSelectCall(slSql)
                        Do While Not myRs.EOF
                            If Len(Trim(myRs!NetworkId)) > 0 Then
                                slDeleteAlmostComplete = slDeletePartial & mFileSafe(Trim$(tgVffInfo(ilVff).sIPumpEventTypeOV)) & mFileSafe(myRs!NetworkId)
                                For c = 0 To imNumberDays - 1
                                    slCurrentDate = DateAdd("d", c, smDate)
                                    slDeleteCommand = slDeleteAlmostComplete & Trim$(Format(slCurrentDate, "yymmdd")) & UCase(Mid(Format$(slCurrentDate, "ddd"), 1, 2)) & "????"
                                    myFacts.WriteFacts slDeleteCommand
                                Next c
                            End If
                            myRs.MoveNext
                        Loop
                    End If ' override!
                End If ' ilvff found
            End If 'llvef found
            rsDeletePrevious!Delete = True
            rsDeletePrevious.MoveNext
        Loop
    End If
Cleanup:
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
        End If
        Set myRs = Nothing
    End If
    mDeleteAllSpots = blRet
    Exit Function
ERRORBOX:
    myErrors.WriteError "mDeleteAllSpots- " & Err.Description, True, True
    blRet = False
    GoTo Cleanup
End Function
'Private Sub mDeleteAllSpots(myRs As ADODB.Recordset, myFacts As CLogger)
'    Dim c As Integer
'    Dim slCurrentDate As String
'    Dim slDeleteCommand As String
'    Dim slDeletePartial As String
'    Dim slPreviousEventType As String
'    Dim slDeleteAlmostComplete As String
'
'    slPreviousEventType = ""
'    myRs.Sort = "EventType"
'    myRs.Filter = adFilterNone
'    If Not myRs.EOF Then
'        slDeletePartial = mFileSafe(myRs!stationid) & " STORAGE DELPLAYLIST "
'        myRs.MoveFirst
'    End If
'    Do While Not myRs.EOF
'        If slPreviousEventType <> myRs!EventType Then
'            slPreviousEventType = myRs!EventType
'            slDeleteAlmostComplete = slDeletePartial & mFileSafe(myRs!EventType) & mFileSafe(myRs!NetworkId)
'            For c = 0 To imNumberDays - 1
'                slCurrentDate = DateAdd("d", c, smDate)
'                slDeleteCommand = slDeleteAlmostComplete & Trim$(Format(slCurrentDate, "yymmdd")) & UCase(Mid(Format$(slCurrentDate, "ddd"), 1, 2)) & "????"
'                myFacts.WriteFacts slDeleteCommand
'            Next c
'        End If
'        myRs.MoveNext
'    Loop
'
'End Sub
Private Function mCreateFile(ByVal slStationID As String, myFacts As CLogger) As String
    Dim slFileName As String
    
    slFileName = smExportDirectory & mFileNameFilter(slStationID) & "_" & Format(smDate, "yyyymmdd")
    myFacts.CleanFile slFileName & ".txt", 0
    myFacts.CleanFile slFileName & ".weg", 0
    myFacts.LogPath = slFileName & ".txt"
    mCreateFile = slFileName
End Function
Private Function mDeleteUnwrittenSpots() As Boolean
    'write out a file for each station that didn't have any exports!.  Wrote to rsDeletePrevious each station/vehicle
    'combo.  When creating spots, marked each as 'deleted'.  So the only ones left are stations that didn't have a file.
    Dim blRet As Boolean
    Dim myRs As ADODB.Recordset
    Dim myClone As ADODB.Recordset
    Dim slSql As String
    Dim myFacts As CLogger
    Dim slFileName As String
    
On Error GoTo ERRORBOX
    blRet = True
    If Not rsDeletePrevious Is Nothing Then
        Set myFacts = New CLogger
        '7548
        myFacts.BlockUserName = True
        Set myClone = rsDeletePrevious.Clone
        myClone.Filter = "Delete = False "
        'we have vehicles to test.  That means stations that need a file.
        Do While Not myClone.EOF
            slSql = "select shttiPumpId from shtt where shttcode = " & CInt(myClone!stationid)
            Set myRs = gSQLSelectCall(slSql)
            If Not myRs.EOF Then
                slFileName = mCreateFile(myRs!shttIPumpID, myFacts)
                If myFacts.isLog Then
                    '7772
                    If Not mDeleteAllSpots(myRs!shttIPumpID, myClone!stationid, "", myFacts) Then
'                    If Not mDeleteAllSpots(myRs!shttIPumpID, myClone!stationid, myClone!NetworkdId, myFacts) Then
                        blRet = False
                        mSetResults "Problem deleting spots for " & myRs!shttIPumpID & ".  See Export log.", MESSAGERED
                    End If
                    If myFacts.myFile.FILEEXISTS(slFileName & ".txt") Then
                        myFacts.myFile.MoveFile slFileName & ".txt", slFileName & ".weg"
                    End If
                End If
            End If
            myClone.MoveNext
        Loop
    End If
Cleanup:
    Set myFacts = Nothing
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
        End If
        Set myRs = Nothing
    End If
    If Not myClone Is Nothing Then
        If (myClone.State And adStateOpen) <> 0 Then
            myClone.Close
        End If
        Set myClone = Nothing
    End If
    mDeleteUnwrittenSpots = blRet
    Exit Function
ERRORBOX:
    myErrors.WriteError "mDeleteUnwrittenSpots- " & Err.Description, True, True
    blRet = False
    GoTo Cleanup
End Function

Private Sub mGetBreakNumbers(rsIPump As ADODB.Recordset)
    Dim ilDat As Integer
    Dim dlStartSavingTime As Date
    Dim dlEndSavingTime As Date
    Dim slAdjustedDate As String
    Dim slAdjustedHour As String
    Dim slAdjustedMinute As String
    Dim llDate As Long
    Dim slHour As String
    Dim slMin As String
    Dim ilBreakNumber As Integer
    Dim ilDay As Integer
    Dim slDate As String
    Dim slMoDate As String
    Dim ilAirDay As Integer
    
    ReDim tgDat(0 To 0) As DAT      'gGetAVails loads tgDat array with avails
    For ilDay = 0 To 6 Step 1
        tgDat(0).iFdDay(ilDay) = 0
        tgDat(0).iPdDay(ilDay) = 0
    Next ilDay
    '7/30/13: Initialize Break Number structure
    If Not rsBreakNumbers Is Nothing Then
        If (rsBreakNumbers.State And adStateOpen) <> 0 Then
            rsBreakNumbers.Close
        End If
    End If
    Set rsBreakNumbers = mPrepRecordsetBreakNumber()
    gGetAvails rsIPump!attCode, rsIPump!shttCode, rsIPump!vefCode, 0, rsIPump!airDate, True
    dlStartSavingTime = mDaylightSavings(True)
    dlEndSavingTime = mDaylightSavings(False)
    slMoDate = gObtainPrevMonday(rsIPump!airDate)
    For ilDat = 0 To UBound(tgDat) - 1 Step 1
        For ilDay = 0 To 7 Step 1
            slDate = DateAdd("d", ilDay, slMoDate)
            If ilDay <= 6 Then
                ilAirDay = tgDat(ilDat).iFdDay(ilDay)
            Else
                ilAirDay = tgDat(ilDat).iFdDay(0)
            End If
            If ilAirDay = 1 Then
                slAdjustedDate = mZoneAdjustTime(slDate, tgDat(ilDat).sFdSTime, rsIPump!ZoneAdjust, rsIPump!DaylightAdjust, dlStartSavingTime, dlEndSavingTime)
                slAdjustedHour = Format(slAdjustedDate, "hh")
                slAdjustedMinute = Format(slAdjustedDate, "nn")
                slAdjustedDate = Format(slAdjustedDate, sgShowDateForm)
                rsBreakNumbers.Filter = "Date = " & gDateValue(slAdjustedDate) & " AND Hour = '" & slAdjustedHour & "' AND Min = '" & slAdjustedMinute & "'"
                If rsBreakNumbers.EOF Then
                    rsBreakNumbers.AddNew Array("Date", "Hour", "Min", "BreakNumber"), Array(gDateValue(slAdjustedDate), slAdjustedHour, slAdjustedMinute, 0)
                End If
            End If
        Next ilDay
    Next ilDat
    rsBreakNumbers.Filter = adFilterNone
    llDate = -1
    Do While Not rsBreakNumbers.EOF
        If (llDate = -1) Or (llDate <> rsBreakNumbers!Date) Then
            ilBreakNumber = 1
            rsBreakNumbers!BreakNumber = ilBreakNumber
        ElseIf slHour <> rsBreakNumbers!Hour Then
            ilBreakNumber = 1
            rsBreakNumbers!BreakNumber = ilBreakNumber
        ElseIf slMin <> rsBreakNumbers!Min Then
            ilBreakNumber = ilBreakNumber + 1
            rsBreakNumbers!BreakNumber = ilBreakNumber
        Else
            rsBreakNumbers!BreakNumber = ilBreakNumber
        End If
        llDate = rsBreakNumbers!Date
        slHour = rsBreakNumbers!Hour
        slMin = rsBreakNumbers!Min
        rsBreakNumbers.MoveNext
    Loop
    
End Sub

Private Function mPrepRecordsetBreakNumber() As Recordset
    Dim myRs As ADODB.Recordset

    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "Date", adInteger
            .Append "Hour", adChar, 2
            .Append "Min", adChar, 2
            .Append "BreakNumber", adInteger
        End With
    myRs.Open
    myRs!Date.Properties("optimize") = True
    myRs.Sort = "Date,Hour,Min"
    Set mPrepRecordsetBreakNumber = myRs
End Function


