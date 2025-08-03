VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExportStationComp 
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   9360
   Begin MSComctlLib.ProgressBar plcGauge 
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   5520
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7320
      Top             =   5760
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
      Caption         =   "Export Station Compensation Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4605
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   8895
      Begin V81Affiliate.CSI_Calendar CalOffAirDate 
         Height          =   270
         Left            =   2520
         TabIndex        =   4
         Top             =   495
         Width           =   1200
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
         CSI_AllowBlankDate=   0   'False
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   0
      End
      Begin V81Affiliate.CSI_Calendar CalOnAirDate 
         Height          =   270
         Left            =   840
         TabIndex        =   3
         Top             =   495
         Width           =   1200
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
         CSI_ForceMondaySelectionOnly=   -1  'True
         CSI_AllowBlankDate=   0   'False
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   2
      End
      Begin VB.CheckBox ckcInclBonus 
         Caption         =   "Include Bonus Spots"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Frame frcCompType 
         Caption         =   "Compensation Type"
         Height          =   915
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   3675
         Begin VB.CheckBox ckcCompType 
            Caption         =   "Barter"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox ckcCompType 
            Caption         =   "Pay Network"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   17
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox ckcCompType 
            Caption         =   "Pay Station"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
      End
      Begin VB.ListBox lbcStations 
         Height          =   1620
         ItemData        =   "AffExportStationComp.frx":0000
         Left            =   4200
         List            =   "AffExportStationComp.frx":0002
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2760
         Width           =   4260
      End
      Begin VB.CheckBox ckcAllStations 
         Caption         =   "All Stations"
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ListBox lbcVehicles 
         Height          =   1620
         ItemData        =   "AffExportStationComp.frx":0004
         Left            =   4200
         List            =   "AffExportStationComp.frx":0006
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   660
         Width           =   4275
      End
      Begin VB.Frame frcFeedOrAir 
         Caption         =   "Use"
         Height          =   675
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3705
         Begin VB.OptionButton rbcFeedOrAir 
            Caption         =   "Feed  Dates"
            Height          =   255
            Index           =   0
            Left            =   75
            TabIndex        =   6
            Top             =   270
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton rbcFeedOrAir 
            Caption         =   "Air Dates"
            Height          =   255
            Index           =   1
            Left            =   1485
            TabIndex        =   7
            Top             =   270
            Width           =   1095
         End
      End
      Begin VB.CheckBox ckcAllVehicles 
         Caption         =   "All Vehicles"
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   285
         Width           =   1320
      End
      Begin VB.Frame frcDates 
         Caption         =   "Dates (Monday-Sunday)"
         Height          =   690
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3735
         Begin VB.Label lacTo 
            Caption         =   "To"
            Height          =   225
            Left            =   2010
            TabIndex        =   14
            Top             =   270
            Width           =   315
         End
         Begin VB.Label lacFrom 
            Caption         =   " From"
            Height          =   225
            Left            =   120
            TabIndex        =   12
            Top             =   270
            Width           =   735
         End
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   5520
      Width           =   1845
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lacProgress 
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   5805
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   2040
      TabIndex        =   20
      Top             =   4920
      Width           =   5580
   End
End
Attribute VB_Name = "frmExportStationComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'*  frmExportStationComp - Report of agreements that are due for renewal based on
'   user entered date spans against the agreement end date
'*
'*  Created July,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'

'****************************************************************************
Option Explicit

Private imChkStationIgnore As Integer
Private imckcAllVehiclesIgnore As Integer
Private rst_att As ADODB.Recordset
Private rst_Shtt As ADODB.Recordset
Private tmStatusOptions As STATUSOPTIONS
Private hmAst As Integer
Private hmFileHandle As Integer



Private Sub CalOffAirDate_CalendarChanged()
    mSetCommands
End Sub

Private Sub CalOffAirDate_Change()
    mSetCommands
End Sub

Private Sub CalOnAirDate_CalendarChanged()
    mSetCommands
End Sub

Private Sub CalOnAirDate_Change()
    mSetCommands
End Sub

Private Sub ckcAllVehicles_Click()
    Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    Dim iLoop As Integer
    
    If imckcAllVehiclesIgnore Then
        Exit Sub
    End If
    If ckcAllVehicles.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcVehicles.ListCount > 0 Then
        imckcAllVehiclesIgnore = True
        lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehicles.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imckcAllVehiclesIgnore = False

    End If
    
    If lbcVehicles.SelCount > 1 Then
        lbcStations.Clear
        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
                lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
            End If
        Next iLoop
    End If
    
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
    mSetCommands
End Sub

Private Sub ckcAllStations_Click()
 Dim i As Integer
    Dim lErr As Long
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imChkStationIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    Screen.MousePointer = vbHourglass
    lErr = LockWindowUpdate(Me.hwnd)
    If lbcStations.ListCount > 0 Then
        imChkStationIgnore = True
        lRg = CLng(lbcStations.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStations.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkStationIgnore = False
    End If
    lErr = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
    mSetCommands
End Sub

Private Sub cmcBrowse_Click()
    Dim slCurDir As String
    
    slCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    'CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    '"(*.txt)|*.txt|CSV Files (*.csv)|*.csv"
    CommonDialog1.Filter = "All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
    'CommonDialog1.Filter = "Files (*.csv)|CSV File (*.csv)|*.csv"
    ' Specify default filter
    CommonDialog1.FilterIndex = 3
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

   ' txtFile.Text = Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub cmdDone_Click()
    Unload frmExportStationComp
End Sub

Private Sub cmdExport_Click()
    Dim iRet As Integer
    Dim sStartDate As String
    Dim sEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slStartYear As String
    Dim slStartMonth As String
    Dim slStartDay As String
    Dim slEndYear As String
    Dim slEndMonth As String
    Dim slEndDay As String
    Dim ilRet As Integer
    Dim dFWeek As Date
    Dim slExportName As String
    Dim blUseAirDAte As Boolean
    Dim slRepeat As String * 1
    Dim slDateTime As String
    Dim slStr As String
    Dim blExportOK As Boolean
    
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
        
        'date must be a monday, back it up
        If Weekday(sStartDate, vbSunday) <> vbMonday Then
            sStartDate = Format$(gObtainPrevMonday(sStartDate), sgShowDateForm)
        End If
        'force to Monday - Sunday week
        If Weekday(sEndDate, vbSunday) <> vbSunday Then
            iRet = gMsgBox("Changing end date to Sunday", vbOKCancel)
            If iRet = vbCancel Then
                Exit Sub
            End If
        End If
        If Weekday(sEndDate, vbSunday) <> vbSunday Then
            sEndDate = Format$(gObtainNextSunday(sEndDate), sgShowDateForm)
            CalOffAirDate.Text = sEndDate
            DoEvents
        End If
        
        llStartDate = gDateValue(sStartDate)
        sStartDate = Format(llStartDate, "m/d/yy")   'make sure string start date has a year appended in case not entered with input
        gObtainYearMonthDayStr sStartDate, True, slStartYear, slStartMonth, slStartDay
        
        llEndDate = gDateValue(sEndDate)
        sEndDate = Format(llEndDate, "m/d/yy")   'make sure string start date has a year appended in case not entered with input
        gObtainYearMonthDayStr sEndDate, True, slEndYear, slEndMonth, slEndDay
        
        Screen.MousePointer = vbHourglass
        
        gUserActivityLog "S", "Station Compensation Export"
        'Determine name of export (.csv file)
        slRepeat = "A"
        Do
            iRet = 0
            On Error GoTo cmcExportDupNameErr:
            slExportName = sgExportDirectory & "Station-Comp-" & slStartMonth & slStartDay & Mid(slStartYear, 3, 2) & "-" & slEndMonth & slEndDay & Mid(slEndYear, 3, 2) & slRepeat & ".csv"
            'slDateTime = FileDateTime(slExportName)
            ilRet = gFileExist(slExportName)
            If ilRet = 0 Then       'if went to mOpenErr , there was a filename that existed with same name. Increment the letter
                slRepeat = Chr(Asc(slRepeat) + 1)
            End If
        Loop While ilRet = 0     'loop while file is found
        
        If rbcFeedOrAir(1).Value = True Then     'use air dates vs fed dates,
            '10-8-14 no need to backu up for air date outside of week; new design has key by air date
            'sStartDate = DateAdd("d", -7, sStartDate)  ' need to backup the week and process extra week spot may air outside the week
            blUseAirDAte = True
            gLogMsg "** Begin Station Compensation export using AirDates: " & slExportName & ", " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportStationComp.txt", False
        Else
            blUseAirDAte = False
            gLogMsg "** Begin Station Compensation export using FeedDates: " & slExportName & ", " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportStationComp.txt", False
        End If
        
        'hmFileHandle = FreeFile
        'iRet = 0
        'Open slExportName For Output As hmFileHandle
        ilRet = gFileOpen(slExportName, "Output", hmFileHandle)
        On Error GoTo 0
        If iRet <> 0 Then
            gLogMsg "Open File Error #" & Str$(Err.Number) & slExportName, "Station-Comp.csv", False
            Close #hmFileHandle
            'imExporting = False
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        plcGauge.Visible = True
        plcGauge.Value = 0

        On Error GoTo WriteColumnHeadingerr
        slStr = "Vehicle ID" & "," & "Vehicle" & "," & "Station ID" & "," & "Station" & "," & "Agreement Ref" & "," & "Comp Flag" & ","
        slStr = slStr & "Pledge Day" & "," & "Pledge Date" & "," & "Pledge Start Time" & "," & "Pledge End Time" & ","
        slStr = slStr & "Day Aired" & "," & "Date Aired" & "," & "Time Aired" & "," & "Advertiser/Product" & "," & "ISCI" & ","
        slStr = slStr & "Contract" & "," & "Len" & "," & "Status" & "," & "Replaced Advt/Prod" & "," & "Replaced Contract" & "," & "Replaced ISCI" & ","
        slStr = slStr & "Missed Date" & "," & "Missed Time"
        Print #hmFileHandle, slStr     'write header description
        On Error GoTo 0

        On Error GoTo WriteColumnHeadingerr
        slStr = "As of " & Format$(gNow(), "mm/dd/yy") & " "
        slStr = slStr & Format$(gNow(), "h:mm:ssAM/PM")
        Print #hmFileHandle, slStr        'write header description
        On Error GoTo 0
    
        tmStatusOptions.iNotReported = False
        tmStatusOptions.iInclNotCarry8 = False
        tmStatusOptions.iInclLive0 = True
        tmStatusOptions.iInclDelay1 = True
        tmStatusOptions.iInclMissed2 = True
        tmStatusOptions.iInclMissed3 = True
        tmStatusOptions.iInclMissed4 = True
        tmStatusOptions.iInclMissed5 = True
        tmStatusOptions.iInclAirOutPledge6 = True
        tmStatusOptions.iInclAiredNotPledge7 = True
        tmStatusOptions.iInclDelayCmmlOnly9 = True
        tmStatusOptions.iInclAirCmmlOnly10 = True
        tmStatusOptions.iInclMG11 = True
        tmStatusOptions.iInclRepl13 = True
        If ckcInclBonus.Value = vbChecked Then
            tmStatusOptions.iInclBonus12 = True
        End If
        tmStatusOptions.iInclResolveMissed = False          'dont show the missed part of a mg or replacement
        tmStatusOptions.iInclMissedMGBypass14 = True        '4-12-17 default to include missed mg bypassed
        
        tmStatusOptions.iCompBarter = True
        tmStatusOptions.iCompPayStation = True
        tmStatusOptions.iCompPayNetwork = True
        
        If ckcCompType(0).Value = vbUnchecked Then            ' barter
            tmStatusOptions.iCompBarter = False
        End If
        If ckcCompType(1).Value = vbUnchecked Then            'pay affiliate
            tmStatusOptions.iCompPayStation = False
        End If
        If ckcCompType(2).Value = vbUnchecked Then            'pay network
            tmStatusOptions.iCompPayNetwork = False
        End If
        
      
        blExportOK = gBuildAstForStationComp(hmAst, hmFileHandle, sStartDate, sEndDate, blUseAirDAte, lbcVehicles, lbcStations, tmStatusOptions)
        
        gUserActivityLog "E", "Station Compensation Export"
        
        Close #hmFileHandle
        If blExportOK Then
            lacResult.Caption = "Export stored in- " & slExportName
            gLogMsg "** Completed Station Compensation export: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportStationComp.txt", False
        Else
            lacResult.Caption = "Export not completed - an error has occurred:  See ExportStationComp.txt"
            gLogMsg "** Station Compensation export terminated: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportStationComp.txt", False

        End If
        lacResult.Visible = True
        cmdExport.Enabled = False
        Screen.MousePointer = vbDefault
        Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmExportStationComp-cmdExport"

WriteColumnHeadingerr:
    iRet = Err.Number
    gLogMsg "** Cannot write columnn heading in ExportStationComp:cmdExport, export Terminated: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportStationComp.csv", False
'    lbcInfo(0).AddItem "Cannot write columnn heading in mInitOutputFiles, export Terminated"
    Exit Sub
    
cmcExportDupNameErr:
     iRet = Err.Number
    Resume Next
End Sub

Private Sub cmdReturn_Click()
    frmReports.Show
    Unload frmExportStationComp
End Sub

Private Sub Form_Activate()
    'grdVehAff.Columns(0).Width = grdVehAff.Width
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.25
    Me.Height = Screen.Height / 1.4
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmExportStationComp
    gCenterForm frmExportStationComp
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim slRepeat As String * 1
    
    
    frmExportStationComp.Caption = "Affiliate Export Station Compensation - " & sgClientName
    
    imckcAllVehiclesIgnore = False
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")

    mInit
    Exit Sub
cmcExportDupNameErr:
    ilRet = 1
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_att.Close
    rst_Shtt.Close
    Set frmExportStationComp = Nothing
End Sub



Private Sub lbcStations_Click()
If imChkStationIgnore Then
        Exit Sub
    End If
    If ckcAllStations.Value = vbChecked Then
        imChkStationIgnore = True
        ckcAllStations.Value = vbUnchecked    'chged from false to 0 10-22-99
        imChkStationIgnore = False
    End If
    mSetCommands
End Sub

Private Sub lbcVehicles_Click()
Dim iLoop As Integer
Dim ilVefCode As Integer
Dim llShfCode As Long
Dim ilRet As Integer

    If imckcAllVehiclesIgnore Then
        Exit Sub
    End If
    If ckcAllVehicles.Value = vbChecked Then
        imckcAllVehiclesIgnore = True
        ckcAllVehicles.Value = vbUnchecked
        imckcAllVehiclesIgnore = False
    End If
'    If lbcVehicles.SelCount > 1 Then
'        lbcStations.Clear
'        For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
'            If tgStationInfo(iLoop).sUsedForATT = "Y" Then
'                lbcStations.AddItem Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
'                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(iLoop).iCode
'            End If
'        Next iLoop
'    Else
'        'only one, show only the ones with agreements
'        lbcStations.Clear
'        ilVefCode = lbcVehicles.ItemData(lbcVehicles.ListIndex)
'        SQLQuery = "Select distinct attshfcode from att where attvefcode = " & ilVefCode
'        Set rst_Shtt = gSQLSelectCall(SQLQuery)
'        While Not rst_Shtt.EOF
'            llShfCode = gBinarySearchStationInfoByCode(rst_Shtt!attshfCode)
'            If llShfCode <> -1 Then
'                lbcStations.AddItem Trim$(tgStationInfoByCode(llShfCode).sCallLetters) & ", " & Trim$(tgStationInfoByCode(llShfCode).sMarket)
'                lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfoByCode(llShfCode).iCode
'            End If
'
'        rst_Shtt.MoveNext
'        Wend
'    End If
'    ckcAllStations.Value = vbUnchecked
    
        'mSetStations
        mSetCommands
End Sub
'
'       After the vehicles are populated, determine which ones are defaulted on for the export.
'       If none on, put them all on
'
Private Sub mInit()
Dim ilLoop As Integer
Dim ilRet As Integer
Dim ilVefCode As Integer
Dim ilVff As Integer
Dim llShfCode As Long


        For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
             lbcVehicles.AddItem Trim$(tgVehicleInfo(ilLoop).sVehicle)
             lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(ilLoop).iCode
         Next ilLoop
         
         ckcAllVehicles.Value = vbUnchecked
         lbcStations.Clear
    
        
        For ilLoop = 0 To lbcVehicles.ListCount - 1 Step 1
            'ilVefCode = tgVehicleInfo(ilLoop).iCode
            ilVefCode = lbcVehicles.ItemData(ilLoop)
            If ilVefCode <= 0 Then
                Exit Sub
            End If
            ilVff = gBinarySearchVff(ilVefCode)
            If ilVff <> -1 Then
                If Trim$(tgVffInfo(ilVff).sStationComp) = "Y" Then
                    lbcVehicles.Selected(ilLoop) = True
                End If
            End If
                 
        Next ilLoop
    
        If lbcVehicles.SelCount <= 0 Then
            ckcAllVehicles.Value = vbChecked
        End If
        
        mSetStations
        'ckcAllStations.Value = vbChecked
        gInitStatusSelections tmStatusOptions               '3-14-12 set all options to exclude before interrogating the list box of selections

End Sub
'
'               set stations on based on vehicle(s) selected
'               if only 1 vehicle selected,get the associated stations
'               if more than 1 vehicle selected, select all stations that are related to agreements
'
Private Sub mSetStations()
Dim ilLoop As Integer
Dim ilVefCode As Integer
Dim llShfCode As Long

        'If lbcVehicles.SelCount > 1 Then
        'set the stations using agreements only once
            lbcStations.Clear
            For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                If tgStationInfo(ilLoop).sUsedForATT = "Y" Then
                    lbcStations.AddItem Trim$(tgStationInfo(ilLoop).sCallLetters) & ", " & Trim$(tgStationInfo(ilLoop).sMarket)
                    lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(ilLoop).iCode
                End If
            Next ilLoop
'        Else
'            'only one, show only the ones with agreements
'            lbcStations.Clear
'            ilVefCode = lbcVehicles.ItemData(lbcVehicles.ListIndex)
'            SQLQuery = "Select distinct attshfcode from att where attvefcode = " & ilVefCode
'            Set rst_Shtt = gSQLSelectCall(SQLQuery)
'            While Not rst_Shtt.EOF
'                llShfCode = gBinarySearchStationInfoByCode(rst_Shtt!attshfCode)
'                If llShfCode <> -1 Then
'                    lbcStations.AddItem Trim$(tgStationInfoByCode(llShfCode).sCallLetters) & ", " & Trim$(tgStationInfoByCode(llShfCode).sMarket)
'                    lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfoByCode(llShfCode).iCode
'                End If
'
'            rst_Shtt.MoveNext
'            Wend
'        End If
        mSetCommands
        Exit Sub
End Sub
Private Sub mSetCommands()
'  dates must be entered and at least one vehicle/station selected before Export button is enabled
    Dim ilEnable As Integer
    Dim ilLoop As Integer

    ilEnable = False
    If (CalOnAirDate.Text <> "") And (CalOffAirDate.Text <> "") Then
        ilEnable = True
        If lbcVehicles.SelCount <= 0 Then
            ilEnable = False
        End If
        If lbcStations.SelCount <= 0 Then
            ilEnable = False
        End If
    End If
    
    cmdExport.Enabled = ilEnable
End Sub
