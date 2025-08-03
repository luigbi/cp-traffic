VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form EngrSourceInUseRpt 
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   8160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8160
   Begin VB.Frame frcOption 
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
      Height          =   3660
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   7575
      Begin VB.CheckBox ckcAllBusNames 
         Caption         =   "All Bus Names"
         Height          =   255
         Left            =   3600
         TabIndex        =   26
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox edcTimeTo 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   22
         Text            =   "23:59:59"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox edcTimeFrom 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   20
         Text            =   "00:00:00"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Frame frcInUse 
         Caption         =   "Sources In-Use/Available"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton optInUse 
            Caption         =   "In-Use"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optInUse 
            Caption         =   "Available"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox edcTo 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox edcFrom 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox ckcAllAudioTypes 
         Caption         =   "All Audio Types"
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.ListBox lbcBus 
         Height          =   1230
         ItemData        =   "EngrSourceInUseRpt.frx":0000
         Left            =   3600
         List            =   "EngrSourceInUseRpt.frx":0002
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2280
         Width           =   3615
      End
      Begin VB.ListBox lbcATE 
         Height          =   1230
         ItemData        =   "EngrSourceInUseRpt.frx":0004
         Left            =   3600
         List            =   "EngrSourceInUseRpt.frx":0006
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   600
         Width           =   3555
      End
      Begin VB.Label lacTimeTo 
         Caption         =   "To"
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lacTimeFrom 
         Caption         =   "From"
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lactimes 
         Caption         =   "Times-"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lacTo 
         Caption         =   "To"
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lacFrom 
         Caption         =   "From"
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lacDates 
         Caption         =   "Dates-"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3240
      Top             =   960
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5670
      FormDesignWidth =   8160
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4455
      TabIndex        =   9
      Top             =   1200
      Width           =   1920
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Report List"
      Height          =   375
      Left            =   4275
      TabIndex        =   8
      Top             =   720
      Width           =   2310
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   4050
      TabIndex        =   7
      Top             =   240
      Width           =   2685
   End
   Begin VB.Frame frcOutput 
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
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox cboFileType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1065
         TabIndex        =   4
         Top             =   690
         Width           =   1725
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Station Preference"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "File"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   720
         Width           =   870
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Print"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   480
         Width           =   2190
      End
      Begin VB.OptionButton optRptDest 
         Caption         =   "Display"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2310
      End
   End
End
Attribute VB_Name = "EngrSourceInUseRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
'*  EngrSourceInUseRpt - Create a report to show Audio Sources in-use.
'                   Selectable by Dates, Times, Audio types and Buses in a schedule.
'                   Multiple audio type selectivity is provided.  If more than 1
'                   audio type is selected and the event has audio sources defined
'                   for more than 1 audio type, that event will be shown for each
'                   audio type.  Bus is single selection.
'                   The report is sorted by date, audio type and time.  New page for
'                   each audio type/date.
'
'
'   Audio Source available option  not yet coded
'*
'*  Created September,  2004
'*
'*  Copyright Counterpoint Software, Inc.
'****************************************************************************
Option Explicit
Private hmSEE As Integer

Dim WithEvents rstSource As ADODB.Recordset
Attribute rstSource.VB_VarHelpID = -1
Dim imAudioListBoxIgnore As Integer        'All audio types check box flag
Dim imBusListBoxIgnore As Integer
Dim tmSHE As SHE


Private Type BUSSELECTED
    iBdeCode As Integer
    sBusName As String
End Type



'
'
'       Populate the Audio Types in global array
'
Private Sub mPopATE()
    Dim ilRet As Integer
    Dim ilLoop As Integer

    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrSourceInUseRpt-mPopulate Audio Types", tgCurrATE())
    lbcATE.Clear
    For ilLoop = 0 To UBound(tgCurrATE) - 1 Step 1
        lbcATE.AddItem Trim$(tgCurrATE(ilLoop).sName)
        lbcATE.ItemData(lbcATE.NewIndex) = tgCurrATE(ilLoop).iCode
    Next ilLoop
    
End Sub
'
'           Populate the Relay Name records in global array
'
Private Sub mPopRNE()
    Dim ilRet As Integer
        Dim ilLoop As Integer
        ilRet = gGetTypeOfRecs_RNE_RelayName("C", sgCurrRNEStamp, "EngrSourceInUseRpt-mPopRNE Relay", tgCurrRNE())
    Exit Sub
End Sub
'
'           Populate the  bus definitions in global array
'           Also, place in list box for selectivity
'
Private Sub mPopBDE()
        Dim ilRet As Integer
        Dim ilLoop As Integer
        ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrSourceInUseRpt-mPopBDE Bus Definition", tgCurrBDE())
        lbcBus.Clear
        For ilLoop = 0 To UBound(tgCurrBDE) - 1 Step 1
            If tgCurrBDE(ilLoop).sState = "A" Then
                lbcBus.AddItem Trim$(tgCurrBDE(ilLoop).sName)
                lbcBus.ItemData(lbcBus.NewIndex) = tgCurrBDE(ilLoop).iCode
            End If
        Next ilLoop
    Exit Sub
End Sub
'
'           Populate the Audio Names in global array
'
Private Sub mPopANE()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrSchd-mPopASE Audio Audio Names", tgCurrANE())
    Exit Sub
End Sub
Private Sub ckcAllAudioTypes_Click()

Dim iValue As Integer
Dim lRg As Long
Dim lRet As Long
    'All check box has been selected or deselected
    If imAudioListBoxIgnore Then
        Exit Sub
    End If
    If ckcAllAudioTypes.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    
    If lbcATE.ListCount > 0 Then        'if at least 1 audio type, set the off or on
        imAudioListBoxIgnore = True
        lRg = CLng(lbcATE.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcATE.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAudioListBoxIgnore = False
    End If
    

End Sub

Private Sub ckcAllBusNames_Click()
Dim iValue As Integer
Dim lRg As Long
Dim lRet As Long
    'All check box has been selected or deselected
    If imBusListBoxIgnore Then
        Exit Sub
    End If
    If ckcAllBusNames.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    
    If lbcBus.ListCount > 0 Then        'if at least 1 audio type, set the off or on
        imBusListBoxIgnore = True
        lRg = CLng(lbcBus.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcBus.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imBusListBoxIgnore = False
    End If

End Sub

Private Sub cmdDone_Click()
    Unload EngrSourceInUseRpt
End Sub

Private Sub cmdReport_Click()

    Dim ilRet As Integer            'return error from subs/functions
    Dim ilExportType As Integer     'SAVE-TO output type
    Dim ilRptDest As Integer        'output to display, print, save to
    Dim slRptName As String         'full report name of crystal .rpt
    Dim slExportName As String      'name given to a SAVE-TO file
    Dim slSQLQuery As String        'formatting of sql query for selective libraries
    Dim ilLoop As Integer           'temp variable
    Dim slDate As String
    Dim slSQLFromDate As String     'user entered full from date for formatting sql call
    Dim slSQLToDAte As String       'user entered full to date for formatting sql call
    Dim llLoopType As Long
    Dim slSQLDateQuery As String    'formttted sql string of user entered dates
    Dim slSqlSubQuery As String     'formattied sql string for subnames selection
    Dim slStr As String             'temp string handling
    Dim ilLoopBus As Integer
    Dim llSEE As Long
    Dim ilASE As Integer
    Dim ilFound As Integer
    Dim llStartDate As Long         'earliest date entered by user
    Dim llEndDate As Long           'latest date entered by user
    Dim llLoopDate As Long          'looping date for earliest/latest date
    Dim tlSEE As SEE                'Schedule event image
    Dim tlCTE As CTE                'Comments image
    Dim ilBdeCode As Integer        'bus definition code selected by user
    Dim ilANE As Integer
    Dim ilANEProt As Integer
    Dim ilANEBU As Integer
    Dim ilRNE As Integer
    Dim slBusName As String         'bus name selected by user (stored for report)
    Dim llStartTime As Long         'earliest time entered by user for time test filtering
    Dim llEndTime As Long           'latest time entered by user for time test filtering
    Dim slStartEndTimes As String   'report header info
    Dim ilPrimATECode As Integer
    Dim ilBUATECode As Integer
    Dim ilProtATECode As Integer
    Dim ilATE As Integer
    Dim slPrimAudioName As String       'primary audio name for event to be reported
    Dim slProtAudioName As String       'protection audio name for event to be reported
    Dim slBUAudioName As String         'backup audio name for event to be reported
    Dim slPrimAudioType As String       'primary audio type for event to be reported
    Dim slProtAudioType As String       'protection audio type for event to be reported
    Dim slBUAudioType As String         'backup audio name for event to be reported
    Dim ilUserATE As Integer
    Dim ilLoopATE As Integer
    Dim slAudioTypeForSort As String    'audio type name for report sorting since more than
                                        'one audio type can be applied to 1 event
    Dim llResult As Long
    Dim ilLoopOnBusSelected As Integer
    Dim tlBusSelected() As BUSSELECTED
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass


    If optRptDest(0).Value = True Then
       ilRptDest = 0
    ElseIf optRptDest(1).Value = True Then
        ilRptDest = 1
    ElseIf optRptDest(2).Value = True Then
        ilRptDest = 2
        ilExportType = cboFileType.ListIndex
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    slSQLFromDate = gEditDateInput(edcFrom.text, "1/1/1970")   'check if from date is valid; if no date entered, set earliest possible
    If slSQLFromDate = "" Or slSQLFromDate = "1/1/1970" Then  'if no returned date, its invalid
        edcFrom.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    sgCrystlFormula2 = "Date(" + Format$(slSQLFromDate, "yyyy") + "," + Format$(slSQLFromDate, "mm") + "," + Format$(slSQLFromDate, "dd") + ")"

    
    slSQLToDAte = gEditDateInput(edcTo.text, "12/31/2069")   'check if from date is valid; if no date entered, set latest possible
    If (slSQLToDAte = "" Or slSQLToDAte = "12/31/2069") Then   'if no returned date , its invalid
        'assume end date is same as start date
        slSQLToDAte = slSQLFromDate
    End If
    sgCrystlFormula3 = "Date(" + Format$(slSQLToDAte, "yyyy") + "," + Format$(slSQLToDAte, "mm") + "," + Format$(slSQLToDAte, "dd") + ")"

    llStartDate = gDateValue(slSQLFromDate)
    llEndDate = gDateValue(slSQLToDAte)
    
    'see if dates entered in correct order
    If llStartDate > llEndDate Then     'start date falls after end date, error
        edcFrom.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'verify start time input
    slStr = edcTimeFrom.text
    If Not gIsTime(slStr) Then
        edcTimeFrom.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    slStr = gFormatTime(slStr)
    llStartTime = gStrTimeInTenthToLong(slStr, False)
    slStartEndTimes = slStr
    
    'verify end time input
    slStr = Trim$(edcTimeTo.text)
    If Not gIsTime(slStr) Then
        edcTimeTo.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    slStr = gFormatTime(slStr)
    llEndTime = gStrTimeInTenthToLong(slStr, True)
    slStartEndTimes = slStartEndTimes & "-" & slStr     'save for report header
    
    If lbcATE.SelCount < 1 Or lbcBus.SelCount < 1 Then              'no audio types or buses selected
        MsgBox "At least 1 Audio Type and Bus Name must be selected"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If optInUse(0).Value = True Then       'select sources in-use
        sgCrystlFormula4 = "'I'"
    Else
        sgCrystlFormula4 = "'A'"              'select sources available
    End If
    
    'loop thru the bus list to determine which one selected from user
'    For ilLoopBus = 0 To lbcBus.ListCount - 1
'        If lbcBus.Selected(ilLoopBus) Then          'test if user selected this entry
'            ilBdeCode = lbcBus.ItemData(ilLoopBus)
'            ilFound = True
'            ilLoop = gBinarySearchBDE(ilBdeCode, tgCurrBDE())
'            If ilLoop <> -1 Then
'                slBusName = Trim$(tgCurrBDE(ilLoop).sName)
'            End If
'            Exit For
'        End If
'    Next ilLoopBus
    
    ReDim tlBusSelected(0 To 0) As BUSSELECTED
    For ilLoopOnBusSelected = 0 To lbcBus.ListCount - 1
        If lbcBus.Selected(ilLoopOnBusSelected) Then
            tlBusSelected(UBound(tlBusSelected)).iBdeCode = lbcBus.ItemData(ilLoopOnBusSelected)
            tlBusSelected(UBound(tlBusSelected)).sBusName = lbcBus.List(ilLoopOnBusSelected)
            ReDim Preserve tlBusSelected(0 To UBound(tlBusSelected) + 1) As BUSSELECTED
        End If
    Next ilLoopOnBusSelected
    
    Set rstSource = New Recordset
    rstSource.Fields.Append "AirDate", adChar, 10           'date of schedule
    rstSource.Fields.Append "StartdateSort", adInteger, 4    'date used for sorting within library name
    rstSource.Fields.Append "Bus", adChar, 20               'bus name
    rstSource.Fields.Append "TimesHdr", adChar, 20          'start & end times entered for header info
    rstSource.Fields.Append "AudioTypeforSort", adChar, 20           'audio type
    rstSource.Fields.Append "PrimSource", adChar, 8         'primary source name
    rstSource.Fields.Append "PrimAudioType", adChar, 20       'primary audio type
    rstSource.Fields.Append "PgmName", adChar, 66            'title 1
    rstSource.Fields.Append "StartTime", adChar, 9           'event time in milliseconds (string)
    rstSource.Fields.Append "StartTimeSort", adInteger, 4    'eent time used for sorting time within day
    rstSource.Fields.Append "ItemID", adChar, 32            'itemID
    rstSource.Fields.Append "ProtSource", adChar, 8        'protection source name
    rstSource.Fields.Append "ProtAudioType", adChar, 20      'protection audio type
    rstSource.Fields.Append "BUSource", adChar, 8          'backup source name
    rstSource.Fields.Append "BUAudioType", adChar, 20        'backup audio type
    rstSource.Fields.Append "Duration", adChar, 10           'event duration in milliseconds
    rstSource.Fields.Append "Relay", adChar, 8              'relay
    rstSource.Open
      
    'build the data definition (.ttx) file in the database path for crystal to access
    llResult = CreateFieldDefFile(rstSource, sgDBPath & "\SourcesUsed.ttx", True)
    
    'loop on user requested start & end dates
    For llLoopDate = llStartDate To llEndDate
        slDate = Format$(llLoopDate, sgShowDateForm)
        ilRet = gGetRec_SHE_ScheduleHeaderByDate(slDate, "EngrSourceInUseReport-cmdReport gGetScheduleHeaderbyDate", tmSHE)
        ilRet = gGetRecs_SEE_ScheduleEventsAPI(hmSEE, sgCurrSEEStamp, -1, tmSHE.lCode, "EngrSourceInUseReport-cmdReport gGetRecs_SEE_ScheduleEvents", tgCurrSEE())
        
        'loop thru the schedules for the day:  must match the audio type and bus and event times must
        'be within user input times
        For llSEE = LBound(tgCurrSEE) To UBound(tgCurrSEE) - 1
            tlSEE = tgCurrSEE(llSEE)

            For ilLoopOnBusSelected = LBound(tlBusSelected) To UBound(tlBusSelected) - 1
                ilBdeCode = tlBusSelected(ilLoopOnBusSelected).iBdeCode
                If ilBdeCode = tlSEE.iBdeCode And tlSEE.lTime >= llStartTime And tlSEE.lTime <= llEndTime Then         'found valid bus to report
                    ilFound = False
                    slPrimAudioName = ""
                    slProtAudioName = ""
                    slBUAudioName = ""
                    slPrimAudioType = ""
                    slProtAudioType = ""
                    slBUAudioType = ""
                    ilPrimATECode = 0
                    ilProtATECode = 0
                    ilBUATECode = 0
                    'find the matching audio source record so that the audio names and audio types can be obtained
                    'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                    '    If tlSEE.iAudioAseCode = tgCurrASE(ilASE).iCode Then
                        ilASE = gBinarySearchASE(tlSEE.iAudioAseCode, tgCurrASE())
                        If ilASE <> -1 Then
    
                            'found the matching audio source record, find the matching audio name record; then from
                            'the audio name record go to the audio type record to get the audio name
                            For ilANE = LBound(tgCurrANE) To UBound(tgCurrANE) - 1
                                If tgCurrANE(ilANE).iCode = tgCurrASE(ilASE).iPriAneCode Then       'found the matching audio source record
                                    slPrimAudioName = Trim$(tgCurrANE(ilANE).sName)
                                    ilPrimATECode = tgCurrANE(ilANE).iAteCode
                                    For ilATE = LBound(tgCurrATE) To UBound(tgCurrATE) - 1
                                        If ilPrimATECode = tgCurrATE(ilATE).iCode Then
                                            slPrimAudioType = tgCurrATE(ilATE).sName
                                            Exit For
                                        End If
                                    Next ilATE
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilANE
                        End If
                    '    If ilFound Then
                    '        Exit For
                    '    End If
                    'Next ilASE
                    
                    
                    'get the protection audio source name and audio source type name if applicable
                    For ilANEProt = LBound(tgCurrANE) To UBound(tgCurrANE) - 1
                        If tgCurrANE(ilANEProt).iCode = tlSEE.iProtAneCode Then        'found the matching audio source record
                            slProtAudioName = Trim$(tgCurrANE(ilANEProt).sName)
                            ilProtATECode = tgCurrANE(ilANEProt).iAteCode
                            For ilATE = LBound(tgCurrATE) To UBound(tgCurrATE) - 1
                                If ilProtATECode = tgCurrATE(ilATE).iCode Then
                                    slProtAudioType = tgCurrATE(ilATE).sName
                                    Exit For
                                End If
                            Next ilATE
                            Exit For
                        End If
                    Next ilANEProt
                    
                    'get the backup audio source name and audio source type name if applicable
                    For ilANEBU = LBound(tgCurrANE) To UBound(tgCurrANE) - 1
                        If tgCurrANE(ilANEBU).iCode = tlSEE.iBkupAneCode Then       'found the matching audio source record
                            slBUAudioName = Trim$(tgCurrANE(ilANEBU).sName)
                            ilBUATECode = tgCurrANE(ilANEBU).iAteCode
                            For ilATE = LBound(tgCurrATE) To UBound(tgCurrATE) - 1
                                If ilBUATECode = tgCurrATE(ilATE).iCode Then
                                    slBUAudioType = tgCurrATE(ilATE).sName
                                    Exit For
                                End If
                            Next ilATE
                            Exit For
                        End If
                    Next ilANEBU
                    
                    
                    'loop thru the list of audio types to see which ones user selected.
                    'for each audio type, see if its valid in the primary, protection or backup audio source.
                    'if the sources for an event exists for more than 1 type selected, show it
                    'in all places.
                    'For example:  Audio Types M & R selected.
                    '              Event consists of  Primary audio source REPL5  (audio type R)
                    '                    and          Protection audio source Line 4 (audio type M)
                    '  this event will be shown twice.
                    For ilLoopATE = 0 To lbcATE.ListCount - 1
                        If lbcATE.Selected(ilLoopATE) Then
                            slAudioTypeForSort = ""
                            For ilUserATE = LBound(tgCurrATE) To UBound(tgCurrATE) - 1
                                If lbcATE.ItemData(ilLoopATE) = tgCurrATE(ilUserATE).iCode Then
                                slAudioTypeForSort = tgCurrATE(ilUserATE).sName     'audio type used for sorting report, inc case
                                                                                    'multiple audio types requested and 1 event has more than 1 audio type assigned
                                End If
                                
                            Next ilUserATE
                            
                            For ilUserATE = 1 To 3
                                ilFound = False
                                If ilUserATE = 1 Then
                                    If lbcATE.ItemData(ilLoopATE) = ilPrimATECode Then
                                       ilFound = True
                                    End If
                                ElseIf ilUserATE = 2 Then
                                    If lbcATE.ItemData(ilLoopATE) = ilProtATECode Then
                                        ilFound = True
                                    End If
                                Else
                                    If lbcATE.ItemData(ilLoopATE) = ilBUATECode Then
                                        ilFound = True
                                    End If
                                End If
                                
                                If ilFound Then
                                    rstSource.AddNew
                                    rstSource.Fields("ItemID") = Trim(tlSEE.sAudioItemID)     'Item  Event ID()
                                    rstSource.Fields("AirDate") = slDate        'tmSHE.sAirDate
                                    rstSource.Fields("TimesHdr") = Trim$(slStartEndTimes)
                                    rstSource.Fields("Bus") = Trim$(tlBusSelected(ilLoopOnBusSelected).sBusName)    '(slBusName)
                                    rstSource.Fields("StartTime") = gFormatTime(gLongToTime(tlSEE.lTime))          'event start time in milliseconds
                                    rstSource.Fields("StartTimeSort") = tlSEE.lTime          'event start time in milliseconds
                                    rstSource.Fields("Duration") = gLongToStrLengthInTenth(tlSEE.lDuration, True)
                                    rstSource.Fields("PgmName") = ""
                                    If tlSEE.l1CteCode > 0 Then
                                        ilRet = gGetRec_CTE_CommtsTitle(tlSEE.l1CteCode, "EngrSourceInUseReport-cmdReport gGetRec_CTE_CommtsTitle", tlCTE)
                                        rstSource.Fields("PgmName") = Trim(tlCTE.sComment)          'title 1 (program name)
                                    End If
                                        
                                    'get the relay from schedule
                                    rstSource.Fields("Relay") = ""
                                    For ilRNE = LBound(tgCurrRNE) To UBound(tgCurrRNE) - 1
                                        If tlSEE.i1RneCode = tgCurrRNE(ilRNE).iCode Then
                                            rstSource.Fields("Relay") = tgCurrRNE(ilRNE).sName
                                            Exit For
                                        End If
                                    Next ilRNE
                                    
                                    rstSource.Fields("AudioTypeforSort") = slAudioTypeForSort     'Audio Type name
                                    'get the Primary Audio source name
                                    rstSource.Fields("PrimSource") = slPrimAudioName     'primary audio source name
                                    rstSource.Fields("ProtSource") = slProtAudioName     'protection audio source name
                                    rstSource.Fields("BUSource") = slBUAudioName         'backup audio source name
                                    rstSource.Fields("PrimAudioType") = slPrimAudioType
                                    rstSource.Fields("ProtAudioType") = slProtAudioType
                                    rstSource.Fields("BUAudiotype") = slBUAudioType
                                End If
                            Next ilUserATE
                        End If              'lbcATE.selected
                    Next ilLoopATE          '0 to lbcate.listcount-1
                    
                        
                End If                  'ilbdecod6e = tlsee.ibdecode
            Next ilLoopOnBusSelected
        Next llSEE
    Next llLoopDate
    
   
    
    igRptSource = vbModeless
    gObtainReportforCrystal slRptName, slExportName     'determine which .rpt to call and setup an export name is user selected output to export
    EngrCrystal.gActiveCrystalReports ilExportType, ilRptDest, slRptName & ".rpt", slExportName, rstSource
    'Set fNewForm.Report = Appl.OpenReport(sgReportDirectory + slRptName & "Sum.rpt")
    'fNewForm.Report.Database.Tables(1).SetDataSource rstLibrary, 3
    'fNewForm.Show igRptSource
    
    Screen.MousePointer = vbDefault
    

    'rstLibrary.Close           'causes error when closed
    Set rstSource = Nothing
    If igRptSource = vbModal Then
        Unload EngrSourceInUseRpt
    End If
    
    Exit Sub
    

    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in User Rpt-cmdReport: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
            Screen.MousePointer = vbDefault
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in User Rpt-cmdReport: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cmdReturn_Click()
    EngrReports.Show
    Unload EngrSourceInUseRpt
End Sub

Private Sub edcFrom_GotFocus()
    gCtrlGotFocus edcFrom
End Sub
Private Sub edcTimeFrom_GotFocus()
   gCtrlGotFocus edcTimeFrom
End Sub
Private Sub edcTimeTo_GotFocus()
   gCtrlGotFocus edcTimeTo
End Sub
Private Sub edcTo_GotFocus()
    If edcTo.text = "" Then
        edcTo.text = edcFrom.text
    End If
    gCtrlGotFocus edcTo
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    gSetFonts EngrSourceInUseRpt
    gCenterForm EngrSourceInUseRpt
End Sub
Private Sub Form_Load()
Dim ilRet As Integer

On Error GoTo ErrHand:
    mInit
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors  'rdoErrors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Bus Definition-Form Load: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Bus Definition-Form Load: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    btrDestroy hmSEE
    
    Set EngrSourceInUseRpt = Nothing
End Sub

Private Sub lbcATETypes_Click()

    If imAudioListBoxIgnore Then
        Exit Sub
    End If
    If ckcAllAudioTypes.Value = vbChecked Then
        imAudioListBoxIgnore = True
        ckcAllAudioTypes.Value = False
        imAudioListBoxIgnore = False
    End If
     
    Exit Sub
End Sub


Private Sub lbcATE_Click()
        If imAudioListBoxIgnore Then
            Exit Sub
        End If
        If ckcAllAudioTypes.Value = vbChecked Then
            imAudioListBoxIgnore = True
            ckcAllAudioTypes.Value = False
            imAudioListBoxIgnore = False
        End If
        Exit Sub
End Sub

Private Sub lbcBus_Click()
 If imBusListBoxIgnore Then
            Exit Sub
        End If
        If ckcAllBusNames.Value = vbChecked Then
            imBusListBoxIgnore = True
            ckcAllBusNames.Value = False
            imBusListBoxIgnore = False
        End If
        Exit Sub
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to adobe
    Else
        cboFileType.Enabled = False
    End If
End Sub

Private Sub mInit()
    Dim ilRet As Integer
    hmSEE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSEE, "", sgDBPath & "SEE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    gPopExportTypes cboFileType  'populate the valid export types
    cboFileType.Enabled = False  'disable export file types until SAVE TO selected
    gChangeCaption frcOption     'show report name caption on selectivity box
    mPopBDE
    mPopANE
    mPopATE
    mPopASE
    mPopRNE
    Exit Sub
End Sub

'
'           mFilterSelectivity - determine if this library has user selected dates,
'           correct subnames and bus groups
Private Function mfilterSelectivity(slSQLFromDate As String, slSQLToDAte As String) As Integer
Dim llLoopTemp As Long
Dim ilValidDates As Integer

    ilValidDates = False
     'If Format$(tmDHE.sEndDate, sgSQLDateForm) > Format$(slSQLFromDate, sgSQLDateForm) And Format$(tmDHE.sStartDate, sgSQLDateForm) <= Format$(slSQLToDAte, sgSQLDateForm) Then
        'insure the dates of this library should be included
        ilValidDates = True
    'End If
    
    If ilValidDates Then      'passed all tests
        mfilterSelectivity = True
    Else
        mfilterSelectivity = False
    End If
End Function

Private Sub mPopASE()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrSourceInUseRpt-mInit", tgCurrASE())
    Exit Sub
End Sub

