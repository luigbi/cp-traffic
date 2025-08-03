VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form EngrActivityRpt 
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   8100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8100
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
      Top             =   1860
      Width           =   7575
      Begin VB.CheckBox ckcAll 
         Caption         =   "All Users"
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Top             =   120
         Width           =   1455
      End
      Begin VB.ListBox lbcUsers 
         Height          =   2985
         ItemData        =   "EngrActRpt.frx":0000
         Left            =   4200
         List            =   "EngrActRpt.frx":0002
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   480
         Width           =   3075
      End
      Begin VB.TextBox edcToTime 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox edcFromTime 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox edcTo 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox edcFrom 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lacChangeFromTime 
         Caption         =   "From"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lacChangeFromDate 
         Caption         =   "From"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lacToText 
         Caption         =   "To"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lacChangeTimes 
         Caption         =   "Enter change times-"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lacChangeTo 
         Caption         =   "To"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lacChangeDates 
         Caption         =   "Enter change dates-"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2055
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
      FormDesignWidth =   8100
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
Attribute VB_Name = "EngrActivityRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  EngrActivityRpt - a report of Activity options
'*
'*  Created September,  2004
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Dim imChkListBoxIgnore As Integer
Dim tmUserInfo() As UIE
Dim tmAIE As AIE

Private Sub ckcAll_Click()
Dim lErr As Long
Dim iValue As Integer
Dim lRg As Long
Dim lRet As Long
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If ckcAll.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    
    If lbcUsers.ListCount > 0 Then
        imChkListBoxIgnore = True
        lRg = CLng(lbcUsers.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcUsers.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imChkListBoxIgnore = False
    End If
     
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDone_Click()
    Unload EngrActivityRpt
End Sub

Private Sub cmdReport_Click()
    Dim iType As Integer
    Dim sOutput As String
    Dim ilRet As Integer
    Dim ilExportType As Integer
    Dim ilRptDest As Integer
    Dim slRptName As String
    Dim slExportName As String
    Dim SQLQuery As String
    Dim ilListIndex As Integer
    Dim slSQLFromDate As String
    Dim slSQLToDAte As String
    Dim slDate As String
    Dim ilLoop As Integer
    Dim slTime As String
    Dim slSQLFromTime As String
    Dim slSQLToTime As String
    Dim slUsers As String       'string of selective users
    Dim slAIEStamp As String
    Dim llLoopAIE As Long
    Dim tlAIE() As AIE
    Dim llResult As Long
    Dim slRptType As String
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass

    '*** 6-8-11 DISABLED, AIE FILE NOT IN USE
    
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
    
    
    slSQLFromDate = gEditDateInput(edcFrom.text, "1/1/1970")   'check if from date is valid; if no date entered set as earliest possible
    If slSQLFromDate = "" Then  'if no returned date, its invalid
        edcFrom.SetFocus
        Exit Sub
    End If

    sgCrystlFormula2 = "Date(" + Format$(slSQLFromDate, "yyyy") + "," + Format$(slSQLFromDate, "mm") + "," + Format$(slSQLFromDate, "dd") + ")"
     
    slSQLToDAte = gEditDateInput(edcTo.text, "12/31/2069")   'check if to date is valid; if no date entered, set as latest possible
    If slSQLToDAte = "" Then  'if no returned date, its invalid
        edcTo.SetFocus
        Exit Sub
    End If
     sgCrystlFormula3 = "Date(" + Format$(slSQLToDAte, "yyyy") + "," + Format$(slSQLToDAte, "mm") + "," + Format$(slSQLToDAte, "dd") + ")"
 
     slTime = edcFromTime.text
     If edcFromTime.text = "" Then
         slSQLFromTime = gConvertTime("12M")
     Else
         If Not gIsTime(slTime) Then
             Beep
             MsgBox "Invalid From Time"
             edcFromTime.SetFocus
             Screen.MousePointer = vbDefault
             Exit Sub
         End If
         slSQLFromTime = gConvertTime(slTime)
     End If

     sgCrystlFormula4 = slSQLFromTime       'formula to pass to crystal

     slTime = edcToTime.text
     If edcToTime.text = "" Then
         slSQLToTime = gConvertTime("11:59:59PM")
     Else
         If Not gIsTime(slTime) Then
             Beep
             MsgBox "Invalid To Time"
             edcToTime.SetFocus
             Screen.MousePointer = vbDefault
             Exit Sub
         End If
         
         slSQLToTime = gConvertTime(slTime)
         If slSQLToTime = "12:00AM" Then
            slSQLToTime = "11:59:59PM"
        End If
     End If

    sgCrystlFormula5 = slSQLToTime     'to time formula to pass to crystal
    
    slUsers = ""
    If Not ckcAll.Value = vbChecked Then    'selective users
        For ilLoop = 0 To lbcUsers.ListCount - 1 Step 1
            If lbcUsers.Selected(ilLoop) Then
                If Len(slUsers) = 0 Then
                    slUsers = "(uieCode = " & lbcUsers.ItemData(ilLoop) & ")"
                Else
                    slUsers = slUsers & " OR (uieCode = " & lbcUsers.ItemData(ilLoop) & ")"
                End If
            End If
        Next ilLoop
    End If
    If slUsers <> "" Then
        slUsers = " and (" & slUsers & ")"
    End If

    Set rstAIERpt = New Recordset
    gGenerateRstAIERpt     'generate the ddfs for report
    
    rstAIERpt.Open
    'build the data definition (.ttx) file in the database path for crystal to access
    llResult = CreateFieldDefFile(rstAIERpt, sgDBPath & "\libAIE.ttx", True)
    
    ilRet = gGetTypeOfRecs_AIE_ActiveInfo("", slSQLFromDate, slSQLToDAte, slAIEStamp, "EngrActivityRpt: cmdReport", tlAIE())
    For llLoopAIE = LBound(tlAIE) To UBound(tlAIE) - 1
        If llLoopAIE = 225 Then
            ilRet = ilRet
        End If
        
        tmAIE = tlAIE(llLoopAIE)
        mProcessAIE
    Next llLoopAIE
    slRptType = ""
    gObtainReportforCrystal slRptName, slExportName     'determine which .rpt to call and setup an export name is user selected output to export
    
    slRptName = slRptName & ".rpt"      'concatenate the crystal report name plus extension
    'SQLQuery = "Select * from AIE_Active_Info, UIE_User_Info where aieuiecode = uiecode"
    'SQLQuery = SQLQuery & " and (aieEnteredDate >= '" & Format$(slSQLFromDate, sgSQLDateForm) & "' AND aieEnteredDate <= '" & Format$(slSQLToDAte, sgSQLDateForm) & "') and "
    'SQLQuery = SQLQuery & " (aieEnteredTime >=  '" & Format$(slSQLFromTime, sgSQLTimeForm) & "' AND aieEnteredTime <= '" & Format$(slSQLToTime, sgSQLTimeForm) & "')"
    'SQLQuery = SQLQuery & slUsers
    'SQLQuery = SQLQuery & " Order By aieEnteredDate desc, aieRefFileName"

    EngrCrystal.gActiveCrystalReports ilExportType, ilRptDest, Trim$(slRptName) & Trim$(slRptType), slExportName, rstAIERpt
    
    Screen.MousePointer = vbDefault
    
    Set rstAIERpt = Nothing
    If igRptSource = vbModal Then
        Unload EngrActivityRpt
    End If
    Exit Sub
    

    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Activity Rpt-cmdReport: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
            Screen.MousePointer = vbDefault
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Activity Rpt-cmdReport: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub cmdReturn_Click()
    EngrReports.Show
    Unload EngrActivityRpt
End Sub

Private Sub edcFrom_GotFocus()
    gCtrlGotFocus edcFrom
End Sub

Private Sub edcFromTime_GotFocus()
    gCtrlGotFocus edcFromTime
End Sub

Private Sub edcTo_GotFocus()
    gCtrlGotFocus edcTo
End Sub

Private Sub edcToTime_GotFocus()
    gCtrlGotFocus edcToTime
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    gSetFonts EngrActivityRpt
    gCenterForm EngrActivityRpt
End Sub

Private Sub Form_Load()
Dim ilLoop As Integer
Dim SQLQuery As String
Dim llMax As Long
Dim gMsg As String
Dim ilRet As Integer
    'EngrUserRpt.Caption = "User - " & sgClientName
    gPopExportTypes cboFileType
    cboFileType.Enabled = False
    gChangeCaption frcOption     'show report name as caption
    
    On Error GoTo ErrHand
    
    ReDim tmUserInfo(0 To 0) As UIE
    'gather only the current users and sort them by show name field to appear in the list box
    ilRet = gGetTypeOfRecs_UIE_UserInfo("C", sgCurrUIEStamp, "EngrActivityRpt", tgCurrUIE())
    
    lbcUsers.Clear
    For ilLoop = 0 To UBound(tgCurrUIE) - 1 Step 1
        lbcUsers.AddItem Trim$(tgCurrUIE(ilLoop).sShowName)
        lbcUsers.ItemData(lbcUsers.NewIndex) = tgCurrUIE(ilLoop).iCode
    Next ilLoop
    imChkListBoxIgnore = False
    Exit Sub
    
ErrHand:
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in EngrActivityRpt: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in EngrActivityRpt: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set EngrActivityRpt = Nothing
End Sub

Private Sub lbcUsers_Click()
    If imChkListBoxIgnore Then
        Exit Sub
    End If
    If ckcAll.Value = vbChecked Then
        imChkListBoxIgnore = True
        ckcAll.Value = False
        imChkListBoxIgnore = False
    End If
End Sub

Private Sub optRptDest_Click(Index As Integer)
    If optRptDest(2).Value Then
        cboFileType.Enabled = True
        cboFileType.ListIndex = 0       'default to adobe
    Else
        cboFileType.Enabled = False
    End If
End Sub
'
'
'           mProcessAIE - Generate an Activity report indicating what key type
'           of change occurred.  AIE (Activity Info Record contains the data.
'           Create prepass record (rstAIERpt) for Crystal to report from.
'
Public Sub mProcessAIE()
Dim llDate As Long
Dim ilLoop As Integer
Dim tlAEE As AEE
Dim tlANE As ANE
Dim tlAPE As APE
Dim tlASE As ASE
Dim tlATE As ATE
Dim tlBDE As BDE
Dim tlBGE As BGE
Dim tlCCE As CCE
Dim tlCTE As CTE
Dim tlDEE As DEE
Dim tlDHE As DHE
Dim tlDNE As DNE
Dim tlDSE As DSE
Dim tlETE As ETE
Dim tlEPE As EPE
Dim tlFNE As FNE
Dim tlMTE As MTE
Dim tlNNE As NNE
Dim tlRNE As RNE
Dim tlSCE As SCE
Dim tlSHE As SHE
Dim tlSEE As SEE
Dim tlSOE As SOE
Dim tlTTE As TTE
Dim tlUie As UIE

Dim slFileStamp As String
Dim ilToFileCode As Integer
Dim llToFileCode As Long
Dim ilRet As Integer
Dim slOffSet As String
Dim slEventType As String
Dim ilTempCode As Integer

    On Error GoTo ErrHand

    rstAIERpt.AddNew
    rstAIERpt.Fields("DescKeyField") = ""
    rstAIERpt.Fields("DescSecField") = ""
    llDate = gDateValue(tmAIE.sEnteredDate)
    rstAIERpt.Fields("Date") = llDate           'entered date
    rstAIERpt.Fields("Time") = Format$(tmAIE.sEnteredTime, sgSQLTimeForm)
    For ilLoop = LBound(tgCurrUIE) To UBound(tgCurrUIE) - 1
        If tgCurrUIE(ilLoop).iCode = tmAIE.iUieCode Then
            rstAIERpt.Fields("User") = Trim$(tgCurrUIE(ilLoop).sShowName)
            Exit For
        End If
    Next ilLoop

    'ilToFileCode = tmAIE.lToFileCode       '5-30-06
    llToFileCode = tmAIE.lToFileCode
    'Determine type of change
    If Trim(tmAIE.sRefFileName) = "AEE" Then          'automation equipment
        rstAIERpt.Fields("TypeOfChange") = "Automation Equipment"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_AEE_AutoEquip(ilToFileCode, slFileStamp, tlAEE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlAEE.sName) & "/" & Trim(tlAEE.sDescription)
    
    ElseIf Trim(tmAIE.sRefFileName) = "ANE" Then          'audio name
        rstAIERpt.Fields("TypeOfChange") = "Audio Name"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_ANE_AudioName(ilToFileCode, slFileStamp, tlANE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlANE.sName) & "/" & Trim(tlANE.sDescription)
  
    ElseIf Trim(tmAIE.sRefFileName) = "APE" Then          'Automation Path
        rstAIERpt.Fields("TypeOfChange") = "Automation Path"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_APE_AutoPath(ilToFileCode, slFileStamp, tlAPE)
        If Trim(tlAPE.sType) = "CE" Then        'client export
            slEventType = "Client Export Path"
        ElseIf Trim(tlAPE.sType) = "CI" Then    'client import
            slEventType = "Client Import Path"
        ElseIf Trim(tlAPE.sType) = "SI" Then    'server import
            slEventType = "Server Import Path"
        ElseIf Trim(slEventType) = "SE" Then    'server export
            slEventType = "Server Export Path"
        Else
            slEventType = Trim(tlAPE.sType) & " Automation Path Unknown"
        End If
        rstAIERpt.Fields("DescKeyField") = slEventType
        ilRet = gGetRec_AEE_AutoEquip(tlAPE.iAeeCode, slFileStamp, tlAEE)
        rstAIERpt.Fields("DescSecField") = tlAEE.sName      'automation name
        
    ElseIf Trim(tmAIE.sRefFileName) = "ASE" Then      'audio source
        rstAIERpt.Fields("TypeOfChange") = "Audio Source"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_ASE_AudioSource(ilToFileCode, slFileStamp, tlASE)
        ilRet = gGetRec_ANE_AudioName(ilToFileCode, slFileStamp, tlANE)     'get the prim audio source
        ilRet = gGetRec_ATE_AudioType(ilToFileCode, slFileStamp, tlATE)     'get the audio type
        rstAIERpt.Fields("DescKeyField") = Trim(tlASE.sDescription)
        rstAIERpt.Fields("DescSecField") = Trim(tlANE.sName) & "/" & Trim(tlANE.sDescription)
        
    ElseIf Trim(tmAIE.sRefFileName) = "ATE" Then      'audio type
        rstAIERpt.Fields("TypeOfChange") = "Audio Type"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_ATE_AudioType(ilToFileCode, slFileStamp, tlATE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlATE.sName) & "/" & Trim(tlATE.sDescription)
    
    ElseIf Trim(tmAIE.sRefFileName) = "BDE" Then      'Bus names
        rstAIERpt.Fields("TypeOfChange") = "Bus Definition"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_BDE_BusDefinition(ilToFileCode, slFileStamp, tlBDE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlBDE.sName) & "/" & Trim(tlBDE.sDescription)
    
    ElseIf Trim(tmAIE.sRefFileName) = "BGE" Then      'bus group
        rstAIERpt.Fields("TypeOfChange") = "Bus Group"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_BGE_BusGroup(ilToFileCode, slFileStamp, tlBGE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlBGE.sName) & "/" & Trim(tlBGE.sDescription)
    
    ElseIf Trim(tmAIE.sRefFileName) = "CCE" Then      'control char
        rstAIERpt.Fields("TypeOfChange") = "Control Character"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_CCE_ControlChar(ilToFileCode, slFileStamp, tlCCE)
        If tlCCE.sType = "A" Then
            rstAIERpt.Fields("DescKeyField") = "Audio Control" & "/" & Trim(tlCCE.sDescription)
        Else
            rstAIERpt.Fields("DescKeyField") = "Bus Control" & "/" & Trim(tlCCE.sDescription)
        End If

    ElseIf Trim(tmAIE.sRefFileName) = "CTE" Then      'comments
        rstAIERpt.Fields("TypeOfChange") = "Comments & Title"
        ilRet = gGetRec_CTE_CommtsTitle(llToFileCode, slFileStamp, tlCTE)
        'rstAIERpt.Fields("DescKeyField") = Trim(tlCTE.sName) & "/" & Trim(tlCTE.sComment)
        rstAIERpt.Fields("DescKeyField") = Trim(tlCTE.sComment)
    
    ElseIf Trim(tmAIE.sRefFileName) = "DEE" Then      'day event
        rstAIERpt.Fields("TypeOfChange") = "Day Event"
        ilRet = gGetRec_DHE_DayHeaderInfo(tmAIE.lOrigFileCode, slFileStamp, tlDHE)
        ilRet = gGetRec_DNE_DayName(tlDHE.lDneCode, slFileStamp, tlDNE)      'get the library name
        ilRet = gGetRec_DSE_DaySubName(tlDHE.lDseCode, slFileStamp, tlDSE)    'get the libray subname
        rstAIERpt.Fields("DescKeyField") = Trim(tlDNE.sDescription) & "/" & Trim(tlDSE.sDescription)    'library name & subname
        ilRet = gGetRec_DEE_DayEvent(llToFileCode, slFileStamp, tlDEE)
        slOffSet = Trim(gLongToStrLengthInTenth(tlDEE.lTime, False))        'event time offset to string
        ilRet = gGetRec_ETE_EventType(tlDEE.iEteCode, slFileStamp, tlETE)    'get the libray subname
        If tlETE.sCategory = "A" Then       'avail
            slEventType = "Avail"
        ElseIf tlETE.sCategory = "P" Then  'program
            slEventType = "Program"
        ElseIf tlETE.sCategory = "S" Then   'spot
            slEventType = "Spot"
        Else
            slEventType = "Unknown Event type"
        End If
        rstAIERpt.Fields("DescSecField") = Trim(slEventType) & " at " & Trim(slOffSet)
    
    ElseIf Trim(tmAIE.sRefFileName) = "DHE" Then      'day event
        rstAIERpt.Fields("TypeOfChange") = "Day Header"
        ilRet = gGetRec_DHE_DayHeaderInfo(tmAIE.lOrigFileCode, slFileStamp, tlDHE)
        ilRet = gGetRec_DNE_DayName(tlDHE.lDneCode, slFileStamp, tlDNE)      'get the library name
        ilRet = gGetRec_DSE_DaySubName(tlDHE.lDseCode, slFileStamp, tlDSE)    'get the libray subname
        rstAIERpt.Fields("DescKeyField") = Trim(tlDNE.sDescription) & "/" & Trim(tlDSE.sDescription)    'library name & subname
         
    ElseIf Trim(tmAIE.sRefFileName) = "DNE" Then      'library name
        rstAIERpt.Fields("TypeOfChange") = "Library Name"
        ilRet = gGetRec_DNE_DayName(llToFileCode, slFileStamp, tlDNE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlDNE.sName) & "/" & Trim(tlDNE.sDescription)
    ElseIf Trim(tmAIE.sRefFileName) = "DSE" Then      'library subname
        rstAIERpt.Fields("TypeOfChange") = "Library subname"
        ilRet = gGetRec_DSE_DaySubName(llToFileCode, slFileStamp, tlDSE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlDSE.sName) & "/" & Trim(tlDSE.sDescription)
    
    ElseIf Trim(tmAIE.sRefFileName) = "EPE" Then      'Event Properties
        rstAIERpt.Fields("TypeOfChange") = "Event Properties"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_EPE_EventProperties(ilToFileCode, slFileStamp, tlEPE)
        ilTempCode = tmAIE.lOrigFileCode
        ilRet = gGetRec_ETE_EventType(ilTempCode, slFileStamp, tlETE)    'get the event type
        If tlEPE.sType = "U" Then
            rstAIERpt.Fields("DescSecField") = "Used"
        Else
            rstAIERpt.Fields("DescSecField") = "Mandatory"
        End If
        
        If tlETE.sCategory = "A" Then       'avail
            slEventType = "Avail"
        ElseIf tlETE.sCategory = "P" Then  'program
            slEventType = "Program"
        ElseIf tlETE.sCategory = "S" Then   'spot
            slEventType = "Spot"
        Else
            slEventType = "Unknown Event type"
        End If
        rstAIERpt.Fields("DescKeyField") = Trim(tlETE.sName) & " for " & slEventType & " type"
        
    ElseIf Trim(tmAIE.sRefFileName) = "ETE" Then      'Event Type
        rstAIERpt.Fields("TypeOfChange") = "Event Type"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_ETE_EventType(ilToFileCode, slFileStamp, tlETE)    'get the event type
       
        rstAIERpt.Fields("DescKeyField") = Trim(tlETE.sName) & "/" & Trim(tlETE.sDescription)
        If tlETE.sCategory = "A" Then       'avail
            slEventType = "Avail"
        ElseIf tlETE.sCategory = "P" Then  'program
            slEventType = "Program"
        ElseIf tlETE.sCategory = "S" Then   'spot
            slEventType = "Spot"
        Else
            slEventType = "Unknown Event type"
        End If
        rstAIERpt.Fields("DescSecField") = slEventType
        
    ElseIf Trim(tmAIE.sRefFileName) = "FNE" Then      'follow name
        rstAIERpt.Fields("TypeOfChange") = "Follow Name"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_FNE_FollowName(ilToFileCode, slFileStamp, tlFNE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlFNE.sName) & "/" & Trim(tlFNE.sDescription)
   
    ElseIf Trim(tmAIE.sRefFileName) = "MTE" Then      'material type
        rstAIERpt.Fields("TypeOfChange") = "Material Type"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_MTE_MaterialType(ilToFileCode, slFileStamp, tlMTE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlMTE.sName) & "/" & Trim(tlMTE.sDescription)
    
    ElseIf Trim(tmAIE.sRefFileName) = "NNE" Then      'net cue
        rstAIERpt.Fields("TypeOfChange") = "Netcue Name"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_NNE_NetcueName(ilToFileCode, slFileStamp, tlNNE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlNNE.sName) & "/" & Trim(tlNNE.sDescription)
   
    ElseIf Trim(tmAIE.sRefFileName) = "RNE" Then      'relay
        rstAIERpt.Fields("TypeOfChange") = "Relay Name"
        ilRet = gGetRec_RNE_RelayName(ilToFileCode, slFileStamp, tlRNE)
        ilToFileCode = llToFileCode
        rstAIERpt.Fields("DescKeyField") = Trim(tlRNE.sName) & "/" & Trim(tlRNE.sDescription)
    
    ElseIf Trim(tmAIE.sRefFileName) = "SCE" Then      'silence
        rstAIERpt.Fields("TypeOfChange") = "Silence Character"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_SCE_SilenceChar(ilToFileCode, slFileStamp, tlSCE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlSCE.sAutoChar) & "/" & Trim(tlSCE.sDescription)
    
    ElseIf Trim(tmAIE.sRefFileName) = "SEE" Then      'Day schedule events
        rstAIERpt.Fields("TypeOfChange") = "Schedule Events"
        ilRet = gGetRec_SHE_ScheduleHeader(tmAIE.lOrigFileCode, slFileStamp, tlSHE)
        ilRet = gGetRec_SEE_ScheduleEvent(llToFileCode, slFileStamp, tlSEE)
        ilRet = gGetRec_AEE_AutoEquip(tlSHE.iAeeCode, slFileStamp, tlAEE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlAEE.sName) + "/" & Trim(tlAEE.sDescription)
        rstAIERpt.Fields("DescSecField") = Trim(tlSHE.sAirDate) & " at " & Trim(gLongToStrLengthInTenth(tlSEE.lTime, True))
    
    ElseIf Trim(tmAIE.sRefFileName) = "SHE" Then        'day Schedule header
        rstAIERpt.Fields("TypeOfChange") = "Schedule Header"
        ilRet = gGetRec_SHE_ScheduleHeader(llToFileCode, slFileStamp, tlSHE)
        ilRet = gGetRec_AEE_AutoEquip(tlSHE.iAeeCode, slFileStamp, tlAEE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlAEE.sName) + "/" & Trim(tlAEE.sDescription)
        rstAIERpt.Fields("DescSecField") = Trim(tlSHE.sAirDate)
        
    ElseIf Trim(tmAIE.sRefFileName) = "SOE" Then      'Site
        rstAIERpt.Fields("TypeOfChange") = "Site Option"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_SOE_SiteOption(ilToFileCode, slFileStamp, tlSOE)
        rstAIERpt.Fields("DescKeyField") = tlSOE.sClientName
        
    ElseIf Trim(tmAIE.sRefFileName) = "TTE" Then      'time type
        rstAIERpt.Fields("TypeOfChange") = "Time Type"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_TTE_TimeType(ilToFileCode, slFileStamp, tlTTE)
        rstAIERpt.Fields("DescKeyField") = Trim(tlTTE.sName) & "/" & Trim(tlTTE.sDescription)
        If tlTTE.sType = "S" Then
            rstAIERpt.Fields("DescSecField") = "Start Time"
        Else
            rstAIERpt.Fields("DescSecField") = "End Time"
        End If
            
    ElseIf Trim(tmAIE.sRefFileName) = "UIE" Then      'user
        rstAIERpt.Fields("TypeOfChange") = "User Information"
        ilToFileCode = llToFileCode
        ilRet = gGetRec_UIE_UserInfo(ilToFileCode, slFileStamp, tlUie)
        rstAIERpt.Fields("DescKeyField") = Trim(tlUie.sShowName)
    
    Else
        rstAIERpt.Fields("TypeOfChange") = "Unknown Type for " & Trim(tmAIE.sRefFileName)
   
    End If
    Exit Sub
    
ErrHand:
    'one of the files may have bad codes for the retrieval (i.e. DHE has missing dhednecode)
    Screen.MousePointer = vbDefault
    gMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            gMsg = "A SQL error has occured in Activity Rpt-mProcessAIE: "
            MsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
            Screen.MousePointer = vbDefault
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in Activity Rpt-mProcessAIE: "
        MsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    Exit Sub
End Sub
