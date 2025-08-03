VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSetCompliants 
   Caption         =   "Set Compliance"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   ControlBox      =   0   'False
   Icon            =   "AffSetCompliants.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4155
      Width           =   900
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4215
      TabIndex        =   8
      Top             =   4965
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmcOK 
      Caption         =   "Process"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2145
      TabIndex        =   7
      Top             =   4965
      Width           =   1890
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1380
      Top             =   5160
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   210
      Left            =   2160
      TabIndex        =   6
      Top             =   4590
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin V81Affiliate.CSI_Calendar edcFromDate 
      Height          =   285
      Left            =   1260
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
      CSI_ForceMondaySelectionOnly=   -1  'True
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin V81Affiliate.CSI_Calendar edcToDate 
      Height          =   285
      Left            =   4215
      TabIndex        =   3
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
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   3180
      ItemData        =   "AffSetCompliants.frx":08CA
      Left            =   120
      List            =   "AffSetCompliants.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   810
      Width           =   4665
   End
   Begin VB.Label lacDate 
      Caption         =   "To Date"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   165
      Width           =   1395
   End
   Begin VB.Label lacDate 
      Caption         =   "From Date"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   165
      Width           =   1395
   End
   Begin VB.Label lacAstCount 
      Height          =   210
      Left            =   1365
      TabIndex        =   9
      Top             =   4575
      Width           =   750
   End
End
Attribute VB_Name = "frmSetCompliants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmSetCompliants
'*
'*  Created October,2005 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc. 2005
'*
'******************************************************
Option Explicit
Option Compare Text

'Private tmAstInfo As ASTINFO


Private hmAst As Integer
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO

Private imAllClick As Integer

Private lmTotalRecords As Long
Private lmProcessedRecords As Long
Private lmPercent As Long

Private rst_att As ADODB.Recordset
Private rst_Cptt As ADODB.Recordset
Private rst_Ast As ADODB.Recordset
Private rst_Lst As ADODB.Recordset
Private rst_DAT As ADODB.Recordset





Private Sub cmcCancel_Click()
    Unload frmSetCompliants
End Sub

Private Sub cmcOK_Click()
    If cmcOK.Caption = "Process" Then
        tmcStart.Enabled = True
        Exit Sub
    End If
    Unload frmSetCompliants
End Sub

Private Sub edcToDate_GotFocus()
    If edcFromDate.Text <> "" Then
        edcToDate.Text = gObtainEndStd(edcFromDate.Text)
        edcToDate.Text = ""
    End If
End Sub

Private Sub Form_Load()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand:
    mFillVehicle
    
    SQLQuery = "SELECT Min(cpttStartDate) FROM Cptt"
    Set rst_Cptt = gSQLSelectCall(SQLQuery)
    If Not rst_Cptt.EOF Then
        edcFromDate.Text = rst_Cptt(0).Value
    End If
    cmcCancel.Visible = True
    cmcOK.Enabled = True
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    gCenterStdAlone frmSetCompliants
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmSetCompliant:Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    Erase tmAstInfo
    Erase tmCPDat
    rst_att.Close
    rst_Cptt.Close
    rst_Ast.Close
    rst_Lst.Close
    rst_DAT.Close
    On Error GoTo 0
    Set frmSetCompliants = Nothing

End Sub




Private Function mSetPostCPs() As Integer
    Dim ilSchdCount As Integer
    Dim ilAiredCount As Integer
    Dim ilPledgeCompliantCount As Integer
    Dim ilAgyCompliantCount As Integer
    Dim llCpttDate As Long
    Dim ilVefCode As Integer
    Dim llAttCode As Long
    Dim llWeekDate As Long
    'Dim slWeek1 As String
    'Dim slWeek60 As String
    Dim slCpttStartDate As String
    Dim slCpttEndDate As String
    Dim ilAdfCode As Integer
    Dim ilShtt As Integer
    Dim ilAst As Integer
    Dim slMoDate As String
    Dim ilRet As Integer
    Dim slTPdETime As String
    Dim slPledgeEndTime As String
    Dim ilTechnique As Integer
    Dim llPrevAttCode As Long
    Dim llDat As Long
    Dim ilLoop As Integer
    Dim slFeedStartTime As String
    'Dim llFeedStartTime As Long
    Dim slPledgeStartTime As String
    'Dim llPledgeStartTime As Long
    Dim ilPdDay As Integer
    Dim ilDayOk As Integer
    Dim ilFdDay As Integer
    Dim llDatIndex As Long
    Dim slPledgeDays As String
    Dim ilAirStatus As Integer
    Dim slFeedDate As String
    Dim slPledgeDate As String
    Dim llAstCount As Long
    Dim ilPostingType As Integer
    Dim blAttOk As Boolean
    Dim ilVefCount As Integer
    Dim slVefCode As String
    Dim ilVef As Integer
    Dim ilSvAirStatus As Integer
    Dim tlAstInfo As ASTINFO
    Dim tlDatPledgeInfo As DATPLEDGEINFO
    Dim tlLST As LST
    Dim slServiceAgreement As String
    
    On Error GoTo ErrHand
    mSetPostCPs = False
    llPrevAttCode = -1
    'slWeek1 = Format$(gNow(), sgShowDateForm)   'sgSQLDateForm)
    'slWeek1 = gObtainNextSunday(gObtainNextMonday(gObtainNextSunday(slWeek1)))
    'slWeek60 = DateAdd("d", -(60 * 7) + 1, slWeek1)
    slCpttStartDate = edcFromDate.Text
    slCpttEndDate = edcToDate.Text
    If slCpttEndDate = "" Then
        slCpttEndDate = "12/31/2069"
    End If
    If (chkAll.Value = vbUnchecked) And (lbcVehicles.ListCount <> lbcVehicles.SelCount) Then
        ilVefCount = lbcVehicles.SelCount
        slVefCode = ""
        For ilVef = 0 To lbcVehicles.ListCount - 1 Step 1
            If lbcVehicles.Selected(ilVef) Then
                If slVefCode = "" Then
                    slVefCode = lbcVehicles.ItemData(ilVef)
                Else
                    slVefCode = slVefCode & "," & lbcVehicles.ItemData(ilVef)
                End If
            End If
        Next ilVef
    Else
        ilVefCount = lbcVehicles.ListCount
    End If
    
    ilRet = gPopAvailNames()
    
    If Not gPopCopy(slCpttStartDate, "Post CPs") Then
        Exit Function
    End If
    
    SQLQuery = "SELECT Count(cpttCode) FROM CPTT WHERE cpttStartDate >= '" & Format(slCpttStartDate, sgSQLDateForm) & "'" & " And cpttStartDate <= '" & Format(slCpttEndDate, sgSQLDateForm) & "'"
    If ilVefCount <> lbcVehicles.ListCount Then
        SQLQuery = SQLQuery & " AND cpttVefCode In (" & slVefCode & ")"
    End If
    Set rst_Cptt = gSQLSelectCall(SQLQuery)
    If Not rst_Cptt.EOF Then
        lmTotalRecords = rst_Cptt(0).Value
        SQLQuery = "SELECT * FROM CPTT WHERE cpttStartDate >= '" & Format(slCpttStartDate, sgSQLDateForm) & "'" & " And cpttStartDate <= '" & Format(slCpttEndDate, sgSQLDateForm) & "'"
        If ilVefCount <> lbcVehicles.ListCount Then
            SQLQuery = SQLQuery & " AND cpttVefCode In (" & slVefCode & ")"
        End If
        SQLQuery = SQLQuery & " ORDER BY cpttAtfCode"
        Set rst_Cptt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Cptt.EOF
            'Set counts
            blAttOk = True
            llCpttDate = gDateValue(rst_Cptt!CpttStartDate)
            ilVefCode = rst_Cptt!cpttvefcode
            llAttCode = rst_Cptt!cpttatfCode
            If llAttCode <> llPrevAttCode Then
                SQLQuery = "SELECT * FROM att"
                SQLQuery = SQLQuery + " WHERE ("
                SQLQuery = SQLQuery & " attCode = " & llAttCode & ")"
                Set rst_att = gSQLSelectCall(SQLQuery)
                If Not rst_att.EOF Then
                    ilPostingType = rst_att!attPostingType
                    slServiceAgreement = rst_att!attServiceAgreement
                Else
                    blAttOk = False
                End If
            End If
            If blAttOk Then
                ilSchdCount = 0
                ilAiredCount = 0
                ilPledgeCompliantCount = 0
                ilAgyCompliantCount = 0
                If ilPostingType = 0 Then  'Receipt
                    ilSchdCount = 0
                    ilAiredCount = 0
                ElseIf ilPostingType = 1 Then  'Count
                    ilSchdCount = rst_Cptt!cpttNoSpotsGen
                    ilAiredCount = rst_Cptt!cpttNoSpotsAired
                    ilPledgeCompliantCount = 0
                    ilAgyCompliantCount = 0
                Else
                    'Spots by date and spots by advertiser
                    llWeekDate = gDateValue(gObtainPrevMonday(gAdjYear(Format$(llCpttDate, "m/d/yy"))))
                    ilTechnique = 1
                    If ilTechnique = 1 Then
                        If llAttCode <> llPrevAttCode Then
                            ReDim tlDat(0 To 30) As DATRST
                            llDat = 0
                            SQLQuery = "SELECT * "
                            SQLQuery = SQLQuery + " FROM dat"
                            SQLQuery = SQLQuery + " WHERE (datatfCode= " & llAttCode & ")"
                            Set rst_DAT = gSQLSelectCall(SQLQuery)
                            Do While Not rst_DAT.EOF
                                gCreateUDTForDat rst_DAT, tlDat(llDat)
                                llDat = llDat + 1
                                If llDat = UBound(tlDat) Then
                                    ReDim Preserve tlDat(0 To UBound(tlDat) + 30) As DATRST
                                End If
                                rst_DAT.MoveNext
                            Loop
                            ReDim Preserve tlDat(0 To llDat) As DATRST
                            llPrevAttCode = llAttCode
                        End If
                        llAstCount = 0
                        ReDim llAstUpdate(0 To 0) As Long
                        SQLQuery = "SELECT * FROM ast"
                        SQLQuery = SQLQuery + " WHERE ("
                        SQLQuery = SQLQuery + " astFeedDate >= '" & Format(llWeekDate, sgSQLDateForm) & "'"
                        SQLQuery = SQLQuery + " AND astFeedDate <= '" & Format(llWeekDate + 6, sgSQLDateForm) & "'"
                        SQLQuery = SQLQuery & " AND astatfCode = " & llAttCode & ")"
                        Set rst_Ast = gSQLSelectCall(SQLQuery)
                        Do While Not rst_Ast.EOF
                            tlDatPledgeInfo.lAttCode = rst_Ast!astAtfCode
                            tlDatPledgeInfo.lDatCode = rst_Ast!astDatCode
                            tlDatPledgeInfo.iVefCode = rst_Ast!astVefCode
                            tlDatPledgeInfo.sFeedDate = Format(rst_Ast!astFeedDate, "m/d/yy")
                            tlDatPledgeInfo.sFeedTime = Format(rst_Ast!astFeedTime, "hh:mm:ssam/pm")
                            ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
                            If tgStatusTypes(gGetAirStatus(tlDatPledgeInfo.iPledgeStatus)).iPledged <> 2 Then
                                llAstCount = llAstCount + 1
                                'Find dat match
                                slFeedStartTime = Format$(rst_Ast!astFeedTime, sgShowTimeWSecForm)
                                ''llFeedStartTime = gTimeToLong(slFeedStartTime, False)
                                'slPledgeStartTime = Format$(rst_Ast!astPledgeStartTime, sgShowTimeWSecForm)
                                ''llPledgeStartTime = gTimeToLong(slPledgeStartTime, False)
                                'If Not IsNull(rst_Ast!astPledgeEndTime) Then
                                '    slPledgeEndTime = Format$(rst_Ast!astPledgeEndTime, sgShowTimeWSecForm)
                                'Else
                                '    slPledgeEndTime = slPledgeStartTime
                                'End If
                                slPledgeStartTime = Format$(tlDatPledgeInfo.sPledgeStartTime, sgShowTimeWSecForm)
                                If Not IsNull(tlDatPledgeInfo.sPledgeEndTime) Then
                                    slPledgeEndTime = Format$(tlDatPledgeInfo.sPledgeEndTime, sgShowTimeWSecForm)
                                Else
                                    slPledgeEndTime = tlDatPledgeInfo.sPledgeStartTime
                                End If
                                slFeedDate = rst_Ast!astFeedDate
                                'slPledgeDate = rst_Ast!astPledgeDate
                                slPledgeDate = tlDatPledgeInfo.sPledgeDate
                                llDatIndex = gMatchAstAndDat(slFeedStartTime, slFeedDate, slPledgeStartTime, slPledgeDate, tlDat())
    '                            For llDat = LBound(tlDat) To UBound(tlDat) - 1 Step 1
    '                                If (gTimeToLong(tlDat(llDat).sFdStTime, False) = llFeedStartTime) And (gTimeToLong(tlDat(llDat).sPdStTime, False) = llPledgeStartTime) Then
    '                                    ilDayOk = False
    '                                    ilFdDay = Weekday(rst_Ast!astFeedDate, vbMonday)
    '                                    Select Case ilFdDay
    '                                        Case 1  'Monday
    '                                            If tlDat(llDat).iFdMon Then
    '                                                ilDayOk = True
    '                                            End If
    '                                        Case 2  'Tuesday
    '                                            If tlDat(llDat).iFdTue Then
    '                                                ilDayOk = True
    '                                            End If
    '                                        Case 3  'Wednesady
    '                                            If tlDat(llDat).iFdWed Then
    '                                                ilDayOk = True
    '                                            End If
    '                                        Case 4  'Thursday
    '                                            If tlDat(llDat).iFdThu Then
    '                                                ilDayOk = True
    '                                            End If
    '                                        Case 5  'Friday
    '                                            If tlDat(llDat).iFdFri Then
    '                                                ilDayOk = True
    '                                            End If
    '                                        Case 6  'Saturday
    '                                            If tlDat(llDat).iFdSat Then
    '                                                ilDayOk = True
    '                                            End If
    '                                        Case 7  'Sunday
    '                                            If tlDat(llDat).iFdSun Then
    '                                                ilDayOk = True
    '                                            End If
    '                                    End Select
    '                                    If ilDayOk Then
    '                                        ilDayOk = False
    '                                        ilPdDay = Weekday(rst_Ast!astPledgeDate, vbMonday)
    '                                        Select Case ilPdDay
    '                                            Case 1  'Monday
    '                                                If tlDat(llDat).iPdMon Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                            Case 2  'Tuesday
    '                                                If tlDat(llDat).iPdTue Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                            Case 3  'Wednesday
    '                                                If tlDat(llDat).iPdWed Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                            Case 4  'Thursday
    '                                                If tlDat(llDat).iPdThu Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                            Case 5  'Friday
    '                                                If tlDat(llDat).iPdFri Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                            Case 6  'Saturday
    '                                                If tlDat(llDat).iPdSat Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                            Case 7  'Sunday
    '                                                If tlDat(llDat).iPdSun Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                        End Select
    '                                    End If
    '                                    If ilDayOk Then
    '                                        llDatIndex = llDat
    '                                        slPledgeDays = String(7, "N")
    '                                        For ilPdDay = 1 To 7
    '                                            Select Case ilPdDay
    '                                                Case 1  'Monday
    '                                                    If tlDat(llDat).iPdMon Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                                Case 2  'Tuesday
    '                                                    If tlDat(llDat).iPdTue Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                                Case 3  'Wednesday
    '                                                    If tlDat(llDat).iPdWed Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                                Case 4  'Thursday
    '                                                    If tlDat(llDat).iPdThu Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                                Case 5  'Friday
    '                                                    If tlDat(llDat).iPdFri Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                                Case 6  'Saturday
    '                                                    If tlDat(llDat).iPdSat Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                                Case 7  'Sunday
    '                                                    If tlDat(llDat).iPdSun Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                            End Select
    '                                        Next ilPdDay
    '                                        Exit For
    '                                    End If
    '                                End If
    '                            Next llDat
                                slPledgeDays = String(7, "N")
                                If llDatIndex <> -1 Then
                                    llDat = llDatIndex
                                    For ilPdDay = 1 To 7
                                        Select Case ilPdDay
                                            Case 1  'Monday
                                                If tlDat(llDat).iPdMon Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                            Case 2  'Tuesday
                                                If tlDat(llDat).iPdTue Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                            Case 3  'Wednesday
                                                If tlDat(llDat).iPdWed Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                            Case 4  'Thursday
                                                If tlDat(llDat).iPdThu Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                            Case 5  'Friday
                                                If tlDat(llDat).iPdFri Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                            Case 6  'Saturday
                                                If tlDat(llDat).iPdSat Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                            Case 7  'Sunday
                                                If tlDat(llDat).iPdSun Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                        End Select
                                    Next ilPdDay
                                Else
                                    Mid(slPledgeDays, Weekday(slPledgeDate, vbMonday), 1) = "Y"
                                End If
                                If (gTimeToLong(slPledgeStartTime, False) = gTimeToLong(slPledgeEndTime, True)) Then
                                    If llDatIndex >= 0 Then
                                        slTPdETime = Format$(gLongToTime(gTimeToLong(Format$(slPledgeStartTime, "h:mm:ssam/pm"), False) + gTimeToLong(tlDat(llDatIndex).sFdEdTime, False) - gTimeToLong(tlDat(llDatIndex).sFdStTime, False)), sgShowTimeWSecForm)
                                    Else
                                        'Add 5 minutes to start time
                                        slTPdETime = Format$(gLongToTime(gTimeToLong(Format$(slPledgeStartTime, "h:mm:ssam/pm"), False) + 300), sgShowTimeWSecForm)
                                    End If
                                Else
                                    slTPdETime = Format$(slPledgeEndTime, "h:mm:ssam/pm")
                                End If
                                ilAirStatus = rst_Ast!astStatus
                                If (gGetAirStatus(ilAirStatus) = 6) Or (gGetAirStatus(ilAirStatus) = 7) Then
                                    ilAirStatus = 1
                                    llAstUpdate(UBound(llAstUpdate)) = rst_Ast!astCode
                                    ReDim Preserve llAstUpdate(0 To UBound(llAstUpdate) + 1) As Long
                                End If
                                ''gIncSpotCounts rst_Ast!astPledgeStatus, ilAirStatus, rst_Ast!astCPStatus, slPledgeDays, Format$(slPledgeDate, "m/d/yy"), Format$(rst_Ast!astAirDate, "m/d/yy"), Format$(slPledgeStartTime, "h:mm:ssAM/PM"), Format$(slTPdETime, "h:mm:ssAM/PM"), Format$(rst_Ast!astAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
                                'gIncSpotCounts rst_Ast!astPledgeStatus, ilAirStatus, rst_Ast!astCPStatus, slPledgeDays, Format$(slPledgeDate, "m/d/yy"), Format$(rst_Ast!astAirDate, "m/d/yy"), Format$(slPledgeStartTime, "h:mm:ssAM/PM"), Format$(slTPdETime, "h:mm:ssAM/PM"), Format$(rst_Ast!astAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
                                tlAstInfo.lCode = rst_Ast!astCode
                                tlAstInfo.lAttCode = rst_Ast!astAtfCode
                                'tlAstInfo.iPledgeStatus = rst_Ast!astPledgeStatus
                                tlAstInfo.iPledgeStatus = tlDatPledgeInfo.iPledgeStatus
                                tlAstInfo.iStatus = ilAirStatus
                                tlAstInfo.iCPStatus = rst_Ast!astCPStatus
                                tlAstInfo.sTruePledgeDays = slPledgeDays
                                tlAstInfo.sPledgeDate = slPledgeDate
                                tlAstInfo.sAirDate = rst_Ast!astAirDate
                                tlAstInfo.sPledgeStartTime = slPledgeStartTime
                                tlAstInfo.sTruePledgeEndTime = slTPdETime
                                tlAstInfo.sAirTime = Format(rst_Ast!astAirTime, sgShowTimeWSecForm)
                                tlAstInfo.lLstCode = rst_Ast!astLsfCode
                                tlAstInfo.lSdfCode = rst_Ast!astSdfCode
                                tlAstInfo.iShttCode = rst_Ast!astShfCode
                                SQLQuery = "SELECT *"
                                SQLQuery = SQLQuery & " FROM LST"
                                SQLQuery = SQLQuery & " WHERE lstCode =" & Str(tlAstInfo.lLstCode)
                                Set rst_Lst = gSQLSelectCall(SQLQuery)
                                If Not rst_Lst.EOF Then
                                    gCreateUDTforLST rst_Lst, tlLST
                                    tlAstInfo.lLstBkoutLstCode = tlLST.lBkoutLstCode
                                    tlAstInfo.sLstStartDate = tlLST.sStartDate
                                    tlAstInfo.sLstEndDate = tlLST.sEndDate
                                    tlAstInfo.iLstSpotsWk = tlLST.iSpotsWk
                                    tlAstInfo.iLstMon = tlLST.iMon
                                    tlAstInfo.iLstTue = tlLST.iTue
                                    tlAstInfo.iLstWed = tlLST.iWed
                                    tlAstInfo.iLstThu = tlLST.iThu
                                    tlAstInfo.iLstFri = tlLST.iFri
                                    tlAstInfo.iLstSat = tlLST.iSat
                                    tlAstInfo.iLstSun = tlLST.iSun
                                    tlAstInfo.iLineNo = tlLST.iLineNo
                                    tlAstInfo.iSpotType = tlLST.iSpotType
                                    tlAstInfo.sSplitNet = tlLST.sSplitNetwork
                                    tlAstInfo.iAgfCode = tlLST.iAgfCode
                                    tlAstInfo.sLstLnStartTime = tlLST.sLnStartTime
                                    tlAstInfo.sLstLnEndTime = tlLST.sLnEndTime
                                    tlAstInfo.iVefCode = tlLST.iLogVefCode
                                    gIncSpotCounts tlAstInfo, ilSchdCount, ilAiredCount, ilPledgeCompliantCount, ilAgyCompliantCount, slServiceAgreement
                                End If
                            End If
                            rst_Ast.MoveNext
                        Loop
                        If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
                            lacAstCount.Caption = llAstCount
                            DoEvents
                        End If
                        For ilAst = 0 To UBound(llAstUpdate) - 1 Step 1
                            SQLQuery = "UPDATE ast SET astStatus = 1 WHERE astCode = " & llAstUpdate(ilAst)
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/12/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "SetCompliants-mSetPostCPs"
                                mSetPostCPs = False
                                Exit Function
                            End If
                        Next ilAst
                    ElseIf ilTechnique = 2 Then
                        ReDim tgCPPosting(0 To 1) As CPPOSTING
                        tgCPPosting(0).lCpttCode = rst_Cptt!cpttCode
                        tgCPPosting(0).iStatus = rst_Cptt!cpttStatus
                        tgCPPosting(0).iPostingStatus = rst_Cptt!cpttPostingStatus
                        tgCPPosting(0).lAttCode = rst_Cptt!cpttatfCode
                        tgCPPosting(0).iAttTimeType = 0 'Not used
                        tgCPPosting(0).iVefCode = rst_Cptt!cpttvefcode  'imVefCode
                        tgCPPosting(0).iShttCode = rst_Cptt!cpttshfcode
                        ilShtt = gBinarySearchStationInfoByCode(tgCPPosting(0).iShttCode)
                        If ilShtt <> -1 Then
                            tgCPPosting(0).sZone = tgStationInfoByCode(ilShtt).sZone
                        Else
                            tgCPPosting(0).sZone = ""
                        End If
                        slMoDate = Format$(llWeekDate, "m/d/yy")
                        tgCPPosting(0).sDate = Format$(slMoDate, sgShowDateForm)
                        tgCPPosting(0).sAstStatus = rst_Cptt!cpttAstStatus
                        igTimes = 1 'By Week
                        ilAdfCode = -1
                        'Dan M 9/26/13  6442
                        ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, True, False, True, False)
                       ' ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, False, False, True, False)
                        ReDim llAstUpdate(0 To 0) As Long
                        For ilAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                            ilAirStatus = tmAstInfo(ilAst).iStatus
                            If (ilAirStatus = 6) Or (ilAirStatus = 7) Then
                                ilAirStatus = 1
                                llAstUpdate(UBound(llAstUpdate)) = rst_Ast!astCode
                                ReDim Preserve llAstUpdate(0 To UBound(llAstUpdate) + 1) As Long
                            End If
                            ''gIncSpotCounts tmAstInfo(ilAst).iPledgeStatus, ilAirStatus, tmAstInfo(ilAst).iCPStatus, tmAstInfo(ilAst).sTruePledgeDays, Format$(tmAstInfo(ilAst).sPledgeDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sAirDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sPledgeStartTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sTruePledgeEndTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
                            'gIncSpotCounts tmAstInfo(ilAst).iPledgeStatus, ilAirStatus, tmAstInfo(ilAst).iCPStatus, tmAstInfo(ilAst).sTruePledgeDays, Format$(tmAstInfo(ilAst).sPledgeDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sAirDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sPledgeStartTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sTruePledgeEndTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
                            ilSvAirStatus = tmAstInfo(ilAst).iStatus
                            tmAstInfo(ilAst).iStatus = ilAirStatus
                            gIncSpotCounts tmAstInfo(ilAst), ilSchdCount, ilAiredCount, ilPledgeCompliantCount, ilAgyCompliantCount, slServiceAgreement
                            tmAstInfo(ilAst).iStatus = ilSvAirStatus
                        Next ilAst
                        For ilAst = 0 To UBound(llAstUpdate) - 1 Step 1
                            SQLQuery = "UPDATE ast SET astStatus = 1 WHERE astCode = " & llAstUpdate(ilAst)
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/12/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "SetCompliants-mSetPostCPs"
                                mSetPostCPs = False
                                Exit Function
                            End If
                        Next ilAst
                    End If
                    SQLQuery = "Update cptt Set "
                    SQLQuery = SQLQuery & "cpttNoSpotsGen = " & ilSchdCount & ", "
                    SQLQuery = SQLQuery & "cpttNoSpotsAired = " & ilAiredCount & ", "
                    SQLQuery = SQLQuery & "cpttNoCompliant = " & ilPledgeCompliantCount & ", "
                    SQLQuery = SQLQuery & "cpttAgyCompliant = " & ilAgyCompliantCount & " "
                    SQLQuery = SQLQuery & " Where cpttCode = " & rst_Cptt!cpttCode
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "SetCompliants-mSetPostCPs"
                        mSetPostCPs = False
                        Exit Function
                    End If
                End If
            End If
            mSetGauge
            rst_Cptt.MoveNext
        Loop
    End If
    
    mSetPostCPs = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSetComplaints-mSetPostCPs"
End Function

Private Sub lbcVehicles_Click()
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = vbChecked Then
        imAllClick = True
        chkAll.Value = vbUnchecked
        imAllClick = False
    End If
End Sub

Private Sub tmcStart_Timer()
    Dim ilTask As Integer
    Dim ilRet As Integer
    Dim ilOk As Integer
    
    tmcStart.Enabled = False
    plcGauge.Visible = True
    lacAstCount.Visible = True
    lmPercent = 0
    ilOk = True
    gLogMsg "Set Post CP: Start", "SetCompliance.Txt", False
    ilRet = mSetPostCPs()
    If ilRet Then
        gLogMsg "Set Post CP: Completed", "SetCompliance.Txt", False
    Else
        gLogMsg "Set Post CP: Stopped", "SetCompliance.Txt", False
    End If
    plcGauge.Visible = False
    lacAstCount.Visible = False
    cmcOK.Caption = "Done"
    cmcOK.Enabled = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSetComplaints-mtmcStart"
End Sub

Private Sub mSetGauge()
    lmProcessedRecords = lmProcessedRecords + 1
    lmPercent = (lmProcessedRecords * CSng(100)) / lmTotalRecords
    If lmPercent >= 100 Then
        If lmProcessedRecords + 1 < lmTotalRecords Then
            lmPercent = 99
        Else
            lmPercent = 100
        End If
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
        DoEvents
    End If
End Sub

Private Sub mFillVehicle()
    Dim iLoop As Integer
    lbcVehicles.Clear
    chkAll.Value = vbUnchecked
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
    chkAll.Value = vbChecked
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
