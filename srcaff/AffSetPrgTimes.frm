VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSetPrgTimes 
   Caption         =   "Set Fields"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   ControlBox      =   0   'False
   Icon            =   "AffSetPrgTimes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lbcVehicles 
      Height          =   3180
      ItemData        =   "AffSetPrgTimes.frx":08CA
      Left            =   225
      List            =   "AffSetPrgTimes.frx":08CC
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   375
      Width           =   4665
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   3675
      Width           =   900
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2910
      TabIndex        =   2
      Top             =   4635
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmcOK 
      Caption         =   "Process"
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   4635
      Width           =   1890
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   75
      Top             =   5100
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   210
      Left            =   855
      TabIndex        =   0
      Top             =   4260
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lacAstCount 
      Height          =   210
      Left            =   60
      TabIndex        =   3
      Top             =   4245
      Width           =   750
   End
End
Attribute VB_Name = "frmSetPrgTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmSetPrgTimes
'*
'*  Created October,2005 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc. 2005
'*
'******************************************************
Option Explicit
Option Compare Text

Private imAllClick As Integer

Private lmTotalRecords As Long
Private lmProcessedRecords As Long
Private lmPercent As Long

Private rst_att As ADODB.Recordset
Private rst_DAT As ADODB.Recordset





Private Sub cmcCancel_Click()
    Unload frmSetPrgTimes
End Sub

Private Sub cmcOK_Click()
    If cmcOK.Caption = "Process" Then
        tmcStart.Enabled = True
        Exit Sub
    End If
    Unload frmSetPrgTimes
End Sub

Private Sub Form_Load()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    mFillVehicle
    
    cmcCancel.Visible = True
    cmcOK.Enabled = True
    
    gCenterStdAlone frmSetPrgTimes
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    rst_att.Close
    rst_DAT.Close
    On Error GoTo 0
    Set frmSetPrgTimes = Nothing

End Sub



Private Function mSetAgreements() As Integer
    Dim ilRet As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilExportType As Integer
    Dim slExportToWeb As String
    Dim slExportToUnivision As String
    Dim slExportToMarketron As String
    Dim slCDStartTime As String
    Dim slPledgeType As String
    Dim ilSeqNo As Integer
    Dim ilVefCount As Integer
    Dim slVefCode As String
    Dim ilVef As Integer
    
    On Error GoTo ErrHand
    mSetAgreements = False
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
    SQLQuery = "SELECT Count(attCode) FROM ATT"
    If ilVefCount <> lbcVehicles.ListCount Then
        SQLQuery = SQLQuery & " WHERE attVefCode In (" & slVefCode & ")"
    End If
    Set rst_att = gSQLSelectCall(SQLQuery)
    If Not rst_att.EOF Then
        lmTotalRecords = rst_att(0).Value
        SQLQuery = "SELECT * FROM ATT  "
        If ilVefCount <> lbcVehicles.ListCount Then
            SQLQuery = SQLQuery & " WHERE attVefCode In (" & slVefCode & ")"
        End If
        Set rst_att = gSQLSelectCall(SQLQuery)
        Do While Not rst_att.EOF
            'If IsNull(rst_att!attStartTime) Then
                slCDStartTime = ""
            'Else
            '    slCDStartTime = Format$(rst_att!attStartTime, "hh:mmA/P")
            'End If
            ilRet = gDetermineAgreementTimes(rst_att!attshfCode, rst_att!attvefCode, Format$(rst_att!attOnAir, "m/d/yy"), Format$(rst_att!attOffAir, "m/d/yy"), Format$(rst_att!attDropDate, "m/d/yy"), slCDStartTime, slStartTime, slEndTime)
            If sgSetFieldCallSource = "S" Then
                slExportToWeb = "N"
                If rst_att!attExportType = 1 Then
                    slExportToWeb = "Y"
                End If
                slExportToUnivision = "N"
                '7701
'                If rst_att!attExportType = 2 Then
'                    slExportToUnivision = "Y"
'                End If
                slExportToMarketron = "N"
                '7701
                If gIsVendorWithAgreement(rst_att!attCode, Vendors.NetworkConnect) Then
                'If rst_att!attExportToMarketron = "Y" Then
                    slExportToMarketron = "Y"
                End If
                slPledgeType = ""
                ilSeqNo = 0
                SQLQuery = "SELECT * "
                SQLQuery = SQLQuery + " FROM dat"
                SQLQuery = SQLQuery + " WHERE (datatfCode= " & rst_att!attCode & ")"
                Set rst_DAT = gSQLSelectCall(SQLQuery)
                Do While Not rst_DAT.EOF
                    Select Case rst_DAT!datDACode
                        Case 0  'Dayprt
                            slPledgeType = "D"
                        Case 1  'Avail
                            slPledgeType = "A"
                        Case 2  'CD or Tape
                            slPledgeType = "C"
                        Case Else
                            slPledgeType = ""
                    End Select
                    SQLQuery = "UPDATE dat Set "
                    SQLQuery = SQLQuery & " datAirPlayNo = " & 1 & ","
                    SQLQuery = SQLQuery & " datEstimatedTime = " & "'N'"
                    SQLQuery = SQLQuery & " Where datCode = " & rst_DAT!datCode
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "SetPrgTimes-mSetAgreements"
                        mSetAgreements = False
                        Exit Function
                    End If

'                    SQLQuery = "Insert Into apt ( "
'                    SQLQuery = SQLQuery & "aptCode, "
'                    SQLQuery = SQLQuery & "aptAttCode, "
'                    SQLQuery = SQLQuery & "aptAirPlayNo, "
'                    SQLQuery = SQLQuery & "aptSeqNo, "
'                    SQLQuery = SQLQuery & "aptBreakoutMo, "
'                    SQLQuery = SQLQuery & "aptBreakoutTu, "
'                    SQLQuery = SQLQuery & "aptBreakoutWe, "
'                    SQLQuery = SQLQuery & "aptBreakoutTh, "
'                    SQLQuery = SQLQuery & "aptBreakoutFr, "
'                    SQLQuery = SQLQuery & "aptBreakoutSa, "
'                    SQLQuery = SQLQuery & "aptBreakoutSu, "
'                    SQLQuery = SQLQuery & "aptStartTime, "
'                    SQLQuery = SQLQuery & "aptOffsetDay, "
'                    SQLQuery = SQLQuery & "aptEstimatedTime, "
'                    SQLQuery = SQLQuery & "aptUnused "
'                    SQLQuery = SQLQuery & ") "
'                    SQLQuery = SQLQuery & "Values ( "
'                    SQLQuery = SQLQuery & 0 & ", "
'                    SQLQuery = SQLQuery & rst_att!attCode & ", "
'                    SQLQuery = SQLQuery & 1 & ", "
'                    SQLQuery = SQLQuery & 1 & ", "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    If slPledgeType = "C" Then
'                        SQLQuery = SQLQuery & "'" & Format$(rst_att!attStartTime, sgSQLTimeForm) & "', "
'                    Else
'                        SQLQuery = SQLQuery & "'" & Format$(slStartTime, sgSQLTimeForm) & "', "
'                    End If
'                    SQLQuery = SQLQuery & 0 & ", "
'                    SQLQuery = SQLQuery & "'" & "N" & "', "
'                    SQLQuery = SQLQuery & "'" & "" & "' "
'                    SQLQuery = SQLQuery & ") "
'
'                    ilSeqNo = ilSeqNo + 1
'                    SQLQuery = SQLQuery & "aptCode, "
'                    SQLQuery = SQLQuery & "aptAttCode, "
'                    SQLQuery = SQLQuery & "aptAirPlayNo, "
'                    SQLQuery = SQLQuery & "aptSeqNo, "
'                    SQLQuery = SQLQuery & "aptPledgeType, "
'                    SQLQuery = SQLQuery & "aptFdStatus, "
'                    SQLQuery = SQLQuery & "aptAirMo, "
'                    SQLQuery = SQLQuery & "aptAirTu, "
'                    SQLQuery = SQLQuery & "aptAirWe, "
'                    SQLQuery = SQLQuery & "aptAirTh, "
'                    SQLQuery = SQLQuery & "aptAirFr, "
'                    SQLQuery = SQLQuery & "aptAirSa, "
'                    SQLQuery = SQLQuery & "aptAirSu, "
'                    SQLQuery = SQLQuery & "aptPledgeStartTime, "
'                    SQLQuery = SQLQuery & "aptOffsetDay, "
'                    SQLQuery = SQLQuery & "aptEStimatedTime, "
'                    SQLQuery = SQLQuery & "aptFeedStartTime, "
'                    SQLQuery = SQLQuery & "aptFeedEndTime, "
'                    SQLQuery = SQLQuery & "aptUnused "
'                    SQLQuery = SQLQuery & ") "
'                    SQLQuery = SQLQuery & "Values ( "
'                    SQLQuery = SQLQuery & 0 & ", "
'                    SQLQuery = SQLQuery & rst_att!attCode & ", "
'                    SQLQuery = SQLQuery & 1 & ", "
'                    SQLQuery = SQLQuery & ilSeqNo & ", "
'                    SQLQuery = SQLQuery & "'" & gFixQuote(slPledgeType) & "', "
'                    SQLQuery = SQLQuery & 0 & ", "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    If slPledgeType = "C" Then
'                        SQLQuery = SQLQuery & "'" & Format$(rst_att!attStartTime, sgSQLTimeForm) & "', "
'                    Else
'                        SQLQuery = SQLQuery & "'" & Format$(slStartTime, sgSQLTimeForm) & "', "
'                    End If
'                    SQLQuery = SQLQuery & -1 & ", "
'                    SQLQuery = SQLQuery & "'" & "N" & "', "
'                    SQLQuery = SQLQuery & "'" & Format$("12:00:00AM", sgSQLTimeForm) & "', "
'                    SQLQuery = SQLQuery & "'" & Format$("12:00:00AM", sgSQLTimeForm) & "', "
'                    SQLQuery = SQLQuery & "'" & "" & "' "
'                    SQLQuery = SQLQuery & ") "
'
'                    'cnn.Execute SQLQuery, rdExecDirect
'                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                        GoSub ErrHand:
'                    End If
                    rst_DAT.MoveNext
                Loop
                SQLQuery = "Update att Set "
                SQLQuery = SQLQuery & "attVehProgStartTime = '" & Format$(slStartTime, sgSQLTimeForm) & "', "
                SQLQuery = SQLQuery & "attVehProgEndTime = '" & Format$(slEndTime, sgSQLTimeForm) & "', "
                'SQLQuery = SQLQuery & "attExportType = " & ilExportType & ", "
                SQLQuery = SQLQuery & "attExportToWeb = '" & slExportToWeb & "', "
                SQLQuery = SQLQuery & "attExportToUnivision = '" & slExportToUnivision & "', "
                SQLQuery = SQLQuery & "attExportToMarketron = '" & slExportToMarketron & "', "
                SQLQuery = SQLQuery & "attExportToCBS = '" & "N" & "', "
                SQLQuery = SQLQuery & "attExportToClearCh = '" & "N" & "', "
                SQLQuery = SQLQuery & "attNoAirPlays = " & 1 & ", "
                SQLQuery = SQLQuery & "attDesignVersion = " & 1 & ", "
                SQLQuery = SQLQuery & "attPledgeType = '" & slPledgeType & "'"
                'ttp 5270 change manual to export
                If slExportToMarketron = "Y" Or slExportToWeb = "Y" Or slExportToUnivision = "Y" Then
                      SQLQuery = SQLQuery & " , " & "attExportType = " & 1
               End If
            Else
                SQLQuery = "Update att Set "
                SQLQuery = SQLQuery & "attVehProgStartTime = '" & Format$(slStartTime, sgSQLTimeForm) & "', "
                SQLQuery = SQLQuery & "attVehProgEndTime = '" & Format$(slEndTime, sgSQLTimeForm) & "'"
            End If
            SQLQuery = SQLQuery & " Where attCode = " & rst_att!attCode
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "SetPrgTimes-mSetAgreements"
                mSetAgreements = False
                Exit Function
            End If
            mSetGauge
            rst_att.MoveNext
        Loop
    End If
    
    mSetAgreements = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSetPrgTimes-mSetAgreements"
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
    lmPercent = 0
    ilOk = True
    gLogMsg "Set Agreements: Start", "SetPrgTimes.Txt", False
    ilRet = mSetAgreements()
    If ilRet Then
        gLogMsg "Set Agreements: Completed", "SetPrgTimes.Txt", False
    Else
        gLogMsg "Set Agreements: Stopped", "SetPrgTimes.Txt", False
    End If
    plcGauge.Visible = False
    cmcOK.Caption = "Done"
    cmcOK.Enabled = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSetPrgTimes-tmcStart"
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
