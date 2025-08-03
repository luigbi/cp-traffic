VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmCPCount 
   Caption         =   "Spot Count"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3525
   Icon            =   "AffCPCount.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   3525
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   2940
      Top             =   1515
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   2730
      FormDesignWidth =   3525
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1710
      TabIndex        =   7
      Top             =   2220
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   135
      TabIndex        =   6
      Top             =   2220
      Width           =   1335
   End
   Begin VB.TextBox txtPCount 
      Height          =   285
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1650
      Width           =   780
   End
   Begin VB.TextBox txtSCount 
      Height          =   285
      Left            =   1530
      MaxLength       =   5
      TabIndex        =   1
      Top             =   930
      Width           =   780
   End
   Begin VB.Label labCount 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   315
      Width           =   3240
   End
   Begin VB.Label labCount 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   105
      Width           =   3240
   End
   Begin VB.Label Label3 
      Caption         =   "Or"
      Height          =   255
      Left            =   210
      TabIndex        =   4
      Top             =   1305
      Width           =   390
   End
   Begin VB.Label Label2 
      Caption         =   "% Aired:"
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   1650
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Aired Spot Count:"
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   930
      Width           =   1365
   End
End
Attribute VB_Name = "frmCPCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmCPCount - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer
Private imCount As Integer
Private tmCPDat() As DAT



Private Sub cmdCancel_Click()
    Unload frmCPCount
End Sub

Private Sub cmdOk_Click()
    Dim sStr As String
    Dim iAired As Integer
    Dim iPosted As Integer
    
    'On Error GoTo ErrHand
    If imFieldChgd = False Then
        Unload frmCPCount
        Exit Sub
    End If
    If sgUstWin(7) <> "I" Then
        Unload frmCPCount
        Exit Sub
    End If
    On Error GoTo ErrHand
    sStr = Trim$(txtSCount.Text)
    If sStr = "" Then
        sStr = Trim$(txtPCount.Text)
        If sStr <> "" Then
            iAired = (CLng(imCount) * Val(sStr) + 50) / 100
        Else
            iAired = 0
        End If
    Else
        iAired = Val(sStr)
    End If
    If iAired = 0 Then
        igCPStatus = 0
        igCPPostingStatus = 1
    Else
        igCPStatus = 1
        igCPPostingStatus = 2
    End If
    SQLQuery = "UPDATE cptt SET "
    SQLQuery = SQLQuery + "cpttNoSpotsGen = " & imCount & ", "
    SQLQuery = SQLQuery + "cpttNoSpotsAired = " & iAired & ", "
    SQLQuery = SQLQuery + "cpttStatus = " & igCPStatus & ", "
    SQLQuery = SQLQuery + "cpttPostingStatus = " & igCPPostingStatus
    SQLQuery = SQLQuery + " WHERE cpttCode = " & tgCPPosting(0).lCpttCode
    
    cnn.BeginTrans
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "CPCount-cmdOk_Click"
        cnn.RollbackTrans
        Exit Sub
    End If
    cnn.CommitTrans
    gFileChgdUpdate "cptt.mkd", True
    Screen.MousePointer = vbHourglass
    
    Unload frmCPCount
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "CP Count-cmdOk"
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / 3
    Me.Height = (Screen.Height) / 3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Dim iUpper As Integer
    Dim sFWkDate As String
    Dim sLWkDate As String
    Dim lLWkDate As Long
    Dim lFWkDate As Long
    Dim iFound As Integer
    Dim lSTime As Long
    Dim lETime As Long
    Dim lTime As Long
    Dim iDat As Integer
    Dim iLoop As Integer
    Dim sZone As String
    Dim iLocalAdj As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim iMatch As Integer
    Dim sStationName As String
    Dim sVefName As String
    Dim lLogDate As Long
    Dim iZoneDefined As Integer
    Dim ilPostingStatus As Integer
    Dim slPledgeType As String
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    frmCPCount.Caption = "Spot Count - " & sgClientName
    sZone = tgCPPosting(0).sZone
    iLocalAdj = 0
    iZoneDefined = False
    If Len(Trim$(tgCPPosting(0).sZone)) <> 0 Then
        'Get zone
        For iVef = 0 To UBound(tgVehicleInfo) - 1 Step 1
            If tgVehicleInfo(iVef).iCode = tgCPPosting(0).iVefCode Then
                For iZone = LBound(tgVehicleInfo(iVef).sZone) To UBound(tgVehicleInfo(iVef).sZone) Step 1
                    If (Trim$(tgVehicleInfo(iVef).sZone(iZone)) <> "") And (Trim$(tgVehicleInfo(iVef).sZone(iZone)) <> "~~~") Then
                        iZoneDefined = True
                    End If
                    If Trim$(tgVehicleInfo(iVef).sZone(iZone)) = Trim$(tgCPPosting(0).sZone) Then
                        If tgVehicleInfo(iVef).sFed(iZone) <> "*" Then
                            sZone = tgVehicleInfo(iVef).sZone(tgVehicleInfo(iVef).iBaseZone(iZone))
                            iLocalAdj = tgVehicleInfo(iVef).iLocalAdj(iZone)
                        End If
                        Exit For
                    End If
                Next iZone
                Exit For
            End If
        Next iVef
    End If
    If Not iZoneDefined Then
        sZone = ""
    End If
    
    'Me.Width = (Screen.Width) / 1.55
    'Me.Height = (Screen.Height) / 2.4
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    SQLQuery = "SELECT attPledgeType FROM att WHERE attCode = " & tgCPPosting(0).lAttCode
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        slPledgeType = rst!attPledgeType
    Else
        slPledgeType = ""
    End If
    imCount = 0
    ReDim tmCPDat(0 To 0) As DAT
    iUpper = 0
    SQLQuery = "SELECT * "
    SQLQuery = SQLQuery + " FROM dat"
    SQLQuery = SQLQuery + " WHERE (datatfCode= " & tgCPPosting(0).lAttCode
    'SQLQuery = SQLQuery + " AND datDACode =" & tgCPPosting(0).iAttTimeType & ")"
    SQLQuery = SQLQuery + " )"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        tmCPDat(iUpper).iStatus = 1
        tmCPDat(iUpper).lCode = rst!datCode    '(0).Value
        tmCPDat(iUpper).lAtfCode = rst!datAtfCode  '(1).Value
        tmCPDat(iUpper).iShfCode = rst!datShfCode  '(2).Value
        tmCPDat(iUpper).iVefCode = rst!datVefCode  '(3).Value
        'tmCPDat(iUpper).iDACode = rst!datDACode    '(4).Value
        tmCPDat(iUpper).iFdDay(0) = rst!datFdMon   '(5).Value
        tmCPDat(iUpper).iFdDay(1) = rst!datFdTue   '(6).Value
        tmCPDat(iUpper).iFdDay(2) = rst!datFdWed   '(7).Value
        tmCPDat(iUpper).iFdDay(3) = rst!datFdThu   '(8).Value
        tmCPDat(iUpper).iFdDay(4) = rst!datFdFri   '(9).Value
        tmCPDat(iUpper).iFdDay(5) = rst!datFdSat   '(10).Value
        tmCPDat(iUpper).iFdDay(6) = rst!datFdSun   '(11).Value
        tmCPDat(iUpper).sFdSTime = Format$(CStr(rst!datFdStTime), sgShowTimeWOSecForm)
        tmCPDat(iUpper).sFdETime = Format$(CStr(rst!datFdEdTime), sgShowTimeWOSecForm)
        tmCPDat(iUpper).iFdStatus = rst!datFdStatus    '(14).Value
        tmCPDat(iUpper).iPdDay(0) = rst!datPdMon   '(15).Value
        tmCPDat(iUpper).iPdDay(1) = rst!datPdTue   '(16).Value
        tmCPDat(iUpper).iPdDay(2) = rst!datPdWed   '(17).Value
        tmCPDat(iUpper).iPdDay(3) = rst!datPdThu   '(18).Value
        tmCPDat(iUpper).iPdDay(4) = rst!datPdFri   '(19).Value
        tmCPDat(iUpper).iPdDay(5) = rst!datPdSat   '(20).Value
        tmCPDat(iUpper).iPdDay(6) = rst!datPdSun   '(21).Value
        tmCPDat(iUpper).sPdDayFed = rst!datPdDayFed
        If tmCPDat(iUpper).iStatus <= 1 Then
            tmCPDat(iUpper).sPdSTime = Format$(CStr(rst!datPdStTime), sgShowTimeWOSecForm)
            If tmCPDat(iUpper).iStatus = 1 Then
                tmCPDat(iUpper).sPdETime = Format$(CStr(rst!datPdEdTime), sgShowTimeWOSecForm)
            Else
                tmCPDat(iUpper).sPdETime = ""
            End If
        Else
            tmCPDat(iUpper).sPdSTime = ""
            tmCPDat(iUpper).sPdETime = ""
        End If
        tmCPDat(iUpper).iAirPlayNo = rst!datAirPlayNo
        iUpper = iUpper + 1
        ReDim Preserve tmCPDat(0 To iUpper) As DAT
        rst.MoveNext
    Wend
    
    sFWkDate = Format$(gObtainPrevMonday(tgCPPosting(0).sDate), sgShowDateForm)
    sLWkDate = Format$(gObtainNextSunday(tgCPPosting(0).sDate), sgShowDateForm)
    lFWkDate = DateValue(gAdjYear(sFWkDate))
    lLWkDate = DateValue(gAdjYear(sLWkDate))
    If UBound(tmCPDat) <= LBound(tmCPDat) Then
        SQLQuery = "SELECT COUNT(lstCode) FROM lst "
        SQLQuery = SQLQuery + " WHERE (lstLogVefCode = " & tgCPPosting(0).iVefCode
        SQLQuery = SQLQuery & " AND lstType = 0"
        SQLQuery = SQLQuery + " AND lstBkoutLstCode = 0"
        If Trim$(sZone) <> "" Then
            SQLQuery = SQLQuery + " AND lstZone = '" & sZone & "'"
        End If
        SQLQuery = SQLQuery + " AND (lstLogDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " AND ((lstStatus <= 1) Or (lstStatus = 7))" & ")"
        'D.S. 061606
        'SQLQuery = SQLQuery + " ORDER BY lstLogDate, lstLogTime"
        
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            imCount = rst(0).Value
        Else
            imCount = 0
        End If
    Else
        SQLQuery = "SELECT lstLogDate, lstLogTime, lstStatus, lstCode FROM lst "
        SQLQuery = SQLQuery + " WHERE (lstLogVefCode = " & tgCPPosting(0).iVefCode
        If Trim$(sZone) <> "" Then
            SQLQuery = SQLQuery + " AND lstZone = '" & sZone & "'"
        End If
        SQLQuery = SQLQuery + " AND lstBkoutLstCode = 0"
        SQLQuery = SQLQuery + " AND (lstLogDate >= '" & Format$(lFWkDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(lLWkDate + 1, sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " AND ((lstStatus <= 1) Or (lstStatus = 7))" & ")"
        SQLQuery = SQLQuery + " ORDER BY lstLogDate, lstLogTime"
        
        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF
            'Find pledged
            iFound = False
            For iDat = 0 To UBound(tmCPDat) - 1 Step 1
                lLogDate = DateValue(gAdjYear(rst!lstLogDate))
                lSTime = gTimeToLong(tmCPDat(iDat).sFdSTime, False)
                
                'If tmCPDat(iDat).iDACode = 0 Or tmCPDat(iDat).iDACode = 2 Then
                If (slPledgeType = "D") Or (slPledgeType = "C") Then
                    lETime = gTimeToLong(tmCPDat(iDat).sFdETime, True)
                Else
                    lETime = lSTime + 1
                End If
                
                'If tmCPDat(iDat).iDACode = 2 Then
                If slPledgeType = "C" Then
                    lTime = gTimeToLong(rst!lstLogTime, False)
                Else
                    lTime = gTimeToLong(rst!lstLogTime, False) + 3600 * iLocalAdj
                    If lTime < 0 Then
                        lTime = lTime + 86400
                        lLogDate = lLogDate - 1
                    ElseIf lTime > 86400 Then
                        lTime = lTime - 86400
                        lLogDate = lLogDate + 1
                    End If
                End If
                If tmCPDat(iDat).iFdDay(Weekday(Format$(lLogDate, "m/d/yyyy"), vbMonday) - 1) Then
                    If (lLogDate >= lFWkDate) And (lLogDate <= lLWkDate) And (lTime >= lSTime) And (lTime < lETime) Then
                        If (tmCPDat(iDat).iFdStatus = 0) Or (tmCPDat(iDat).iFdStatus = 1) Or (tmCPDat(iDat).iFdStatus = 7) Or (tmCPDat(iDat).iFdStatus = 9) Or (tmCPDat(iDat).iFdStatus = 10) Then  '2=Not Carried
                            iFound = True
                        End If
                        Exit For
                    End If
                End If
            Next iDat
            If ((Not iFound) And (UBound(tmCPDat) <= LBound(tmCPDat))) Or (iFound) Then   'Treat as if live broadcast
                imCount = imCount + 1
            End If
            rst.MoveNext
        Wend
    End If
    
    SQLQuery = "SELECT cpttNoSpotsGen, cpttNoSpotsAired, cpttPostingStatus FROM cptt"
    SQLQuery = SQLQuery + " WHERE (cpttCode= " & tgCPPosting(0).lCpttCode & ")"
    
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        txtSCount.Text = rst!cpttNoSpotsAired
        ilPostingStatus = rst!cpttPostingStatus
    Else
        txtSCount.Text = ""
        ilPostingStatus = 0
    End If
    
    sVefName = ""
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        If tgVehicleInfo(iLoop).iCode = tgCPPosting(0).iVefCode Then
            sVefName = Trim$(tgVehicleInfo(iLoop).sVehicle)
            Exit For
        End If
    Next iLoop
        
    sStationName = ""
    For iLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(iLoop).iCode = tgCPPosting(0).iShttCode Then
            If Trim$(tgStationInfo(iLoop).sMarket) <> "" Then
                sStationName = Trim$(tgStationInfo(iLoop).sCallLetters) & ", " & Trim$(tgStationInfo(iLoop).sMarket)
            Else
                sStationName = Trim$(tgStationInfo(iLoop).sCallLetters)
            End If
            Exit For
        End If
    Next iLoop
        
    
    labCount(0).Caption = sStationName
    labCount(1).Caption = Trim$(Str$(imCount)) & " Spots Generated on " & sFWkDate & " for " & sVefName
    If ilPostingStatus = 0 Then
        txtSCount.Text = Trim$(Str$(imCount))
    End If
    If sgUstWin(7) <> "I" Then
        txtPCount.Enabled = False
        txtSCount.Enabled = False
        cmdOK.Enabled = False
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase tmCPDat
    Set frmCPCount = Nothing
End Sub
Private Sub txtSCount_Change()
    txtPCount.Text = ""
    imFieldChgd = True
End Sub

Private Sub txtSCount_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtPCount_Change()
    txtSCount.Text = ""
    imFieldChgd = True
End Sub

