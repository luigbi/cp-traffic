VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmAstCheckUtil 
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Action"
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   5175
      Begin VB.OptionButton rbcAction 
         Caption         =   "Correct  Records"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton rbcAction 
         Caption         =   "Check Only"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.ListBox lbcResults 
      Height          =   3960
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   7935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   5400
      Width           =   3375
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Execute"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   5400
      Width           =   3375
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   0
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6255
      FormDesignWidth =   10245
   End
End
Attribute VB_Name = "frmAstCheckUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmAstCcheckUtil - See notes below.
'*
'*  Created December,2007 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************

Option Explicit
Option Compare Text

Private tmAstChk() As AST_INFO
Private imTerminate As Integer
Private imChecking As Integer


' All cases general results:
' Affiliate Pledge status is Not Carried (astPledgeStatus) and the
' Affiliate Air Status (astStatus) indicates that the spot aired.
' We split the general cases into four types of groups as follows:
'
' Group #1: Affiliate Feed Time matched Agreement Pledge time and
' Affiliate Pledge Status does not match the Agreement Pledge Status
'
' Group #2: Affiliate Feed Time matched Agreement Pledge time and
' Affiliate Pledge Status matches Agreement Pledge Status
'
' Group #3:
' 3A: Affiliate Feed time fell within a Agreement Daypart Pledge times
' and Affiliate Pledge Status does not match the Agreement Pledge Status
'
' 3B: Affiliate Feed time fell within a Agreement Daypart Pledge times
' and Affiliate Pledge Status matches Agreement Pledge Status
'
' 3C: Affiliate Feed time does not fall within any Agreement Daypart time
'
' Group #4: No pledge times defined (spots treated as if Agreement Pledge
' status is defined as Live)
'
' Affiliate Feed Time:  astFeedTime
' Affiliate Pledge status: astPledgeStatus
' Affiliate Air Status: astStatus
'
' Agreement Pledge time:  datFdStTime
' Agreement Pledge status: datFdStatus
' Agreement Daypart Pledge times: datFdStTime and datFdEdTime
'
' How to treat each group:
' Group #1: Change Affiliate Pledge status to match Agreement Pledge status
'
' Group #2: Change the Affiliate Air Status to match the Affiliate Pledge Status
'
' Group #3:
'  3A: Change Affiliate Pledge status to match Agreement Pledge Status
'  3B: Change the Affiliate Air Status to match the Affiliate Pledge Status
'  3C: Change the Affiliate Air Status to match the Affiliate Pledge Status
'
' Group #4: Change Affiliate Pledge status to zero (aired live).
'
' Note:  To test Times, the Agreement Pledge Day(datFdMon, datFdTue,..)
' must match the Feed day (astFeedDate translated to week day).

Private Function mProcessAstRecs()

    Dim ast_Recs As ADODB.Recordset
    Dim cptt_Recs As ADODB.Recordset
    Dim rst_DAT As ADODB.Recordset
    
    Dim slUpdateStr As String
    Dim llPartial() As Long
    Dim ilAiredCnt As Integer
    Dim ilPartialCnt As Integer
    Dim llPartialWksCnt As Long
    Dim llNoPostWksCnt As Long
    Dim slEndDate As String
    Dim llLoop As Long
    
    Dim llIdx As Long
    Dim llIdx2 As Long
    Dim ilDatIdx As Integer
    Dim llMaxCode As Long
    Dim ilRecMeetsCondition As Integer
    Dim llTempIdx As Long
    Dim llTempIdx2 As Long
    Dim ilFound1 As Integer
    Dim ilFound2 As Integer
    Dim ilFound3A As Integer
    Dim ilFound3B As Integer
    Dim ilFound3C As Integer
    Dim ilFound4 As Integer
    Dim llFound As Long
    Dim llNoDat As Long
    Dim llFdTimeMatchStatusDoesnt As Long
    Dim llFdTimeMatchStatusMatch As Long
    Dim llNoFdTimesMatchA As Long
    Dim llNoFdTimesMatchB As Long
    Dim llNoFdTimesMatchC As Long
    Dim ilDayOk As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slDay As String
    Dim slVefName As String
    Dim ilVefCode As Integer
    Dim slStaName As String
    Dim ilFdPledge1 As Integer
    Dim ilFdPledge2 As Integer
    Dim ilFdPledge3a As Integer
    Dim ilFdPledge3b As Integer
    Dim tlDatPledgeInfo As DATPLEDGEINFO
    Dim ilRet As Integer
    
    llMaxCode = 20000
    ReDim tmAstChk(0 To llMaxCode) As AST_INFO

    On Error GoTo Err_Handler
    
    SetResults "Starting AST Check Utility", 0
    Call gLogMsg("Starting AST Check Utility", "ASTCheckUtility.txt", False)
    
    mProcessAstRecs = False

    'D.S. 07/28/08 Start
    'Look for partial weeks where either the only spots showing received are spots that were not supposed to air
    'and partial weeks that show no spots as being received.
    SQLQuery = "select * from cptt where cpttPostingStatus = 1"
    Set cptt_Recs = gSQLSelectCall(SQLQuery)
    llPartialWksCnt = 0
    llNoPostWksCnt = 0

    While Not cptt_Recs.EOF
        DoEvents
        If imTerminate Then
            cptt_Recs.Close
            Exit Function
        End If
        slEndDate = DateAdd("d", 6, cptt_Recs!CpttStartDate)
        SQLQuery = ""
        SQLQuery = SQLQuery + "Select * from ast WHERE (astAtfCode = " & cptt_Recs!cpttatfCode
        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(cptt_Recs!CpttStartDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slEndDate, sgSQLDateForm) & "')" & ")"
            
        Set ast_Recs = gSQLSelectCall(SQLQuery)
        ReDim llPartial(0 To 5000) As Long
        ilPartialCnt = 0
        ilAiredCnt = 0
        While Not ast_Recs.EOF
            'Get count of spots that show aired, but were not supposed to be carried or aired
            '12/13/13: Obtain Pledge information from Dat
            tlDatPledgeInfo.lAttCode = ast_Recs!astAtfCode
            tlDatPledgeInfo.lDatCode = ast_Recs!astDatCode
            tlDatPledgeInfo.iVefCode = ast_Recs!astVefCode
            tlDatPledgeInfo.sFeedDate = Format(ast_Recs!astFeedDate, "m/d/yy")
            tlDatPledgeInfo.sFeedTime = Format(ast_Recs!astFeedTime, "hh:mm:ssam/pm")
            ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
            
            If tgStatusTypes(gGetAirStatus(tlDatPledgeInfo.iPledgeStatus)).iPledged = 2 And ast_Recs!astCPStatus = 1 Then
                llPartial(ilPartialCnt) = CLng(ast_Recs!astCode)
                ilPartialCnt = ilPartialCnt + 1
            End If
            'Get count of spots that were supposed to air, and show received
            If tgStatusTypes(gGetAirStatus(tlDatPledgeInfo.iPledgeStatus)).iPledged <> 2 And ast_Recs!astCPStatus = 1 Then
                ilAiredCnt = ilAiredCnt + 1
            End If
            ast_Recs.MoveNext
        Wend
        
        'Case where NOT aired spot(s) show received, but no spots that were supposed to air are received
        If ilPartialCnt > 0 And ilAiredCnt = 0 Then
            For llLoop = 0 To ilPartialCnt - 1 Step 1
                slUpdateStr = "UPDATE ast SET astCPStatus = 0 where astCode = " & llPartial(llLoop)
                'cnn.Execute slUpdateStr
                If gSQLWaitNoMsgBox(slUpdateStr, False) <> 0 Then
                    mProcessAstRecs = False
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
            Next llLoop
            slUpdateStr = "UPDATE cptt SET cpttPostingStatus = 0 where cpttCode = " & cptt_Recs!cpttCode
            'cnn.Execute slUpdateStr
            If gSQLWaitNoMsgBox(slUpdateStr, False) <> 0 Then
                mProcessAstRecs = False
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            llPartialWksCnt = llPartialWksCnt + 1
        End If
        
        'Case where nothing shows posted, but the cptt is set to partial
        If (ilPartialCnt = 0 And ilAiredCnt = 0) Then
            slUpdateStr = "UPDATE cptt SET cpttPostingStatus = 0 where cpttCode = " & cptt_Recs!cpttCode
            'cnn.Execute slUpdateStr
            If gSQLWaitNoMsgBox(slUpdateStr, False) <> 0 Then
                mProcessAstRecs = False
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            llNoPostWksCnt = llNoPostWksCnt + 1
        End If
        
        ReDim llPartial(0 To 5000) As Long
        cptt_Recs.MoveNext
    Wend
    gFileChgdUpdate "cptt.mkd", True
    'D.S. 07/28/08 End
    
    SetResults "Starting Pass One of Three", 0
    Call gLogMsg("Starting Pass One of Three", "ASTCheckUtility.txt", False)
    
    'Pass One - gather all of the spots pledged to NOT AIR
    SQLQuery = "select * from ast where astCPStatus = 1"
    Set ast_Recs = gSQLSelectCall(SQLQuery)
    llIdx = 0
    llTempIdx = 0
    While Not ast_Recs.EOF
        DoEvents

        If imTerminate Then
            ast_Recs.Close
            Exit Function
        End If
        
        '12/13/13: Obtain Pledge information from Dat
        tlDatPledgeInfo.lAttCode = ast_Recs!astAtfCode
        tlDatPledgeInfo.lDatCode = ast_Recs!astDatCode
        tlDatPledgeInfo.iVefCode = ast_Recs!astVefCode
        tlDatPledgeInfo.sFeedDate = Format(ast_Recs!astFeedDate, "m/d/yy")
        tlDatPledgeInfo.sFeedTime = Format(ast_Recs!astFeedTime, "hh:mm:ssam/pm")
        ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)

        'Pledged to NOT Air
        If tgStatusTypes(gGetAirStatus(tlDatPledgeInfo.iPledgeStatus)).iPledged = 2 And tgStatusTypes(gGetAirStatus(ast_Recs!astStatus)).iPledged <> 2 Then
            ilRecMeetsCondition = True
        Else
            ilRecMeetsCondition = False
        End If

        If ilRecMeetsCondition Then
            tmAstChk(llIdx).lCode = ast_Recs!astCode
            tmAstChk(llIdx).lCode2 = -1
            tmAstChk(llIdx).lAttCode = ast_Recs!astAtfCode
            tmAstChk(llIdx).lSdfCode = ast_Recs!astSdfCode
            tmAstChk(llIdx).sAirDate = Format$(ast_Recs!astAirDate, sgShowDateForm)
            tmAstChk(llIdx).sAirTime = Format$(ast_Recs!astAirTime, sgSQLTimeForm)
            tmAstChk(llIdx).sAirTime2 = ""
            tmAstChk(llIdx).iStatus = ast_Recs!astStatus
            tmAstChk(llIdx).iStatus2 = -1
            tmAstChk(llIdx).lSdfCode = ast_Recs!astSdfCode
            tmAstChk(llIdx).lSdfCode2 = -1
            tmAstChk(llIdx).iPledgeStatus = tlDatPledgeInfo.iPledgeStatus

            slDate = gAdjYear(Format$(ast_Recs!astFeedDate, sgShowDateForm))
            tmAstChk(llIdx).lFeedDate = DateValue(slDate)

            slTime = Format(ast_Recs!astFeedTime, sgShowTimeWSecForm)
            tmAstChk(llIdx).lFeedTime = gTimeToLong(slTime, False)

            tmAstChk(llIdx).iFeedDay = Weekday(slDate)

            llIdx = llIdx + 1
            'Increase the array size if needed
            If llIdx = llMaxCode Then
                llMaxCode = llMaxCode + 5000
                ReDim Preserve tmAstChk(0 To llMaxCode) As AST_INFO
            End If
        End If
        ast_Recs.MoveNext
    Wend

    If llIdx > 0 Then
        ReDim Preserve tmAstChk(0 To llIdx) As AST_INFO
        ArraySortTyp fnAV(tmAstChk(), 0), UBound(tmAstChk), 0, LenB(tmAstChk(0)), 0, -2, 0

        Call gLogMsg("Pass One of Three has completed.", "ASTCheckUtility.txt", False)
        SetResults "Pass One of Three has completed.", 0

    Else
        Call gLogMsg("No records met pass one criteria. Nothing to do.", "ASTCheckUtility.txt", False)
        SetResults "No records met pass one criteria. Nothing to do.", 0
        Call gLogMsg("See file ASTCheckUtility.txt in your messages folder for results.", "ASTCheckUtility.txt", False)
        mProcessAstRecs = True
        Screen.MousePointer = vbDefault
        
        Exit Function
    End If
    
    'Pass Two - gather all of the spots pledged TO AIR
    Call gLogMsg("Starting Pass Two of Three.", "ASTCheckUtility.txt", False)
    SetResults "Starting Pass Two of Three.", 0
    SQLQuery = "Select * from ast where astCPStatus = 1"
    Set ast_Recs = gSQLSelectCall(SQLQuery)
    llTempIdx = 0

    While Not ast_Recs.EOF
        DoEvents
        
        If imTerminate Then
            ast_Recs.Close
            Exit Function
        End If
        
        '12/13/13: Obtain Pledge information from Dat
        tlDatPledgeInfo.lAttCode = ast_Recs!astAtfCode
        tlDatPledgeInfo.lDatCode = ast_Recs!astDatCode
        tlDatPledgeInfo.iVefCode = ast_Recs!astVefCode
        tlDatPledgeInfo.sFeedDate = Format(ast_Recs!astFeedDate, "m/d/yy")
        tlDatPledgeInfo.sFeedTime = Format(ast_Recs!astFeedTime, "hh:mm:ssam/pm")
        ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
        
        'Pledged to Aired
        If tgStatusTypes(gGetAirStatus(tlDatPledgeInfo.iPledgeStatus)).iPledged <> 2 Then
            ilRecMeetsCondition = True
        Else
            ilRecMeetsCondition = False
        End If

        If ilRecMeetsCondition Then
            llFound = gBinarySearchSdf(ast_Recs!astSdfCode, llIdx - 1)
            If llFound >= 0 Then
                Do
                    If tmAstChk(llFound).lSdfCode = ast_Recs!astSdfCode Then
                        If DateValue(gAdjYear((Format$(ast_Recs!astFeedDate, sgShowDateForm)))) = tmAstChk(llFound).lFeedDate Then
                            If ast_Recs!astAtfCode = tmAstChk(llFound).lAttCode Then
                                llTempIdx = llTempIdx + 1
                                tmAstChk(llFound).lCode2 = ast_Recs!astCode
                                tmAstChk(llFound).iStatus2 = ast_Recs!astStatus
                                tmAstChk(llFound).sAirTime2 = Trim$(Format$(ast_Recs!astAirTime, sgSQLTimeForm))
                                tmAstChk(llFound).lSdfCode2 = ast_Recs!astSdfCode
                                Exit Do
                            End If
                        End If
                    Else
                        Exit Do
                    End If
                    llFound = llFound + 1
                Loop While llFound < UBound(tmAstChk)
            End If
        End If
        ast_Recs.MoveNext
    Wend

    Call gLogMsg("Pass Two of Three has completed.", "ASTCheckUtility.txt", False)
    SetResults "Pass Two of Three has completed.", 0
    Call gLogMsg("Starting Final Pass Three.", "ASTCheckUtility.txt", False)
    SetResults "Starting Final Pass Three.", 0
    
    'Pass Three - Check or Fix Problems Found
    
    llTempIdx = 0
    llTempIdx2 = 0
    llNoDat = 0
    llFdTimeMatchStatusDoesnt = 0
    llFdTimeMatchStatusMatch = 0
    llNoFdTimesMatchA = 0
    llNoFdTimesMatchB = 0
    llNoFdTimesMatchC = 0

    For llIdx = 0 To UBound(tmAstChk) - 1 Step 1
        DoEvents
        If imTerminate Then
            Exit Function
        End If
        If tmAstChk(llIdx).lCode2 <> -1 Then
            llIdx2 = tmAstChk(llIdx).lCode2
            llTempIdx = llTempIdx + 1
            
            If rbcAction(1).Value Then
                SQLQuery = "UPDATE ast SET"
                SQLQuery = SQLQuery & " astAirTime = " & "'" & Format$(tmAstChk(llIdx).sAirTime2, sgSQLTimeForm) & "',"
                SQLQuery = SQLQuery & " astStatus = " & tmAstChk(llIdx).iPledgeStatus
                SQLQuery = SQLQuery & " WHERE astCode = " & tmAstChk(llIdx).lCode
                'cnn.Execute SQLQuery
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    mProcessAstRecs = False
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
    
                SQLQuery = "UPDATE ast SET"
                SQLQuery = SQLQuery & " astAirTime = " & "'" & Format(tmAstChk(llIdx).sAirTime, sgSQLTimeForm) & "',"
                SQLQuery = SQLQuery & " astStatus = " & tmAstChk(llIdx).iStatus
                SQLQuery = SQLQuery & " WHERE astCode = " & tmAstChk(llIdx).lCode2
                'cnn.Execute SQLQuery
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    mProcessAstRecs = False
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
            End If
            
            Call gLogMsg("Grp 0, AST Code = " & CStr(tmAstChk(llIdx).lCode) & " SDF Code = " & CStr(tmAstChk(llIdx).lSdfCode) & " ATT Code = " & CStr(tmAstChk(llIdx).lAttCode) & " changing astStatus from " & CStr(tmAstChk(llIdx).iStatus) & " to " & CStr(tmAstChk(llIdx).iStatus2) & " and astAirTime from " & Format(tmAstChk(llIdx).sAirTime, sgShowTimeWSecForm) & " to " & Format(tmAstChk(llIdx).sAirTime2, sgShowTimeWSecForm) & " Air Date " & Format(tmAstChk(llIdx).sAirDate, sgShowDateForm), "ASTCheckUtility.txt", False)
            Call gLogMsg("Grp 0, AST Code = " & CStr(tmAstChk(llIdx).lCode2) & " SDF Code = " & CStr(tmAstChk(llIdx).lSdfCode2) & " ATT Code = " & CStr(tmAstChk(llIdx).lAttCode) & " changing astStatus from " & CStr(tmAstChk(llIdx).iStatus2) & " to " & CStr(tmAstChk(llIdx).iStatus) & " and astAirTime from " & Format(tmAstChk(llIdx).sAirTime2, sgShowTimeWSecForm) & " to " & Format(tmAstChk(llIdx).sAirTime, sgShowTimeWSecForm), "ASTCheckUtility.txt", False)
        Else
            ilFound1 = False
            ilFound2 = False
            ilFound3A = False
            ilFound3B = False
            ilFound3C = False
            ilFound4 = False
            SQLQuery = "Select * FROM DAT WHERE datAtfCode = " & tmAstChk(llIdx).lAttCode
            Set rst_DAT = gSQLSelectCall(SQLQuery)
            If Not rst_DAT.EOF Then
                DoEvents
                Do While Not rst_DAT.EOF
                    ilDayOk = False
                    Select Case tmAstChk(llIdx).iFeedDay
                    Case vbMonday
                        If rst_DAT!datFdMon = 1 Then
                            ilDayOk = True
                            slDay = "Mon"
                        End If
                    Case vbTuesday
                        If rst_DAT!datFdTue = 1 Then
                            ilDayOk = True
                            slDay = "Tue"
                        End If
                    Case vbWednesday
                        If rst_DAT!datFdWed = 1 Then
                            ilDayOk = True
                            slDay = "Wed"
                        End If
                    Case vbThursday
                        If rst_DAT!datFdThu = 1 Then
                            ilDayOk = True
                            slDay = "Thu"
                        End If
                    Case vbFriday
                        If rst_DAT!datFdFri = 1 Then
                            ilDayOk = True
                            slDay = "Fri"
                        End If
                    Case vbSaturday
                        If rst_DAT!datFdSat = 1 Then
                            ilDayOk = True
                            slDay = "Sat"
                        End If
                    Case vbSunday
                        If rst_DAT!datFdSun = 1 Then
                            ilDayOk = True
                            slDay = "Sun"
                        End If
                    End Select
                    
                    If ilDayOk Then
                        If tmAstChk(llIdx).lFeedTime = gTimeToLong(Format(rst_DAT!datFdStTime, sgShowTimeWSecForm), False) And tmAstChk(llIdx).iPledgeStatus <> rst_DAT!datFdStatus Then
                            ilFdPledge1 = rst_DAT!datFdStatus
                            ilFound1 = True
                        End If
                        If tmAstChk(llIdx).lFeedTime = gTimeToLong(Format(rst_DAT!datFdStTime, sgShowTimeWSecForm), False) And tmAstChk(llIdx).iPledgeStatus = rst_DAT!datFdStatus Then
                            ilFdPledge2 = rst_DAT!datFdStatus
                            ilFound2 = True
                            Exit Do
                        End If
                        If (tmAstChk(llIdx).lFeedTime >= gTimeToLong(Format(rst_DAT!datFdStTime, sgShowTimeWSecForm), False)) And (tmAstChk(llIdx).lFeedTime < gTimeToLong(Format(rst_DAT!datFdEdTime, sgShowTimeWSecForm), False)) And tmAstChk(llIdx).iPledgeStatus <> rst_DAT!datFdStatus Then
                            ilFdPledge3a = rst_DAT!datFdStatus
                            ilFound3A = True
                        End If
                        If (tmAstChk(llIdx).lFeedTime >= gTimeToLong(Format(rst_DAT!datFdStTime, sgShowTimeWSecForm), False)) And (tmAstChk(llIdx).lFeedTime < gTimeToLong(Format(rst_DAT!datFdEdTime, sgShowTimeWSecForm), False)) And tmAstChk(llIdx).iPledgeStatus = rst_DAT!datFdStatus Then
                            ilFdPledge3b = rst_DAT!datFdStatus
                            ilFound3B = True
                        End If
                    End If
                    rst_DAT.MoveNext
                Loop
                If (ilFound1) Or (ilFound2) Then
                    ilFound3A = False
                    ilFound3B = False
                    ilFound3B = False
                End If
                
                If ilFound3B Then
                    ilFound3A = False
                End If
                
                If (ilFound1 = False) And (ilFound2 = False) And (ilFound3A = False) And (ilFound3B = False) Then
                    ilFound3C = True
                End If
            Else
                ilFound4 = True
            End If
            
            slDate = gAdjYear(Format$(tmAstChk(llIdx).lFeedDate, sgShowDateForm))
            slTime = Format(gLongToTime(tmAstChk(llIdx).lFeedTime), sgShowTimeWSecForm)
            slStaName = gGetCallLettersByAttCode(tmAstChk(llIdx).lAttCode)
            ilVefCode = gGetVehCodeFromAttCode(CStr(tmAstChk(llIdx).lAttCode))
            slVefName = Trim$(gGetVehNameByVefCode(ilVefCode))
            
            If ilFound2 = True Then
                If rbcAction(1).Value Then
                    SQLQuery = "UPDATE ast SET"
                    SQLQuery = SQLQuery & " astStatus = " & ilFdPledge2
                    SQLQuery = SQLQuery & " WHERE astCode = " & tmAstChk(llIdx).lCode
                    'cnn.Execute SQLQuery
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        mProcessAstRecs = False
                        Screen.MousePointer = vbDefault
                        Exit Function
                    End If
                End If
                llFdTimeMatchStatusMatch = llFdTimeMatchStatusMatch + 1
                Call gLogMsg("Grp 2, AST: " & CStr(tmAstChk(llIdx).lCode) & ", SDF: " & CStr(tmAstChk(llIdx).lSdfCode) & ", ATT: " & CStr(tmAstChk(llIdx).lAttCode) & ", AST Sts: " & CStr(tmAstChk(llIdx).iStatus) & ", AST AirTm: " & Format(tmAstChk(llIdx).sAirTime, sgShowTimeWSecForm) & ", AST AirDt: " & Format(tmAstChk(llIdx).sAirDate, sgShowDateForm) & ", AST FdDt: " & slDate & ", AST FdTm: " & slTime & ", Day: " & slDay & ", AstPlgSts: " & tmAstChk(llIdx).iPledgeStatus & ", " & slStaName & ", " & slVefName, "ASTCheckUtility.txt", False)
            Else
                If ilFound1 = True Then
                    'Change Affiliate Pledge status to match Agreement Pledge status
                    If rbcAction(1).Value Then
                        '12/13/13: Pledge now obtained from DAT
                        'SQLQuery = "UPDATE ast SET"
                        'SQLQuery = SQLQuery & " astPledgeStatus = " & ilFdPledge1
                        'SQLQuery = SQLQuery & " WHERE astCode = " & tmAstChk(llIdx).lCode
                        'cnn.Execute SQLQuery
                    '12/13/13: Pledge now obtained from DAT
                    Else
                        llFdTimeMatchStatusDoesnt = llFdTimeMatchStatusDoesnt + 1
                        Call gLogMsg("Grp 1, AST: " & CStr(tmAstChk(llIdx).lCode) & ", SDF: " & CStr(tmAstChk(llIdx).lSdfCode) & ", ATT: " & CStr(tmAstChk(llIdx).lAttCode) & ", AST Sts: " & CStr(tmAstChk(llIdx).iStatus) & ", AST AirTm: " & Format(tmAstChk(llIdx).sAirTime, sgShowTimeWSecForm) & ", AST AirDt: " & Format(tmAstChk(llIdx).sAirDate, sgShowDateForm) & ", AST FdDt: " & slDate & ", AST FdTm: " & slTime & ", Day: " & slDay & ", AstPlgSts: " & tmAstChk(llIdx).iPledgeStatus & ", " & slStaName & ", " & slVefName, "ASTCheckUtility.txt", False)
                    End If
                    'llFdTimeMatchStatusDoesnt = llFdTimeMatchStatusDoesnt + 1
                    'Call gLogMsg("Grp 1, AST: " & CStr(tmAstChk(llIdx).lCode) & ", SDF: " & CStr(tmAstChk(llIdx).lSdfCode) & ", ATT: " & CStr(tmAstChk(llIdx).lAttCode) & ", AST Sts: " & CStr(tmAstChk(llIdx).iStatus) & ", AST AirTm: " & Format(tmAstChk(llIdx).sAirTime, sgShowTimeWSecForm) & ", AST AirDt: " & Format(tmAstChk(llIdx).sAirDate, sgShowDateForm) & ", AST FdDt: " & slDate & ", AST FdTm: " & slTime & ", Day: " & slDay & ", AstPlgSts: " & tmAstChk(llIdx).iPledgeStatus & ", " & slStaName & ", " & slVefName, "ASTCheckUtility.txt", False)
                End If
            End If
            
            If ilFound3A = True Then
                If rbcAction(1).Value Then
                    '12/13/13: Pledge now obtained from DAT
                    'SQLQuery = "UPDATE ast SET"
                    'SQLQuery = SQLQuery & " astPledgeStatus = " & ilFdPledge3a
                    'SQLQuery = SQLQuery & " WHERE astCode = " & tmAstChk(llIdx).lCode
                    'cnn.Execute SQLQuery
                '12/13/13: Pledge now obtained from DAT
                Else
                    llNoFdTimesMatchA = llNoFdTimesMatchA + 1
                    Call gLogMsg("Grp 3A, AST: " & CStr(tmAstChk(llIdx).lCode) & ", SDF: " & CStr(tmAstChk(llIdx).lSdfCode) & ", ATT: " & CStr(tmAstChk(llIdx).lAttCode) & ", AST Sts: " & CStr(tmAstChk(llIdx).iStatus) & ", AST AirTm: " & Format(tmAstChk(llIdx).sAirTime, sgShowTimeWSecForm) & ", AST AirDt: " & Format(tmAstChk(llIdx).sAirDate, sgShowDateForm) & ", AST FdDt: " & slDate & ", AST FdTm: " & slTime & ", Day: " & slDay & ", AstPlgSts: " & tmAstChk(llIdx).iPledgeStatus & ", " & slStaName & ", " & slVefName, "ASTCheckUtility.txt", False)
                End If
                'llNoFdTimesMatchA = llNoFdTimesMatchA + 1
                'Call gLogMsg("Grp 3A, AST: " & CStr(tmAstChk(llIdx).lCode) & ", SDF: " & CStr(tmAstChk(llIdx).lSdfCode) & ", ATT: " & CStr(tmAstChk(llIdx).lAttCode) & ", AST Sts: " & CStr(tmAstChk(llIdx).iStatus) & ", AST AirTm: " & Format(tmAstChk(llIdx).sAirTime, sgShowTimeWSecForm) & ", AST AirDt: " & Format(tmAstChk(llIdx).sAirDate, sgShowDateForm) & ", AST FdDt: " & slDate & ", AST FdTm: " & slTime & ", Day: " & slDay & ", AstPlgSts: " & tmAstChk(llIdx).iPledgeStatus & ", " & slStaName & ", " & slVefName, "ASTCheckUtility.txt", False)
            End If
            
            If ilFound3B = True Then
                If rbcAction(1).Value Then
                    SQLQuery = "UPDATE ast SET"
                    SQLQuery = SQLQuery & " astStatus = " & ilFdPledge3b
                    SQLQuery = SQLQuery & " WHERE astCode = " & tmAstChk(llIdx).lCode
                    'cnn.Execute SQLQuery
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        mProcessAstRecs = False
                        Screen.MousePointer = vbDefault
                        Exit Function
                    End If
                End If
                llNoFdTimesMatchB = llNoFdTimesMatchB + 1
                Call gLogMsg("Grp 3B, AST: " & CStr(tmAstChk(llIdx).lCode) & ", SDF: " & CStr(tmAstChk(llIdx).lSdfCode) & ", ATT: " & CStr(tmAstChk(llIdx).lAttCode) & ", AST Sts: " & CStr(tmAstChk(llIdx).iStatus) & ", AST AirTm: " & Format(tmAstChk(llIdx).sAirTime, sgShowTimeWSecForm) & ", AST AirDt: " & Format(tmAstChk(llIdx).sAirDate, sgShowDateForm) & ", AST FdDt: " & slDate & ", AST FdTm: " & slTime & ", Day: " & slDay & ", AstPlgSts: " & tmAstChk(llIdx).iPledgeStatus & ", " & slStaName & ", " & slVefName, "ASTCheckUtility.txt", False)
            End If
            
            If ilFound3C = True Then
                If rbcAction(1).Value Then
                    SQLQuery = "UPDATE ast SET"
                    SQLQuery = SQLQuery & " astStatus = " & tmAstChk(llIdx).iPledgeStatus
                    SQLQuery = SQLQuery & " WHERE astCode = " & tmAstChk(llIdx).lCode
                    'cnn.Execute SQLQuery
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        mProcessAstRecs = False
                        Screen.MousePointer = vbDefault
                        Exit Function
                    End If
                End If
                llNoFdTimesMatchC = llNoFdTimesMatchC + 1
                Call gLogMsg("Grp 3C, AST: " & CStr(tmAstChk(llIdx).lCode) & ", SDF: " & CStr(tmAstChk(llIdx).lSdfCode) & ", ATT: " & CStr(tmAstChk(llIdx).lAttCode) & ", AST Sts: " & CStr(tmAstChk(llIdx).iStatus) & ", AST AirTm: " & Format(tmAstChk(llIdx).sAirTime, sgShowTimeWSecForm) & ", AST AirDt: " & Format(tmAstChk(llIdx).sAirDate, sgShowDateForm) & ", AST FdDt: " & slDate & ", AST FdTm: " & slTime & ", Day: " & slDay & ", AstPlgSts: " & tmAstChk(llIdx).iPledgeStatus & ", " & slStaName & ", " & slVefName, "ASTCheckUtility.txt", False)
            End If
            
            If ilFound4 = True Then
                If rbcAction(1).Value Then
                    '12/13/13: Pledge now obtained from DAT
                    'SQLQuery = "UPDATE ast SET"
                    'SQLQuery = SQLQuery & " astPledgeStatus = 0"
                    'SQLQuery = SQLQuery & " WHERE astCode = " & tmAstChk(llIdx).lCode
                    'cnn.Execute SQLQuery
                '12/13/13: Pledge now obtained from DAT
                Else
                    'llNoDat = llNoDat + 1
                    'Call gLogMsg("Grp 4, AST: " & CStr(tmAstChk(llIdx).lCode) & ", SDF: " & CStr(tmAstChk(llIdx).lSdfCode) & ", ATT: " & CStr(tmAstChk(llIdx).lAttCode) & ", AST Sts: " & CStr(tmAstChk(llIdx).iStatus) & ", AST AirTm: " & Format(tmAstChk(llIdx).sAirTime, sgShowTimeWSecForm) & ", AST AirDt: " & Format(tmAstChk(llIdx).sAirDate, sgShowDateForm) & ", AST FdDt: " & slDate & ", AST FdTm: " & slTime & ", Day: " & slDay & ", AstPlgSts: " & tmAstChk(llIdx).iPledgeStatus & ", " & slStaName & ", " & slVefName, "ASTCheckUtility.txt", False)
                End If
                'llNoDat = llNoDat + 1
                'Call gLogMsg("Grp 4, AST: " & CStr(tmAstChk(llIdx).lCode) & ", SDF: " & CStr(tmAstChk(llIdx).lSdfCode) & ", ATT: " & CStr(tmAstChk(llIdx).lAttCode) & ", AST Sts: " & CStr(tmAstChk(llIdx).iStatus) & ", AST AirTm: " & Format(tmAstChk(llIdx).sAirTime, sgShowTimeWSecForm) & ", AST AirDt: " & Format(tmAstChk(llIdx).sAirDate, sgShowDateForm) & ", AST FdDt: " & slDate & ", AST FdTm: " & slTime & ", Day: " & slDay & ", AstPlgSts: " & tmAstChk(llIdx).iPledgeStatus & ", " & slStaName & ", " & slVefName, "ASTCheckUtility.txt", False)
            End If
        End If
    Next llIdx

    'Show Results on the Screen
    SetResults "Total: AST Switched: " & CStr(llTempIdx), 0
    SetResults "Total: CPTT Weeks Set to Oustanding.  Only Not Aired Spots Were Showing Received: " & CStr(llPartialWksCnt), 0
    SetResults "Total: CPTT Weeks Set to Oustanding.  No Aired Spots Were Showing Received: " & CStr(llNoPostWksCnt), 0
    
    SetResults "Total: Ast FeedTimes Match, but Ast PledgeStatus Does not: " & CStr(llFdTimeMatchStatusDoesnt), 0
    SetResults "Total: Ast FeedTimes Match and Ast PledgeStatus Matches: " & CStr(llFdTimeMatchStatusMatch), 0
    SetResults "Total: No Matching Pledge Time Grp 3A: " & CStr(llNoFdTimesMatchA), 0
    SetResults "Total: No Matching Pledge Time Grp 3B: " & CStr(llNoFdTimesMatchB), 0
    SetResults "Total: No Matching Pledge Time Grp 3C: " & CStr(llNoFdTimesMatchC), 0
    SetResults "Total: No DAT record Found: " & CStr(llNoDat), 0
    SetResults "See file ASTCheckUtility.txt in your messages folder for results.", 0
    SetResults "Press Done to end the program.", 0

    'Write the Results to the log file
    Call gLogMsg(" ", "ASTCheckUtility.txt", False)
    Call gLogMsg("CPTT Weeks Set to Oustanding.  Only Not Aired Spots Were Showing Received: " & CStr(llPartialWksCnt), "ASTCheckUtility.txt", False)
    Call gLogMsg("CPTT Weeks Set to Oustanding.  No Aired Spots Were Showing Received: " & CStr(llNoPostWksCnt), "ASTCheckUtility.txt", False)
    Call gLogMsg("Grp 0: Total AST Switched: " & CStr(llTempIdx), "ASTCheckUtility.txt", False)
    Call gLogMsg("Grp 1: Total Ast FeedTimes Match, but Ast Pledge Status Does Not Match: " & CStr(llFdTimeMatchStatusDoesnt), "ASTCheckUtility.txt", False)
    Call gLogMsg("Grp 2: Ast FeedTimes Match and Ast Pledge Status Matches: " & CStr(llFdTimeMatchStatusMatch), "ASTCheckUtility.txt", False)
    Call gLogMsg("Grp 3A: Total Ast FeedTimes Match Daypart, but AST Pledge Status Does Not Match: " & CStr(llNoFdTimesMatchA), "ASTCheckUtility.txt", False)
    Call gLogMsg("Grp 3B: Total Ast FeedTimes Match Daypart, but AST Pledge Status Matches: " & CStr(llNoFdTimesMatchB), "ASTCheckUtility.txt", False)
    Call gLogMsg("Grp 3C: Total No Matching Pledge Times, but Pledge Exists: " & CStr(llNoFdTimesMatchC), "ASTCheckUtility.txt", False)
    Call gLogMsg("Grp 4: Total No DAT record Found: " & CStr(llNoDat), "ASTCheckUtility.txt", False)

    mProcessAstRecs = True
    Screen.MousePointer = vbDefault
    ast_Recs.Close
    rst_DAT.Close
    cmdCancel.Caption = "&Done"
    Exit Function
    
Err_Handler:
    mProcessAstRecs = False
    gHandleError "AffErrorLog.txt", "frmAstCheckUtil-mProcessAstRecs"
    Screen.MousePointer = vbDefault
End Function


Private Function gBinarySearchSdf(llCode As Long, lMax As Long) As Long
         
     'D.S. 01/16/06
     'Returns the index number of tmAstChk that matches the lstCode that was passed in
     'Note: for this to work tglsttInfo was previously sorted
     
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim llCheck As Long
             
    On Error GoTo ErrHand
    
If llCode = 120531 Then
    llMin = llMin
End If
    
    
    llMin = LBound(tmAstChk)
    llMax = lMax
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tmAstChk(llMiddle).lSdfCode Then
             'found the match
            'gBinarySearchSdf = llMiddle
            'Backup to first match
            If llMiddle > 0 Then
                llCheck = llMiddle - 1
                Do While llCode = tmAstChk(llCheck).lSdfCode
                    llCheck = llCheck - 1
                    If llCheck < LBound(tmAstChk) Then
                        Exit Do
                    End If
                Loop
                gBinarySearchSdf = llCheck + 1
                Exit Function
            Else
                gBinarySearchSdf = llMiddle
                Exit Function
            End If
        ElseIf llCode < tmAstChk(llMiddle).lSdfCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchSdf = -1
    Exit Function
     
ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in gBinarySearchSdf: "
        Call gLogMsg("Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "AffErrorLog.Txt", False)
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    gBinarySearchSdf = -1
    Exit Function
End Function

Private Sub cmdCancel_Click()
    If imChecking Then
        imTerminate = True
        Exit Sub
    End If
    imTerminate = True
    Unload frmAstCheckUtil
End Sub

Private Sub cmdUpdate_Click()
    
    Dim ilRet As Integer
    
    If imChecking Then
        Exit Sub
    End If
    
    ilRet = False
    
    CSPWord.Show vbModal
    If igPasswordOk Then
        cmdUpdate.Enabled = True
    Else
        Call gLogMsg("User Failed to Provide Correct Password", "ASTCheckUtility.txt", False)
    End If
    
    If igPasswordOk Then
        imChecking = True
        Screen.MousePointer = vbHourglass
        ilRet = mProcessAstRecs
        imChecking = False
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Sub Form_Load()
    
    igPasswordOk = False
    imTerminate = False
    imChecking = False
    cmdUpdate.Enabled = True
    frmAstCheckUtil.Caption = "Spot AST Check Utility - " & sgClientName

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    gLogMsg "", "ASTCheckUtility.txt", False
    If imTerminate And cmdCancel.Caption = "&Cancel" Then
        gLogMsg "   *** User Terminated Program.  Ending Ast Check Utility Program   ***", "ASTCheckUtility.txt", False
    Else
        gLogMsg "   *** Ending Ast Check Utility Program   ***", "ASTCheckUtility.txt", False
    End If
    
    Erase tmAstChk
    igPasswordOk = False
    Unload frmAstCheckUtil

End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmAstCheckUtil
    gCenterForm frmAstCheckUtil

End Sub


Private Sub SetResults(sMsg As String, lFGC As Long)
    lbcResults.AddItem sMsg
    lbcResults.ListIndex = lbcResults.ListCount - 1
    lbcResults.ForeColor = lFGC
    DoEvents
End Sub

Private Sub Option1_Click(Index As Integer)

End Sub
