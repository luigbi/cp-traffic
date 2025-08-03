Attribute VB_Name = "modAgmnt"
'******************************************************
'*  modAgmnt - various global declarations
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Type CPTTARRAY
    sCpttStartDate As String * 10
    lCpttCode As Long
End Type

Type ADJUSTDATES
    lAttCode As Long
    sAttAgreeStart As String * 10
    sAttAgreeEnd As String * 10
End Type

Type FASTADDATTCOUNT
    iShttCode As Integer
    iShttCount As Integer
End Type
    
Type DAT
    iStatus As Integer        '0=New; 1=Used
    lCode As Long             'Auto Code
    lAtfCode As Long          'AttCode (Agreement)
    iShfCode As Integer       'ShttCode (Station)
    iVefCode As Integer       'VefCode (Vehicle)
    'iDACode As Integer        '0=Live Daypart; 1=Avail; 2 = CD/Tape Daypart
    iFdDay(0 To 6) As Integer 'Feed Day 0=
    sFdSTime As String * 10   'Feed Start Time
    sFdETime As String * 10   'Feed End Time
    iFdStatus As Integer
    iPdDay(0 To 6) As Integer 'Pledged Days
    sPdDayFed As String * 1
    sPdSTime As String * 10   'Pledged Start Time
    sPdETime As String * 10   'Pleged End Time
    iAirPlayNo As Integer     'Air play number
    sEstimatedTime As String * 1         ' Estimated Time allowed to be defined (Y or N). Test for Y.  This is only valid for Dayparts
    sEmbeddedOrROS As String * 1      ' Delivery Embedded spots or ROS spots (E/R).  Test for E. Blank is the same as R.
    iFirstET As Integer
    iRdfCode As Integer
End Type

Public tgDat() As DAT

'5/24/16: Add gathering of program times
Type PRGTIMES
    iDay As Integer '0=Mon; 1=Tues
    lPrgStartTime As Long   'Program start time
    lPrgEndTime As Long
End Type

Public tgPrgTimes() As PRGTIMES

Type SSFAVAIL
    iRecType As Integer
    iTime0 As Integer
    iTime1 As Integer
    iLtfCode As Integer
    iAvInfo As Integer
    iLen As Integer
    iAnfCode As Integer
    iNoSpotsThis As Integer
    iUnused1 As Integer
    iUnused2 As Integer
End Type

Type AVINFO
    iIndex0 As Integer
    iIndex1 As Integer
    iDay(0 To 6) As Integer
End Type

Type UNDOAVTIME
    lAttCode As Long
    iShfCode As Integer
    iVefCode As Integer
    sOffDate As String
End Type

Type AGMNTOVERLAPINFO
    lAttCode As Long
    lOnAirDate As Long
    lOffAirDate As Long
    lDropDate As Long
    iShfCode As Integer
End Type

Type AIRPLAYSPEC
    iAirPlayNo As Integer
    sAction As String * 1   'A=Add; R=Replace
    iFirstBO As Integer
End Type

Type BREAKOUTSPEC
    sType As String * 8
    sStatus As String * 15
    iFirstDP As Integer
    sPledgeStartTime As String * 12
    sPledgeEndTime As String * 12
    sEstimatedTime As String * 3
    sDays As String * 30
    sPledgeTime As String * 12
    iPledgeOffsetDay As Integer
    sPartialStartTime As String * 12
    sPartialEndTime As String * 12
    iNextBO As Integer
End Type

Type DPSELECTION
    sName As String * 20
    iIndex As Integer
    iRdfCode As Integer
    iNextDP As Integer
End Type

Public tgAirPlaySpec() As AIRPLAYSPEC
Public tgBreakoutSpec() As BREAKOUTSPEC
Public tgDPSelection() As DPSELECTION

'Estimated Time information
Type ETAVAILINFO
    sFdDay As String * 2    'Avail Feed Day name
    sFdTime As String * 12  'Avail Feed Time
    sETDay As String * 2    'Estimated Day name
    sETTime As String * 12  'Estimated Time
    lEptCode As Long
    iNextET As Integer      'Next Avail information
End Type

'I went ahead and made this a type in the event later we needed to add more information
Type LATESTRATECARD
    iLatestRCFCode As Integer
End Type
Public tgLatestRateCard() As LATESTRATECARD

Type TIME
    lSTime() As Long
    lETime() As Long
End Type

Type REMAPINFO
    iVefCode As Integer
    sStartDate As String
End Type

Type ATTINFO
    lAttCode As Long
    iStnCode As Integer
    sStnName As String
    iSelected As Integer
End Type

Type STANAMECODE
    sStationName As String * 40
    sInfo As String * 100
    iStationCode As Integer
End Type

Public tgStaNameAndCode() As STANAMECODE
Public tgAttInfo() As ATTINFO
Public tgRemapInfo As REMAPINFO
Public tmFDDayTime(0 To 6) As TIME
Public tmPDDayTime(0 To 6) As TIME
Public tgRdfCodes() As Integer
Public igLiveDayPart As Integer
Public igCDTapeDayPart As Integer
Public sgAirPlay1TimeType As String * 1
Public lgPledgeAttCode As Long
Public igNoAirPlays As Integer
Public igDefaultAirPlayNo As Integer
Public sgVehProgStartTime As String
Public sgVehProgEndTime As String
Public igAvails As Integer
Public igAgmntReturn As Integer
Public igDayPartShttCode As Integer
Public igDayPartVefCode As Integer
Public igDayPartVefCombo As Integer
Public igOkToRemap As Integer
Public igReload As Integer
Public igPledgeExist As Integer
Public sgCDStartTime As String
Public igCDStartTimeOK As Integer
Public igFastAddContinue As Integer
'D.S. 5/22/18 added cancel button
Public bgFastAddCancelButton As Boolean
Public igReturnPledgeStatus As Integer 'Used to indicate Pledge status from Daypart
Public lgLine As Long

Private rst_att As ADODB.Recordset
Private rst_lcf As ADODB.Recordset
Private rst_lvf As ADODB.Recordset
Private rst_Vpf As ADODB.Recordset
Private rst_abf As ADODB.Recordset
Private rst_Cptt As ADODB.Recordset

Private dlf_rst As ADODB.Recordset
Public DlfInfo_rst As ADODB.Recordset
Public bgDlfExist As Boolean
Private bmBypassZeroUnits As Boolean




Public Function gPopSelRemap(frm As Form, ilListCount As Integer, iIndex As Integer) As Integer

    Dim att_rst As ADODB.Recordset
    Dim shtt_rst As ADODB.Recordset
    Dim ilIdx As Integer
    
    On Error GoTo ErrHand
    
    'Screen.MousePointer = vbHourglass
    SQLQuery = "SELECT attCode, attShfCode"
    SQLQuery = SQLQuery + " FROM att"
    SQLQuery = SQLQuery + " WHERE (attVefCode = " & tgRemapInfo.iVefCode
    SQLQuery = SQLQuery + " AND attOffAir >= '" & Format$(tgRemapInfo.sStartDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery + " AND attDropDate >= '" & Format$(tgRemapInfo.sStartDate, sgSQLDateForm) & "'" & ")"
    Set att_rst = gSQLSelectCall(SQLQuery)

    ReDim tgAttInfo(0 To 0) As ATTINFO
    While Not att_rst.EOF
        tgAttInfo(UBound(tgAttInfo)).lAttCode = att_rst!attCode
        'We only need stn code is we allow users to pick affiliates
        'If frmAvRemap!lbcAvail(1).ListCount > 0 Then
        If ilListCount > 0 Then
            tgAttInfo(UBound(tgAttInfo)).iStnCode = att_rst!attshfcode
        End If
        'default value in the case that the New Times column had no times left in it.
        'If New Times is empty we don't allow users to select which affiliates to remap.
        'We do them all.  If New Times was not empty we allow the users to select which
        'affiliates to remap and set the iSelected element at that time.
        'If frmAvRemap!lbcAvail(1).ListCount = 0 Then
        If ilListCount = 0 Then
            tgAttInfo(UBound(tgAttInfo)).iSelected = True
        End If
        ReDim Preserve tgAttInfo(0 To (UBound(tgAttInfo) + 1))
        att_rst.MoveNext
    Wend
    
    'We don't need this info if we are not going to allow users to select affiliates to remap
    'If frmAvRemap!lbcAvail(1).ListCount > 0 Then
    If ilListCount > 0 Then
        For ilIdx = 0 To UBound(tgAttInfo) - 1 Step 1
            'SQLQuery = "SELECT shttCallLetters, shttMarket"
            'SQLQuery = SQLQuery + " FROM shtt"
            SQLQuery = "SELECT shttCallLetters, mktName"
            SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode"
            SQLQuery = SQLQuery + " WHERE (shttCode = " & tgAttInfo(ilIdx).iStnCode & ")"
            Set shtt_rst = gSQLSelectCall(SQLQuery)
            'tgAttInfo(ilIdx).sStnName = Trim$(shtt_rst!shttCallLetters) & ", " & Trim$(shtt_rst!shttMarket)
            If IsNull(shtt_rst!mktName) = True Then
                tgAttInfo(ilIdx).sStnName = Trim$(shtt_rst!shttCallLetters)
            Else
                tgAttInfo(ilIdx).sStnName = Trim$(shtt_rst!shttCallLetters) & ", " & Trim$(shtt_rst!mktName)
            End If
        Next ilIdx
    End If
    
    'Index 0 populate the to be remapped side; Index 1 populate the Not remapped side
    If iIndex = 0 Then
        'frmSelRemap!lbcSelRemap(1).Clear
        frm!lbcSelRemap(1).Clear
    Else
        'frmSelRemap!lbcSelRemap(0).Clear
        frm!lbcSelRemap(0).Clear
    End If
    
    'Screen.MousePointer = vbDefault
    For ilIdx = 0 To UBound(tgAttInfo) - 1 Step 1
        'frmSelRemap!lbcSelRemap(iIndex).AddItem tgAttInfo(ilIdx).sStnName
        'frmSelRemap!lbcSelRemap(iIndex).ItemData(frmSelRemap!lbcSelRemap(iIndex).NewIndex) = tgAttInfo(ilIdx).lAttCode
        frm!lbcSelRemap(iIndex).AddItem tgAttInfo(ilIdx).sStnName
        frm!lbcSelRemap(iIndex).ItemData(frm!lbcSelRemap(iIndex).NewIndex) = tgAttInfo(ilIdx).lAttCode
    Next ilIdx
    gPopSelRemap = True
    Exit Function
    
ErrHand:
    gHandleError "AffErrorLog.txt", "modAgmnt-gPopSelRemap"
    Screen.MousePointer = vbDefault
End Function



'Public tgCPDat() As DAT
Public Sub gCleanUpAtt()
    Dim llAtt As Long
    Dim SQLRequest As String
    Dim cleanuprst As ADODB.Recordset
    
    
    On Error GoTo ErrHand
    ReDim llAttCode(0 To 0) As Long
    SQLRequest = "SELECT *"
    SQLRequest = SQLRequest + " FROM att"
    SQLRequest = SQLRequest + " WHERE (attOffAir <  attOnAir)"
    Set cleanuprst = gSQLSelectCall(SQLRequest)
    While Not cleanuprst.EOF
        llAttCode(UBound(llAttCode)) = cleanuprst!attCode
        ReDim Preserve llAttCode(0 To UBound(llAttCode) + 1) As Long
        cleanuprst.MoveNext
    Wend
    For llAtt = 0 To UBound(llAttCode) - 1 Step 1
        DoEvents
        SQLRequest = "SELECT *"
        SQLRequest = SQLRequest + " FROM cptt"
        SQLRequest = SQLRequest + " WHERE (cpttAtfCode = " & llAttCode(llAtt) & ")"
        Set cleanuprst = gSQLSelectCall(SQLRequest)
        If cleanuprst.EOF Then
            ' JD 12-18-2006 Added new function to properly remove an agreement.
            If Not gDeleteAgreement(llAttCode(llAtt), "AffAgreementLog.Txt") Then
                gLogMsg "FAIL: gCleanUpAtt - Unable to delete att code " & llAttCode(llAtt), "AffErrorLog.Txt", False
            End If
'            cnn.BeginTrans
'            SQLRequest = "DELETE FROM dat WHERE (datAtfCode = " & llAttCode(llAtt) & ")"
'            cnn.Execute SQLRequest, rdExecDirect
'            SQLRequest = "DELETE FROM Att WHERE (AttCode = " & llAttCode(llAtt) & ")"
'            cnn.Execute SQLRequest, rdExecDirect
'            cnn.CommitTrans
        End If
    Next llAtt
    Erase llAttCode
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modAgmnt-CleanUpAtt"
    Exit Sub

End Sub



Public Sub gGetAvails(lAtfCode As Long, iShfCode As Integer, iInVefCode As Integer, iInVefCombo As Integer, sDate As String, ilAdjForZone As Integer)
    Dim sMinDate As String
    Dim sMaxDate As String
    Dim sSDate As String
    Dim sEDate As String
    Dim iWkDay As Integer
    Dim iPass As Integer
    Dim iVefCode As Integer
    Dim iVefCombo As Integer
    Dim ilVef As Integer
    Dim ilSetValue As Integer
    Dim ilVff As Integer
    Dim sAgrMoDate As String
    Dim sAgrSuDate As String
    
    '7/27/14:
    Dim ilShtt As Integer
    Dim slZone As String
    Dim ilRet As Integer
    Dim iTimeAdj As Integer
    Dim slFed As String
    
    Dim lefrst As ADODB.Recordset
    On Error GoTo ErrHand
    
    '5/24/16: Add gathering of program times
    ReDim tgPrgTimes(0 To 0) As PRGTIMES
    
    'sSDate = Format$(txtOnAirDate.Text, "mm/dd/yyyy")
    'sEDate = Format$(DateValue(sSDate) + 6, "mm/dd/yyyy")
    'SQLQuery = "SELECT * "
    'SQLQuery = SQLQuery + " FROM SSF_Spot_Summary SSF"
    'SQLQuery = SQLQuery + " WHERE (ssf.ssfType= 'O'"
    'SQLQuery = SQLQuery & " AND ssf.ssfVefCode = " & imVefCode
    'SQLQuery = SQLQuery + " AND ssf.ssfDate BETWEEN '" & sSDate & "' And '" & sEDate & "')"
    'Set rst = gSQLSelectCall(SQLQuery)
    'While Not rst.EOF
    '    iIndex = 2
    '    sStr = Mid$(rst!ssfPAS, iIndex, 16)
    '    rst.MoveNext
    'Wend
    'SQLQuery = "SELECT Min(lcfLogDate) FROM LCF_Log_Calendar lcf WHERE lcfVefCode = " & iVefCode & " AND lcfLogDate > '1/1/1970'"
    iVefCode = iInVefCode
    iVefCombo = iInVefCombo
    ReDim ilVehArray(0 To 0) As Integer
    ilVef = gBinarySearchVef(CLng(iVefCode))
    If ilVef <> -1 Then
        If tgVehicleInfo(ilVef).sVehType = "L" Then
            For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) Step 1
                If tgVehicleInfo(ilVef).iVefCode = iVefCode Then
                    ilSetValue = False
                    For ilVff = 0 To UBound(tgVffInfo) - 1 Step 1
                        If tgVehicleInfo(ilVef).iCode = tgVffInfo(ilVff).iVefCode Then
                            If tgVffInfo(ilVff).sMergeAffiliate <> "S" Then
                                ilSetValue = True
                                Exit For
                            End If
                        End If
                    Next ilVff
                    If ilSetValue = True Then
                        ilVehArray(UBound(ilVehArray)) = tgVehicleInfo(ilVef).iCode
                        ReDim Preserve ilVehArray(0 To (UBound(ilVehArray) + 1))
                    End If
                End If
            Next ilVef
        Else
            If iVefCombo > 0 Then
                ReDim ilVehArray(0 To 2) As Integer
                ilVehArray(0) = iInVefCode
                ilVehArray(1) = iVefCombo
            Else
                ReDim ilVehArray(0 To 1) As Integer
                ilVehArray(0) = iInVefCode
            End If
        End If
    Else
        If iVefCombo > 0 Then
            ReDim ilVehArray(0 To 2) As Integer
            ilVehArray(0) = iInVefCode
            ilVehArray(1) = iVefCombo
        Else
            ReDim ilVehArray(0 To 1) As Integer
            ilVehArray(0) = iInVefCode
        End If
    End If
    For iPass = 0 To UBound(ilVehArray) - 1 Step 1
        iVefCode = ilVehArray(iPass)
        
        bmBypassZeroUnits = False
        SQLQuery = "SELECT vffHonorZeroUnits, vefType "
        SQLQuery = SQLQuery & "From vff_Vehicle_Features Left Outer Join vef_Vehicles On vffVefCode = vefCode "
        SQLQuery = SQLQuery & "Where vffVefCode = " & iVefCode
        Set rst = gSQLSelectCall(SQLQuery)
        If Not rst.EOF Then
            If (rst!vefType = "A") And (rst!vffHonorZeroUnits = "Y") Then
                bmBypassZeroUnits = True
            End If
        End If
        
        SQLQuery = "SELECT Min(lcfLogDate) "
        SQLQuery = SQLQuery + " FROM LCF_Log_Calendar"
        SQLQuery = SQLQuery & " WHERE lcfVefCode = " & iVefCode & " AND lcfLogDate > " & "'" & Format$("1/1/1970", sgSQLDateForm) & "'"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst.EOF Then
            Exit Sub
        End If
        If IsNull(rst(0).Value) Then
            Exit Sub
        End If
        sMinDate = rst(0).Value
        'SQLQuery = "SELECT Max(lcfLogDate) FROM LCF_Log_Calendar lcf WHERE lcfVefCode = " & iVefCode & " AND lcfLogDate > '1/1/1970'" & " AND lcfType = 'O' AND lcfStatus = 'C'"
        SQLQuery = "SELECT Max(lcfLogDate) "
        SQLQuery = SQLQuery + " FROM LCF_Log_Calendar"
        '6/12/06-  changed lcfType from 'O' to 0 (zero).  this was changed to handle games in traffic.  User will be unable to post games using avails.
        'SQLQuery = SQLQuery & " WHERE lcfVefCode = " & iVefCode & " AND lcfLogDate > " & "'" & Format$("1/1/1970", sgSQLDateForm) & "'" & " AND lcfType = 'O' AND lcfStatus = 'C'"
        SQLQuery = SQLQuery & " WHERE lcfVefCode = " & iVefCode & " AND lcfLogDate > " & "'" & Format$("1/1/1970", sgSQLDateForm) & "'" & " AND lcfType = 0 AND lcfStatus = 'C'"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst.EOF Then
            Exit Sub
        End If
        If IsNull(rst(0).Value) Then
            Exit Sub
        End If
        sMaxDate = rst(0).Value
        If DateValue(gAdjYear(sDate)) < DateValue(gAdjYear(sMinDate)) Then
            sSDate = Format$(gObtainNextMonday(sMinDate), sgShowDateForm)
        ElseIf DateValue(gAdjYear(sDate)) > DateValue(gAdjYear(sMaxDate)) Then
            sSDate = Format$(gObtainPrevMonday(sMaxDate), sgShowDateForm)
        Else
            sSDate = Format$(gObtainPrevMonday(sDate), sgShowDateForm)
        End If
        sEDate = Format$(DateValue(gAdjYear(sSDate)) + 6, sgShowDateForm)
        
        '7/27/14: Determine if Delivery links exist and if so populate array
        Set DlfInfo_rst = gInitDlfInfo()
        ilShtt = gBinarySearchStationInfoByCode(iShfCode)
        If ilShtt <> -1 Then
            slZone = tgStationInfoByCode(ilShtt).sZone
            iTimeAdj = gGetTimeAdj(iShfCode, iVefCode, slFed)
        Else
            slZone = ""
            slFed = "*"
        End If
        'lRet = gBuildDlfInfo(iVefCode, sSDate, sEDate, slZone)
        sAgrMoDate = gObtainPrevMonday(sDate)
        sAgrSuDate = Format$(DateValue(gAdjYear(sAgrMoDate)) + 6, sgShowDateForm)
        ilRet = gBuildDlfInfo(iVefCode, sAgrMoDate, sAgrSuDate, slZone)
        If slFed <> "*" Then
            ilRet = gBuildDlfInfo(iVefCode, sAgrMoDate, sAgrSuDate, slFed & "ST")
        End If
        
        SQLQuery = "SELECT * "
        'SQLQuery = SQLQuery + " FROM LCF_Log_Calendar lcf"
        SQLQuery = SQLQuery + " FROM LCF_Log_Calendar"
        '6/12/06-  changed lcfType from 'O' to 0 (zero).  this was changed to handle games in traffic.  User will be unable to post games using avails.
        'SQLQuery = SQLQuery + " WHERE (lcfType = 'O'"
        SQLQuery = SQLQuery + " WHERE (lcfType = 0"
        SQLQuery = SQLQuery & " AND lcfStatus = 'C'"
        SQLQuery = SQLQuery + " AND lcfLogDate >= '" & Format$(sSDate, sgSQLDateForm) & "' And lcfLogDate <= '" & Format$(sEDate, sgSQLDateForm) & "'"
        ''D.S. 10/27/05 Add if statement below
        'If iVefCombo = 0 Then
            SQLQuery = SQLQuery & " AND lcfVefCode = " & iVefCode & ")"
        'Else
        '    SQLQuery = SQLQuery & " AND (lcfVefCode = " & iVefCode & " Or lcfVefCode = " & iVefCombo & "))"
        'End If
        Set rst = gSQLSelectCall(SQLQuery)
        If rst.EOF Then
            'SQLQuery = "SELECT Min(lcfLogDate) FROM LCF_Log_Calendar lcf WHERE lcfVefCode = " & iVefCode & " AND lcfLogDate > '" & Format$(sDate, "mm/dd/yyyy") & "'"
            SQLQuery = "SELECT Min(lcfLogDate)"
            SQLQuery = SQLQuery + " FROM LCF_Log_Calendar"
            SQLQuery = SQLQuery & " WHERE lcfVefCode = " & iVefCode & " AND lcfLogDate > '" & Format$(sDate, sgSQLDateForm) & "'"
            Set rst = gSQLSelectCall(SQLQuery)
            If rst.EOF Then
                Exit Sub
            End If
            sMaxDate = rst(0).Value
            sSDate = Format$(gObtainPrevMonday(sMaxDate), sgShowDateForm)
            sEDate = Format$(DateValue(gAdjYear(sSDate)) + 6, sgShowDateForm)
            SQLQuery = "SELECT * "
            'SQLQuery = SQLQuery + " FROM LCF_Log_Calendar lcf"
            SQLQuery = SQLQuery + " FROM LCF_Log_Calendar"
            '6/12/06-  changed lcfType from 'O' to 0 (zero).  this was changed to handle games in traffic.  User will be unable to post games using avails.
            'SQLQuery = SQLQuery + " WHERE (lcfType = 'O'"
            SQLQuery = SQLQuery + " WHERE (lcfType = 0"
            SQLQuery = SQLQuery & " AND lcfStatus = 'C'"
            SQLQuery = SQLQuery & " AND lcfVefCode = " & iVefCode
            SQLQuery = SQLQuery + " AND lcfLogDate >= '" & Format$(sSDate, sgSQLDateForm) & "' And lcfLogDate <= '" & Format$(sEDate, sgSQLDateForm) & "')"
            Set rst = gSQLSelectCall(SQLQuery)
        End If
        While Not rst.EOF
            iWkDay = Weekday(Format$(DateValue(gAdjYear(rst!lcfLogDate)), "m/d/yyyy"))
            mGetEvents rst!lcfLvf1, rst!lcfTime1, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf2, rst!lcfTime2, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf3, rst!lcfTime3, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf4, rst!lcfTime4, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf5, rst!lcfTime5, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf6, rst!lcfTime6, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf7, rst!lcfTime7, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf8, rst!lcfTime8, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf9, rst!lcfTime9, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf10, rst!lcfTime10, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf11, rst!lcfTime11, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf12, rst!lcfTime12, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf13, rst!lcfTime13, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf14, rst!lcfTime14, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf15, rst!lcfTime15, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf16, rst!lcfTime16, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf17, rst!lcfTime17, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf18, rst!lcfTime18, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf19, rst!lcfTime19, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf20, rst!lcfTime20, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf21, rst!lcfTime21, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf22, rst!lcfTime22, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf23, rst!lcfTime23, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf24, rst!lcfTime24, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf25, rst!lcfTime25, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf26, rst!lcfTime26, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf27, rst!lcfTime27, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf28, rst!lcfTime28, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf29, rst!lcfTime29, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf30, rst!lcfTime30, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf31, rst!lcfTime31, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf32, rst!lcfTime32, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf33, rst!lcfTime33, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf34, rst!lcfTime34, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf35, rst!lcfTime35, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf36, rst!lcfTime36, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf37, rst!lcfTime37, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf38, rst!lcfTime38, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf39, rst!lcfTime39, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf40, rst!lcfTime40, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf41, rst!lcfTime41, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf42, rst!lcfTime42, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf43, rst!lcfTime43, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf44, rst!lcfTime44, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf45, rst!lcfTime45, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf46, rst!lcfTime46, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf47, rst!lcfTime47, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf48, rst!lcfTime48, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf49, rst!lcfTime49, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            mGetEvents rst!lcfLvf50, rst!lcfTime50, lAtfCode, iShfCode, iInVefCode, iWkDay, ilAdjForZone
            rst.MoveNext
        Wend
        'If iVefCombo <> 0 Then
        '    iVefCode = iVefCombo
        '    iVefCombo = 0
        'Else
        '    Exit For
        'End If
    Next iPass

    '7/27/14
    gCloseDlfInfo

    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modAgmnt-gGetAvails"
End Sub

'*******************************************************************************************************
'
'   mGetEvents - called by gGetAvails
'                takes the results obtained by gGetAvails one record at a time and
'                loads tgDat array with avails
'
'********************************************************************************************************
Private Sub mGetEvents(lLvfCode As Integer, sStartTime As String, lAtfCode As Long, iShfCode As Integer, iVefCode As Integer, iInDay As Integer, ilAdjForZone As Integer)
    Dim lStartTime As Long
    '5/24/16
    Dim lEndTime As Long
    Dim lOffTime As Long
    Dim lTime As Long
    Dim iUpper As Integer
    Dim iFound As Integer
    Dim iLoop As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim sTime As String
    Dim iTimeAdj As Integer
    Dim iDay As Integer
    Dim slPledgeType As String
    Dim lefrst As ADODB.Recordset
    Dim rstDat As ADODB.Recordset
    Dim attrst As ADODB.Recordset
    Dim ilAnf As Integer
    Dim blAvailOk As Boolean
    Dim llDlfTime As Long
    Dim llFdTime As Long
    '11/26/14: Set before if using delivery links
    Dim ilSvDay As Integer
    Dim slPdDayFed As String
    
    '12/12/14: Match time not adjusted (Air Time from Delivery Links)
    Dim llSearchTime As Long
    Dim slFed As String
    Dim slZone As String
    Dim ilShtt As Integer
    
    '5/24/16: Add getting Program times and building array
    Dim ilPrgDay As Integer
    Dim ilPrg As Integer
    Dim blFound As Boolean
    
    Dim slLength As String
    Dim llLength As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    iDay = iInDay
    iTimeAdj = 0
    ''8/27/14: Ignoe time adjustments if delivery links defined as they contain the adjustments
    
    'If (ilAdjForZone) And (Not bgDlfExist) Then
    '12/12/14: Adjust by zone time to obtain the Pledge Feed time when using Delivery links
    If (ilAdjForZone) Then 'And (Not bgDlfExist) Then
    
        'If (iShfCode > 0) And (iVefCode > 0) Then
        '    For iLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
        '        If tgStationInfo(iLoop).iCode = iShfCode Then
        '            For iVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
        '                If tgVehicleInfo(iVef).iCode = iVefCode Then
        '                    For iZone = LBound(tgVehicleInfo(iVef).sZone) To UBound(tgVehicleInfo(iVef).sZone) Step 1
        '                        If StrComp(tgStationInfo(iLoop).sZone, tgVehicleInfo(iVef).sZone(iZone), 1) = 0 Then
        '                            iTimeAdj = tgVehicleInfo(iVef).iVehLocalAdj(iZone)
        '                            Exit For
        '                        End If
        '                    Next iZone
        '                    Exit For
        '                End If
        '            Next iVef
        '            Exit For
        '        End If
        '    Next iLoop
        'End If
        iTimeAdj = gGetTimeAdj(iShfCode, iVefCode, slFed)
    End If
    SQLQuery = "SELECT attPledgeType FROM att WHERE attCode = " & lAtfCode
    Set attrst = gSQLSelectCall(SQLQuery)
    If Not attrst.EOF Then
        slPledgeType = attrst!attPledgeType
    Else
        slPledgeType = ""
    End If
    If lLvfCode > 0 Then
        lStartTime = gTimeToLong(Format(sStartTime, "h:mm:ssAM/PM"), False)
        
        '5/24/16: Dtermine end time of program
        On Error GoTo mGetEventTimeErr
        llLength = 0
        SQLQuery = "SELECT lvfLen "
        SQLQuery = SQLQuery + " FROM LVF_Library_Version"
        SQLQuery = SQLQuery & " WHERE lvfCode = " & lLvfCode
        Set rst_lvf = gSQLSelectCall(SQLQuery)
        If Not rst_lvf.EOF Then
            ilRet = 0
            slLength = Format$(rst_lvf!lvfLen, "h:mm:ssAM/PM")
            If ilRet = 0 Then
                llLength = gTimeToLong(slLength, True)
            End If
        End If
        On Error GoTo ErrHand
        
        SQLQuery = "SELECT * "
        SQLQuery = SQLQuery + " FROM LEF_Library_Events"
        SQLQuery = SQLQuery + " WHERE (lefLvfCode = " & lLvfCode & ")"
        SQLQuery = SQLQuery & " ORDER BY lefStartTime, lefSeqNo"
        Set lefrst = gSQLSelectCall(SQLQuery)
        '5/24/16: Build Program time array
        If Not lefrst.EOF Then
            On Error GoTo mGetEventTimeErr
            
            lEndTime = lStartTime + llLength
            
            On Error GoTo ErrHand
            Select Case iInDay
                Case vbMonday
                    ilPrgDay = 0
                Case vbTuesday
                    ilPrgDay = 1
                Case vbWednesday
                    ilPrgDay = 2
                Case vbThursday
                    ilPrgDay = 3
                Case vbFriday
                    ilPrgDay = 4
                Case vbSaturday
                    ilPrgDay = 5
                Case vbSunday
                    ilPrgDay = 6
            End Select
            blFound = False
            For ilPrg = 0 To UBound(tgPrgTimes) - 1 Step 1
                If (tgPrgTimes(ilPrg).iDay = ilPrgDay) And (tgPrgTimes(ilPrg).lPrgStartTime = lStartTime) And (tgPrgTimes(ilPrg).lPrgEndTime = lEndTime) Then
                    blFound = True
                    Exit For
                End If
            Next ilPrg
            If Not blFound Then
                ilPrg = UBound(tgPrgTimes)
                tgPrgTimes(ilPrg).iDay = ilPrgDay
                tgPrgTimes(ilPrg).lPrgStartTime = lStartTime
                tgPrgTimes(ilPrg).lPrgEndTime = lEndTime
                ReDim Preserve tgPrgTimes(0 To ilPrg + 1) As PRGTIMES
            End If
        End If
        While Not lefrst.EOF
            If (lefrst!lefetfCode >= 2) And (lefrst!lefetfCode <= 9) Then
                blAvailOk = True
                ilAnf = gBinarySearchAnf(lefrst!lefAnfCode)
                If ilAnf <> -1 Then
                    If tgAvailNamesInfo(ilAnf).sTrafToAff = "N" Then
                        blAvailOk = False
                    End If
                End If
                
                If (blAvailOk) And (bmBypassZeroUnits) Then
                    If (lefrst!lefMaxUnits <= 0) Then
                        blAvailOk = False
                    Else
                        If gTimeToLong(lefrst!lefLen, False) <= 0 Then
                            blAvailOk = False
                        End If
                    End If
                End If
                                
                If blAvailOk Then
                    iDay = iInDay
                    lOffTime = gTimeToLong(lefrst!lefStartTime, False)
                    iFound = False
                    iUpper = UBound(tgDat)
                    lTime = lStartTime + lOffTime
                    
                    If bgDlfExist Then
                        '12/12/14: Save time to be used in the Find Delivery (Air Time)
                        llSearchTime = lTime
                        
                        If slFed <> "*" Then
                            lTime = gFindDlf(iDay, llSearchTime, lefrst!lefetfCode, lefrst!lefenfCode, slFed)
                        End If
                    End If
                    
                    lTime = lTime + 3600 * iTimeAdj
                    If lTime < 0 Then
                        lTime = lTime + 86400
                        If iDay = vbSunday Then
                            iDay = vbSaturday
                        Else
                            iDay = iDay - 1
                        End If
                    ElseIf lTime > 86400 Then
                        lTime = lTime - 86400
                        If iDay = vbSaturday Then
                            iDay = vbSunday
                        Else
                            iDay = iDay + 1
                        End If
                    End If
    
                    '11/26/14
                    ilSvDay = iDay
                    
                    '12/12/14: Use unadjusted time for serach
                    'llDlfTime = gFindDlf(iDay, lTime, lefrst!lefetfCode, lefrst!lefenfCode)
                    If bgDlfExist Then
                        ilShtt = gBinarySearchStationInfoByCode(iShfCode)
                        If ilShtt <> -1 Then
                            slZone = tgStationInfoByCode(ilShtt).sZone
                        Else
                            slZone = ""
                        End If
                        llDlfTime = gFindDlf(iDay, llSearchTime, lefrst!lefetfCode, lefrst!lefenfCode, slZone)
                        If slFed = "*" Then
                            lTime = llDlfTime
                        End If
                    Else
                        llDlfTime = lTime
                    End If
                    
                    If llDlfTime <> -1 Then
                        '11/26/14: Set Pledge Before/After to before if days different
                        slPdDayFed = ""
                        If (bgDlfExist) And (ilSvDay <> iDay) Then
                            If (iDay < ilSvDay) Or ((iDay = vbSaturday) And (ilSvDay = vbSunday)) Then
                                slPdDayFed = "B"
                            End If
                        End If
                        '11/21/14: Retain Feed time so that it can be compared to Delivery affiliate time (pledge time)
                        llFdTime = lTime
                        lTime = llDlfTime
                    
                        For iLoop = LBound(tgDat) To UBound(tgDat) - 1 Step 1
                            '11/26/14: Don't combine if Before
                            If slPdDayFed = "B" Then
                                Exit For
                            End If
                            '12/2/14: Bypass Before days
                            If (tgDat(iLoop).sPdDayFed <> "B") Then
                                'If (lTime = gTimeToLong(tgDat(iLoop).sFdSTime, False)) And (tgDat(iLoop).iDACode = 1) Then
                                If (llFdTime = gTimeToLong(tgDat(iLoop).sFdSTime, False)) Then
                                    '6/9/14: Add test of feed end time
                                    If gTimeToLong(tgDat(iLoop).sFdETime, True) = llFdTime + gTimeToLong(lefrst!lefLen, False) Then
                                        '11/21/14: Add Delivery additional time test
                                        If (lTime = gTimeToLong(tgDat(iLoop).sPdSTime, False)) Then
                                            iFound = True
                                            iUpper = iLoop
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        Next iLoop
                        tgDat(iUpper).iStatus = 1
                        tgDat(iUpper).lCode = 0
                        tgDat(iUpper).lAtfCode = lAtfCode
                        tgDat(iUpper).iShfCode = iShfCode
                        tgDat(iUpper).iVefCode = iVefCode
                        'tgDat(iUpper).iDACode = 1
                        
                        '11/26/14: if Before set Pdedge one day prior to Feed
                        If slPdDayFed <> "B" Then
                            Select Case iDay
                                Case vbMonday
                                    tgDat(iUpper).iFdDay(0) = 1
                                    tgDat(iUpper).iPdDay(0) = 1
                                Case vbTuesday
                                    tgDat(iUpper).iFdDay(1) = 1
                                    tgDat(iUpper).iPdDay(1) = 1
                                Case vbWednesday
                                    tgDat(iUpper).iFdDay(2) = 1
                                    tgDat(iUpper).iPdDay(2) = 1
                                Case vbThursday
                                    tgDat(iUpper).iFdDay(3) = 1
                                    tgDat(iUpper).iPdDay(3) = 1
                                Case vbFriday
                                    tgDat(iUpper).iFdDay(4) = 1
                                    tgDat(iUpper).iPdDay(4) = 1
                                Case vbSaturday
                                    tgDat(iUpper).iFdDay(5) = 1
                                    tgDat(iUpper).iPdDay(5) = 1
                                Case vbSunday
                                    tgDat(iUpper).iFdDay(6) = 1
                                    tgDat(iUpper).iPdDay(6) = 1
                            End Select
                        Else
                            '12/2/14: Use from day
                            'Select Case iDay
                            Select Case ilSvDay
                                Case vbMonday
                                    tgDat(iUpper).iFdDay(0) = 1
                                    tgDat(iUpper).iPdDay(6) = 1
                                Case vbTuesday
                                    tgDat(iUpper).iFdDay(1) = 1
                                    tgDat(iUpper).iPdDay(0) = 1
                                Case vbWednesday
                                    tgDat(iUpper).iFdDay(2) = 1
                                    tgDat(iUpper).iPdDay(1) = 1
                                Case vbThursday
                                    tgDat(iUpper).iFdDay(3) = 1
                                    tgDat(iUpper).iPdDay(2) = 1
                                Case vbFriday
                                    tgDat(iUpper).iFdDay(4) = 1
                                    tgDat(iUpper).iPdDay(3) = 1
                                Case vbSaturday
                                    tgDat(iUpper).iFdDay(5) = 1
                                    tgDat(iUpper).iPdDay(4) = 1
                                Case vbSunday
                                    tgDat(iUpper).iFdDay(6) = 1
                                    tgDat(iUpper).iPdDay(5) = 1
                            End Select
                        End If
                        '11/26/14: Set pdDayFed
                        'tgDat(iUpper).sPdDayFed = ""
                        tgDat(iUpper).sPdDayFed = slPdDayFed
                        '11/21/14: Determine if Feed Time and Pledge time match.  Might be difference if using Delivery links
                        If (llFdTime = lTime) Or (Not bgDlfExist) Then
                            sTime = Format$(gLongToTime(lTime), sgShowTimeWSecForm)
                            If Second(sTime) = 0 Then
                                sTime = Format$(gLongToTime(lTime), sgShowTimeWOSecForm)
                            End If
                            tgDat(iUpper).sFdSTime = sTime
                            sTime = Format$(gLongToTime(lTime + gTimeToLong(lefrst!lefLen, False)), sgShowTimeWSecForm)
                            If Second(sTime) = 0 Then
                                sTime = Format$(sTime, sgShowTimeWOSecForm)
                            End If
                            tgDat(iUpper).sFdETime = sTime
                            tgDat(iUpper).iFdStatus = 0    '(14).Value
                            tgDat(iUpper).sPdSTime = tgDat(iUpper).sFdSTime
                            tgDat(iUpper).sPdETime = tgDat(iUpper).sFdETime
                        Else
                            '11/21/14: Set as delay
                            sTime = Format$(gLongToTime(llFdTime), sgShowTimeWSecForm)
                            If Second(sTime) = 0 Then
                                sTime = Format$(gLongToTime(llFdTime), sgShowTimeWOSecForm)
                            End If
                            tgDat(iUpper).sFdSTime = sTime
                            sTime = Format$(gLongToTime(llFdTime + gTimeToLong(lefrst!lefLen, False)), sgShowTimeWSecForm)
                            If Second(sTime) = 0 Then
                                sTime = Format$(sTime, sgShowTimeWOSecForm)
                            End If
                            tgDat(iUpper).sFdETime = sTime
                            tgDat(iUpper).iFdStatus = 1    '(14).Value
                            sTime = Format$(gLongToTime(lTime), sgShowTimeWSecForm)
                            If Second(sTime) = 0 Then
                                sTime = Format$(gLongToTime(lTime), sgShowTimeWOSecForm)
                            End If
                            tgDat(iUpper).sPdSTime = sTime
                            sTime = Format$(gLongToTime(lTime + gTimeToLong(lefrst!lefLen, False)), sgShowTimeWSecForm)
                            If Second(sTime) = 0 Then
                                sTime = Format$(sTime, sgShowTimeWOSecForm)
                            End If
                            tgDat(iUpper).sPdETime = sTime
                        End If
                        tgDat(iUpper).iAirPlayNo = tgDat(iUpper).iAirPlayNo
                        tgDat(iUpper).sEstimatedTime = tgDat(iUpper).sEstimatedTime
                        '7/15/14
                        tgDat(iUpper).sEmbeddedOrROS = "R"
                        If Not iFound Then
                            iUpper = iUpper + 1
                            ReDim Preserve tgDat(0 To iUpper) As DAT
                            For iDay = 0 To 6 Step 1
                                tgDat(iUpper).iFdDay(iDay) = 0
                                tgDat(iUpper).iPdDay(iDay) = 0
                            Next iDay
                        End If
                        
                    End If
                    
                End If
            End If
            lefrst.MoveNext
        Wend
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modAgmnt-gGetEvents"
mGetEventTimeErr:
    If lStartTime = 0 Then
        llLength = 86400
        Resume Next
    End If
    Exit Sub
End Sub

Public Function gGetLatestRatecard() As Integer
'************************************************************************************************
'*    Get the code, name, year and startdate out of rcf. Rules to find the latest ratecard:
'*    1. Look at the name. If the name starts with "~" then exclude that record.
'*    2. Look at the year. If there is one record with the latest year then take that one.
'*    3. If there is more than one record with the same latest year then take
'*       the one with the latest startdate.
'************************************************************************************************
    
    Dim rst_Ratecard As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    ReDim tgLatestRateCard(0 To 0) As LATESTRATECARD
    
    'Let SQL call order by Year then StartDate both in decsending order
    SQLQuery = "SELECT rcfCode, rcfName, rcfYear, rcfStartDate "
    SQLQuery = SQLQuery & " FROM RCF_Rate_Card "
    SQLQuery = SQLQuery + " ORDER BY rcfYear DESC, rcfStartDate DESC"
    
    Set rst_Ratecard = gSQLSelectCall(SQLQuery)
    If rst_Ratecard.EOF Then
        gGetLatestRatecard = False
        Exit Function
    End If
    
    While Not rst_Ratecard.EOF
        If Left(rst_Ratecard!rcfName, 1) <> "~" Then
            tgLatestRateCard(0).iLatestRCFCode = rst_Ratecard!rcfCode
            gGetLatestRatecard = True
            Exit Function
        End If
        rst_Ratecard.MoveNext
    Wend
       
    gGetLatestRatecard = False
    Exit Function
    
ErrHand:
    gHandleError "AffErrorLog.txt", "modAgmnt-gGetLatestRatecard"
End Function
Public Function gDetermineAgreementTimes(ilInShttCode As Integer, ilInVefCode As Integer, slOnAir As String, slOffAir As String, slDropDate As String, slCDStartTime As String, slOutStartTime As String, slOutEndTime As String)
    Dim ilVefCode As Integer
    Dim ilShttCode As Integer
    Dim ilVef As Integer
    Dim slDate As String
    Dim slMinDate As String
    Dim slMaxDate As String
    Dim slSDate As String
    Dim slEDate As String
    Dim slZone As String
    Dim ilPass As Integer
    Dim ilLoop As Integer
    Dim ilZone As Integer
    Dim llTimeAdj As Long
    Dim llAllStartTime As Long
    Dim llAllEndTime As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilTimeAdj As Long
    Dim llCDStartTime As Long
    Dim llTimeOffset As Long
    
    On Error GoTo ErrHand
    gDetermineAgreementTimes = False
    slOutStartTime = "12AM"
    slOutEndTime = "12AM"

    'SQLQuery = "SELECT * FROM ATT WHERE attCode = " & llInAttCode
    'Set rst_att = gSQLSelectCall(SQLQuery)
    'If rst_att.EOF Then
    '    Exit Function
    'End If
    ilVefCode = ilInVefCode 'rst_att!attvefCode
    ilShttCode = ilInShttCode   'rst_att!attshfCode
    slDate = Trim$(slOnAir)    'rst_att!attOnAir
    If slDate = "" Then
        gDetermineAgreementTimes = True
        Exit Function
    End If
    'If gDateValue(gAdjYear(Trim$(rst_att!attOffAir))) <= gDateValue(gAdjYear(Trim$(rst_att!attDropDate))) Then
    '    If gDateValue(gAdjYear(Trim$(rst_att!attOffAir))) < gDateValue("12/31/2069") Then
    '        slDate = Trim$(rst_att!attOffAir)
    If (Trim$(slOffAir) <> "") And (Trim$(slDropDate) <> "") Then
        If gDateValue(gAdjYear(Trim$(slOffAir))) <= gDateValue(gAdjYear(Trim$(slDropDate))) Then
            If gDateValue(gAdjYear(Trim$(slOffAir))) < gDateValue("12/31/2069") Then
                slDate = slOffAir   'Trim$(rst_att!attOffAir)
            End If
        Else
            slDate = slDropDate 'Trim$(rst_att!attDropDate)
        End If
    ElseIf (Trim$(slOffAir) <> "") Then
        If gDateValue(gAdjYear(Trim$(slOffAir))) < gDateValue("12/31/2069") Then
            slDate = slOffAir   'Trim$(rst_att!attOffAir)
        End If
    ElseIf (Trim$(slDropDate) <> "") Then
        If gDateValue(gAdjYear(Trim$(slDropDate))) < gDateValue("12/31/2069") Then
            slDate = slDropDate   'Trim$(rst_att!attOffAir)
        End If
    End If
    'llCDStartTime = gTimeToLong(rst_att!attStartTime, False)
    If slCDStartTime <> "" Then
        llCDStartTime = gTimeToLong(slCDStartTime, False)
    Else
        llCDStartTime = -1
    End If
    slZone = ""
    For ilLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(ilLoop).iCode = ilShttCode Then
            slZone = tgStationInfo(ilLoop).sZone
            Exit For
        End If
    Next ilLoop
    ReDim ilVehArray(0 To 0) As Integer
    ilVef = gBinarySearchVef(CLng(ilVefCode))
    If ilVef <> -1 Then
        If tgVehicleInfo(ilVef).sVehType = "L" Then
            For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) Step 1
                If tgVehicleInfo(ilVef).iVefCode = ilVefCode Then
                    ilVehArray(UBound(ilVehArray)) = tgVehicleInfo(ilVef).iCode
                    ReDim Preserve ilVehArray(0 To (UBound(ilVehArray) + 1))
                End If
            Next ilVef
        Else
            ReDim ilVehArray(0 To 1) As Integer
            ilVehArray(0) = ilVefCode
        End If
    Else
        ReDim ilVehArray(0 To 1) As Integer
        ilVehArray(0) = ilVefCode
    End If
    llAllStartTime = 86401
    llAllEndTime = -1
    For ilPass = 0 To UBound(ilVehArray) - 1 Step 1
        llStartTime = -1
        llEndTime = llStartTime
        ilVefCode = ilVehArray(ilPass)
        llTimeAdj = 0
        For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
            If tgVehicleInfo(ilVef).iCode = ilVefCode Then
                For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                    If StrComp(slZone, tgVehicleInfo(ilVef).sZone(ilZone), 1) = 0 Then
                        llTimeAdj = CLng(3600) * tgVehicleInfo(ilVef).iVehLocalAdj(ilZone)
                        Exit For
                    End If
                Next ilZone
                Exit For
            End If
        Next ilVef
        SQLQuery = "SELECT Min(lcfLogDate) "
        SQLQuery = SQLQuery + " FROM LCF_Log_Calendar"
        SQLQuery = SQLQuery & " WHERE lcfVefCode = " & ilVefCode & " AND lcfLogDate > " & "'" & Format$("1/1/1970", sgSQLDateForm) & "'"
        Set rst_lcf = gSQLSelectCall(SQLQuery)
        If rst_lcf.EOF Then
            Exit Function
        End If
        If IsNull(rst_lcf(0).Value) Then
            Exit Function
        End If
        slMinDate = rst_lcf(0).Value
        SQLQuery = "SELECT Max(lcfLogDate) "
        SQLQuery = SQLQuery + " FROM LCF_Log_Calendar"
        SQLQuery = SQLQuery & " WHERE lcfVefCode = " & ilVefCode & " AND lcfLogDate > " & "'" & Format$("1/1/1970", sgSQLDateForm) & "'" & " AND lcfType = 0 AND lcfStatus = 'C'"
        Set rst_lcf = gSQLSelectCall(SQLQuery)
        If rst_lcf.EOF Then
            Exit Function
        End If
        If IsNull(rst_lcf(0).Value) Then
            Exit Function
        End If
        slMaxDate = rst_lcf(0).Value
        If gDateValue(gAdjYear(slDate)) <> gDateValue("1/1/1970") Then
            If gDateValue(gAdjYear(slDate)) < gDateValue(gAdjYear(slMinDate)) Then
                slSDate = Format$(gObtainNextMonday(slMinDate), sgShowDateForm)
            ElseIf gDateValue(gAdjYear(slDate)) > gDateValue(gAdjYear(slMaxDate)) Then
                slSDate = Format$(gObtainPrevMonday(slMaxDate), sgShowDateForm)
            Else
                slSDate = Format$(gObtainPrevMonday(slDate), sgShowDateForm)
            End If
        Else
            slSDate = Format$(gObtainNextMonday(slMinDate), sgShowDateForm)
        End If
        slEDate = Format$(gDateValue(gAdjYear(slSDate)) + 6, sgShowDateForm)
        SQLQuery = "SELECT * "
        SQLQuery = SQLQuery + " FROM LCF_Log_Calendar"
        SQLQuery = SQLQuery + " WHERE (lcfType = 0"
        SQLQuery = SQLQuery & " AND lcfStatus = 'C'"
        SQLQuery = SQLQuery + " AND lcfLogDate >= '" & Format$(slSDate, sgSQLDateForm) & "' And lcfLogDate <= '" & Format$(slEDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND lcfVefCode = " & ilVefCode & ")"
        Set rst_lcf = gSQLSelectCall(SQLQuery)
        If rst_lcf.EOF Then
            SQLQuery = "SELECT Min(lcfLogDate)"
            SQLQuery = SQLQuery + " FROM LCF_Log_Calendar"
            SQLQuery = SQLQuery & " WHERE lcfVefCode = " & ilVefCode & " AND lcfLogDate > '" & Format$(slDate, sgSQLDateForm) & "'"
            Set rst_lcf = gSQLSelectCall(SQLQuery)
            If rst_lcf.EOF Then
                Exit Function
            End If
            slMaxDate = rst_lcf(0).Value
            slSDate = Format$(gObtainPrevMonday(slMaxDate), sgShowDateForm)
            slEDate = Format$(DateValue(gAdjYear(slSDate)) + 6, sgShowDateForm)
            SQLQuery = "SELECT * "
            SQLQuery = SQLQuery + " FROM LCF_Log_Calendar"
            SQLQuery = SQLQuery + " WHERE (lcfType = 0"
            SQLQuery = SQLQuery & " AND lcfStatus = 'C'"
            SQLQuery = SQLQuery & " AND lcfVefCode = " & ilVefCode
            SQLQuery = SQLQuery + " AND lcfLogDate >= '" & Format$(slSDate, sgSQLDateForm) & "' And lcfLogDate <= '" & Format$(slEDate, sgSQLDateForm) & "')"
            Set rst_lcf = gSQLSelectCall(SQLQuery)
        End If
        Do While Not rst_lcf.EOF
            mGetEventTimes rst_lcf!lcfLvf1, Format$(rst_lcf!lcfTime1, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf2, Format$(rst_lcf!lcfTime2, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf3, Format$(rst_lcf!lcfTime3, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf4, Format$(rst_lcf!lcfTime4, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf5, Format$(rst_lcf!lcfTime5, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf6, Format$(rst_lcf!lcfTime6, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf7, Format$(rst_lcf!lcfTime7, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf8, Format$(rst_lcf!lcfTime8, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf9, Format$(rst_lcf!lcfTime9, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf10, Format$(rst_lcf!lcfTime10, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf11, Format$(rst_lcf!lcfTime11, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf12, Format$(rst_lcf!lcfTime12, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf13, Format$(rst_lcf!lcfTime13, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf14, Format$(rst_lcf!lcfTime14, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf15, Format$(rst_lcf!lcfTime15, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf16, Format$(rst_lcf!lcfTime16, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf17, Format$(rst_lcf!lcfTime17, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf18, Format$(rst_lcf!lcfTime18, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf19, Format$(rst_lcf!lcfTime19, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf20, Format$(rst_lcf!lcfTime20, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf21, Format$(rst_lcf!lcfTime21, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf22, Format$(rst_lcf!lcfTime22, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf23, Format$(rst_lcf!lcfTime23, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf24, Format$(rst_lcf!lcfTime24, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf25, Format$(rst_lcf!lcfTime25, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf26, Format$(rst_lcf!lcfTime26, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf27, Format$(rst_lcf!lcfTime27, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf28, Format$(rst_lcf!lcfTime28, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf29, Format$(rst_lcf!lcfTime29, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf30, Format$(rst_lcf!lcfTime30, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf31, Format$(rst_lcf!lcfTime31, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf32, Format$(rst_lcf!lcfTime32, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf33, Format$(rst_lcf!lcfTime33, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf34, Format$(rst_lcf!lcfTime34, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf35, Format$(rst_lcf!lcfTime35, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf36, Format$(rst_lcf!lcfTime36, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf37, Format$(rst_lcf!lcfTime37, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf38, Format$(rst_lcf!lcfTime38, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf39, Format$(rst_lcf!lcfTime39, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf40, Format$(rst_lcf!lcfTime40, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf41, Format$(rst_lcf!lcfTime41, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf42, Format$(rst_lcf!lcfTime42, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf43, Format$(rst_lcf!lcfTime43, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf44, Format$(rst_lcf!lcfTime44, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf45, Format$(rst_lcf!lcfTime45, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf46, Format$(rst_lcf!lcfTime46, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf47, Format$(rst_lcf!lcfTime47, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf48, Format$(rst_lcf!lcfTime48, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf49, Format$(rst_lcf!lcfTime49, "h:mm:ssAM/PM"), llStartTime, llEndTime
            mGetEventTimes rst_lcf!lcfLvf50, Format$(rst_lcf!lcfTime50, "h:mm:ssAM/PM"), llStartTime, llEndTime
            rst_lcf.MoveNext
        Loop
        If llStartTime <> -1 Then
            llStartTime = llStartTime + llTimeAdj
            If llStartTime < 0 Then
                llStartTime = 86400 + llStartTime
            End If
            llEndTime = llEndTime + llTimeAdj
            If llEndTime > 86400 Then
                llEndTime = llEndTime - 86400
            End If
            If llStartTime < llAllStartTime Then
                llAllStartTime = llStartTime
            End If
            If llEndTime > llAllEndTime Then
                llAllEndTime = llEndTime
            End If
        End If
    Next ilPass
    llTimeOffset = 0
    If (llAllStartTime <> 86401) And (llCDStartTime >= 0) Then
        If llAllStartTime > llCDStartTime Then
            llTimeOffset = ((llAllStartTime - llCDStartTime) * -1)
        Else
            'Pledged to air later then cd was supposed to
            llTimeOffset = llCDStartTime - llAllStartTime
        End If
    End If
    If llAllStartTime <> 86401 Then
        slOutStartTime = gLongToTime(llAllStartTime + llTimeOffset)
    End If
    If llAllEndTime >= 0 Then
        slOutEndTime = gLongToTime(llAllEndTime + llTimeOffset)
    End If
    gDetermineAgreementTimes = True
    Exit Function
    
ErrHand:
    gHandleError "AffErrorLog.txt", "modAgmnt-gDetermineAgreementTimes"
End Function

Private Sub mGetEventTimes(llLvfCode As Long, slLcfStartTime As String, llOutStartTime As Long, llOutEndTime As Long)
    Dim llStartTime As Long
    Dim llEndDate As Long
    Dim llLength As Long
    Dim slLength As String
    Dim ilRet As Integer
    
    On Error GoTo mGetEventTimeErr
    If llLvfCode <= 0 Then
        Exit Sub
    End If
    llStartTime = gTimeToLong(slLcfStartTime, False)
    SQLQuery = "SELECT lvfLen "
    SQLQuery = SQLQuery + " FROM LVF_Library_Version"
    SQLQuery = SQLQuery & " WHERE lvfCode = " & llLvfCode
    Set rst_lvf = gSQLSelectCall(SQLQuery)
    If Not rst_lvf.EOF Then
        ilRet = 0
        slLength = Format$(rst_lvf!lvfLen, "h:mm:ssAM/PM")
        If ilRet = 0 Then
            llLength = gTimeToLong(slLength, False)
        End If
        If llLength > 0 Then
            If (llOutStartTime = -1) And (llOutEndTime = -1) Then
                llOutStartTime = llStartTime
                llOutEndTime = llStartTime + llLength
            Else
                If llStartTime + llLength > llOutEndTime Then
                    llOutEndTime = llStartTime + llLength
                End If
                If llStartTime < llOutStartTime Then
                    llOutStartTime = llStartTime
                End If
            End If
        End If
    End If
    Exit Sub
mGetEventTimeErr:
    If llStartTime = 0 Then
        llLength = 86400
        ilRet = 1
        Resume Next
    End If
    Exit Sub
End Sub


Public Function gCompactTime(slInTime As String) As String
    gCompactTime = slInTime
    If Second(slInTime) = 0 Then
        If Minute(slInTime) = 0 Then
            gCompactTime = Format(slInTime, "ha/p")
        Else
            gCompactTime = Format(slInTime, "h:mma/p")
        End If
    End If

End Function

Public Function gGetTimeAdj(ilShttCode As Integer, ilVefCode As Integer, slFed As String, Optional slZone As String) As Integer
    Dim ilTimeAdj As Integer
    Dim ilVef As Integer
    Dim ilLoop As Integer
    Dim ilZone As Integer
    
    ilTimeAdj = 0
    slFed = ""
    slZone = ""
    If (ilShttCode > 0) And (ilVefCode > 0) Then
        For ilLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(ilLoop).iCode = ilShttCode Then
                For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                    If tgVehicleInfo(ilVef).iCode = ilVefCode Then
                        For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                            If StrComp(tgStationInfo(ilLoop).sZone, tgVehicleInfo(ilVef).sZone(ilZone), 1) = 0 Then
                                ilTimeAdj = tgVehicleInfo(ilVef).iVehLocalAdj(ilZone)
                                slFed = tgVehicleInfo(ilVef).sFed(ilZone)
                                slZone = tgVehicleInfo(ilVef).sZone(ilZone)
                                Exit For
                            End If
                        Next ilZone
                        Exit For
                    End If
                Next ilVef
                Exit For
            End If
        Next ilLoop
    End If
    gGetTimeAdj = ilTimeAdj
End Function


Public Sub gAdjustEventTime(ilTimeAdj As Integer, slDate As String, slTime As String)
    Dim llDate As Long
    Dim llTime As Long
    
    llDate = DateValue(gAdjYear(slDate))
    llTime = gTimeToLong(slTime, False)
    
    llTime = llTime + 3600 * ilTimeAdj
    If llTime < 0 Then
        llTime = llTime + 86400
        llDate = llDate - 1
    ElseIf llTime > 86400 Then
        llTime = llTime - 86400
        llDate = llDate + 1
    End If
    slTime = Format$(gLongToTime(llTime), sgShowTimeWSecForm)
    slDate = Format$(llDate, sgShowDateForm)

End Sub

Public Sub gSetStationSpotBuilder(slSource As String, ilVefCode As Integer, ilShttCode As Integer, llInGenStartDate As Long, llInGenEndDate As Long)
    'Source: P=Post Log; A=Agreement; F=Fast add
    Dim slLLD As String
    Dim slMoDate As String
    Dim llDate As Long
    Dim llAbfCode As Long
    Dim blInsert As Boolean
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llGenStartDate As Long
    Dim llGenEndDate As Long
    Dim llLLD As Long
    Dim llVpf As Long
    
    'Check if last log date is within Gneration dates
    On Error GoTo ErrHand
    llGenStartDate = llInGenStartDate
    llGenEndDate = llInGenEndDate
'    SQLQuery = "SELECT vpfLLD"
'    SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
'    SQLQuery = SQLQuery + " WHERE (vpfvefKCode =" & ilVefCode & ")"
'    Set rst_Vpf = gSQLSelectCall(SQLQuery)
'    If Not rst_Vpf.EOF Then
'        If IsNull(rst_Vpf!vpfLLD) Then
'            Exit Sub
'        Else
'            If Not gIsDate(rst_Vpf!vpfLLD) Then
'                Exit Sub
'            Else
'                'set sLLD to last log date
'                slLLD = Format$(rst_Vpf!vpfLLD, sgShowDateForm)
'                llLLD = gDateValue(slLLD)
'            End If
'        End If
'    Else
'        Exit Sub
'    End If
    
    'TTP 10051 - Fast Add grid slowness
    llVpf = gBinarySearchVpf(CLng(ilVefCode))
    If llVpf = -1 Then
        gPopVehicleOptions
        llVpf = gBinarySearchVpf(CLng(ilVefCode))
        If llVpf = -1 Then
            Exit Sub
        End If
    Else
        If Not gIsDate(Trim(tgVpfOptions(llVpf).sLLD)) Then
            Exit Sub
        Else
            slLLD = Trim(tgVpfOptions(llVpf).sLLD)
            llLLD = gDateValue(slLLD)
        End If
    End If
    
    If llGenStartDate > llLLD Then
        Exit Sub
    End If
    llDate = llGenStartDate
    slMoDate = Format(llDate, "m/d/yy")
    'Determine if generate record should be inserted or updated
    Do While Weekday(slMoDate, vbSunday) <> vbMonday
        slMoDate = DateAdd("d", -1, slMoDate)
    Loop
    llGenEndDate = gDateValue(DateAdd("d", 6, slMoDate))
    If llGenEndDate > llInGenEndDate Then
        llGenEndDate = llInGenEndDate
    End If
    If llGenEndDate > llLLD Then
        llGenEndDate = llLLD
    End If
    Do
        blInsert = False
        If llGenEndDate < llGenStartDate Then
            Exit Do
        End If
        SQLQuery = "Select abfGenStartDate, abfGenEndDate, Count(*) as Match From abf_Ast_Build_Queue"
        SQLQuery = SQLQuery & " Where abfStatus = 'G' And abfVefCode = " & ilVefCode
        SQLQuery = SQLQuery & " And abfShttCode = " & ilShttCode
        SQLQuery = SQLQuery & " And abfMondayDate = '" & Format(slMoDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " Group By abfGenStartDate, abfGenEndDate"
        Set rst_abf = gSQLSelectCall(SQLQuery)
        If Not rst_abf.EOF Then
            If rst_abf!Match <= 0 Then
                blInsert = True
            Else
                'Test dates
                llStartDate = gDateValue(Format(rst_abf!abfGenStartDate, "m/d/yy"))
                llEndDate = gDateValue(Format(rst_abf!abfGenEndDate, "m/d/yy"))
                If (llStartDate > llGenStartDate) Or (llEndDate < llGenEndDate) Then
                    If llGenStartDate < llStartDate Then
                        If llGenEndDate >= llStartDate Then
                            llAbfCode = mAbfInsert(slSource, ilVefCode, ilShttCode, slMoDate, llGenStartDate, llStartDate - 1)
                        Else
                            'llAbfCode = mAbfInsert(slSource, ilVefcode, ilShttCode, slMoDate, llGenStartDate, llGenEndDate)
                            blInsert = True
                        End If
                    End If
                    If llGenEndDate > llEndDate Then
                        If llGenStartDate <= llEndDate + 1 Then
                            llAbfCode = mAbfInsert(slSource, ilVefCode, ilShttCode, slMoDate, llEndDate + 1, llGenEndDate)
                        Else
                            'llAbfCode = mAbfInsert(slSource, ilVefcode, ilShttCode, slMoDate, llGenStartDate, llGenEndDate)
                            blInsert = True
                        End If
                    End If
                End If
            End If
        Else
            blInsert = True
        End If
        If blInsert Then
            llAbfCode = mAbfInsert(slSource, ilVefCode, ilShttCode, slMoDate, llGenStartDate, llGenEndDate)
        End If
        slMoDate = DateAdd("d", 7, slMoDate)
        llDate = gDateValue(slMoDate)
        llGenStartDate = llDate
        llGenEndDate = llDate + 6
        If llGenEndDate > llInGenEndDate Then
            llGenEndDate = llInGenEndDate
        End If
        If llGenEndDate > llLLD Then
            llGenEndDate = llLLD
        End If
    Loop While llDate <= llLLD
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "gSetStationSpotBuilder"
    'Resume Next
    Exit Sub
ErrHand1:
    gHandleError "AffErrorLog.txt", "gSetStationSpotBuilder"
    Return
End Sub


Private Function mAbfInsert(slSource As String, ilVefCode As Integer, ilShttCode As Integer, slMoDate As String, llGenStartDate As Long, llGenEndDate As Long) As Long
    Dim llAbfCode As Long

    On Error GoTo ErrHand
    SQLQuery = "Insert Into abf_Ast_Build_Queue ( "
    SQLQuery = SQLQuery & "abfCode, "
    SQLQuery = SQLQuery & "abfSource, "
    SQLQuery = SQLQuery & "abfVefCode, "
    SQLQuery = SQLQuery & "abfShttCode, "
    SQLQuery = SQLQuery & "abfStatus, "
    SQLQuery = SQLQuery & "abfMondayDate, "
    SQLQuery = SQLQuery & "abfGenStartDate, "
    SQLQuery = SQLQuery & "abfGenEndDate, "
    SQLQuery = SQLQuery & "abfEnteredDate, "
    SQLQuery = SQLQuery & "abfEnteredTime, "
    SQLQuery = SQLQuery & "abfCompletedDate, "
    SQLQuery = SQLQuery & "abfCompletedTime, "
    SQLQuery = SQLQuery & "abfUrfCode, "
    SQLQuery = SQLQuery & "abfUstCode, "
    SQLQuery = SQLQuery & "abfUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slSource) & "', "
    SQLQuery = SQLQuery & ilVefCode & ", "
    SQLQuery = SQLQuery & ilShttCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote("G") & "', "
    SQLQuery = SQLQuery & "'" & Format$(slMoDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(llGenStartDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(llGenEndDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(Now, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(Now, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("12/31/2069", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & igUstCode & ", "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    llAbfCode = gInsertAndReturnCode(SQLQuery, "abf_Ast_Build_Queue", "abfCode", "Replace")
    If llAbfCode > 0 Then
        If ilShttCode <= 0 Then
            'Update CPTT
            SQLQuery = "UPDATE cptt SET "
            SQLQuery = SQLQuery + "cpttAstStatus = " & "'R'"
            SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & ilVefCode
            SQLQuery = SQLQuery + " AND cpttStartDate = '" & Format(slMoDate, sgSQLDateForm) & "'"
            SQLQuery = SQLQuery + " AND cpttAstStatus <> 'N'" & ")"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand1:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "modAgmnt-mAbfInsert"
                mAbfInsert = -1
                Exit Function
            End If
        Else
            'Update CPTT
            SQLQuery = "UPDATE cptt SET "
            SQLQuery = SQLQuery + "cpttAstStatus = " & "'R'"
            SQLQuery = SQLQuery + " WHERE (cpttVefCode = " & ilVefCode
            SQLQuery = SQLQuery + " AND cpttShfCode = " & ilShttCode
            SQLQuery = SQLQuery + " AND cpttStartDate = '" & Format(slMoDate, sgSQLDateForm) & "'"
            SQLQuery = SQLQuery + " AND cpttAstStatus <> 'N'" & ")"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand1:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "modAgmnt-mAbfInsert"
                mAbfInsert = -1
                Exit Function
            End If
        End If
    End If
    mAbfInsert = llAbfCode
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "modAgmnt-mAbfInsert"
    mAbfInsert = -1
    Exit Function
'ErrHand1:
'    gHandleError "AffErrorLog.txt", "gSetStationSpotBuilder"
'    Return
End Function

Public Function gInitDlfInfo() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    gCloseDlfInfo
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "AirDay", adChar, 1
        .Append "AirTime", adInteger
        .Append "LocalTime", adChar, 12
        .Append "EtfCode", adInteger
        .Append "EnfCode", adInteger
        .Append "Zone", adChar, 1
    End With
    rst.Open
    Set gInitDlfInfo = rst
End Function

Public Sub gCloseDlfInfo()
    On Error Resume Next
    If Not DlfInfo_rst Is Nothing Then
        If (DlfInfo_rst.State And adStateOpen) <> 0 Then
            DlfInfo_rst.Close
        End If
        Set DlfInfo_rst = Nothing
    End If
End Sub

Public Function gBuildDlfInfo(ilVefCode As Integer, slStartDate As String, slEndDate As String, slZone As String) As Integer
    Dim slDay As String
    Dim ilDay As Integer
    Dim llDate As Long
    ReDim blDayTest(0 To 2) As Boolean
    ReDim slDate(0 To 2) As String
    
    On Error GoTo ErrHand
    gBuildDlfInfo = False
    bgDlfExist = False
    If sgDelNet <> "Y" Then
        Exit Function
    End If
    blDayTest(0) = False
    blDayTest(1) = False
    blDayTest(2) = False
    For llDate = gDateValue(slStartDate) To gDateValue(slEndDate) Step 1
        Select Case gWeekDayLong(llDate)
            Case 0, 1, 2, 3, 4
                slDay = "0"
                ilDay = 0
            Case 5
                slDay = "6"
                ilDay = 1
            Case 6
                slDay = "7"
                ilDay = 2
        End Select
        If Not blDayTest(ilDay) Then
            SQLQuery = "SELECT Max(dlfStartDate) FROM Dlf_Delivery_Links"
            SQLQuery = SQLQuery & " WHERE dlfVefCode = " & ilVefCode
            SQLQuery = SQLQuery & " AND dlfAirDay = '" & slDay & "'"
            SQLQuery = SQLQuery & " AND dlfStartDate <= '" & Format(llDate, sgSQLDateForm) & "'"
            SQLQuery = SQLQuery & " AND UCase(dlfZone) = '" & UCase(slZone) & "'"
            SQLQuery = SQLQuery & " AND ((dlfCmmlSched = 'Y') Or (dlfmnfSubfeed > 0)) "
            Set dlf_rst = gSQLSelectCall(SQLQuery)
            If Not dlf_rst.EOF Then
                If Not IsNull(dlf_rst(0).Value) Then
                    blDayTest(ilDay) = True
                    slDate(ilDay) = Format(CStr(dlf_rst(0).Value), sgSQLDateForm)
                Else
                    SQLQuery = "SELECT Max(dlfStartDate) FROM Dlf_Delivery_Links"
                    SQLQuery = SQLQuery & " WHERE dlfVefCode = " & ilVefCode
                    SQLQuery = SQLQuery & " AND dlfAirDay = '" & slDay & "'"
                    SQLQuery = SQLQuery & " AND dlfStartDate <= '" & Format("12/31/2069", sgSQLDateForm) & "'"
                    SQLQuery = SQLQuery & " AND UCase(dlfZone) = '" & UCase(slZone) & "'"
                    SQLQuery = SQLQuery & " AND ((dlfCmmlSched = 'Y') Or (dlfmnfSubfeed > 0)) "
                    Set dlf_rst = gSQLSelectCall(SQLQuery)
                    If Not dlf_rst.EOF Then
                        If Not IsNull(dlf_rst(0).Value) Then
                            blDayTest(ilDay) = True
                            slDate(ilDay) = Format(CStr(dlf_rst(0).Value), sgSQLDateForm)
                        Else
                        End If
                    End If
                End If
            Else
                SQLQuery = "SELECT Max(dlfStartDate) FROM Dlf_Delivery_Links"
                SQLQuery = SQLQuery & " WHERE dlfVefCode = " & ilVefCode
                SQLQuery = SQLQuery & " AND dlfAirDay = '" & slDay & "'"
                SQLQuery = SQLQuery & " AND dlfStartDate <= '" & Format("12/31/2069", sgSQLDateForm) & "'"
                SQLQuery = SQLQuery & " AND UCase(dlfZone) = '" & UCase(slZone) & "'"
                SQLQuery = SQLQuery & " AND ((dlfCmmlSched = 'Y') Or (dlfmnfSubfeed > 0)) "
                Set dlf_rst = gSQLSelectCall(SQLQuery)
                If Not dlf_rst.EOF Then
                    If Not IsNull(dlf_rst(0).Value) Then
                        blDayTest(ilDay) = True
                        slDate(ilDay) = Format(CStr(dlf_rst(0).Value), sgSQLDateForm)
                    Else
                    End If
                End If
            End If
        End If
    Next llDate
            
    For ilDay = 0 To 2 Step 1
        If blDayTest(ilDay) Then
            If ilDay = 0 Then
                slDay = "0"
            ElseIf ilDay = 1 Then
                slDay = "6"
            Else
                slDay = "7"
            End If
            SQLQuery = "SELECT * FROM Dlf_Delivery_Links"
            SQLQuery = SQLQuery & " WHERE dlfVefCode = " & ilVefCode
            SQLQuery = SQLQuery & " AND dlfAirDay = '" & slDay & "'"
            SQLQuery = SQLQuery & " AND dlfStartDate = '" & slDate(ilDay) & "'"
            SQLQuery = SQLQuery & " AND UCase(dlfZone) = '" & UCase(slZone) & "'"
            SQLQuery = SQLQuery & " AND ((dlfCmmlSched = 'Y') Or (dlfmnfSubfeed > 0)) "
            SQLQuery = SQLQuery & " Order By dlfAirTime"
            Set dlf_rst = gSQLSelectCall(SQLQuery)
            Do While Not dlf_rst.EOF
                bgDlfExist = True
                DlfInfo_rst.AddNew Array("AirDay", "AirTime", "LocalTime", "EtfCode", "EnfCode", "Zone"), Array(dlf_rst!dlfAirDay, gTimeToLong(Format(CStr(dlf_rst!dlfAirTime), sgShowTimeWSecForm), False), Format(CStr(dlf_rst!dlfLocalTime), sgShowTimeWSecForm), dlf_rst!dlfEtfCode, dlf_rst!dlfEnfCode, Left(slZone, 1))
                dlf_rst.MoveNext
            Loop
        End If
    Next ilDay
    gBuildDlfInfo = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.Txt", "modAgreement-gBuildDlfInfo"
    Resume Next
End Function

Public Function gFindDlf(ilDay As Integer, llTime As Long, ilEtfCode As Integer, ilEnfCode As Integer, slZone As String) As Long
    Dim llDate As Long
    Dim slDay As String
    Dim llDlfTime As Long
    
    On Error GoTo ErrHand
    gFindDlf = llTime
    If Not bgDlfExist Then
        Exit Function
    End If
    
    Select Case ilDay
        Case vbMonday, vbTuesday, vbWednesday, vbThursday, vbFriday
            slDay = "0"
        Case vbSaturday
            slDay = "6"
        Case vbSunday
            slDay = "7"
    End Select

    '1/25/18: Test if any delivery links defined for the day.  If not, treat as links not defined
    DlfInfo_rst.Filter = "AirDay = '" & slDay & "' "
    If DlfInfo_rst.EOF Then
        Exit Function
    End If
    
    DlfInfo_rst.Filter = "AirDay = '" & slDay & "' And AirTime = " & llTime & " And EtfCode = " & ilEtfCode & " And EnfCode = " & ilEnfCode & " And Zone = '" & Left(slZone, 1) & "'"
    If Not DlfInfo_rst.EOF Then
        'gFindDlf = gTimeToLong(Format(CStr(DlfInfo_rst!LocalTime), "h:mm:ssAM/PM"), False)
        llDlfTime = gTimeToLong(Format(CStr(DlfInfo_rst!LocalTime), "h:mm:ssAM/PM"), False)
        '9/6/14: Add test that Delivery time between 9p-12m
        'If (llTime >= 0) And (llTime < 21600) And (llDlfTime < 86400) Then
'12/12/14: Remove
'        If (llTime >= 0) And (llTime < 21600) And (llDlfTime > 75600) And (llDlfTime < 86400) Then
'            If ilDay = vbSunday Then
'                ilDay = vbSaturday
'            Else
'                ilDay = ilDay - 1
'            End If
'        '9/6/14: Add test that avail time between 9p-12m
'        'ElseIf (llTime < 86400) And (llDlfTime >= 0) And (llDlfTime < 21600) Then
'        ElseIf (llDlfTime > 75600) And (llTime < 86400) And (llDlfTime >= 0) And (llDlfTime < 21600) Then
'            If ilDay = vbSaturday Then
'                ilDay = vbSunday
'            Else
'                ilDay = ilDay + 1
'            End If
'        End If
        gFindDlf = llDlfTime
    Else
        gFindDlf = -1
    End If
    DlfInfo_rst.Filter = adFilterNone
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.Txt", "modAgreement-gFindDlf"
    Resume Next
End Function

Public Sub gDlfExist(ilVefCode As Integer, slStartDate As String, slEndDate As String, slZone As String)
    Dim slDay As String
    Dim ilDay As Integer
    Dim llDate As Long
    Dim blMoChecked As Boolean
    
    On Error GoTo ErrHand
    bgDlfExist = False
    If sgDelNet <> "Y" Then
        Exit Sub
    End If
    
    blMoChecked = False
    For llDate = gDateValue(slStartDate) To gDateValue(slEndDate) Step 1
        Select Case gWeekDayLong(llDate)
            Case 0, 1, 2, 3, 4
                slDay = "0"
                ilDay = 0
            Case 5
                slDay = "6"
                ilDay = 1
            Case 6
                slDay = "7"
                ilDay = 2
        End Select
        If Not blMoChecked Then
            If ilDay = 0 Then
                blMoChecked = True
            End If
            SQLQuery = "SELECT Max(dlfStartDate) FROM Dlf_Delivery_Links"
            SQLQuery = SQLQuery & " WHERE dlfVefCode = " & ilVefCode
            SQLQuery = SQLQuery & " AND dlfAirDay = '" & slDay & "'"
            SQLQuery = SQLQuery & " AND dlfStartDate <= '" & Format(llDate, sgSQLDateForm) & "'"
            SQLQuery = SQLQuery & " AND UCase(dlfZone) = '" & UCase(slZone) & "'"
            SQLQuery = SQLQuery & " AND ((dlfCmmlSched = 'Y') Or (dlfmnfSubfeed > 0)) "
            Set dlf_rst = gSQLSelectCall(SQLQuery)
            If Not dlf_rst.EOF Then
                If Not IsNull(dlf_rst(0).Value) Then
                    bgDlfExist = True
                    Exit Sub
                Else
                    SQLQuery = "SELECT Max(dlfStartDate) FROM Dlf_Delivery_Links"
                    SQLQuery = SQLQuery & " WHERE dlfVefCode = " & ilVefCode
                    SQLQuery = SQLQuery & " AND dlfAirDay = '" & slDay & "'"
                    SQLQuery = SQLQuery & " AND dlfStartDate <= '" & Format("12/31/2069", sgSQLDateForm) & "'"
                    SQLQuery = SQLQuery & " AND UCase(dlfZone) = '" & UCase(slZone) & "'"
                    SQLQuery = SQLQuery & " AND ((dlfCmmlSched = 'Y') Or (dlfmnfSubfeed > 0)) "
                    Set dlf_rst = gSQLSelectCall(SQLQuery)
                    If Not dlf_rst.EOF Then
                        If Not IsNull(dlf_rst(0).Value) Then
                            bgDlfExist = True
                            Exit Sub
                        End If
                    End If
                End If
            Else
                SQLQuery = "SELECT Max(dlfStartDate) FROM Dlf_Delivery_Links"
                SQLQuery = SQLQuery & " WHERE dlfVefCode = " & ilVefCode
                SQLQuery = SQLQuery & " AND dlfAirDay = '" & slDay & "'"
                SQLQuery = SQLQuery & " AND dlfStartDate <= '" & Format("12/31/2069", sgSQLDateForm) & "'"
                SQLQuery = SQLQuery & " AND UCase(dlfZone) = '" & UCase(slZone) & "'"
                SQLQuery = SQLQuery & " AND ((dlfCmmlSched = 'Y') Or (dlfmnfSubfeed > 0)) "
                Set dlf_rst = gSQLSelectCall(SQLQuery)
                If Not dlf_rst.EOF Then
                    If Not IsNull(dlf_rst(0).Value) Then
                        bgDlfExist = True
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next llDate
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.Txt", "modAgreement-gDlfExist"
    Exit Sub
End Sub
Public Sub gSendKeys(slString As String, blWait As Boolean)
    On Error Resume Next
    '2/16/16: SendKeys fails in Exe mode
    '''2/5/16
    '''Sendkeys slString, blWait
    ''2/16/16
    ''CreateObject("WScript.Shell").Sendkeys slString, blWait
    'If gRunningInIDE() Then
        CreateObject("WScript.Shell").SendKeys slString, blWait
    'Else
    '    Sendkeys slString, blWait
    'End If

End Sub


Public Function gRunningInIDE() As Boolean
   Dim llValue As Long
   
   llValue = 0
   Debug.Assert Not mTestIDE(llValue)
   gRunningInIDE = llValue = 1

End Function

Private Function mTestIDE(llValue As Long) As Boolean
    llValue = 1
End Function
