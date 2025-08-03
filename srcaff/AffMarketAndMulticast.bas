Attribute VB_Name = "modMarketAndMulticast"
'******************************************************
'*  modMarketAndMulticast - contains various general routines
'*
'*  Copyright Counterpoint Software, Inc. 2006
'*
'*  Helper functions to support Market, Cluster and Multicast
'*
'*  Doug Smith 12/9/05
'******************************************************
Option Explicit
Option Compare Text

Private tmDat() As DAT


Public Function gGetStaMulticastGroupID(shttCode As Integer) As Long

    'D.S. 12/9/05
    'Get a station's multicast group ID, if it has one otherwise return 0
    Dim tmp_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select shttMultiCastGroupID FROM shtt where shttCode = " & shttCode
    Set tmp_rst = gSQLSelectCall(SQLQuery)
    If Not tmp_rst.EOF Then
        If IsNull(tmp_rst!shttMultiCastGroupID) = True Then
            gGetStaMulticastGroupID = 0
        Else
            gGetStaMulticastGroupID = tmp_rst!shttMultiCastGroupID
        End If
    Else
        gGetStaMulticastGroupID = 0
    End If
    
Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gGetStaMulticastGroupID"
    gGetStaMulticastGroupID = 0
    Exit Function
End Function

Public Function gGetStaMarketCode(shttCode As Integer) As Long

    'D.S. 12/9/05
    'Get a station's market code, if it has one otherwise return 0
    Dim tmp_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select shttMktCode FROM shtt where shttCode = " & shttCode
    Set tmp_rst = gSQLSelectCall(SQLQuery)
    If Not tmp_rst.EOF Then
        If IsNull(tmp_rst!shttMktCode) = True Then
            gGetStaMarketCode = 0
        Else
            gGetStaMarketCode = tmp_rst!shttMktCode
        End If
    End If
    
Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gGetStaMarketCode"
    gGetStaMarketCode = 0
    Exit Function
End Function

Public Function gGetStaMarketName(mktCode As Long) As String

    'D.S. 12/9/05
    'Get a station's market name, if it has one otherwise return 0
    Dim tmp_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select mktName FROM mkt where mktCode = " & mktCode
    Set tmp_rst = gSQLSelectCall(SQLQuery)
    If Not tmp_rst.EOF Then
        If IsNull(tmp_rst!mktName) = True Then
            gGetStaMarketName = ""
        Else
            gGetStaMarketName = Trim$(tmp_rst!mktName)
        End If
    End If
    
Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gGetStaMarketName"
    gGetStaMarketName = ""
    Exit Function
End Function

Public Function gGetCallLettersByShttCode(shttCode As Integer) As String

    'D.S. 12/9/05
    'Get a station's call letters from it's Shtt Code
    Dim tmp_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select shttCallLetters FROM shtt where shttCode = " & shttCode
    Set tmp_rst = gSQLSelectCall(SQLQuery)
    If Not tmp_rst.EOF Then
        If IsNull(tmp_rst!shttCallLetters) = True Then
            gGetCallLettersByShttCode = ""
        Else
            gGetCallLettersByShttCode = Trim$(tmp_rst!shttCallLetters)
        End If
    End If
    
Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gGetCallLettersByShttCode"
    gGetCallLettersByShttCode = ""
    Exit Function
End Function
Public Function gGetVehNameByVefCode(vefCode As Integer) As String

    'D.S. 3/14/06
    'Get a vehicles name from it's vef Code
    Dim tmp_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select vefName FROM VEF_Vehicles WHERE vefCode = " & vefCode
    Set tmp_rst = gSQLSelectCall(SQLQuery)
    If Not tmp_rst.EOF Then
        If IsNull(tmp_rst!vefName) = True Then
            gGetVehNameByVefCode = ""
        Else
            gGetVehNameByVefCode = Trim$(tmp_rst!vefName)
        End If
    End If
    
Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gGetVehNameByVefCode"
    gGetVehNameByVefCode = ""
    Exit Function
End Function


Public Function gGetCallLettersByAttCode(lAttCode As Long) As String

    'D.S. 12/9/05
    'Get a station's call letters from the AttCode
    Dim tmp_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select attShfCode FROM att where attCode = " & lAttCode
    Set tmp_rst = gSQLSelectCall(SQLQuery)
    If Not tmp_rst.EOF Then
        If IsNull(tmp_rst!attshfCode) = True Then
            gGetCallLettersByAttCode = ""
        Else
            gGetCallLettersByAttCode = gGetCallLettersByShttCode(Trim$(tmp_rst!attshfCode))
        End If
    End If
    
Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gGetCallLettersByAttCode"
    gGetCallLettersByAttCode = ""
    Exit Function
End Function


Public Function gIsMulticast(shttCode As Integer) As Integer

    'D.S. 12/9/05
    'Returns True or False if a stations is defined as part of a multicast
    Dim tmp_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select shttMultiCastGroupID FROM shtt where shttCode = " & shttCode
    Set tmp_rst = gSQLSelectCall(SQLQuery)
    If Not tmp_rst.EOF Then
        If IsNull(tmp_rst!shttMultiCastGroupID) = True Then
            gIsMulticast = False
            Exit Function
        End If
        
        If tmp_rst!shttMultiCastGroupID = 0 Then
            gIsMulticast = False
        Else
            gIsMulticast = True
        End If
    Else
        gIsMulticast = False
    End If
    
Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gIsMulticast"
    gIsMulticast = False
    Exit Function
End Function


Public Function gMulticastMaxGroupID() As Long

    'D.S. 12/9/05
    'Returns the MAX groupID from mgt.  Used to get next group ID for new muticast

    Dim max_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(shttMultiCastGroupID) from shtt"
    Set max_rst = gSQLSelectCall(SQLQuery)
        If IsNull(max_rst(0).Value) Then
            gMulticastMaxGroupID = 0
        Else
            gMulticastMaxGroupID = max_rst(0).Value
        Exit Function
    End If

    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gMulticastMaxGroupID"
    gMulticastMaxGroupID = 0
End Function

Public Function gMarketClusterMaxGroupID() As Long

    'D.S. 12/9/05
    'Returns the MAX groupID from mgt.  Used to get next group ID for new muticast

    Dim max_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(shttClusterGroupID) from shtt"
    Set max_rst = gSQLSelectCall(SQLQuery)
        If IsNull(max_rst(0).Value) Then
            gMarketClusterMaxGroupID = 0
        Else
            gMarketClusterMaxGroupID = max_rst(0).Value
        Exit Function
    End If

    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gMarketClusterMaxGroupID"
    gMarketClusterMaxGroupID = 0
End Function
Public Function gGetDatByAttCode(llAttCode As Long) As Integer
    
    'D.S. 12/13/05 Get the DAT agreement and load it into tmDat
    
    Dim ilUpper As Integer
    Dim dat_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    ReDim tmDat(0 To 0) As DAT
    ilUpper = 0
    SQLQuery = "SELECT * "
    SQLQuery = SQLQuery + " FROM dat"
    SQLQuery = SQLQuery + " WHERE (datAtfCode = " & llAttCode & ")"
    SQLQuery = SQLQuery & " ORDER BY datFdStTime"
    Set dat_rst = gSQLSelectCall(SQLQuery)
    
    If Not dat_rst.EOF Then
        While Not dat_rst.EOF
            tmDat(ilUpper).iStatus = 1
            tmDat(ilUpper).lCode = dat_rst!datCode    '(0).Value
            tmDat(ilUpper).lAtfCode = dat_rst!datAtfCode  '(1).Value
            tmDat(ilUpper).iShfCode = dat_rst!datShfCode  '(2).Value
            tmDat(ilUpper).iVefCode = dat_rst!datVefCode  '(3).Value
            'tmDat(ilUpper).iDACode = dat_rst!datDACode    '(4).Value
            tmDat(ilUpper).iFdDay(0) = dat_rst!datFdMon   '(5).Value
            tmDat(ilUpper).iFdDay(1) = dat_rst!datFdTue   '(6).Value
            tmDat(ilUpper).iFdDay(2) = dat_rst!datFdWed   '(7).Value
            tmDat(ilUpper).iFdDay(3) = dat_rst!datFdThu   '(8).Value
            tmDat(ilUpper).iFdDay(4) = dat_rst!datFdFri   '(9).Value
            tmDat(ilUpper).iFdDay(5) = dat_rst!datFdSat   '(10).Value
            tmDat(ilUpper).iFdDay(6) = dat_rst!datFdSun   '(11).Value
            If Second(dat_rst!datFdStTime) = 0 Then
                tmDat(ilUpper).sFdSTime = Format$(CStr(dat_rst!datFdStTime), sgShowTimeWOSecForm)
            Else
                tmDat(ilUpper).sFdSTime = Format$(CStr(dat_rst!datFdStTime), sgShowTimeWSecForm)
            End If
            If Second(dat_rst!datFdEdTime) = 0 Then
                tmDat(ilUpper).sFdETime = Format$(CStr(dat_rst!datFdEdTime), sgShowTimeWOSecForm)
            Else
                tmDat(ilUpper).sFdETime = Format$(CStr(dat_rst!datFdEdTime), sgShowTimeWSecForm)
            End If
            tmDat(ilUpper).iFdStatus = dat_rst!datFdStatus    '(14).Value
            tmDat(ilUpper).iPdDay(0) = dat_rst!datPdMon   '(15).Value
            tmDat(ilUpper).iPdDay(1) = dat_rst!datPdTue   '(16).Value
            tmDat(ilUpper).iPdDay(2) = dat_rst!datPdWed   '(17).Value
            tmDat(ilUpper).iPdDay(3) = dat_rst!datPdThu   '(18).Value
            tmDat(ilUpper).iPdDay(4) = dat_rst!datPdFri   '(19).Value
            tmDat(ilUpper).iPdDay(5) = dat_rst!datPdSat   '(20).Value
            tmDat(ilUpper).iPdDay(6) = dat_rst!datPdSun   '(21).Value
            tmDat(ilUpper).sPdDayFed = dat_rst!datPdDayFed
            If (tmDat(ilUpper).iFdStatus <= 1) Or (tmDat(ilUpper).iFdStatus = 9) Or (tmDat(ilUpper).iFdStatus = 10) Then
                If Second(dat_rst!datPdStTime) = 0 Then
                    tmDat(ilUpper).sPdSTime = Format$(CStr(dat_rst!datPdStTime), sgShowTimeWOSecForm)
                Else
                    tmDat(ilUpper).sPdSTime = Format$(CStr(dat_rst!datPdStTime), sgShowTimeWSecForm)
                End If
                If Second(dat_rst!datPdEdTime) = 0 Then
                    tmDat(ilUpper).sPdETime = Format$(CStr(dat_rst!datPdEdTime), sgShowTimeWOSecForm)
                Else
                    tmDat(ilUpper).sPdETime = Format$(CStr(dat_rst!datPdEdTime), sgShowTimeWSecForm)
                End If
            Else
                tmDat(ilUpper).sPdSTime = ""
                tmDat(ilUpper).sPdETime = ""
            End If
            ilUpper = ilUpper + 1
            ReDim Preserve tmDat(0 To ilUpper) As DAT
            dat_rst.MoveNext
        Wend
    End If
    gGetDatByAttCode = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gGetDatByAttCode"
    gGetDatByAttCode = False
    Exit Function
End Function

Public Function gCompare2AgrmntsPledges(frm As Form, lTestAgainst As Long, iVefCode) As String

    'D.S. 12/13/05
    'Compare two agreement by passing in there attCodes
    
    Dim tlDat() As DAT
    Dim ilLoop As Integer
    Dim ilLoop2 As Integer
    Dim ilRet As Integer
    Dim ilIdx As Integer
    Dim ilTimeType As Integer
    Dim ilFound As Integer
    Dim ilPledgeType As Integer
    Dim attrst As ADODB.Recordset
    
    On Error GoTo ErrHand

    gCompare2AgrmntsPledges = ""
    
    ilFound = False
    
    SQLQuery = "SELECT attPledgeType FROM att WHERE attCode = " & lTestAgainst
    Set attrst = gSQLSelectCall(SQLQuery)
    If Not attrst.EOF Then
        If attrst!attPledgeType = "D" Then
            ilPledgeType = 0
        ElseIf attrst!attPledgeType = "A" Then
            ilPledgeType = 1
        ElseIf attrst!attPledgeType = "C" Then
            ilPledgeType = 2
        Else
            ilPledgeType = -1
        End If
    Else
        ilPledgeType = -1
    End If

    'Get the first agreement
    'Load the first agreement in the local DAT array
    ReDim tlDat(0 To UBound(tgDat)) As DAT
    For ilLoop = 0 To UBound(tgDat) - 1 Step 1
        tlDat(ilLoop) = tgDat(ilLoop)
    Next ilLoop


    'Get the second agreement and compare it to the first agreement
    ilRet = gGetDatByAttCode(lTestAgainst)
    
    'Did they chose dayparts, avails or cd/tape
    If frm!optTimeType(0).Value Then
        ilTimeType = 0
    End If
    If frm!optTimeType(1).Value Then
        ilTimeType = 1
    End If
    If frm!optTimeType(2).Value Then
        ilTimeType = 2
    End If

    'If tmDat(0).iDACode <> ilTimeType Then
    If ilPledgeType <> ilTimeType Then
        'If tmDat(0).iDACode = 0 Then
        If ilPledgeType = 0 Then
            gCompare2AgrmntsPledges = " Pledge Type should be set to Dayparts "
            Exit Function
        End If
        'If tmDat(0).iDACode = 1 Then
        If ilPledgeType = 1 Then
            gCompare2AgrmntsPledges = " Pledge Type should be set to Avails "
            Exit Function
        End If
        'If tmDat(0).iDACode = 2 Then
        If ilPledgeType = 2 Then
            gCompare2AgrmntsPledges = " Pledge Type should be set to CD/Tape "
            Exit Function
        End If
    End If
    
    
    For ilLoop = 0 To UBound(tgDat) - 1 Step 1
        For ilLoop2 = 0 To UBound(tmDat) - 1 Step 1
            ilFound = True
        For ilIdx = 0 To 6 Step 1
                If tlDat(ilLoop).iFdDay(ilIdx) <> tmDat(ilLoop2).iFdDay(ilIdx) Then
                gCompare2AgrmntsPledges = " One of the Feed Days is incorrect. "
                    ilFound = False
                    'Exit Function
            End If
        Next ilIdx

            If tlDat(ilLoop).iFdStatus <> tmDat(ilLoop2).iFdStatus Then
            gCompare2AgrmntsPledges = "One of the Feed Status is incorrect. "
                ilFound = False
                'Exit Function
        End If

        For ilIdx = 0 To 6 Step 1
                If tlDat(ilLoop).iPdDay(ilIdx) <> tmDat(ilLoop2).iPdDay(ilIdx) Then
                gCompare2AgrmntsPledges = " One of Pledge Days is incorrect. "
                    ilFound = False
                    'Exit Function
            End If
        Next ilIdx

            If iVefCode <> tmDat(ilLoop2).iVefCode Then
            gCompare2AgrmntsPledges = " The vehicle is incorrect. "
                ilFound = False
                'Exit Function
        End If


            If gTimeToLong(tlDat(ilLoop).sFdETime, False) <> gTimeToLong(tmDat(ilLoop2).sFdETime, False) Then
            gCompare2AgrmntsPledges = " One of Feed End Times is incorrect. "
                ilFound = False
                'Exit Function
        End If

            If gTimeToLong(tlDat(ilLoop).sFdSTime, False) <> gTimeToLong(tmDat(ilLoop2).sFdSTime, False) Then
            gCompare2AgrmntsPledges = " One of Feed Start Times is incorrect. "
                ilFound = False
                'Exit Function
        End If

            If gTimeToLong(tlDat(ilLoop).sPdETime, False) <> gTimeToLong(tmDat(ilLoop2).sPdETime, False) Then
            gCompare2AgrmntsPledges = " One of Pledge End Times is incorrect. "
                ilFound = False
                'Exit Function
        End If

            If gTimeToLong(tlDat(ilLoop).sPdSTime, False) <> gTimeToLong(tmDat(ilLoop2).sPdSTime, False) Then
            gCompare2AgrmntsPledges = " One of Pledge Start Times is incorrect. "
                ilFound = False
                'Exit Function
            End If
            If ilFound Then
                Exit For
            Else
                ilFound = ilFound
        End If
        Next ilLoop2
        If Not ilFound Then
            Exit For
        End If
    Next ilLoop
    
    '6/19/11: Handle case of no pledge info
    If (UBound(tlDat) <= LBound(tlDat)) And (UBound(tmDat) <= LBound(tmDat)) And (Not ilFound) Then
        ilFound = True
    End If
    
    If ilFound Then
        gCompare2AgrmntsPledges = ""
    Else
        gCompare2AgrmntsPledges = " Error on Pledge Line " & ilLoop + 1
    End If
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gCompre2AgrmntsPledges"
    gCompare2AgrmntsPledges = "Error"
    Exit Function
End Function


'This function needs to return an array of att codes.  There may be more than one agreement for
'the veh/sta combination.
'Public Function gGetAttCodeFromStnAndVefCode(iStnCode As Integer, iVefCode As Integer) As Long
'
'    'D.S. Get an agreements attCode by passing in the station and vehicle codes
'
'    Dim att_rst As ADODB.Recordset
'
'    On Error GoTo ErrHand
'
'    gGetAttCodeFromStnAndVefCode = 0
'
'    SQLQuery = "SELECT attCode "
'    SQLQuery = SQLQuery + " FROM att"
'    SQLQuery = SQLQuery + " WHERE (attShfCode = " & iStnCode & " And attVefCode = " & iVefCode & ")"
'    Set att_rst = gSQLSelectCall(SQLQuery)
'
'    If Not att_rst.EOF Then
'        gGetAttCodeFromStnAndVefCode = att_rst!attCode
'    Else
'        gGetAttCodeFromStnAndVefCode = 0
'    End If
'
'    Exit Function
'

Public Function gGetVehCodeFromAttCode(sATTCode As String) As Long

    'D.S. 12/13/05
    'Get the vehicle code by passing in the attCode
    
    Dim att_rst As ADODB.Recordset
    On Error GoTo ErrHand:
    
    SQLQuery = "SELECT attVefCode"
    SQLQuery = SQLQuery + " FROM att"
    SQLQuery = SQLQuery + " WHERE (attCode = " & sATTCode & ")"
    Set att_rst = gSQLSelectCall(SQLQuery)
    
    If Not att_rst.EOF Then
        gGetVehCodeFromAttCode = att_rst!attvefCode
    Else
        gGetVehCodeFromAttCode = 0
        Exit Function
    End If
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gGetVehCodeFromAttCode"
    gGetVehCodeFromAttCode = 0
    Exit Function
End Function


Public Function gGetShttCodeFromAttCode(sATTCode As String) As Long

    'D.S. 12/13/05
    'Get the station code by passing in the attCode
    
    Dim shtt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "SELECT attshfCode"
    SQLQuery = SQLQuery + " FROM att"
    SQLQuery = SQLQuery + " WHERE (attCode = " & sATTCode & ")"
    Set shtt_rst = gSQLSelectCall(SQLQuery)
    
    If Not shtt_rst.EOF Then
        gGetShttCodeFromAttCode = shtt_rst!attshfCode
    Else
        gGetShttCodeFromAttCode = 0
        Exit Function
    End If
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gGetShttCodeFromAttCode"
    gGetShttCodeFromAttCode = 0
    Exit Function
End Function

Public Function gGetShttCodeFromCallLetters(sCallLetters As String) As String

    'D.S. 12/10/08
    'Get the station code by passing in the call letters
    
    Dim shtt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "SELECT shttCode"
    SQLQuery = SQLQuery + " FROM shtt"
    SQLQuery = SQLQuery + " WHERE (shttCallLetters = '" & sCallLetters & "')"
    Set shtt_rst = gSQLSelectCall(SQLQuery)
    
    If Not shtt_rst.EOF Then
        gGetShttCodeFromCallLetters = shtt_rst!shttCode
    Else
        gGetShttCodeFromCallLetters = "0"
        Exit Function
    End If
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.Txt", "modMarketAndMulticast-gGetShttCodeFromCallLetters"
    gGetShttCodeFromCallLetters = "0"
    Exit Function
End Function


'Public Function gCompare2AgrmntsByPostAndLogType(frm As Form, lAttCode As Long) As Integer
' Dan 7701 not used, so not fixed!
'    Dim tmp_rst As ADODB.Recordset
'    Dim ilBasePost As Integer
'    Dim ilPost As Integer
'    Dim ilBaseLog As Integer
'    Dim ilLog As Integer
'    Dim ilBaseExp As Integer
'    Dim ilExp As Integer
'    Dim ilIdx As Integer
'    Dim slBaseExportToWeb As String
'    Dim slBaseExportToCumulus As String
'    Dim slBaseExportToUnivision As String
'    Dim slBaseExportToMarketron As String
'    Dim slExportToWeb As String
'    Dim slExportToCumulus As String
'    Dim slExportToUnivision As String
'    Dim slExportToMarketron As String
'
'    On Error GoTo ErrHand
'
'    gCompare2AgrmntsByPostAndLogType = False
'
'    If frm!rbcPostType(0).Value Then
'        ilBasePost = 0
'    End If
'    If frm!rbcPostType(1).Value Then
'        ilBasePost = 1
'    End If
'    If frm!rbcPostType(2).Value Then
'        ilBasePost = 2
'    End If
'
'    If frm!rbcLogType(0).Value Then
'        ilBaseLog = 0
'    End If
'    If frm!rbcLogType(1).Value Then
'        ilBaseLog = 1
'    End If
'    If frm!rbcLogType(2).Value Then
'        ilBaseLog = 2
'    End If
'
'    If frm!rbcExportType(0).Value Then
'        ilBaseExp = 0
'    End If
'    If frm!rbcExportType(1).Value Then
'        ilBaseExp = 1
'        slBaseExportToWeb = "N"
'        slBaseExportToCumulus = "N"
'        slBaseExportToUnivision = "N"
'        slBaseExportToMarketron = "N"
'        If frm!ckcExportTo(0).Value = vbChecked Then
'            slBaseExportToWeb = "Y"
'        End If
'        '7701
'        With frm!cboLogDelivery
'            If .ListIndex > -1 Then
'                Select Case .ItemData(.ListIndex)
'                    Case Vendors.Cumulus
'                         slBaseExportToCumulus = "Y"
'                    Case Vendors.NetworkConnect
'                        slBaseExportToMarketron = "Y"
'                End Select
'            End If
'        End With
'
''         If frm!ckcExportTo(1).Value = vbChecked Then
''            slBaseExportToCumulus = "Y"
''        End If
''        If frm!ckcExportTo(2).Value = vbChecked Then
''            slBaseExportToUnivision = "Y"
''        End If
''        If frm!ckcExportTo(3).Value = vbChecked Then
''            slBaseExportToMarketron = "Y"
''        End If
'    End If
'    'If frm!rbcExportType(2).Value Then
'    '    ilBaseExp = 2
'    'End If
'    '7701  Dan added attExportToWeb to make function work
'    SQLQuery = "SELECT attLogType, attPostType, attExportType, attExportToWeb, vatWvtIdCodeLog from att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode WHERE attCode = " & lAttCode
''    SQLQuery = "SELECT attLogType, attPostType, attExportType, attWebInterface, attExportToUnivision, attExportToMarketron from att WHERE attCode = " & lAttCode
'    Set tmp_rst = gSQLSelectCall(SQLQuery)
'    If Not tmp_rst.EOF Then
'        ilPost = tmp_rst!attPostType
'        ilLog = tmp_rst!attLogType
'        ilExp = tmp_rst!attExportType
'        slExportToWeb = "N"
'        slExportToCumulus = "N"
'        slExportToUnivision = "N"
'        slExportToMarketron = "N"
'        If tmp_rst!attExportToWeb = "Y" Then
'            slExportToWeb = "Y"
'        End If
'        '7701
'        If gIfNullInteger(tmp_rst!vatWvtIdCodeLog) = Vendors.Cumulus Then
'            slExportToCumulus = "Y"
'        End If
'        If gIfNullInteger(tmp_rst!vatWvtIdCodeLog) = Vendors.NetworkConnect Then
'            slExportToMarketron = "Y"
'        End If
''        If tmp_rst!attWebInterface = "C" Then
''            slExportToCumulus = "Y"
''        End If
''        If tmp_rst!attExportToUnivision = "Y" Then
''            slExportToUnivision = "Y"
''        End If
''        If tmp_rst!attExportToMarketron = "Y" Then
''            slExportToMarketron = "Y"
''        End If
'    End If
'
'    If StrComp(ilBaseExp, ilExp, 1) <> 0 Then
'        gCompare2AgrmntsByPostAndLogType = 3
'        Exit Function
'    End If
'    If ilBaseExp = 1 Then
'        If slBaseExportToWeb <> slExportToWeb Then
'            gCompare2AgrmntsByPostAndLogType = 3
'            Exit Function
'        End If
'        If slBaseExportToCumulus <> slExportToCumulus Then
'            gCompare2AgrmntsByPostAndLogType = 3
'            Exit Function
'        End If
'        If slBaseExportToUnivision <> slExportToUnivision Then
'            gCompare2AgrmntsByPostAndLogType = 3
'            Exit Function
'        End If
'        If slBaseExportToMarketron <> slExportToMarketron Then
'            gCompare2AgrmntsByPostAndLogType = 3
'            Exit Function
'        End If
'
'    End If
'
'    If StrComp(ilBasePost, ilPost, 1) <> 0 Then
'        gCompare2AgrmntsByPostAndLogType = 1
'        Exit Function
'    End If
'
'    If StrComp(ilBaseLog, ilLog, 1) <> 0 Then
'        gCompare2AgrmntsByPostAndLogType = 2
'        Exit Function
'    End If
'
'    gCompare2AgrmntsByPostAndLogType = True
'    Exit Function
'End Function

Public Sub gAlignMulticastStations(ilVefCode As Integer, slStationOrAgreementItemData As String, lbcStationsAlign As ListBox, lbcStationsSource As ListBox, Optional llAttStartDate As Long = -1, Optional llAttEndDate As Long = -1)
    '9/12/18: Added date as an option.  It was added to handle call from Spot Utilty
    Dim ilLoop1 As Integer
    Dim ilShttCode1 As Integer
    Dim llGroupID As Long
    Dim slMulticast1 As String
    Dim llLastDate1 As Long
    Dim ilLoop2 As Integer
    Dim ilShttCode2 As Integer
    Dim slMulticast2 As String
    Dim llLastDate2 As Long
    Dim ilShttCode3 As Integer
    Dim llOnAir As Long
    Dim ilFound As Integer
    Dim ilLoopIndex As Integer
    Dim att_rst As ADODB.Recordset
    Dim att_rst2 As ADODB.Recordset
    Dim shtt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
        
    For ilLoop1 = 0 To lbcStationsAlign.ListCount - 1 Step 1
        If slStationOrAgreementItemData = "S" Then
            ilShttCode1 = lbcStationsAlign.ItemData(ilLoop1)
        Else
            SQLQuery = "SELECT attShfCode FROM att"
            SQLQuery = SQLQuery + " WHERE (attCode = " & lbcStationsAlign.ItemData(ilLoop1) & ")"
            Set att_rst = gSQLSelectCall(SQLQuery)
            If Not att_rst.EOF Then
                ilShttCode1 = att_rst!attshfCode
            Else
                Exit Sub
            End If
        End If
        If gIsMulticast(ilShttCode1) Then
            'Determine if agreement exist, if not then can determine if any other station needs to be multicast with it
            llLastDate1 = 0
            llOnAir = 0
            slMulticast1 = ""
            SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate, attMulticast FROM att"
            SQLQuery = SQLQuery + " WHERE (attShfCode = " & ilShttCode1 & " AND attVefCode = " & ilVefCode & ")"
            Set att_rst = gSQLSelectCall(SQLQuery)
            While Not att_rst.EOF
                '9/12/18: Added date as an option.
                llOnAir = gDateValue(att_rst!attOnAir)
                If gDateValue(att_rst!attOffAir) <= gDateValue(att_rst!attDropDate) Then
                    If gDateValue(att_rst!attOffAir) > llLastDate1 Then
                        llLastDate1 = gDateValue(att_rst!attOffAir)
                        slMulticast1 = att_rst!attMulticast
                    End If
                Else
                    If gDateValue(att_rst!attDropDate) > llLastDate1 Then
                        llLastDate1 = gDateValue(att_rst!attDropDate)
                        slMulticast1 = att_rst!attMulticast
                    End If
                End If
                att_rst.MoveNext
            Wend
            '9/12/18: Added date as an option.
            If (slMulticast1 = "Y") And ((llAttEndDate >= llOnAir) Or (llAttEndDate = -1)) And ((llAttStartDate <= llLastDate1) Or (llAttStartDate = -1)) Then
                'Obtain list of other multicast stations
                llGroupID = gGetStaMulticastGroupID(ilShttCode1)
                SQLQuery = "Select shttCode FROM shtt where shttMultiCastGroupID = " & llGroupID
                Set shtt_rst = gSQLSelectCall(SQLQuery)
                While Not shtt_rst.EOF
                    ilShttCode2 = shtt_rst!shttCode
                    llLastDate2 = 0
                    slMulticast2 = ""
                    SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate, attMulticast FROM att"
                    SQLQuery = SQLQuery + " WHERE (attShfCode = " & ilShttCode2 & " AND attVefCode = " & ilVefCode & ")"
                    Set att_rst = gSQLSelectCall(SQLQuery)
                    While Not att_rst.EOF
                        If gDateValue(att_rst!attOffAir) <= gDateValue(att_rst!attDropDate) Then
                            If gDateValue(att_rst!attOffAir) > llLastDate2 Then
                                llLastDate2 = gDateValue(att_rst!attOffAir)
                                slMulticast2 = att_rst!attMulticast
                            End If
                        Else
                            If gDateValue(att_rst!attDropDate) > llLastDate2 Then
                                llLastDate2 = gDateValue(att_rst!attDropDate)
                                slMulticast2 = att_rst!attMulticast
                            End If
                        End If
                        att_rst.MoveNext
                    Wend
                    If (slMulticast2 = "Y") And (llLastDate1 = llLastDate2) Then
                        ilFound = False
                        If slStationOrAgreementItemData = "S" Then
                            For ilLoop2 = 0 To lbcStationsAlign.ListCount - 1 Step 1
                                If ilShttCode2 = lbcStationsAlign.ItemData(ilLoop2) Then
                                    SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate, attMulticast FROM att"
                                    SQLQuery = SQLQuery + " WHERE (attShfCode = " & ilShttCode2 & " AND attVefCode = " & ilVefCode & ")"
                                    Set att_rst = gSQLSelectCall(SQLQuery)
                                    While Not att_rst.EOF
                                        If gDateValue(att_rst!attOffAir) <= gDateValue(att_rst!attDropDate) Then
                                            If llLastDate2 = gDateValue(att_rst!attOffAir) Then
                                                ilFound = True
                                                Exit For
                                            End If
                                        Else
                                            If llLastDate2 = gDateValue(att_rst!attDropDate) Then
                                                ilFound = True
                                                Exit For
                                            End If
                                        End If
                                        att_rst.MoveNext
                                    Wend
                                End If
                            Next ilLoop2
                            If Not ilFound Then
                                'See if in other list: if so move it;
                                ilFound = False
                                For ilLoop2 = 0 To lbcStationsSource.ListCount - 1 Step 1
                                    If ilShttCode2 = lbcStationsSource.ItemData(ilLoop2) Then
                                        SQLQuery = "SELECT attCode, attOnAir, attOffAir, attDropDate, attMulticast FROM att"
                                        SQLQuery = SQLQuery + " WHERE (attShfCode = " & ilShttCode2 & " AND attVefCode = " & ilVefCode & ")"
                                        Set att_rst = gSQLSelectCall(SQLQuery)
                                        While Not att_rst.EOF
                                            If gDateValue(att_rst!attOffAir) <= gDateValue(att_rst!attDropDate) Then
                                                If llLastDate2 = gDateValue(att_rst!attOffAir) Then
                                                    ilFound = True
                                                    ilLoopIndex = ilLoop2
                                                    Exit For
                                                End If
                                            Else
                                                If llLastDate2 = gDateValue(att_rst!attDropDate) Then
                                                    ilFound = True
                                                    ilLoopIndex = ilLoop2
                                                    Exit For
                                                End If
                                            End If
                                            att_rst.MoveNext
                                        Wend
                                    End If
                                Next ilLoop2
                                If ilFound Then
                                    'Move
                                    lbcStationsAlign.AddItem lbcStationsSource.List(ilLoopIndex)
                                    lbcStationsAlign.ItemData(lbcStationsAlign.NewIndex) = lbcStationsSource.ItemData(ilLoopIndex)
                                    lbcStationsSource.RemoveItem ilLoopIndex
                                End If
                            End If
                        Else
                            For ilLoop2 = 0 To lbcStationsAlign.ListCount - 1 Step 1
                                SQLQuery = "SELECT attShfCode, attOffAir, attDropDate FROM att"
                                SQLQuery = SQLQuery + " WHERE (attCode = " & lbcStationsAlign.ItemData(ilLoop2) & ")"
                                Set att_rst2 = gSQLSelectCall(SQLQuery)
                                If Not att_rst2.EOF Then
                                    ilShttCode3 = att_rst2!attshfCode
                                Else
                                    ilShttCode3 = -1
                                End If
                                If ilShttCode2 = ilShttCode3 Then
                                    If gDateValue(att_rst2!attOffAir) <= gDateValue(att_rst2!attDropDate) Then
                                        If llLastDate2 = gDateValue(att_rst2!attOffAir) Then
                                            ilFound = True
                                            Exit For
                                        End If
                                    Else
                                        If llLastDate2 = gDateValue(att_rst2!attDropDate) Then
                                            ilFound = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next ilLoop2
                            If Not ilFound Then
                                'See if in other list: if so move it;
                                ilFound = False
                                For ilLoop2 = 0 To lbcStationsSource.ListCount - 1 Step 1
                                    SQLQuery = "SELECT attShfCode, attOffAir, attDropDate FROM att"
                                    SQLQuery = SQLQuery + " WHERE (attCode = " & lbcStationsSource.ItemData(ilLoop2) & ")"
                                    Set att_rst2 = gSQLSelectCall(SQLQuery)
                                    If Not att_rst2.EOF Then
                                        ilShttCode3 = att_rst2!attshfCode
                                    Else
                                        ilShttCode3 = -1
                                    End If
                                    If ilShttCode2 = ilShttCode3 Then
                                        If gDateValue(att_rst2!attOffAir) <= gDateValue(att_rst2!attDropDate) Then
                                            If llLastDate2 = gDateValue(att_rst2!attOffAir) Then
                                                ilFound = True
                                                ilLoopIndex = ilLoop2
                                                Exit For
                                            End If
                                        Else
                                            If llLastDate2 = gDateValue(att_rst2!attDropDate) Then
                                                ilFound = True
                                                ilLoopIndex = ilLoop2
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next ilLoop2
                                If ilFound Then
                                    'Move
                                    lbcStationsAlign.AddItem lbcStationsSource.List(ilLoopIndex)
                                    lbcStationsAlign.ItemData(lbcStationsAlign.NewIndex) = lbcStationsSource.ItemData(ilLoopIndex)
                                    lbcStationsSource.RemoveItem ilLoopIndex
                                End If
                            End If
                        End If
                    End If
                    shtt_rst.MoveNext
                Wend
            End If
        End If
    Next ilLoop1
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "gAlignMulticastStations"
    Resume Next
ErrHand1:
    gHandleError "AffErrorLog.txt", "gAlignMulticastStations"
    Return
End Sub


Public Sub gAlignAllMulticastStations(lbcStationsAlign As ListBox, lbcStationsSource As ListBox)
    Dim ilLoop1 As Integer
    Dim ilShttCode1 As Integer
    Dim llGroupID As Long
    Dim ilLoop2 As Integer
    Dim ilShttCode2 As Integer
    Dim ilFound As Integer
    Dim shtt_rst As ADODB.Recordset

    On Error GoTo ErrHand
        
    For ilLoop1 = 0 To lbcStationsAlign.ListCount - 1 Step 1
        ilShttCode1 = lbcStationsAlign.ItemData(ilLoop1)
        llGroupID = gGetStaMulticastGroupID(ilShttCode1)
        If llGroupID > 0 Then
            SQLQuery = "Select shttCode FROM shtt where shttMultiCastGroupID = " & llGroupID
            Set shtt_rst = gSQLSelectCall(SQLQuery)
            While Not shtt_rst.EOF
                ilFound = False
                ilShttCode2 = shtt_rst!shttCode
                For ilLoop2 = 0 To lbcStationsAlign.ListCount - 1 Step 1
                    If ilShttCode2 = lbcStationsAlign.ItemData(ilLoop2) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop2
                If Not ilFound Then
                    For ilLoop2 = 0 To lbcStationsSource.ListCount - 1 Step 1
                        If ilShttCode2 = lbcStationsSource.ItemData(ilLoop2) Then
                            'Move
                            lbcStationsAlign.AddItem lbcStationsSource.List(ilLoop2)
                            lbcStationsAlign.ItemData(lbcStationsAlign.NewIndex) = lbcStationsSource.ItemData(ilLoop2)
                            lbcStationsSource.RemoveItem ilLoop2
                            Exit For
                        End If
                    Next ilLoop2
                End If
                shtt_rst.MoveNext
            Wend
        End If
    Next ilLoop1
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "gAlignAllMulticastStations"
    Resume Next
ErrHand1:
    gHandleError "AffErrorLog.txt", "gAlignAllMulticastStations"
    Return
End Sub
