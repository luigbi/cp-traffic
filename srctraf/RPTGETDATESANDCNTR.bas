Attribute VB_Name = "RPTGETDATESANDCNTR"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptsubs.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  hmRaf                         imShfRecLen                   tmShf                     *
'*  hmShf                         tmShfSrchKey                  imMktRecLen               *
'*  tmMkt                         hmMkt                         tmMktSrchKey              *
'*                                                                                        *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  gGetMonthsForYr               gYearForCorpStartMonth                                  *
'******************************************************************************************

Option Explicit
Option Compare Text

Dim tmVsf As VSF

'
'
'               gActivityCntr - Given a contract #, read all versions
'               of the contract header and find the header with
'               the date entered that is between a specified date entered.
'               This is a variation of an Activity report.
'               and allows the user to request a report with the
'               same date and always get the same results.
'
'               <input>  llCntrNo = Contract #
'                        llStart - user entered effective (find contrs entered between start & end date parms)
'                        llEnd  - end date to use to find cntrs entered between
'               <output> hlChf - contract handle
'                        tlChf - contract image to use
'                        Contract code if one found, else zero
'
Function gActivityCntr(llCntrNo As Long, llStart As Long, llEnd As Long, hlChf As Integer, tlChf As CHF) As Long
Dim ilRet As Integer
Dim tlChfSrchKey1 As CHFKEY1
Dim llEntryDate As Long
Dim slTemp As String
    gActivityCntr = 0               'assume no contract found yet

    tlChfSrchKey1.lCntrNo = llCntrNo
    tlChfSrchKey1.iCntRevNo = 32000
    tlChfSrchKey1.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hlChf, tlChf, Len(tlChf), tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlChf.lCntrNo = llCntrNo)
        gUnpackDate tlChf.iOHDDate(0), tlChf.iOHDDate(1), slTemp            'convert date entered
        llEntryDate = gDateValue(slTemp)
        'ignore any contract that isnt an unsched/sch order or hold , or whose entered date is outside the
        'start and end dates passed to this rtn
        If (llEntryDate < llStart) Then
            Exit Function                       'nothing found
        Else
            If (llEntryDate >= llStart And llEntryDate <= llEnd) And (tlChf.sStatus = "G" Or tlChf.sStatus = "H" Or tlChf.sStatus = "O" Or tlChf.sStatus = "N") Then                'found a contract to use
                gActivityCntr = tlChf.lCode
                Exit Function
            End If
        End If
        ilRet = btrGetNext(hlChf, tlChf, Len(tlChf), BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Function
'
'
'           gBuildStartDates - Build array of start dates for the
'           Corporate, standard or weekly calendar
'
'           Created:  11/3/97  D.hosaka
'
'           <input> slStart - 1st date to start the array. If Mon date reqd on weekly,
'                   need to send this routine a Monday date.  It is not backed up to aMonday.
'                   ilCalType - 1 = std, 2 = corp, 3 = week
'                   ilMaxDates - # of dates to build (normally ask for 1 extra
'                               since its an array of start dates.  To gather data
'                               for 12 months need a 13th month to test
'           <output> llStartDates() - array of start dates
'
Sub gBuildStartDates(slStart As String, ilCalType As Integer, ilMaxDates As Integer, llStartDates() As Long)
    Dim slDate As String
    Dim ilLoop As Integer
    Dim llDate As Long
    slDate = slStart
    Debug.Print " - gBuildStartDates(" & ilCalType & "):";
    If ilCalType = 2 Then                       'corp
        For ilLoop = 1 To ilMaxDates Step 1
            slDate = gObtainStartCorp(slDate, True)
            llStartDates(ilLoop) = gDateValue(slDate)
            Debug.Print slDate & ",";
            slDate = gObtainEndCorp(slDate, True)
            llDate = gDateValue(slDate) + 1                      'increment for next month
            slDate = Format$(llDate, "m/d/yy")
        Next ilLoop
    ElseIf ilCalType = 1 Then                           'std
        For ilLoop = 1 To ilMaxDates Step 1
            slDate = gObtainStartStd(slDate)
            llStartDates(ilLoop) = gDateValue(slDate)
            Debug.Print slDate & ",";
            slDate = gObtainEndStd(slDate)
            llDate = gDateValue(slDate) + 1                      'increment for next month
            slDate = Format$(llDate, "m/d/yy")
        Next ilLoop
    ElseIf ilCalType = 4 Then                       'calendar month
       For ilLoop = 1 To ilMaxDates Step 1
            slDate = gObtainStartCal(slDate)
            llStartDates(ilLoop) = gDateValue(slDate)
            Debug.Print slDate & ",";
            slDate = gObtainEndCal(slDate)
            llDate = gDateValue(slDate) + 1                      'increment for next month
            slDate = Format$(llDate, "m/d/yy")
        Next ilLoop

    Else
        For ilLoop = 1 To ilMaxDates Step 1
            llStartDates(ilLoop) = gDateValue(slDate)
            Debug.Print slDate & ",";
            'slDate = Format$(llDate + 7, "m/d/yy")
            slDate = Format$(llStartDates(ilLoop) + 7, "m/d/yy")
        Next ilLoop
    End If
    Debug.Print ""
    Exit Sub
End Sub
'**************************************************************
'*                                                            *
'*      Procedure Name:gGetCorpInxByDate                      *
'*                                                            *
'*             Created:9/21/98       By:D. Hosaka             *
'*            Modified:              By:                      *
'*                                                            *
'*            Comments:Determine index into tmMCof            *
'*              based on a date                               *
'*            <Input> llDate to determine year                *
'*            <output> llYearStartDate - start of corp year found *
'*                     llYearEndDate - end dae of cpr year fnd*
'*                                                            *
'**************************************************************
Function gGetCorpInxByDate(llDate As Long, llYearStartDate As Long, llYearEndDate As Long) As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    ilRet = gObtainCorpCal()
    For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
        'gUnpackDateLong tgMCof(ilLoop).iStartDate(0, 1), tgMCof(ilLoop).iStartDate(1, 1), llYearStartDate         'convert beginning of corp year to long
        'gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 12), tgMCof(ilLoop).iEndDate(1, 12), llYearEndDate         'convert beginning of corp year to long
        gUnpackDateLong tgMCof(ilLoop).iStartDate(0, 0), tgMCof(ilLoop).iStartDate(1, 0), llYearStartDate         'convert beginning of corp year to long
        gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iEndDate(1, 11), llYearEndDate         'convert beginning of corp year to long
        If llDate >= llYearStartDate And llDate <= llYearEndDate Then
        'If tgMCof(ilLoop).iYear = ilYear Then
            gGetCorpInxByDate = ilLoop
            Exit Function
        End If
    Next ilLoop
    gGetCorpInxByDate = -1
    llYearStartDate = 0
    llYearEndDate = 0
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainCntrForOHD              *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain an array of contracts    *
'*                     salesperson code                *
'*                                                     *
'*         7/15/98 Remove filtering on ChfDelete
'*******************************************************
Function gObtainCntrForOHD(frm As Form, slStartDate As String, slEndDate As String, slStatus As String, slCntrType As String, ilHOType As Integer, tlChfAdvtExt() As CHFADVTEXT) As Integer
'
'   ilRet = gObtainCntrForOHD (MainForm, slStartDate, slEndDate, slStatus, slCntrType, ilHOType, tlChfAdvtExt() )
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       slStartDate(I)- Start Date for filtering on OHD Date
'       slEndDate(I)- End Date  for filtering on OHD Date
'       slStatus (I)- chfStatus value or blank
'                         W=Working; D=Dead; C=Completed; I=Incomplete; H=Hold; O=Order
'                         Multiple status can be specified (WDI)
'       slCntrType (I)- chfType value or blank
'                       C=Standard; V=Reservation; T=Remnant; R=DR; Q=PI; S=PSA; M=Promo
'       ilHOType (I)-  1=H or O only; 2=H or O or G or N (if G or N exists show it over H or O);
'                      3=H or O or G or N or W or C or I (if G or N or W or C or I exists show it over H or O)
'                        Note: G or N can't exist at the same time as W or C or I for an order
'                              G or N or W or C or I CntrRev > 0
'       tlChfAdvtExt(O)- Array of contracts of the structure CHFADVTEXT which match selection
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'CHF date/time stamp
    Dim hlChf As Integer        'CHF handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlChf As CHF
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim llTodayDate As Long
    Dim hlVef As Integer        'Vef handle
    Dim tlVef As VEF
    Dim ilVefRecLen As Integer     'Record length
    Dim hlVsf As Integer        'Vsf handle
    'Dim tlVsf As VSF
    Dim ilVsfReclen As Integer     'Record length
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilOffSet As Integer
    Dim ilOper As Integer
    Dim slStr As String
    Dim ilCntUpper As Integer
    Dim ilSlfCode As Integer
    Dim ilCurrent As Integer
    Dim ilTestCntrNo As Integer
    Dim slCntrStatus As String
    Dim slHOStatus As String
    If slStatus = "" Then
        slCntrStatus = "WCIDHO"
    Else
        slCntrStatus = slStatus
    End If
    slHOStatus = ""
    If ilHOType = 1 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "H"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "O"
        End If
    ElseIf ilHOType = 2 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
    ElseIf ilHOType = 3 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
        If (InStr(1, slCntrStatus, "H", 1) <> 0) Or (InStr(1, slCntrStatus, "O", 1) <> 0) Then
            slHOStatus = slHOStatus & "WCI"
        End If
    End If
    slStamp = gFileDateTime(sgDBPath & "Chf.Btr") & slStartDate & slEndDate & Trim$(slCntrStatus) & Trim$(slCntrType) & Trim$(str$(ilCurrent)) & Trim$(str$(ilHOType))
    If sgCntrForDateStamp <> "" Then
        If StrComp(slStamp, sgCntrForDateStamp, 1) = 0 Then
            gObtainCntrForOHD = CP_MSG_NOPOPREQ
            Exit Function
        End If
    End If
    gObtainCntrForOHD = CP_MSG_POPREQ
    ilCurrent = 1   'All
    llTodayDate = gDateValue(gNow())
    'gObtainVehComboList
    hlChf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrOpen):" & "Chf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlChf) 'btrRecordLength(hlChf)  'Get and save record length
    hlVsf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrOpen):" & "Vsf.Btr", frm
    On Error GoTo 0
    ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlSlf)  'Get and save record length
    hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrOpen):" & "Vef.Btr", frm
    On Error GoTo 0
    ilVefRecLen = Len(tlVef) 'btrRecordLength(hlSlf)  'Get and save record length
    sgCntrForDateStamp = slStamp
    'ReDim tlChfAdvtExt(1 To 1) As CHFADVTEXT
    ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
    'ilCntUpper = 1
    ilCntUpper = 0
    ilExtLen = Len(tlChfAdvtExt(ilCntUpper))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlChf) 'Obtain number of records
    btrExtClear hlChf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlChf, tlChf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
        ilRet = btrClose(hlChf)
        On Error GoTo gObtainCntrForOHDErr
        gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrReset):" & "Chf.Btr", frm
        On Error GoTo 0
        btrDestroy hlChf
        ilRet = btrClose(hlVsf)
        On Error GoTo gObtainCntrForOHDErr
        gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrReset):" & "Vsf.Btr", frm
        On Error GoTo 0
        btrDestroy hlVsf
        Exit Function
    Else
        On Error GoTo gObtainCntrForOHDErr
        gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrGetFirst):" & "Chf.Btr", frm
        On Error GoTo 0
    End If
    Call btrExtSetBounds(hlChf, llNoRec, -1, "UC", "CHFADVTEXT", CHFADVTEXTPK) 'Set extract limits (all records)
    ilSlfCode = tgUrf(0).iSlfCode
    ' chfEndDate >= InputStartDate And chfStartDate <= InputEndDate
    gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    'ilOffset = gFieldOffset("Chf", "ChfEndDate")
    ilOffSet = gFieldOffset("Chf", "ChfOHDDate")
    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    'ilOffset = gFieldOffset("Chf", "ChfStartDate")
    ilOffSet = gFieldOffset("Chf", "ChfOHDDate")
    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    'tlCharTypeBuff.sType = "Y"
    'ilOffset = gFieldOffset("Chf", "ChfDelete")
    'If (slCntrStatus = "") And (slCntrType = "") Then
    If (slCntrType = "") Then
        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        'ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
    Else
        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        'ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
        'If slCntrStatus <> "" Then
        '    ilOper = BTRV_EXT_OR
        '    slStr = slCntrStatus
        '    Do While slStr <> ""
        '        If Len(slStr) = 1 Then
        '            If slCntrType <> "" Then
        '                ilOper = BTRV_EXT_AND
        '            Else
        '                ilOper = BTRV_EXT_LAST_TERM
        '            End If
        '        End If
        '        tlCharTypeBuff.sType = Left$(slStr, 1)
        '        ilOffset = gFieldOffset("Chf", "ChfStatus")
        '        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1)
        '        slStr = Mid$(slStr, 2)
        '    Loop
        'End If
        If slCntrType <> "" Then
            ilOper = BTRV_EXT_OR
            slStr = slCntrType
            Do While slStr <> ""
                If Len(slStr) = 1 Then
                    ilOper = BTRV_EXT_LAST_TERM
                End If
                tlCharTypeBuff.sType = Left$(slStr, 1)
                ilOffSet = gFieldOffset("Chf", "ChfType")
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1)
                slStr = Mid$(slStr, 2)
            Loop
        End If
    End If
    ilOffSet = gFieldOffset("Chf", "ChfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract iCode field
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfCntrNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract Contract number
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfExtRevNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfCntRevNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfType")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfAdfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfProduct")
    ilRet = btrExtAddField(hlChf, ilOffSet, 35) 'Extract Product
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfAgfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfSlfCode1")
    ilRet = btrExtAddField(hlChf, ilOffSet, 20) 'Extract salesperson code
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfMnfDemo1")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract salesperson code
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfCxfInt")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfPropVer")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract end date
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfStatus")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfMnfPotnType")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract SellNet
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfStartDate")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfEndDate")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract end date
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfVefCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfSifCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0

     '8-21-05 add pct of trade to array
    ilOffSet = gFieldOffset("Chf", "ChfPctTrade")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'pct trade
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    
    '7/12/10
    ilOffSet = gFieldOffset("Chf", "ChfCBSOrder")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0

    '2/24/12
    ilOffSet = gFieldOffset("Chf", "ChfBillCycle")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0

    'ilRet = btrExtGetNextExt(hlChf)    'Extract record
    ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo gObtainCntrForOHDErr
        gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrExtGetNextExt):" & "Chf.Btr", frm
        On Error GoTo 0
        ilExtLen = Len(tlChfAdvtExt(ilCntUpper))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilFound = True
            ilTestCntrNo = False
            'For Proposal CntRevNo = 0; For Orders CntRevNo >= 0 (for W, C, I CntRevNo > 0)
            If (tlChfAdvtExt(ilCntUpper).iCntRevNo = 0) And ((tlChfAdvtExt(ilCntUpper).sStatus <> "H") And (tlChfAdvtExt(ilCntUpper).sStatus <> "O") And (tlChfAdvtExt(ilCntUpper).sStatus <> "G") And (tlChfAdvtExt(ilCntUpper).sStatus <> "N")) Then  'Proposal
                If (InStr(1, slCntrStatus, tlChfAdvtExt(ilCntUpper).sStatus) = 0) Then
                    ilFound = False
                End If
            Else    'Order
                If (InStr(1, slHOStatus, tlChfAdvtExt(ilCntUpper).sStatus) <> 0) Then
                    If (ilHOType = 2) Or (ilHOType = 3) Then
                        ilTestCntrNo = True
                    End If
                Else
                    ilFound = False
                End If
            End If
            If ilFound Then
                ilFound = gTestChfAdvtExt(frm, ilSlfCode, tlChfAdvtExt(ilCntUpper), hlVsf, ilCurrent)
            End If
            If ilFound Then
                ilFound = False
                If ilTestCntrNo Then
                    'For ilLoop = 1 To ilCntUpper - 1 Step 1
                    For ilLoop = 0 To ilCntUpper - 1 Step 1
                        If tlChfAdvtExt(ilLoop).lCntrNo = tlChfAdvtExt(ilCntUpper).lCntrNo Then
                            If tlChfAdvtExt(ilLoop).iCntRevNo < tlChfAdvtExt(ilCntUpper).iCntRevNo Then
                                tlChfAdvtExt(ilLoop) = tlChfAdvtExt(ilCntUpper)
                            End If
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                End If
                If Not ilFound Then
                    ilCntUpper = ilCntUpper + 1
                    'ReDim Preserve tlChfAdvtExt(1 To ilCntUpper) As CHFADVTEXT
                    ReDim Preserve tlChfAdvtExt(0 To ilCntUpper) As CHFADVTEXT
                End If
            End If
            ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    ilRet = btrClose(hlChf)
    On Error GoTo gObtainCntrForOHDErr
    gBtrvErrorMsg ilRet, "gObtainCntrForOHD (btrReset):" & "Chf.Btr", frm
    On Error GoTo 0
    btrDestroy hlChf
    Exit Function
gObtainCntrForOHDErr:
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    ilRet = btrClose(hlChf)
    btrDestroy hlChf
    gDbg_HandleError "RptSubs: gObtainCntrForOHD"
'    gObtainCntrForOHD = CP_MSG_NOSHOW
'    Exit Function
End Function
'
'
'               gPaceCntr - Given a contract #, read all versions
'               of the contract header and find the header with
'               the date entered that is not greater than the
'               user entered effective date.  So is called Pacing
'               and allows the user to request a report with the
'               same date and always get the same results.
'
'               <input>  llCntrNo = Contract #
'                        llWeekTYStart - user entered effective to pace from
'               <output> hlChf - contract handle
'                        tlChf - contract image to use
'                        Contract code if one found, else zero
'
Function gPaceCntr(llCntrNo As Long, llWeekTYStart As Long, hlChf As Integer, tlChf As CHF) As Long
Dim ilRet As Integer
Dim tlChfSrchKey1 As CHFKEY1
Dim llEntryDate As Long
Dim slTemp As String
    gPaceCntr = 0               'assume no contract found yet

    tlChfSrchKey1.lCntrNo = llCntrNo
    tlChfSrchKey1.iCntRevNo = 32000
    tlChfSrchKey1.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hlChf, tlChf, Len(tlChf), tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlChf.lCntrNo = llCntrNo)
        gUnpackDate tlChf.iOHDDate(0), tlChf.iOHDDate(1), slTemp            'convert date entered
        llEntryDate = gDateValue(slTemp)
        'ignore any contract that isnt an unsched/sch order or hold , or whose entered date is greater than the
        'effective pacing date
        If (llEntryDate <= llWeekTYStart) And (tlChf.sStatus = "G" Or tlChf.sStatus = "H" Or tlChf.sStatus = "O" Or tlChf.sStatus = "N") Then                'found a contract to use
            gPaceCntr = tlChf.lCode
            Exit Function
        End If
        ilRet = btrGetNext(hlChf, tlChf, Len(tlChf), BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Function
'****************************************************************
'*                                                              *
'*      Procedure Name:gCntrForOrigOHD                          *
'*                                                              *
'*             Created:8/4/02     By:D. Hosaka                  *
'*             copied from gcontrforActiveOHD  changed          *
'*             to use original entry date, not latest mod date  *
'*                                                              *
'*            Comments:Obtain an array of contracts based       *
'*            on the original entry date whose cnt start/       *
'*            end dates span the active dates passed            *
'****************************************************************
Function gCntrForOrigOHD(frm As Form, slAStartDate As String, slAEndDate As String, slOHDStartDate As String, slOHDEndDate As String, slStatus As String, slCntrType As String, ilHOType As Integer, tlChfAdvtExt() As CHFADVTEXT) As Integer
'
'   ilRet = gCntrForOrigOHD (MainForm, slAStartDate, slAEndDate, slOHDStartDate, slOHDEndDate, slStatus, slCntrType, ilHOType, tlChfAdvtExt() )
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       slAStartDate(I)- Active Start Date
'       slAEndDate(I)- Active End Date
'       slOHDStartDate(I)- OHD Start Date
'       slOHDEndDate(I)- OHD End Date
'       slStatus (I)- chfStatus value or blank
'                         W=Working; D=Dead; C=Completed; I=Incomplete; H=Hold; O=Order
'                         Multiple status can be specified (WDI)
'       slCntrType (I)- chfType value or blank
'                       C=Standard; V=Reservation; T=Remnant; R=DR; Q=PI; S=PSA; M=Promo
'       ilHOType (I)-  1=H or O only; 2=H or O or G or N (if G or N exists show it over H or O);
'                      3=H or O or G or N or W or C or I (if G or N or W or C or I exists show it over H or O)
'                        Note: G or N can't exist at the same time as W or C or I for an order
'                              G or N or W or C or I CntrRev > 0
'                      4=Return all matching contracts not deleted
'       tlChfAdvtExt(O)- Array of contracts of the structure CHFADVTEXT which match selection
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'CHF date/time stamp
    Dim hlChf As Integer        'CHF handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlChf As CHF
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim llTodayDate As Long
    Dim slDate As String
    Dim hlVef As Integer        'Vef handle
    Dim tlVef As VEF
    Dim ilVefRecLen As Integer     'Record length
    Dim hlVsf As Integer        'Vsf handle
    'Dim tlVsf As VSF
    Dim ilVsfReclen As Integer     'Record length
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilOffSet As Integer
    Dim ilOper As Integer
    Dim slStr As String
    Dim ilCntUpper As Integer
    Dim ilSlfCode As Integer
    Dim ilCurrent As Integer
    Dim ilTestCntrNo As Integer
    Dim slCntrStatus As String
    Dim slHOStatus As String

    If slStatus = "" Then
        slCntrStatus = "WCIDHO"
    Else
        slCntrStatus = slStatus
    End If
    slHOStatus = ""
    If ilHOType = 1 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "H"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "O"
        End If
    ElseIf ilHOType = 2 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
    ElseIf ilHOType = 3 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
        If (InStr(1, slCntrStatus, "H", 1) <> 0) Or (InStr(1, slCntrStatus, "O", 1) <> 0) Then
            slHOStatus = slHOStatus & "WCI"
        End If
    ElseIf ilHOType = 4 Then
        slHOStatus = slCntrStatus
    End If

    slStamp = gFileDateTime(sgDBPath & "Chf.Btr") & slAStartDate & slAEndDate & slOHDStartDate & slOHDEndDate & Trim$(slCntrStatus) & Trim$(slCntrType) & Trim$(str$(ilCurrent)) & Trim$(str$(ilHOType))
    If sgCntrForDateStamp <> "" Then
        If StrComp(slStamp, sgCntrForDateStamp, 1) = 0 Then
            gCntrForOrigOHD = CP_MSG_NOPOPREQ
            Exit Function
        End If
    End If
    gCntrForOrigOHD = CP_MSG_POPREQ
    ilCurrent = 1   'All
    llTodayDate = gDateValue(gNow())
    'gObtainVehComboList
    hlChf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrOpen):" & "Chf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlChf) 'btrRecordLength(hlChf)  'Get and save record length
    hlVsf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrOpen):" & "Vsf.Btr", frm
    On Error GoTo 0
    ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlSlf)  'Get and save record length
    hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrOpen):" & "Vef.Btr", frm
    On Error GoTo 0
    ilVefRecLen = Len(tlVef) 'btrRecordLength(hlSlf)  'Get and save record length
    sgCntrForDateStamp = slStamp
    'ReDim tlChfAdvtExt(1 To 1) As CHFADVTEXT
    ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
    'ilCntUpper = 1
    ilCntUpper = 0
    ilExtLen = Len(tlChfAdvtExt(ilCntUpper))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlChf) 'Obtain number of records
    btrExtClear hlChf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlChf, tlChf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
        ilRet = btrClose(hlChf)
        On Error GoTo gCntrForOrigOHDErr
        gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrReset):" & "Chf.Btr", frm
        On Error GoTo 0
        btrDestroy hlChf
        ilRet = btrClose(hlVsf)
        On Error GoTo gCntrForOrigOHDErr
        gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrReset):" & "Vsf.Btr", frm
        On Error GoTo 0
        btrDestroy hlVsf
        Exit Function
    Else
        On Error GoTo gCntrForOrigOHDErr
        gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrGetFirst):" & "Chf.Btr", frm
        On Error GoTo 0
    End If
    'Call btrExtSetBounds(hlChf, llNoRec, -1, "UC") 'Set extract limits (all records)
    Call btrExtSetBounds(hlChf, llNoRec, -1, "UC", "CHFADVTEXT", CHFADVTEXTPK) 'Set extract limits (all records)

    ilSlfCode = tgUrf(0).iSlfCode
    If (tgUrf(0).iGroupNo > 0) Then     'And (tgUrf(0).iSlfCode <= 0) Then
        ilRet = gObtainUrf()
        ilRet = gObtainSalesperson()
    End If
    ' chfEndDate >= InputStartDate And chfStartDate <= InputEndDate
    If slAStartDate = "" Then
        slDate = "1/1/1970"
    Else
        slDate = slAStartDate
    End If
    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Chf", "ChfEndDate")
    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    If slAEndDate = "" Then
        slDate = "12/31/2069"
    Else
        slDate = slAEndDate
    End If
    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Chf", "ChfStartDate")
    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    If slOHDStartDate = "" Then
        slDate = "1/1/1970"
    Else
        slDate = slOHDStartDate
    End If
    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Chf", "ChfOHDDate")
    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    If slOHDEndDate = "" Then
        slDate = "12/31/2069"
    Else
        slDate = slOHDEndDate
    End If
    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Chf", "ChfOHDDate")
    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

    'tlCharTypeBuff.sType = "Y"
    'ilOffset = gFieldOffset("Chf", "ChfDelete")

    '8-1-02 must be original version entered
    tlIntTypeBuff.iType = 0     'revision must be 0, orig order
    ilOffSet = gFieldOffset("Chf", "ChfCntRevNo")

    If (slCntrType = "") Then
        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 1)
    Else
        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 1)
        If slCntrType <> "" Then
            ilOper = BTRV_EXT_OR
            slStr = slCntrType
            Do While slStr <> ""
                If Len(slStr) = 1 Then
                    ilOper = BTRV_EXT_LAST_TERM
                End If
                tlCharTypeBuff.sType = Left$(slStr, 1)
                ilOffSet = gFieldOffset("Chf", "ChfType")
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1)
                slStr = Mid$(slStr, 2)
            Loop
        End If
    End If
    ilOffSet = gFieldOffset("Chf", "ChfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract iCode field
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfCntrNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract Contract number
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfExtRevNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfCntRevNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfType")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfAdfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfProduct")
    ilRet = btrExtAddField(hlChf, ilOffSet, 35) 'Extract Product
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfAgfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfSlfCode1")
    ilRet = btrExtAddField(hlChf, ilOffSet, 20) 'Extract salesperson code
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfMnfDemo1")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract salesperson code
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfCxfInt")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfPropVer")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract end date
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfStatus")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfMnfPotnType")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract SellNet
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfStartDate")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfEndDate")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract end date
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfVefCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfSifCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0

     '8-21-05 add pct of trade to array
    ilOffSet = gFieldOffset("Chf", "ChfPctTrade")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'pct trade
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    
    '7/12/10
    ilOffSet = gFieldOffset("Chf", "ChfCBSOrder")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0

    '2/24/12
    ilOffSet = gFieldOffset("Chf", "ChfBillCycle")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0

    'ilRet = btrExtGetNextExt(hlChf)    'Extract record
    ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo gCntrForOrigOHDErr
        gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrExtGetNextExt):" & "Chf.Btr", frm
        On Error GoTo 0
        ilExtLen = Len(tlChfAdvtExt(ilCntUpper))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilFound = True
            ilTestCntrNo = False
            'For Proposal CntRevNo = 0; For Orders CntRevNo >= 0 (for W, C, I CntRevNo > 0)
            If (tlChfAdvtExt(ilCntUpper).iCntRevNo = 0) And ((tlChfAdvtExt(ilCntUpper).sStatus <> "H") And (tlChfAdvtExt(ilCntUpper).sStatus <> "O") And (tlChfAdvtExt(ilCntUpper).sStatus <> "G") And (tlChfAdvtExt(ilCntUpper).sStatus <> "N")) Then  'Proposal
                If (InStr(1, slCntrStatus, tlChfAdvtExt(ilCntUpper).sStatus) = 0) Then
                    ilFound = False
                End If
            Else    'Order
                If (InStr(1, slHOStatus, tlChfAdvtExt(ilCntUpper).sStatus) <> 0) Then
                    If (ilHOType = 2) Or (ilHOType = 3) Then
                        ilTestCntrNo = True
                    End If
                Else
                    ilFound = False
                End If
            End If
            If ilFound Then
                ilFound = gTestChfAdvtExt(frm, ilSlfCode, tlChfAdvtExt(ilCntUpper), hlVsf, ilCurrent)
            End If
            If ilFound Then
                ilFound = False
                If ilTestCntrNo Then
                    'For ilLoop = 1 To ilCntUpper - 1 Step 1
                    For ilLoop = 0 To ilCntUpper - 1 Step 1
                        If tlChfAdvtExt(ilLoop).lCntrNo = tlChfAdvtExt(ilCntUpper).lCntrNo Then
                            If tlChfAdvtExt(ilLoop).iCntRevNo < tlChfAdvtExt(ilCntUpper).iCntRevNo Then
                                tlChfAdvtExt(ilLoop) = tlChfAdvtExt(ilCntUpper)
                            End If
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                End If
                If Not ilFound Then
                    ilCntUpper = ilCntUpper + 1
                    'ReDim Preserve tlChfAdvtExt(1 To ilCntUpper) As CHFADVTEXT
                    ReDim Preserve tlChfAdvtExt(0 To ilCntUpper) As CHFADVTEXT
                End If
            End If
            ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    ilRet = btrClose(hlChf)
    On Error GoTo gCntrForOrigOHDErr
    gBtrvErrorMsg ilRet, "gCntrForOrigOHD (btrReset):" & "Chf.Btr", frm
    On Error GoTo 0
    btrDestroy hlChf
    Exit Function
gCntrForOrigOHDErr:
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    ilRet = btrClose(hlChf)
    btrDestroy hlChf
    gDbg_HandleError "RptSubs: gCntrForOrigOHD"
    gCntrForOrigOHD = CP_MSG_NOSHOW
    Exit Function
End Function
'
'           gBuildCntTypesForAll - test the user to see if allowed to see the type of contract
'           (including psa/promos).  gBuildCntTypes always excludes psa/promos
'
Public Function gBuildCntTypesForAll()
Dim slCntrType As String
    slCntrType = "C"                        'everyone gets to orders
    If tgUrf(0).sResvType <> "H" Then
        slCntrType = slCntrType & "V"
    End If
    If tgUrf(0).sRemType <> "H" Then
        slCntrType = slCntrType & "T"
    End If
    If tgUrf(0).sDRType <> "H" Then
        slCntrType = slCntrType & "R"
    End If
    If tgUrf(0).sPIType <> "H" Then
        slCntrType = slCntrType & "Q"
    End If
    If tgUrf(0).sPSAType <> "H" Then
        slCntrType = slCntrType & "S"
    End If
    If tgUrf(0).sPromoType <> "H" Then
        slCntrType = slCntrType & "M"
    End If
    If slCntrType = "CVTRQSM" Then
        slCntrType = ""                 'all types
    End If
    gBuildCntTypesForAll = slCntrType
End Function
'
'           gSetupCntTypesForGet - setup the selected contract types for retrieval
'           gObtainCntForDate routine
'           <input> valid contract header types to include (holds, order, standard, PI, etc)
'           <output> slCntrType = string of cnttypes to send to gObtainCntForDate (standard, PI, etc)
'                    slStatus = string of contract statuses to send to gobtaincntrForDate (holds, orders)
'
Public Sub gSetupCntTypesForGet(tlCntTypes As CNTTYPES, slCntrType As String, slCntrStatus As String)
    slCntrStatus = ""                 'statuses: hold, order, unsch hold, uns order
    slCntrStatus = ""
    If tlCntTypes.iHold = True Then     'incl holds and uns holds
        slCntrStatus = "HG"
    End If
   
    If tlCntTypes.iOrder = True Then  'incl order and uns oeswe
        slCntrStatus = slCntrStatus & "ON"
    End If

    slCntrType = ""
    If tlCntTypes.iStandard = True Then      'std
        slCntrType = "C"
    End If
    If tlCntTypes.iReserv = True Then      'resv
        slCntrType = slCntrType & "V"
    End If
    If tlCntTypes.iRemnant = True Then      'remnant
        slCntrType = slCntrType & "T"
    End If
    If tlCntTypes.iDR = True Then      'DR
        slCntrType = slCntrType & "R"
    End If
    If tlCntTypes.iPI = True Then      'PI
        slCntrType = slCntrType & "Q"
    End If
    If tlCntTypes.iPSA = True Then      'PSA
        slCntrType = slCntrType & "S"
    End If
    If tlCntTypes.iPromo = True Then      'Promo
        slCntrType = slCntrType & "M"
    End If
    If slCntrType = "CVTRQSM" Then          'all types: PI, DR, etc.  except PSA(p) and Promo(m)
        slCntrType = ""                     'blank out string for "All"
    End If

End Sub
