Attribute VB_Name = "RPTCRALLOC"
'*******************************************************************
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
'
' Description:
' This file contains the code for gathering the prepass data for the
' Revenue Allocation Report (12-13-18)
'*******************************************************************
Option Explicit
Option Compare Text
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length

Dim lmCashGrossBilling As Long          'total std month gross cash billing
Dim lmCashNetBilling As Long            'total std month net cash billing
Dim lmCashCommBilling As Long           'total std month comm cash billing
Dim lmTradeGrossBilling As Long         'total std month gross trade billing
Dim lmTradeNetBilling As Long           'total std month net trade billing
Dim lmTradeCommBilling As Long          'total std month comm trade billing
Dim lmSingleCntr As Long                'single contract selection

'accumulated value of all $ distributed
Dim lmMaxCashGross As Long      'cumulative Cash gross $ allocated, maintain as to not exceed max billing
Dim lmMaxCashNet As Long        'cumulative Cash net $ allocated, maintain as to not exceed max billing
Dim lmMaxTradeGross As Long     'cumulative trade gross $ allocated, maintain as to not exceed max billing
Dim lmMaxTradeNet As Long       'cumulative trade net $ allocated, maintain as to not exceed max billing


Dim rst_Temp As ADODB.Recordset         'contract header data
Dim rst_Billing As ADODB.Recordset      'receivables/history data
Dim rst_ActiveAtt As ADODB.Recordset    'all unique station active ageements for month
Dim rst_AST As ADODB.Recordset          'affiliate spot data
Dim rst_StationAtt As ADODB.Recordset       'active station agreements for std month
Dim rst_StationCPTT As ADODB.Recordset      'active station cptts from agreements for std month

Dim tmLongSrchKey0 As LONGKEY0    'Key record image
Dim ilMnfcodes() As Integer     'array of valid vehicle groups to gather
Type ALLOC_CNTLIST
    lCntrNo As Long
    iPctTrade As Integer
End Type
Dim tmContractList() As ALLOC_CNTLIST

Type ALLOC_MKTLIST
    iCode As Integer
    sName As String * 60
    iRank As Integer
    lCashGrimps As Long         '9-1-20 Overflow issue, change to string
    sCashGrimps As String * 15
    lTradeGrimps As Long
End Type
Dim tmMarketList() As ALLOC_MKTLIST

Type ALLOC_REVLIST
    iShttCode As Integer
    sCallLetters As String * 40
    lAudP12Plus As Long
    lCashSpotsOrdered As Long
    lCashSpotsAired As Long
    sCashGrimps As String * 15          '9-1-20
    lCashGrimps As Long
    lCashPctAired As Long
    lCashGrossAlloc As Long
    lCashNet As Long
    lCashComm As Long
    lTradeSpotsOrdered As Long
    lTradeSpotsAired As Long
    lTradeGrimps As Long
    lTradePctAired As Long
    lTradeGrossAlloc As Long
    lTradeNet As Long
    lTradeComm As Long
    bP12AudExists As Boolean
'    sMktName As String * 60
    iMktCode As Integer
    iMktRank As Integer
End Type
Dim tmRevList() As ALLOC_REVLIST

Type STATUSTYPES
    sName As String * 30
    iPledged As Integer '0=Live; 1=Delayed; 2=Not Carried; 3=No Pledge
    iStatus As Integer
End Type
Public tmStatusTypes(0 To 14) As STATUSTYPES
Dim imCurrentMnfGroup As Integer

'*********************************************************************
'       gCreateAlloc()
'       Created 10/1/09
'       Create Alert report prepass file because of decryption of urfname
'
'*********************************************************************
Function gCreateRevAlloc() As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim slSQLQuery As String
    
    Dim llCashGross As Long
    Dim llCashNet As Long
    Dim llTradeGross As Long
    Dim llTradeNet As Long
    Dim llCashGrossHistory As Long
    Dim llCashNetHistory As Long
    Dim llTradeGrossHistory As Long
    Dim llTradeNetHistory As Long
    Dim slStartStd As String
    Dim slEndStd As String
    Dim llStartStd As Long
    Dim llEndStd As Long
    Dim ilUpper As Integer
    Dim blFound As Boolean
    Dim illoop As Integer
    Dim ilShttInx As Integer
    Dim ilLoopOnCnt As Integer
    Dim ilPrevShttCode As Integer
    Dim ilAstPledgeStatus As Integer
    Dim llOrderedCash As Long
    Dim llOrderedTrade As Long
    Dim llAiredCash As Long
    Dim llAiredTrade As Long
    Dim ilPctTrade As Integer
    Dim blItsRvf As Boolean
    Dim ilCntInx As Integer
    Dim ilAdjTradeSpotCount As Integer
    Dim ilAdjCashSpotCount As Integer
    Dim ilStatus As Integer
    Dim llFTCashGrimps As Long
    Dim slFTCashGrimps As String
    Dim llFTTradeGrimps As Long
    Dim slFTTradeGrimps As String

    On Error GoTo gCreateRevAllocErr

    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
     
    'Obtain the standard months start and end dates
    slStr = Trim$(str(igMonthOrQtr)) & "/15/" & Trim$(str(igYear))     'form mm/dd/yy
    slStartStd = gObtainStartStd(slStr)               'obtain std start date for month
    llStartStd = gDateValue(slStartStd)
    slEndStd = gObtainEndStd(slStr)                 'obtain std end date for month
    llEndStd = gDateValue(slEndStd)
    
    lmSingleCntr = Val(RptSelALLOC!edcContract.Text)

    ilRet = gObtainMarkets()
    ilRet = gObtainStations()
    
    '-------------------------------------------------------------------
    'TTP 10376 - Revenue Allocation report: update to use vehicle groups
    Dim blUseVehGroups As Boolean
    Dim ilVehGroupLoop As Integer
    ReDim Preserve ilMnfcodes(0 To 0) 'Veh Group codes
    Dim ilTemp As Integer
    Dim slNameCode As String
    Dim slCode As String
    
    Dim ilSelectedGroup As Integer '2=vefMnfVehGp2, 3=vefMnfVehGp3Mkt, 4=vefMnfVehGp4Fmt, 5=vefMnfVehGp5Rsch, 6=vefMnfVehGp6Sub
    blUseVehGroups = False
    If RptSelALLOC!rbcSortBy(0).Value Then      'sort by Cash/Trade, Market Rank, Station
    ElseIf RptSelALLOC!rbcSortBy(1).Value Then  'sort by Cash/Trade, Market Name, Station
    ElseIf RptSelALLOC!rbcSortBy(2).Value Then  'sort by Cash/Trade, Station
    ElseIf RptSelALLOC!rbcSortBy(3).Value Then  'sory by Vehicle Group, Cash/Trade, Market Rank, Station
        blUseVehGroups = True
    ElseIf RptSelALLOC!rbcSortBy(4).Value Then  'sort by Vehicle Group, Cash/Trade, Market Name, Station
        blUseVehGroups = True
    ElseIf RptSelALLOC!rbcSortBy(5).Value Then  'sort by Vehicle Group, Cash/Trade, Station
        blUseVehGroups = True
    End If
    
    '-------------------------------------------------------------------
    'TTP 10376 - get Selected vehicle Group
    ilSelectedGroup = 0
    If blUseVehGroups Then
        'ilSelectedGroup = '2=vefMnfVehGp2, 3=vefMnfVehGp3Mkt, 4=vefMnfVehGp4Fmt, 5=vefMnfVehGp5Rsch, 6=vefMnfVehGp6Sub
        If RptSelALLOC!cbcSet.ListIndex = 0 Then ilSelectedGroup = 2
        If RptSelALLOC!cbcSet.ListIndex = 1 Then ilSelectedGroup = 3
        If RptSelALLOC!cbcSet.ListIndex = 2 Then ilSelectedGroup = 4
        If RptSelALLOC!cbcSet.ListIndex = 3 Then ilSelectedGroup = 5
        If RptSelALLOC!cbcSet.ListIndex = 4 Then ilSelectedGroup = 6
    End If
    
    '-------------------------------------------------------------------
    'TTP 10376 - Gather selected Vehicle Groups into an Array of mnfCodes
    If blUseVehGroups Then
        For ilTemp = 0 To RptSelALLOC!lbcSelection(0).ListCount - 1 Step 1
            If RptSelALLOC!lbcSelection(0).Selected(ilTemp) Then
                slNameCode = tgSOCodeAA(ilTemp).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilMnfcodes(UBound(ilMnfcodes)) = Val(slCode)
                ReDim Preserve ilMnfcodes(0 To UBound(ilMnfcodes) + 1)
            End If
        Next ilTemp
    End If
    If blUseVehGroups = False Then
        'for this Vehicle Groups to work like the old way, we just need it to run one loop by adding a dummy
        ReDim ilMnfcodes(0 To 1)
    End If
    
    '-------------------------------------------------------------------
    'TTP 10376 - Revenue Allocation report: update to use vehicle groups
    For ilVehGroupLoop = 0 To UBound(ilMnfcodes) - 1
        imCurrentMnfGroup = ilMnfcodes(ilVehGroupLoop)
        slFTCashGrimps = ""
        slFTTradeGrimps = ""
        slFTCashGrimps = ""
        slFTTradeGrimps = ""
        lmCashGrossBilling = 0 'total std month gross cash billing
        lmCashNetBilling = 0 'total std month net cash billing
        lmCashCommBilling = 0 'total std month comm cash billing
        lmTradeGrossBilling = 0 'total std month gross trade billing
        lmTradeNetBilling = 0 'total std month net trade billing
        lmTradeCommBilling = 0 'total std month comm trade billing
        lmMaxCashGross = 0 'cumulative Cash gross $ allocated, maintain as to not exceed max billing
        lmMaxCashNet = 0 'cumulative Cash net $ allocated, maintain as to not exceed max billing
        lmMaxTradeGross = 0 'cumulative trade gross $ allocated, maintain as to not exceed max billing
        lmMaxTradeNet = 0 'cumulative trade net $ allocated, maintain as to not exceed max billing
        
        'markets in mkt code order (gathering using key0)
        ReDim tmMarketList(0 To 0) As ALLOC_MKTLIST
        For illoop = LBound(tgMarkets) To UBound(tgMarkets) - 1
             tmMarketList(illoop).iCode = tgMarkets(illoop).iCode
             tmMarketList(illoop).iRank = tgMarkets(illoop).iRank
             tmMarketList(illoop).sName = tgMarkets(illoop).sName
             tmMarketList(illoop).sCashGrimps = ""                       '9-1-20
             tmMarketList(illoop).lTradeGrimps = 0
             ReDim Preserve tmMarketList(0 To UBound(tmMarketList) + 1) As ALLOC_MKTLIST
        Next illoop
        llCashGross = 0
        llCashNet = 0
        llTradeGross = 0
        llTradeNet = 0
        llCashGrossHistory = 0
        llCashNetHistory = 0
        llTradeGrossHistory = 0
        llTradeNetHistory = 0
        lmMaxCashGross = 0
        lmMaxCashNet = 0
        lmMaxTradeGross = 0
        lmMaxTradeNet = 0
        
        blItsRvf = True
        'Fix TTP 10839 - Single Contract wasnt working
        'mGetBilling slStartStd, slEndStd, blItsRvf, llCashGross, llCashNet, llTradeGross, llTradeNet, ilSelectedGroup, imCurrentMnfGroup
        mGetBilling slStartStd, slEndStd, blItsRvf, llCashGross, llCashNet, llTradeGross, llTradeNet, ilSelectedGroup, imCurrentMnfGroup, lmSingleCntr
        
        'Gather $ for cash and trade gross and net from receivables; include IN and AN only
        blItsRvf = False
        'Fix TTP 10839 - Single Contract wasnt working
        'mGetBilling slStartStd, slEndStd, blItsRvf, llCashGrossHistory, llCashNetHistory, llTradeGrossHistory, llTradeNetHistory, ilSelectedGroup, imCurrentMnfGroup
        mGetBilling slStartStd, slEndStd, blItsRvf, llCashGrossHistory, llCashNetHistory, llTradeGrossHistory, llTradeNetHistory, ilSelectedGroup, imCurrentMnfGroup, lmSingleCntr
        
        'combine receivables and history cash/trad gross/net & comm amts
        lmCashGrossBilling = llCashGross + llCashGrossHistory
        lmCashNetBilling = llCashNet + llCashNetHistory
        lmCashCommBilling = lmCashGrossBilling - lmCashNetBilling
        lmTradeGrossBilling = llTradeGross + llTradeGrossHistory
        lmTradeNetBilling = llTradeNet + llTradeNetHistory
        lmTradeCommBilling = lmTradeGrossBilling - lmTradeNetBilling
         
        'pass formulas to crystal to print the Cash and trade billing amounts at the top of the report
'        If Not gSetFormula("CashGrossBilling", lmCashGrossBilling) Then
'            gCreateRevAlloc = -1
'        End If
'        If Not gSetFormula("CashNetBilling", lmCashNetBilling) Then
'            gCreateRevAlloc = -1
'        End If
'        If Not gSetFormula("TradeGrossBilling", lmTradeGrossBilling) Then
'            gCreateRevAlloc = -1
'        End If
'        If Not gSetFormula("TradeNetBilling", lmTradeNetBilling) Then
'            gCreateRevAlloc = -1
'        End If
        'Fix TTP 10839 - cleaning up
'        slStr = ""
'        If lmSingleCntr > 0 Then
'            slStr = " and rvfcntrno = " & lmSingleCntr
'        End If
        'Create list of unique contract #s from Receivables
        ilUpper = 0
        ReDim tmContractList(0 To 0) As ALLOC_CNTLIST
        slSQLQuery = ""
        slSQLQuery = slSQLQuery & "SELECT distinct rvfCntrno "
        slSQLQuery = slSQLQuery & "FROM rvf_receivables "
        'Filter Results to vehicle Groups?
        Select Case ilSelectedGroup
            Case 0 'Do nothing
            Case 1 'participants (N/A)
            Case 2, 3, 4, 5, 6 'Sub-Totals,Market,Format,Research,Sub-Company.
                slSQLQuery = slSQLQuery & "JOIN VEF_Vehicles on vefCode=rvfAirVefCode "
        End Select
        slSQLQuery = slSQLQuery & "WHERE (rvftrantype = 'IN' or rvftrantype = 'AN') "
        slSQLQuery = slSQLQuery & " and rvftrandate >= '" & Format$(slStartStd, sgSQLDateForm) & "' "
        slSQLQuery = slSQLQuery & " and rvftrandate <= '" & Format$(slEndStd, sgSQLDateForm) & "' "
        'slSQLQuery = slSQLQuery & " " & slStr & " "
        'Filter Results to vehicle Groups?
        Select Case ilSelectedGroup
            Case 0 'Do nothing
            Case 1 'participants (N/A)
            Case 2 'Sub-Totals
                slSQLQuery = slSQLQuery & " and vefMnfVehGp2 = " & imCurrentMnfGroup
            Case 3 'Market
                slSQLQuery = slSQLQuery & " and vefMnfVehGp3Mkt = " & imCurrentMnfGroup
            Case 4 'Format
                slSQLQuery = slSQLQuery & " and vefMnfVehGp4Fmt = " & imCurrentMnfGroup
            Case 5 'Research
                slSQLQuery = slSQLQuery & " and vefMnfVehGp5Rsch = " & imCurrentMnfGroup
            Case 6 'Sub-Company.
                slSQLQuery = slSQLQuery & " and vefMnfVehGp6Sub = " & imCurrentMnfGroup
        End Select
        If lmSingleCntr > 0 Then
            slSQLQuery = slSQLQuery & " and rvfcntrno = " & lmSingleCntr
        End If
        slSQLQuery = slSQLQuery & " ORDER BY rvfcntrno"
        Set rst_Billing = gSQLSelectCall(slSQLQuery)
        While Not rst_Billing.EOF
            tmContractList(ilUpper).lCntrNo = rst_Billing!rvfCntrno
            ilPctTrade = mGetPctTrade(rst_Billing!rvfCntrno)
            tmContractList(ilUpper).iPctTrade = ilPctTrade
            ilUpper = ilUpper + 1
            ReDim Preserve tmContractList(0 To ilUpper) As ALLOC_CNTLIST
            rst_Billing.MoveNext
        Wend
        
        'get unique contract #s from payment history and merge into the contract list from receivables .  Need to know if any of the contracts are trade
'        slStr = ""
        'Fix TTP 10839 - cleaning up
'        If lmSingleCntr > 0 Then
'            slStr = " and phfcntrno = " & lmSingleCntr
'        End If
        slSQLQuery = ""
        slSQLQuery = slSQLQuery & "SELECT distinct phfCntrno "
        slSQLQuery = slSQLQuery & "FROM phf_payment_history "
        'Filter Results to vehicle Groups?
        Select Case ilSelectedGroup
            Case 0 'Do nothing
            Case 1 'participants (N/A)
            Case 2, 3, 4, 5, 6 'Sub-Totals,Market,Format,Research,Sub-Company.
                slSQLQuery = slSQLQuery & "JOIN VEF_Vehicles on vefCode=phfAirVefCode "
        End Select
        slSQLQuery = slSQLQuery & "WHERE (phftrantype = 'IN' or phftrantype = 'AN' or phfTrantype = 'HI') "
        slSQLQuery = slSQLQuery & " and phftrandate >= '" & Format$(slStartStd, sgSQLDateForm) & "' "
        slSQLQuery = slSQLQuery & " and phftrandate <= '" & Format$(slEndStd, sgSQLDateForm) & "' "
        'slSQLQuery = slSQLQuery & slStr
        'Filter Results to vehicle Groups?
        Select Case ilSelectedGroup
            Case 0 'Do nothing
            Case 1 'participants (N/A)
            Case 2 'Sub-Totals
                slSQLQuery = slSQLQuery & " and vefMnfVehGp2 = " & imCurrentMnfGroup
            Case 3 'Market
                slSQLQuery = slSQLQuery & " and vefMnfVehGp3Mkt = " & imCurrentMnfGroup
            Case 4 'Format
                slSQLQuery = slSQLQuery & " and vefMnfVehGp4Fmt = " & imCurrentMnfGroup
            Case 5 'Research
                slSQLQuery = slSQLQuery & " and vefMnfVehGp5Rsch = " & imCurrentMnfGroup
            Case 6 'Sub-Company.
                slSQLQuery = slSQLQuery & " and vefMnfVehGp6Sub = " & imCurrentMnfGroup
        End Select
        If lmSingleCntr > 0 Then
            slSQLQuery = slSQLQuery & " and phfcntrno = " & lmSingleCntr
        End If
        slSQLQuery = slSQLQuery & " Order by phfcntrno"
        Set rst_Billing = gSQLSelectCall(slSQLQuery)
        While Not rst_Billing.EOF
            blFound = False
            For illoop = 0 To UBound(tmContractList) - 1
                If tmContractList(illoop).lCntrNo = rst_Billing!phfcntrno Then
                    ilPctTrade = mGetPctTrade(rst_Billing!phfcntrno)
                    tmContractList(illoop).iPctTrade = ilPctTrade
                    blFound = True
                    Exit For
                End If
            Next illoop
            If Not blFound Then         'insert new contract # notin list yet
                tmContractList(UBound(tmContractList)).lCntrNo = rst_Billing!phfcntrno
                ilPctTrade = mGetPctTrade(rst_Billing!phfcntrno)
                tmContractList(illoop).iPctTrade = ilPctTrade
                ReDim Preserve tmContractList(0 To UBound(tmContractList) + 1) As ALLOC_CNTLIST
            End If
            rst_Billing.MoveNext
        Wend
        
        'sort the contract list by contract # to easily find the reference to determine if trade
        If UBound(tmContractList) - 1 > 0 Then
            ArraySortTyp fnAV(tmContractList(), 0), UBound(tmContractList), 0, LenB(tmContractList(0)), 0, -2, 0
        End If
        
        'Create table of all unique stations to process for output.  Initialze all the fields within the array.
        ilUpper = 0
        ReDim tmRevList(0 To 0) As ALLOC_REVLIST
        slSQLQuery = "Select Distinct shttCode from att Left Outer Join shtt on attShfCode = shttCode Where attOnAir <= '" & Format$(slEndStd, sgSQLDateForm) & "' and attOffAir >= '" & Format$(slStartStd, sgSQLDateForm) & "' and  attDropDate >= '" & Format$(slStartStd, sgSQLDateForm) & "' and shttCode <> '' Order by shttcode"
        Set rst_ActiveAtt = gSQLSelectCall(slSQLQuery)
        While Not rst_ActiveAtt.EOF
            'stations are in sorted internal code order
            ilShttInx = gBinarySearchStation(rst_ActiveAtt!shttCode)
            If ilShttInx = -1 Then
                'missing station
            Else
                'Create new station entry and initialize fields
                tmRevList(ilUpper).iShttCode = rst_ActiveAtt!shttCode
                tmRevList(ilUpper).iMktCode = tgStations(ilShttInx).iMktCode
                tmRevList(ilUpper).bP12AudExists = False
                tmRevList(ilUpper).lAudP12Plus = tgStations(ilShttInx).lAudP12Plus
                If tmRevList(ilUpper).lAudP12Plus > 0 Then          'P12 audience should exist; otherwise list them on report
                    tmRevList(ilUpper).bP12AudExists = True
                End If
                tmRevList(ilUpper).sCallLetters = tgStations(ilShttInx).sCallLetters
                tmRevList(ilUpper).lCashSpotsOrdered = 0
                tmRevList(ilUpper).lCashSpotsAired = 0
                tmRevList(ilUpper).sCashGrimps = ""                     '9-1-20
                tmRevList(ilUpper).lCashPctAired = 0
                tmRevList(ilUpper).lCashGrossAlloc = 0
                tmRevList(ilUpper).lCashNet = 0
                tmRevList(ilUpper).lCashComm = 0
                
                tmRevList(ilUpper).lTradeSpotsOrdered = 0
                tmRevList(ilUpper).lTradeSpotsAired = 0
                tmRevList(ilUpper).lTradeGrimps = 0
                tmRevList(ilUpper).lTradePctAired = 0
                tmRevList(ilUpper).lTradeGrossAlloc = 0
                tmRevList(ilUpper).lTradeNet = 0
                tmRevList(ilUpper).lTradeComm = 0
                
                llOrderedCash = 0
                llOrderedTrade = 0
                llAiredCash = 0
                llAiredTrade = 0
                
        
                'read all the agreements for this station
                slSQLQuery = ""
                'Fix TTP 10839 - query doesnt need to return vehicle table
                'slSQLQuery = slSQLQuery & "Select * from att WITH(INDEX(KEY2)) "
                slSQLQuery = slSQLQuery & "Select att.* from att WITH(INDEX(KEY2)) "
                
                'Fix TTP 10839 - This broke the groups
                'Filter Results to vehicle Groups?
'                Select Case ilSelectedGroup
'                    Case 0 'Do nothing
'                    Case 1 'participants (N/A)
'                    Case 2, 3, 4, 5, 6 'Sub-Totals,Market,Format,Research,Sub-Company.
'                        slSQLQuery = slSQLQuery & "JOIN VEF_Vehicles on vefCode=attVefCode "
'                End Select
                slSQLQuery = slSQLQuery & "Where  attshfcode = " & tmRevList(ilUpper).iShttCode & " "
                slSQLQuery = slSQLQuery & " and attOnAir <= '" & Format$(slEndStd, sgSQLDateForm) & "' "
                slSQLQuery = slSQLQuery & " and attOffAir >= '" & Format$(slStartStd, sgSQLDateForm) & "' "
                slSQLQuery = slSQLQuery & " and  attDropDate >= '" & Format$(slStartStd, sgSQLDateForm) & "'"
                'Fix TTP 10839 - This broke the groups
'                'Filter Results to vehicle Groups?
'                Select Case ilSelectedGroup
'                    Case 0 'Do nothing
'                    Case 1 'participants (N/A)
'                    Case 2 'Sub-Totals
'                        slSQLQuery = slSQLQuery & " and vefMnfVehGp2 = " & imCurrentMnfGroup
'                    Case 3 'Market
'                        slSQLQuery = slSQLQuery & " and vefMnfVehGp3Mkt = " & imCurrentMnfGroup
'                    Case 4 'Format
'                        slSQLQuery = slSQLQuery & " and vefMnfVehGp4Fmt = " & imCurrentMnfGroup
'                    Case 5 'Research
'                        slSQLQuery = slSQLQuery & " and vefMnfVehGp5Rsch = " & imCurrentMnfGroup
'                    Case 6 'Sub-Company.
'                        slSQLQuery = slSQLQuery & " and vefMnfVehGp6Sub = " & imCurrentMnfGroup
'                End Select
                
                Set rst_StationAtt = gSQLSelectCall(slSQLQuery)
                While Not rst_StationAtt.EOF
                    'obtain all the cppts for this station for the month
                    slSQLQuery = ""
                    slSQLQuery = slSQLQuery & "Select cpttPostingStatus,cpttStartDate,cpttatfcode "
                    slSQLQuery = slSQLQuery & "from cptt "
                    slSQLQuery = slSQLQuery & "Where  cpttatfcode = " & rst_StationAtt!attcode & " "
                    slSQLQuery = slSQLQuery & " and cpttStartDate >= '" & Format$(slStartStd, sgSQLDateForm) & "' "
                    slSQLQuery = slSQLQuery & " and cpttStartDate <= '" & Format$(slEndStd, sgSQLDateForm) & "'"
                    Set rst_StationCPTT = gSQLSelectCall(slSQLQuery)
                    While Not rst_StationCPTT.EOF
                        'rbcOrderAir(0) = Use ordered so process all cptts; rbcorderair(1)- use aired so must be posted
                         If (rst_StationCPTT!cpttPostingStatus = 2 And RptSelALLOC!rbcOrderAir(1).Value) Or (RptSelALLOC!rbcOrderAir(0).Value) Then          'use if Using Aired and posted, or Using Ordered take everything
                            'all ast have to be read regardless if the ordered and aired spot counts have been updated in cptt since it must be
                            'determined if a spot is a trade or not
                            'get ast for week of cptt
                            slSQLQuery = ""
                            slSQLQuery = slSQLQuery & "Select AstcpStatus, astStatus,astcntrno,astAirDate,astAtfCode "
                            slSQLQuery = slSQLQuery & "FROM ast "
                            'Fix TTP 10839 - this broke the groups
                            'Filter Results to vehicle Groups?
'                            Select Case ilSelectedGroup
'                                Case 0 'Do nothing
'                                Case 1 'participants (N/A)
'                                Case 2, 3, 4, 5, 6 'Sub-Totals,Market,Format,Research,Sub-Company.
'                                    slSQLQuery = slSQLQuery & "JOIN VEF_Vehicles on vefCode=astVefCode "
'                            End Select
                            slSQLQuery = slSQLQuery & "WHERE astairDate >= '" & Format$(rst_StationCPTT!cpttStartDate, sgSQLDateForm) & "' "
                            slSQLQuery = slSQLQuery & " and astAirDate <= '" & Format$(rst_StationCPTT!cpttStartDate + 6, sgSQLDateForm) & "' "
                            slSQLQuery = slSQLQuery & " and astAtfCode = " & rst_StationCPTT!cpttatfcode
                            'Fix TTP 10839 - this broke the groups
                            'Filter Results to vehicle Groups?
'                            Select Case ilSelectedGroup
'                                Case 0 'Do nothing
'                                Case 1 'participants (N/A)
'                                Case 2 'Sub-Totals
'                                    slSQLQuery = slSQLQuery & " and vefMnfVehGp2 = " & imCurrentMnfGroup
'                                Case 3 'Market
'                                    slSQLQuery = slSQLQuery & " and vefMnfVehGp3Mkt = " & imCurrentMnfGroup
'                                Case 4 'Format
'                                    slSQLQuery = slSQLQuery & " and vefMnfVehGp4Fmt = " & imCurrentMnfGroup
'                                Case 5 'Research
'                                    slSQLQuery = slSQLQuery & " and vefMnfVehGp5Rsch = " & imCurrentMnfGroup
'                                Case 6 'Sub-Company.
'                                    slSQLQuery = slSQLQuery & " and vefMnfVehGp6Sub = " & imCurrentMnfGroup
'                            End Select
                            Set rst_AST = gSQLSelectCall(slSQLQuery)
                            While Not rst_AST.EOF
                                ilCntInx = mBinarySearchCntList(rst_AST!astcntrno)
                                If ilCntInx < 0 Then
                                    'missing contract #, assume no trade
                                    ilPctTrade = 0
                                Else
                                    ilPctTrade = tmContractList(ilCntInx).iPctTrade
                                    If ilPctTrade > 0 Then
                                    ilPctTrade = ilPctTrade
                                    End If
                                End If
                                'calc % of trade spots
                                ilAdjCashSpotCount = 10         'carry the spot counts in tenths (10 = 1.0)
                                ilAdjTradeSpotCount = 0
                                If ilPctTrade > 0 Then      'trade
                                    ilAdjTradeSpotCount = ilPctTrade / 10
                                    ilAdjCashSpotCount = (100 - ilPctTrade) / 10
                                End If
                                'if posted, use it; if not posted, must be using ordered
                                If (rst_AST!AstcpStatus = 1 Or rst_AST!AstcpStatus = 2) And (rst_AST!astcntrno = lmSingleCntr Or lmSingleCntr = 0) Then       '1 = posted, 2 =posted, none aired
                                    ilStatus = mGetAirStatus(rst_AST!astStatus)
                                    'If ilStatus < ASTEXTENDED_MG Or ilStatus = ASTAIR_MISSED_MG_BYPASS Then
                                        '0 = live, 1 =delay
                                        '6 = aired outside pledge, 7 = aired notpledged, 9 delayed, air coml only,10 = live air comlonly
                                        If ilStatus = 0 Or ilStatus = 1 Or ilStatus = 6 Or ilStatus = 7 Or ilStatus = 9 Or ilStatus = 10 Then
                                            llOrderedCash = llOrderedCash + ilAdjCashSpotCount
                                            llOrderedTrade = llOrderedTrade + ilAdjTradeSpotCount
                                            llAiredCash = llAiredCash + ilAdjCashSpotCount
                                            llAiredTrade = llAiredTrade + ilAdjTradeSpotCount
                                        ElseIf ilStatus = ASTEXTENDED_MG Or ilStatus = ASTEXTENDED_REPLACEMENT Or ilStatus = ASTEXTENDED_BONUS Then
                                            llAiredCash = llAiredCash + ilAdjCashSpotCount
                                            llAiredTrade = llAiredTrade + ilAdjTradeSpotCount
                                        ElseIf (ilStatus >= 2 And ilStatus <= 5) Or ilStatus = ASTAIR_MISSED_MG_BYPASS Then     'missed, or missed- bypass mg: count the ordered
                                            llOrderedCash = llOrderedCash + ilAdjCashSpotCount
                                            llOrderedTrade = llOrderedTrade + ilAdjTradeSpotCount
                                        End If
                                Else            'Not Posted.  using ordered , process all as if aired
                                    If (RptSelALLOC!rbcOrderAir(0).Value) And (rst_AST!astcntrno = lmSingleCntr Or lmSingleCntr = 0) Then       'using ordered, match single cnt #
                                        If ilStatus = ASTEXTENDED_MG Or ilStatus = ASTEXTENDED_REPLACEMENT Or ilStatus = ASTEXTENDED_BONUS Then     'aired
                                        Else
                                            llOrderedCash = llOrderedCash + ilAdjCashSpotCount
                                            llOrderedTrade = llOrderedTrade + ilAdjTradeSpotCount
                                        End If
                                    End If
                                End If              'if rstAST!astcpstatus =1
                    
                                rst_AST.MoveNext                'get next spot within the stationcptt
                            Wend
                        End If          'cpttpostingstatus =2
                        rst_StationCPTT.MoveNext            'next cptt for station agreement
                    Wend
        
                rst_StationAtt.MoveNext                     'next vehicle cptt for same staion
            Wend
            End If                          'if shttinx = -1
            
            'udpate the spot counts and station grimps
            mUpdateStationGrimps ilUpper, llOrderedCash, llAiredCash, llOrderedTrade, llAiredTrade
        
            ReDim Preserve tmRevList(0 To ilUpper + 1) As ALLOC_REVLIST
            ilUpper = ilUpper + 1
        
            rst_ActiveAtt.MoveNext          'next unique station
        Wend
        
        'gather grimps by market
        mAccumGrimpsByMkt slFTCashGrimps, slFTTradeGrimps
        
        'calc each stations share of gross impressions
        mCalcStationPct slFTCashGrimps, slFTTradeGrimps
        
        mCalcStationAlloc       'loop thru all stations in list and calculate each stations share of billing
        mUpdateGrf              'loop thru stations are create a cash record and trade record
        
        DoEvents
        If ilVehGroupLoop > 0 Then
           RptSelALLOC.Caption = "Revenue Allocation Report " & Format((ilVehGroupLoop / UBound(ilMnfcodes)) * 100, "###") & "%"
        End If
    Next ilVehGroupLoop
    
    If RptSelALLOC!ckcShowMissingAudience.Value = vbChecked Then mUpdateGrf (True) 'Write stations w/o RSCH
    
    RptSelALLOC.Caption = "Revenue Allocation Report"
    On Error GoTo gCreateRevAllocCloseErr
    'Cleanup
    rst_ActiveAtt.Close
    rst_Temp.Close
    rst_Billing.Close
    rst_AST.Close
    rst_StationAtt.Close
    rst_StationCPTT.Close
    On Error GoTo 0
    Erase tmContractList, tmMarketList, tmRevList
    Exit Function

gCreateRevAllocCloseErr:
    Resume Next

gCreateRevAllocErr:
    gDbg_HandleError "RptCrAlloc.bas: gCreateRevAlloc"
    Exit Function
    
End Function
Private Function mBinarySearchRevList(ilShttCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim ilRet As Integer
    
    On Error GoTo mBinarySearchErr
    
    ilMin = LBound(tmRevList)
    ilMax = UBound(tmRevList)
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilShttCode = tmRevList(ilMiddle).iShttCode Then
            'found the match
            mBinarySearchRevList = ilMiddle
            Exit Function
        ElseIf ilShttCode < tmRevList(ilMiddle).iShttCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchRevList = -1
    Exit Function
mBinarySearchErr:
    mBinarySearchRevList = -1
    Exit Function
End Function

Private Sub mCreateStatustype()
    'Agreement only shows status- 1:; 2:; 5: and 9:
    'All other screens show all the status
    tmStatusTypes(0).sName = "1-Aired Live"        'In Agreement and Pre_Log use 'Air Live'
    tmStatusTypes(0).iPledged = 0
    tmStatusTypes(0).iStatus = 0
    tmStatusTypes(1).sName = "2-Aired Delay B'cast" '"2-Aired In Daypart"  'In Agreement and Pre-Log use 'Air In Daypart'
    tmStatusTypes(1).iPledged = 1
    tmStatusTypes(1).iStatus = 1
    tmStatusTypes(2).sName = "3-Not Aired Tech Diff"
    tmStatusTypes(2).iPledged = 2
    tmStatusTypes(2).iStatus = 2
    tmStatusTypes(3).sName = "4-Not Aired Blackout"
    tmStatusTypes(3).iPledged = 2
    tmStatusTypes(3).iStatus = 3
    tmStatusTypes(4).sName = "5-Not Aired Other"
    tmStatusTypes(4).iPledged = 2
    tmStatusTypes(4).iStatus = 4
    tmStatusTypes(5).sName = "6-Not Aired Product"
    tmStatusTypes(5).iPledged = 2
    tmStatusTypes(5).iStatus = 5
    tmStatusTypes(6).sName = "7-Aired Outside Pledge"  'In Pre-Log use 'Air-Outside Pledge'
    tmStatusTypes(6).iPledged = 3
    tmStatusTypes(6).iStatus = 6
    tmStatusTypes(7).sName = "8-Aired Not Pledged"  'in Pre-Log use 'Air-Not Pledged'
    tmStatusTypes(7).iPledged = 3
    tmStatusTypes(7).iStatus = 7
    'D.S. 11/6/08 remove the "or Aired" from the status 9 description
    'Affiliate Meeting Decisions item 5) f-iv
    'tmStatusTypes(8).sName = "9-Not Carried or Aired"
    tmStatusTypes(8).sName = "9-Not Carried"
    tmStatusTypes(8).iPledged = 2
    tmStatusTypes(8).iStatus = 8
    tmStatusTypes(9).sName = "10-Delay Cmml/Prg"  'In Agreement and Pre-Log use 'Air In Daypart'
    tmStatusTypes(9).iPledged = 1
    tmStatusTypes(9).iStatus = 9
    tmStatusTypes(10).sName = "11-Air Cmml Only"  'In Agreement and Pre-Log use 'Air In Daypart'
    tmStatusTypes(10).iPledged = 1
    tmStatusTypes(10).iStatus = 10
    tmStatusTypes(ASTEXTENDED_MG).sName = "MG"
    tmStatusTypes(ASTEXTENDED_MG).iPledged = 3
    tmStatusTypes(ASTEXTENDED_MG).iStatus = ASTEXTENDED_MG
    tmStatusTypes(ASTEXTENDED_BONUS).sName = "Bonus"
    tmStatusTypes(ASTEXTENDED_BONUS).iPledged = 3
    tmStatusTypes(ASTEXTENDED_BONUS).iStatus = ASTEXTENDED_BONUS
    tmStatusTypes(ASTEXTENDED_REPLACEMENT).sName = "Replacement"
    tmStatusTypes(ASTEXTENDED_REPLACEMENT).iPledged = 3
    tmStatusTypes(ASTEXTENDED_REPLACEMENT).iStatus = ASTEXTENDED_REPLACEMENT
    tmStatusTypes(ASTAIR_MISSED_MG_BYPASS).sName = "15-Missed MG Bypassed"
    tmStatusTypes(ASTAIR_MISSED_MG_BYPASS).iPledged = 2
    tmStatusTypes(ASTAIR_MISSED_MG_BYPASS).iStatus = ASTAIR_MISSED_MG_BYPASS
End Sub
Public Function mGetAirStatus(ilAstStatus) As Integer
    mGetAirStatus = ilAstStatus Mod 100
End Function
'
'       mGetPctTrade -Gather all unique contract #s for the std month requested
'       Maintain the % of trade for each contract so separate cash/trade allocations can be calculated
'       <input>  llCntrNo - 0 if all contracts to process; else single contract #
'       Return - % of trade
Public Function mGetPctTrade(llCntrNo As Long) As Integer
Dim ilPctTrade As Integer
Dim slSQLQuery As String

        ilPctTrade = 0
        slSQLQuery = "Select * from chf_contract_header where chfcntrno = " & llCntrNo & " and chfDelete = 'N' "
        Set rst_Temp = gSQLSelectCall(slSQLQuery)
        If rst_Temp.EOF Then
            'none exists
        Else
            ilPctTrade = rst_Temp!chfPctTrade
        End If
        mGetPctTrade = ilPctTrade
End Function
'
'           mGetBilling -Read either rvf or phf and get the sum total of cash and trade IN & AN transactions.
'           If history, include the HI tranasctions.  Match on transaction date falling between the indicated start/end dates
'           <input> slStartStd - start date to filter transaction date
'                   slEndDate - end date to filter transaction dates
'                   blItsRvf - true if reading RVF, else false to get payment history
'           <output> llCashGross - Summed $ of cash gross transactions included
'                   llNetGross - summed $ of cash net transactions included
'                   llTradeGrss - summed $ of trade gross transactions included
'                   llTradeNet - summed $ of trade net transactions included
'
'  TTP 10376 - Revenue Allocation report: update to use vehicle groups, added optional mnfVehGroup and mnfVehGroupFilter
'Fix TTP 10839 - Single Contract mode wasnt working
Public Sub mGetBilling(slStartStd As String, slEndStd As String, blItsRvf As Boolean, llCashGross As Long, llCashNet As Long, llTradeGross As Long, llTradeNet As Long, Optional mnfVehGroup As Integer = 0, Optional mnfVehGroupFilter = 0, Optional llSingleCntrNo As Long = 0)
    Dim slSQLQuery As String
    If blItsRvf Then
        slSQLQuery = ""
        slSQLQuery = slSQLQuery & "Select Sum(If(rvfCashTrade = 'C', rvfGross, 0)) as SumCashGross, "
        slSQLQuery = slSQLQuery & " Sum(If(rvfCashTrade = 'C', rvfNet, 0)) as SumCashNet, "
        slSQLQuery = slSQLQuery & " Sum(If(rvfCashTrade = 'T', rvfGross, 0)) as SumTradeGross, "
        slSQLQuery = slSQLQuery & " Sum(If(rvfCashTrade = 'T', rvfNet, 0)) as SumTradeNet "
        slSQLQuery = slSQLQuery & "From rvf_Receivables "
        'Filter Results to vehicle Groups?
        Select Case mnfVehGroup
            Case 0 'Do nothing
            Case 1 'participants (N/A)
            Case 2, 3, 4, 5, 6 'Sub-Totals,Market,Format,Research,Sub-Company.
                slSQLQuery = slSQLQuery & "JOIN VEF_Vehicles on vefCode=rvfAirVefCode "
        End Select
        slSQLQuery = slSQLQuery & "Where (rvftrantype = 'IN' or rvftrantype = 'AN') "
        slSQLQuery = slSQLQuery & " and rvfsbfcode = 0 "
        slSQLQuery = slSQLQuery & " and rvftrandate >= '" & Format$(slStartStd, sgSQLDateForm) & "' "
        slSQLQuery = slSQLQuery & " and rvftrandate <= '" & Format$(slEndStd, sgSQLDateForm) & "' "
        'Filter Results to vehicle Groups?
        Select Case mnfVehGroup
            Case 0 'Do nothing
            Case 1 'participants (N/A)
            Case 2 'Sub-Totals
                slSQLQuery = slSQLQuery & " and vefMnfVehGp2 = " & mnfVehGroupFilter
            Case 3 'Market
                slSQLQuery = slSQLQuery & " and vefMnfVehGp3Mkt = " & mnfVehGroupFilter
            Case 4 'Format
                slSQLQuery = slSQLQuery & " and vefMnfVehGp4Fmt = " & mnfVehGroupFilter
            Case 5 'Research
                slSQLQuery = slSQLQuery & " and vefMnfVehGp5Rsch = " & mnfVehGroupFilter
            Case 6 'Sub-Company.
                slSQLQuery = slSQLQuery & " and vefMnfVehGp6Sub = " & mnfVehGroupFilter
        End Select
        'Fix TTP 10839 - Single Contract Mode
        If llSingleCntrNo <> 0 Then
            slSQLQuery = slSQLQuery & " and rvfCntrNo = " & llSingleCntrNo
        End If
        Set rst_Billing = gSQLSelectCall(slSQLQuery)
        If rst_Billing.EOF Then
            'No receivables
        Else
            'convert string $ to long variable
            If Not IsNull(rst_Billing!sumCashNet) Then
                llCashGross = gStrDecToLong(rst_Billing!SumCashGross, 2)
                llCashNet = gStrDecToLong(rst_Billing!sumCashNet, 2)
            End If
            If Not IsNull(rst_Billing!sumTradeNet) Then
                llTradeGross = gStrDecToLong(rst_Billing!SumTradeGross, 2)
                llTradeNet = gStrDecToLong(rst_Billing!sumTradeNet, 2)
            End If
        End If
    Else
        'Gather $ for cash and trade gross and net from payment history; include IN , HI and AN only
        slSQLQuery = ""
        slSQLQuery = slSQLQuery & "Select Sum(If(phfCashTrade = 'C', phfGross, 0)) as SumCashGross, "
        slSQLQuery = slSQLQuery & " Sum(If(phfCashTrade = 'C', phfNet, 0)) as SumCashNet, "
        slSQLQuery = slSQLQuery & " Sum(If(phfCashTrade = 'T', phfGross, 0)) as SumTradeGross, "
        slSQLQuery = slSQLQuery & " Sum(If(phfCashTrade = 'T', phfNet, 0)) as SumTradeNet "
        slSQLQuery = slSQLQuery & "From phf_Payment_History "
        'Filter Results to vehicle Groups?
        Select Case mnfVehGroup
            Case 0 'Do nothing
            Case 1 'participants (N/A)
            Case 2, 3, 4, 5, 6 'Sub-Totals,Market,Format,Research,Sub-Company.
                slSQLQuery = slSQLQuery & "JOIN VEF_Vehicles on vefCode=phfAirVefCode "
        End Select
        
        slSQLQuery = slSQLQuery & "Where (phftrantype = 'IN' or phftrantype = 'AN' or phftrantype = 'HI') "
        slSQLQuery = slSQLQuery & " and phfsbfcode = 0 and phftrandate >= '" & Format$(slStartStd, sgSQLDateForm) & "' "
        slSQLQuery = slSQLQuery & " and phftrandate <= '" & Format$(slEndStd, sgSQLDateForm) & "'"
        
        'Filter Results to vehicle Groups?
        Select Case mnfVehGroup
            Case 0 'Do nothing
            Case 1 'participants (N/A)
            Case 2 'Sub-Totals
                slSQLQuery = slSQLQuery & " and vefMnfVehGp2 = " & mnfVehGroupFilter
            Case 3 'Market
                slSQLQuery = slSQLQuery & " and vefMnfVehGp3Mkt = " & mnfVehGroupFilter
            Case 4 'Format
                slSQLQuery = slSQLQuery & " and vefMnfVehGp4Fmt = " & mnfVehGroupFilter
            Case 5 'Research
                slSQLQuery = slSQLQuery & " and vefMnfVehGp5Rsch = " & mnfVehGroupFilter
            Case 6 'Sub-Company.
                slSQLQuery = slSQLQuery & " and vefMnfVehGp6Sub = " & mnfVehGroupFilter
        End Select
        'Fix TTP 10839 - Single Contract mode
        If llSingleCntrNo <> 0 Then
            slSQLQuery = slSQLQuery & " and phfCntrNo = " & llSingleCntrNo
        End If
        
        Set rst_Billing = gSQLSelectCall(slSQLQuery)
        If rst_Billing.EOF Then
            'No receivables
        Else
            'convert string $ to long variable
            If Not IsNull(rst_Billing!sumCashNet) Then
                llCashGross = gStrDecToLong(rst_Billing!SumCashGross, 2)
                llCashNet = gStrDecToLong(rst_Billing!sumCashNet, 2)
            End If
            If Not IsNull(rst_Billing!sumTradeNet) Then
                llTradeGross = gStrDecToLong(rst_Billing!SumTradeGross, 2)
                llTradeNet = gStrDecToLong(rst_Billing!sumTradeNet, 2)
            End If
        End If
    End If
    'Exit Sub
End Sub
'
'           mBinarySearchCntList - list of contracts has been created to keep track of trade contracts and their % of trade
'           <input> llCntrno - 0 if all contracts processed; else single contract #
Private Function mBinarySearchCntList(llCntrNo As Long) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim ilRet As Integer
    
    On Error GoTo mBinarySearchCntListErr
    
    ilMin = LBound(tmContractList)
    ilMax = UBound(tmContractList)
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If llCntrNo = tmContractList(ilMiddle).lCntrNo Then
            'found the match
            mBinarySearchCntList = ilMiddle
            Exit Function
        ElseIf llCntrNo < tmContractList(ilMiddle).lCntrNo Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchCntList = -1
    Exit Function
mBinarySearchCntListErr:
    mBinarySearchCntList = -1
    Exit Function
End Function
'
'       Accumlate each stations cash & trade grimps by mkt
'       <output> slFTCashGrimps - accumulation of total Cash Grimps required to compute the stations allocation%
'               slFTTradeGrimps - accumulation of total Trade Grimps required to compute the stations allocation%
'
Public Sub mAccumGrimpsByMkt(slFTCashGrimps As String, slFTTradeGrimps As String)
    Dim ilLoopOnStation As Integer
    Dim ilLoopOnMkt As Integer
    Dim slTemp As String
    Dim slRevListCashGrimps As String
    Dim ilNoDecPlaces As Integer
    ilNoDecPlaces = 1
    
    For ilLoopOnStation = LBound(tmRevList) To UBound(tmRevList) - 1
        For ilLoopOnMkt = LBound(tmMarketList) To UBound(tmMarketList) - 1
            If ilLoopOnStation = 76 Then
                ilLoopOnStation = ilLoopOnStation
            End If
            If tmRevList(ilLoopOnStation).iMktCode = tmMarketList(ilLoopOnMkt).iCode Then
                '9-1-20 change the cash grimps to string computations
                'the grimps calculated # spots (held in frctions i.e 1.5 spot = 15)
'                    tmMarketList(ilLoopOnMkt).lCashGrimps = tmMarketList(ilLoopOnMkt).lCashGrimps + (tmRevList(ilLoopOnStation).lCashGrimps / 10)
'                    slTemp = gLongToStrDec(tmRevList(ilLoopOnStation).lCashGrimps,1)
                slRevListCashGrimps = gDivStr(tmRevList(ilLoopOnStation).sCashGrimps, "10")
                tmMarketList(ilLoopOnMkt).sCashGrimps = gAddStr(tmMarketList(ilLoopOnMkt).sCashGrimps, slRevListCashGrimps)
                tmMarketList(ilLoopOnMkt).lTradeGrimps = tmMarketList(ilLoopOnMkt).lTradeGrimps + (tmRevList(ilLoopOnStation).lTradeGrimps / 10)
                slTemp = gFormatStrDec(tmRevList(ilLoopOnStation).sCashGrimps, 1)
'                    slTemp = tmRevList(ilLoopOnStation).sCashGrimps
'                    If Len(slTemp) >= 1 Then
'                        slTemp = Left$(slTemp, Len(slTemp) - ilNoDecPlaces) & "." & right$(slTemp, ilNoDecPlaces)
'                    Else
'                        Do While Len(slTemp) < ilNoDecPlaces
'                            slTemp = "0" & slTemp
'                        Loop
'                        slTemp = "." & slTemp
'                    End If
                slFTCashGrimps = gAddStr(slFTCashGrimps, slTemp)
                slTemp = gLongToStrDec(tmRevList(ilLoopOnStation).lTradeGrimps, 1)
                slFTTradeGrimps = gAddStr(slFTTradeGrimps, slTemp)
                Exit For
            End If
        Next ilLoopOnMkt
    Next ilLoopOnStation
    'Exit Sub
End Sub
'
'       mUpdateStationGrimps - Accumulate the Ordered and Aired, Cash and trade spots, and Audience into Station array
'       <input> ilUpper - index into station to accumulate
'               llOrderedCash - cash ordered spots for station for the month
'               llAiredCash - cash aired spots for month
'               llOrderedTrade - trade ordered spots for station for month
'               llAiredTrade - trade aired spots for station for month
Public Sub mUpdateStationGrimps(ilUpper As Integer, llOrderedCash As Long, llAiredCash As Long, llOrderedTrade As Long, llAiredTrade As Long)

        'spots are carried in fractions due to split cash trade contracts, but saved in whole numbers.  ie. 1.5 spots stored as 15
        tmRevList(ilUpper).lCashSpotsOrdered = tmRevList(ilUpper).lCashSpotsOrdered + llOrderedCash
        tmRevList(ilUpper).lCashSpotsAired = tmRevList(ilUpper).lCashSpotsAired + llAiredCash
        tmRevList(ilUpper).lTradeSpotsOrdered = tmRevList(ilUpper).lTradeSpotsOrdered + llOrderedTrade
        tmRevList(ilUpper).lTradeSpotsAired = tmRevList(ilUpper).lTradeSpotsAired + llAiredTrade
    If ilUpper = 76 Then
    ilUpper = ilUpper
    End If
        If RptSelALLOC!rbcOrderAir(0).Value Then           'use ordered spots to calc the grimps
        '9-1-20 change to string due to overflow
'                tmRevList(ilUpper).lCashGrimps = tmRevList(ilUpper).lCashSpotsOrdered * tmRevList(ilUpper).lAudP12Plus
            tmRevList(ilUpper).sCashGrimps = gMulStr(gLongToStrDec(tmRevList(ilUpper).lCashSpotsOrdered, 0), gLongToStrDec(tmRevList(ilUpper).lAudP12Plus, 0))
            tmRevList(ilUpper).lTradeGrimps = tmRevList(ilUpper).lTradeSpotsOrdered * tmRevList(ilUpper).lAudP12Plus
        Else                                                'use aired spots to calc grimps
'                tmRevList(ilUpper).lCashGrimps = tmRevList(ilUpper).lCashSpotsAired * tmRevList(ilUpper).lAudP12Plus
            tmRevList(ilUpper).sCashGrimps = gMulStr(gLongToStrDec(tmRevList(ilUpper).lCashSpotsAired, 0), gLongToStrDec(tmRevList(ilUpper).lAudP12Plus, 0))
            tmRevList(ilUpper).lTradeGrimps = tmRevList(ilUpper).lTradeSpotsAired * tmRevList(ilUpper).lAudP12Plus
        End If
            
        Exit Sub

End Sub
'
'           mCalcStationPct - Calc the stations % of allocation based on total grimps
'           % allocation = station grimps / total grimps - carry to 4 decimal places
'           <input> slFTCashGrimps - overall Cash Grimps for the month
'                   slFTTradeGrimps - overall Trade Grimps for the month
Public Sub mCalcStationPct(slFTCashGrimps As String, slFTTradeGrimps As String)
    Dim ilLoopOnStation As Integer '
    Dim slCashGrimps As String
'    Dim slFTTradeGrimps as strug
    Dim slTradeGrimps As String
'    Dim slFTCashGrimps As String
    Dim slCashPct As String
    Dim slTradePct As String
    Dim llCashMax As Long
    Dim llTradeMax As Long
    Dim slTemp As String

'        slFTCashGrimps = gLongToStrDec(llFTCashGrimps, 1)
'        slFTTradeGrimps = gLongToStrDec(llFTTradeGrimps, 1)
        For ilLoopOnStation = LBound(tmRevList) To UBound(tmRevList) - 1
        '9-1-20 change to string computation for the cash grimps due to overflow
'            slCashGrimps = gLongToStrDec(tmRevList(ilLoopOnStation).lCashGrimps, 0)
            slCashGrimps = Trim$(tmRevList(ilLoopOnStation).sCashGrimps)
            slCashGrimps = gMulStr(slCashGrimps, "10000")

            'slTradeGrimps = gLongToStrDec((tmRevList(ilLoopOnStation).lTradeGrimps * 100), 1)
            slTradeGrimps = gLongToStrDec(tmRevList(ilLoopOnStation).lTradeGrimps, 0)
            slTradeGrimps = gMulStr(slTradeGrimps, "10000")
            
'            slCashPct = gMulStr(gDivStr(slCashGrimps, slFTCashGrimps), "1000")
'            slTemp = gFormatStrDec(slFTCashGrimps, 0)
            slTemp = Trim$(slFTCashGrimps)
            slCashPct = gMulStr(gDivStr(slCashGrimps, slTemp), "1000")
            slTradePct = gMulStr(gDivStr(slTradeGrimps, slFTTradeGrimps), "1000")
            
            tmRevList(ilLoopOnStation).lCashPctAired = gStrDecToLong(slCashPct, "0")
            tmRevList(ilLoopOnStation).lTradePctAired = gStrDecToLong(slTradePct, "0")
        Next ilLoopOnStation
    Exit Sub
End Sub
'
'       mCalcStationAlloc - calculate the stations $ amount of Cash and trade allocations
'           for gross, net and comm.  Carry out to 4 decimal places for most accuracy.
'           When allocations are distributed, the total $ billing for gross and net
'           will not be exceeded.
'
Public Sub mCalcStationAlloc()

    Dim ilLoopOnStation As Integer
    Dim slPct As String
    Dim slGross As String
    Dim slNet As String
    Dim slComm As String
    Dim slGrossAlloc As String
    Dim slNetAlloc As String
'    Dim llMaxCashGross As Long      'cumulative Cash gross $ allocated, maintain as to not exceed max billing
'    Dim llMaxCashNet As Long        'cumulative Cash net $ allocated, maintain as to not exceed max billing
'    Dim llMaxTradeGross As Long     'cumulative trade gross $ allocated, maintain as to not exceed max billing
'    Dim llMaxTradeNet As Long       'cumulative trade net $ allocated, maintain as to not exceed max billing
    Dim llTempGross As Long
    Dim llTempNet As Long
    
        'Loop thru each station and calculate its gross/net/comm $ for cash and trade
        'carry out to 4 decimal places.
        'Never exceed distributing total $ from billing
        For ilLoopOnStation = LBound(tmRevList) To UBound(tmRevList) - 1
            slGross = gLongToStrDec(lmCashGrossBilling, 2)     'receivables+ history cash gross
            slNet = gLongToStrDec(lmCashNetBilling, 2)         'receivables + history cash gross
            'Calculate the CASH Gross, Net and Commission billing for single station at a time
            
            slPct = gLongToStrDec(tmRevList(ilLoopOnStation).lCashPctAired, 4) 'pct of station allocation
            slGrossAlloc = gMulStr(slGross, slPct)
            slNetAlloc = gMulStr(slNet, slPct)
            llTempGross = gStrDecToLong(slGrossAlloc / 100, 0)
            llTempNet = gStrDecToLong(slNetAlloc / 100, 0)
            If llTempGross + lmMaxCashGross > lmCashGrossBilling Then     'do not exceed the max billing, last one process gets whats left
                llTempGross = lmCashGrossBilling - lmMaxCashGross
            End If
            If llTempNet + lmMaxCashNet > lmCashNetBilling Then     'do not exceed the max billing, last one process gets whats left
                llTempNet = lmCashNetBilling - lmMaxCashNet
            End If
            lmMaxCashGross = lmMaxCashGross + llTempGross            'accum total gross allocations so far
            lmMaxCashNet = lmMaxCashNet + llTempNet                 'accum total net allocations so far
            
            tmRevList(ilLoopOnStation).lCashGrossAlloc = llTempGross
            tmRevList(ilLoopOnStation).lCashNet = llTempNet
            tmRevList(ilLoopOnStation).lCashComm = llTempGross - llTempNet
            
             'Calculate the TRADE Gross, Net and Commission billing for single station at a time
            slGross = gLongToStrDec(lmTradeGrossBilling, 2)     'receivables+ history cash gross
            slNet = gLongToStrDec(lmTradeNetBilling, 2)         'receivables + history cash gross
            slPct = gLongToStrDec(tmRevList(ilLoopOnStation).lTradePctAired, 4)  'pct of station allocation
            slGrossAlloc = gMulStr(slGross, slPct)
            slNetAlloc = gMulStr(slNet, slPct)
            llTempGross = gStrDecToLong(slGrossAlloc / 100, 0)
            llTempNet = gStrDecToLong(slNetAlloc / 100, 0)
            If llTempGross + lmMaxTradeGross > lmTradeGrossBilling Then     'do not exceed the max billing, last one process gets whats left
                llTempGross = lmTradeGrossBilling - lmMaxTradeGross
            End If
            If llTempNet + lmMaxTradeNet > lmTradeNetBilling Then     'do not exceed the max billing, last one process gets whats left
                llTempNet = lmTradeNetBilling - lmMaxTradeNet
            End If
            lmMaxTradeGross = lmMaxTradeGross + llTempGross            'accum total gross allocations so far
            lmMaxTradeNet = lmMaxTradeNet + llTempNet                 'accum total net allocations so far
            
            tmRevList(ilLoopOnStation).lTradeGrossAlloc = llTempGross
            tmRevList(ilLoopOnStation).lTradeNet = llTempNet
            tmRevList(ilLoopOnStation).lTradeComm = llTempGross - llTempNet
           
        Next ilLoopOnStation
        Exit Sub
End Sub

Public Sub mUpdateGrf(Optional WriteMissingStns As Boolean = False)
    Dim ilLoopOnStation As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim slSQLQuery As String
    Dim llTempGross As Long
    Dim llTempNet As Long
    Dim ilLastCashInx As Integer
    Dim ilLastTradeInx As Integer
    Dim ilLastCashStn As Integer
    Dim ilLastTradeStn As Integer
    
    On Error GoTo 0
    'grfGenDate - Generation date for filtering (sGenDate)
    'grfGenTime - Generation time for filtering (sGenTime)
    'grfSofCode - Station code
    'grfBktType - C = Cash, T = Trade, Z = station without aud values (which sort to end of report)
    'grfPer7 - % of Audience allocation (xx.xxxx%)
    'grfPer1 - Cash or Trade Spots ordered
    'grfPer2 - Cash or Trade Spots aired
    'grfper3 - Cash or Trade Grimps
    'grfPer4 - C/T Gross $ Allocation (with pennies)
    'grfPer5 - C/T Comm $ allocation
    'grfPer6 - C/T Net $ allocation
    'TTP 10376 - Revenue Allocation report: update to use vehicle groups (was formula passed in once per report, now written each record)
    'grfPer10 - CashGrossBilling
    'grfPer11 - CashNetBilling
    'grfPer12 - TradeGrossBilling
    'grfPer13 - TradeNetBilling
    'grfPer1Gen -LastCashStation
    'grfPer2Genl -LastTradeStation
    
    'determine the last station with spots to process
    For ilLastCashInx = UBound(tmRevList) - 1 To LBound(tmRevList) Step -1
        If tmRevList(ilLastCashInx).lCashSpotsOrdered > 0 And tmRevList(ilLastCashInx).bP12AudExists Then
            Exit For
        End If
    Next ilLastCashInx
'    If Not gSetFormula("LastCashStation", "'  '") Then
'        llTempGross = llTempGross
'    End If
    For ilLastTradeInx = UBound(tmRevList) - 1 To LBound(tmRevList) Step -1
        If tmRevList(ilLastTradeInx).lTradeSpotsOrdered > 0 And tmRevList(ilLastTradeInx).bP12AudExists Then
            Exit For
        End If
    Next ilLastTradeInx
'    If Not gSetFormula("LastTradeStation", "' '") Then
'        llTempGross = llTempGross
'    End If
    gUnpackDateLong igNowDate(0), igNowDate(1), llDate
    slDate = Format$(llDate, "ddddd")
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    For ilLoopOnStation = LBound(tmRevList) To UBound(tmRevList) - 1
        If tmRevList(ilLoopOnStation).bP12AudExists Then                            'no audience , show station at end of report in a list
            If WriteMissingStns = False Then
                'Cash
                If tmRevList(ilLoopOnStation).lCashSpotsOrdered > 0 Then                'if nothing ordered, then nothing has been returned ; ignore the station
                    If RptSelALLOC!ckcExtraFundsToBal.Value = vbChecked Then
                        If ilLoopOnStation = ilLastCashInx Then                     'processing last station, give them  any remaining funds
                            llTempGross = lmCashGrossBilling - lmMaxCashGross
                            tmRevList(ilLoopOnStation).lCashGrossAlloc = tmRevList(ilLoopOnStation).lCashGrossAlloc + llTempGross
                            llTempNet = lmCashNetBilling - lmMaxCashNet
                            tmRevList(ilLoopOnStation).lCashNet = tmRevList(ilLoopOnStation).lCashNet + llTempNet
                            tmRevList(ilLoopOnStation).lCashComm = tmRevList(ilLoopOnStation).lCashComm + (llTempGross - llTempNet)
                            'TTP 10376 - Revenue Allocation report: update to use vehicle groups
                            ilLastCashStn = tmRevList(ilLoopOnStation).iShttCode
'                            If Not gSetFormula("LastCashStation", "'" & tmRevList(ilLoopOnStation).sCallLetters & "'") Then
'                                llTempGross = llTempGross
'                            End If
                        End If
                    Else
'                        If Not gSetFormula("LastCashStation", "' '") Then
'                            llTempGross = llTempGross
'                        End If
                    End If
                    slSQLQuery = "INSERT INTO GRF_Generic_Report"
                    slSQLQuery = slSQLQuery & "(grfvefCode, "
                    slSQLQuery = slSQLQuery & "grfSofCode, "
                    slSQLQuery = slSQLQuery & "grfBktType, "
                    slSQLQuery = slSQLQuery & "grfPer7, "
                    slSQLQuery = slSQLQuery & "grfPer1, "
                    slSQLQuery = slSQLQuery & "grfPer2, "
                    slSQLQuery = slSQLQuery & "grfGenDesc, "
                    slSQLQuery = slSQLQuery & "grfPer4, "
                    slSQLQuery = slSQLQuery & "grfPer5, "
                    slSQLQuery = slSQLQuery & "grfPer6, "
                    slSQLQuery = slSQLQuery & "grfPer10, " 'TTP 10376 -lmCashGrossBilling (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer11, " 'TTP 10376 -lmCashNetBilling (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer12, " 'TTP 10376 -lmTradeGrossBilling (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer13, " 'TTP 10376 -lmTradeNetBilling (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer1Genl, " 'TTP 10376 -LastCashStation (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer2Genl, " 'TTP 10376 -LastTradeStation (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfGendate, "
                    slSQLQuery = slSQLQuery & "grfGenTime) "
                    slSQLQuery = slSQLQuery & " Values("
                    slSQLQuery = slSQLQuery & imCurrentMnfGroup 'grfvefCode - Vehicle Group MNFCode
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).iShttCode ' grfSofCode
                    slSQLQuery = slSQLQuery & ", 'C'" 'grfBktType
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lCashPctAired 'grfPer7
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lCashSpotsOrdered 'grfPer1
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lCashSpotsAired
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).sCashGrimps
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lCashGrossAlloc
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lCashComm
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lCashNet
                    'TTP 10376 - Revenue Allocation report: update to use vehicle groups
                    slSQLQuery = slSQLQuery & ", " & lmCashGrossBilling
                    slSQLQuery = slSQLQuery & ", " & lmCashNetBilling
                    slSQLQuery = slSQLQuery & ", " & lmTradeGrossBilling
                    slSQLQuery = slSQLQuery & ", " & lmTradeNetBilling
                    slSQLQuery = slSQLQuery & ", " & ilLastCashStn
                    slSQLQuery = slSQLQuery & ", " & ilLastTradeStn
                    
                    slSQLQuery = slSQLQuery & ", '" & Format$(slDate, sgSQLDateForm) & "'"
                    slSQLQuery = slSQLQuery & ", " & lgNowTime & ")"
                    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                        gHandleError "TrafficErrors.txt", "RptcrAlloc: mUpdateGrf"
                    End If
                End If 'Cash
                
                'Trade
                If tmRevList(ilLoopOnStation).lTradeSpotsOrdered > 0 Then                 'if nothing ordered, then nothing has been returned ; ignore the station
                    If RptSelALLOC!ckcExtraFundsToBal.Value = vbChecked Then
                        If ilLoopOnStation = ilLastTradeInx Then                     'processing last station, give them  any remaining funds
                            llTempGross = lmTradeGrossBilling - lmMaxTradeGross
                            tmRevList(ilLoopOnStation).lTradeGrossAlloc = tmRevList(ilLoopOnStation).lTradeGrossAlloc + llTempGross
                            llTempNet = lmTradeNetBilling - lmMaxTradeNet
                            tmRevList(ilLoopOnStation).lTradeNet = tmRevList(ilLoopOnStation).lTradeNet + llTempNet
                            tmRevList(ilLoopOnStation).lTradeComm = tmRevList(ilLoopOnStation).lTradeComm + (llTempGross - llTempNet)
                           
                            'TTP 10376 - Revenue Allocation report: update to use vehicle groups
                            ilLastTradeStn = tmRevList(ilLoopOnStation).iShttCode
'                            If Not gSetFormula("LastTradeStation", "'" & tmRevList(ilLoopOnStation).sCallLetters & "'") Then
'                                llTempGross = llTempGross
'                            End If
                        End If
                    Else
'                        If Not gSetFormula("LastTradeStation", "'  '") Then
'                            llTempGross = llTempGross
'                        End If
                    End If
            
                    slSQLQuery = "INSERT INTO GRF_Generic_Report "
                    slSQLQuery = slSQLQuery & "(grfvefCode, " ''TTP 10376 - Vehicle Group (mnfCode)
                    slSQLQuery = slSQLQuery & "grfSofCode, "
                    slSQLQuery = slSQLQuery & "grfBktType, "
                    slSQLQuery = slSQLQuery & "grfPer7, "
                    slSQLQuery = slSQLQuery & "grfPer1, "
                    slSQLQuery = slSQLQuery & "grfPer2, "
                    slSQLQuery = slSQLQuery & "grfPer3, "
                    slSQLQuery = slSQLQuery & "grfPer4, "
                    slSQLQuery = slSQLQuery & "grfPer5, "
                    slSQLQuery = slSQLQuery & "grfPer6, "
                    slSQLQuery = slSQLQuery & "grfPer10, " 'TTP 10376 -lmCashGrossBilling (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer11, " 'TTP 10376 -lmCashNetBilling (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer12, " 'TTP 10376 -lmTradeGrossBilling (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer13, " 'TTP 10376 -lmTradeNetBilling (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer1Genl, " 'TTP 10376 -LastCashStation (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer2Genl, " 'TTP 10376 -LastTradeStation (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfGendate, "
                    slSQLQuery = slSQLQuery & "grfGenTime) "
                    slSQLQuery = slSQLQuery & " Values("
                    slSQLQuery = slSQLQuery & imCurrentMnfGroup 'Vehicle Group MNFCode
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).iShttCode
                    slSQLQuery = slSQLQuery & ", 'T'"
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lTradePctAired
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lTradeSpotsOrdered
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lTradeSpotsAired
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lTradeGrimps
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lTradeGrossAlloc
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lTradeComm
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lTradeNet
                    'TTP 10376 - Revenue Allocation report: update to use vehicle groups
                    slSQLQuery = slSQLQuery & ", " & lmCashGrossBilling
                    slSQLQuery = slSQLQuery & ", " & lmCashNetBilling
                    slSQLQuery = slSQLQuery & ", " & lmTradeGrossBilling
                    slSQLQuery = slSQLQuery & ", " & lmTradeNetBilling
                    slSQLQuery = slSQLQuery & ", " & ilLastCashStn
                    slSQLQuery = slSQLQuery & ", " & ilLastTradeStn
                    slSQLQuery = slSQLQuery & ", " & "'" & Format$(slDate, sgSQLDateForm) & "'"
                    slSQLQuery = slSQLQuery & ", " & lgNowTime & ")"
                    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                        gHandleError "TrafficErrors.txt", "RptcrAlloc: mUpdateGrf"
                    End If
                End If 'Trade
                
            End If 'WriteMissingStns = False
            
        Else        'station without audience
            If RptSelALLOC!ckcShowMissingAudience.Value = vbChecked Then
                If WriteMissingStns = True Then
                    slSQLQuery = "INSERT INTO GRF_Generic_Report "
                    slSQLQuery = slSQLQuery & "(grfvefCode, "
                    slSQLQuery = slSQLQuery & "grfSofCode, "
                    slSQLQuery = slSQLQuery & "grfBktType, "
                    slSQLQuery = slSQLQuery & "grfPer7,"
                    slSQLQuery = slSQLQuery & "grfPer1, "
                    slSQLQuery = slSQLQuery & "grfPer2, "
                    slSQLQuery = slSQLQuery & "grfGenDesc, "
                    slSQLQuery = slSQLQuery & "grfPer4, "
                    slSQLQuery = slSQLQuery & "grfPer5, "
                    slSQLQuery = slSQLQuery & "grfPer6, "
                    slSQLQuery = slSQLQuery & "grfPer10, " 'TTP 10376 -lmCashGrossBilling (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer11, " 'TTP 10376 -lmCashNetBilling (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer12, " 'TTP 10376 -lmTradeGrossBilling (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer13, " 'TTP 10376 -lmTradeNetBilling (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer1Genl, " 'TTP 10376 -LastCashStation (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfPer2Genl, " 'TTP 10376 -LastTradeStation (was formula passed in once per report)
                    slSQLQuery = slSQLQuery & "grfGendate, "
                    slSQLQuery = slSQLQuery & "grfGenTime) "
                    slSQLQuery = slSQLQuery & " Values("
                    slSQLQuery = slSQLQuery & "-1"  'imCurrentMnfGroup 'Vehicle Group MNFCode
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).iShttCode
                    slSQLQuery = slSQLQuery & ", 'Z'"
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lCashPctAired
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lCashSpotsOrdered
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lCashSpotsAired
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).sCashGrimps
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lCashGrossAlloc
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lCashComm
                    slSQLQuery = slSQLQuery & ", " & tmRevList(ilLoopOnStation).lCashNet
                    'TTP 10376 - Revenue Allocation report: update to use vehicle groups
                    slSQLQuery = slSQLQuery & ", " & lmCashGrossBilling
                    slSQLQuery = slSQLQuery & ", " & lmCashNetBilling
                    slSQLQuery = slSQLQuery & ", " & lmTradeGrossBilling
                    slSQLQuery = slSQLQuery & ", " & lmTradeNetBilling
                    slSQLQuery = slSQLQuery & ", " & ilLastCashStn
                    slSQLQuery = slSQLQuery & ", " & ilLastTradeStn
                    slSQLQuery = slSQLQuery & ", '" & Format$(slDate, sgSQLDateForm) & "'"
                    slSQLQuery = slSQLQuery & ", " & lgNowTime & ")"
                    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                        gHandleError "TrafficErrors.txt", "RptcrAlloc: mUpdateGrf"
                    End If
                End If 'WriteMissingStns
            End If 'ckcShowMissingAudience
        End If 'bP12AudExists
    Next ilLoopOnStation
    Exit Sub
End Sub
