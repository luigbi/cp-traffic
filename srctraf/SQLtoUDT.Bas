Attribute VB_Name = "SQLtoUDTSUBS"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SQLtoUDT.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the transfer of SQL to UDT
Option Explicit
Option Compare Text

Sub gVefSQLtoUDT(ilVefCode As Integer, tlVef As VEF)

    Dim rst As ADODB.Recordset
    Dim slSQLQuery As String

    slSQLQuery = "Select * from VEF_Vehicles where vefCode = " & ilVefCode
    Set rst = cnn.Execute(slSQLQuery)
    If rst.EOF Then
        rst.Close
        tlVef.iCode = 0
        Exit Sub
    End If

    tlVef.iCode = rst!vefCode
    tlVef.sName = rst!VEFNAME
    tlVef.sAddr(0) = rst!vefAddr1
    tlVef.sAddr(1) = rst!vefAddr2
    tlVef.sAddr(2) = rst!vefAddr3
    tlVef.sPhone = rst!vefPhone
    tlVef.sFax = rst!vefFax
    tlVef.sUnused1 = rst!vefUnused1
    tlVef.sDialPos = rst!vefDialPos
    tlVef.lPvfCode = rst!vefPvfCode
    tlVef.iReallDnfCode = rst!vefReallDnfCode
    tlVef.sUpdateRVF(0) = rst!vefUpdateRvf1
    tlVef.sUpdateRVF(1) = rst!vefUpdateRvf2
    tlVef.sUpdateRVF(2) = rst!vefUpdateRvf3
    tlVef.sUpdateRVF(3) = rst!vefUpdateRvf4
    tlVef.sUpdateRVF(4) = rst!vefUpdateRvf5
    tlVef.sUpdateRVF(5) = rst!vefUpdateRvf6
    tlVef.sUpdateRVF(6) = rst!vefUpdateRvf7
    tlVef.sUpdateRVF(7) = rst!vefUpdateRvf8
    tlVef.iCombineVefCode = rst!vefCombineVefCode
    tlVef.iMnfHubCode = rst!vefMnfHubCode
    tlVef.iTrfCode = rst!vefTrfCode
    tlVef.sType = rst!vefType
    tlVef.sCodeStn = rst!vefCodeStn
    tlVef.iVefCode = rst!vefvefCode
    tlVef.iOwnerMnfCode = rst!vefOwnerMnfCode
    tlVef.iProdPct(0) = rst!vefProdPct
    tlVef.iProdPct(1) = rst!vefProdPct2
    tlVef.iProdPct(2) = rst!vefProdPct3
    tlVef.iProdPct(3) = rst!vefProdPct4
    tlVef.iProdPct(4) = rst!vefProdPct5
    tlVef.iProdPct(5) = rst!vefProdPct6
    tlVef.iProdPct(6) = rst!vefProdPct7
    tlVef.iProdPct(7) = rst!vefProdPct8
    tlVef.sState = rst!vefState
    tlVef.iMnfGroup(0) = rst!vefmnfGroup
    tlVef.iMnfGroup(1) = rst!vefmnfGroup2
    tlVef.iMnfGroup(2) = rst!vefmnfGroup3
    tlVef.iMnfGroup(3) = rst!vefmnfGroup4
    tlVef.iMnfGroup(4) = rst!vefmnfGroup5
    tlVef.iMnfGroup(5) = rst!vefmnfGroup6
    tlVef.iMnfGroup(6) = rst!vefmnfGroup7
    tlVef.iMnfGroup(7) = rst!vefmnfGroup8
    tlVef.iSort = rst!vefSort
    tlVef.iDnfCode = rst!vefdnfCode
    tlVef.iMnfDemo = rst!vefmnfDemo
    tlVef.iMnfSSCode(0) = rst!vefmnfSSCode1
    tlVef.iMnfSSCode(1) = rst!vefmnfSSCode2
    tlVef.iMnfSSCode(2) = rst!vefmnfSSCode3
    tlVef.iMnfSSCode(3) = rst!vefmnfSSCode4
    tlVef.iMnfSSCode(4) = rst!vefmnfSSCode5
    tlVef.iMnfSSCode(5) = rst!vefmnfSSCode6
    tlVef.iMnfSSCode(6) = rst!vefmnfSSCode7
    tlVef.iMnfSSCode(7) = rst!vefmnfSSCode8
    tlVef.sUnused3 = rst!vefUnused3
    tlVef.lVsfCode = rst!vefVsfCode
    tlVef.lRateAud = rst!vefRateAud
    tlVef.lCPPCPM = rst!vefCPPCPM
    tlVef.lYearAvails = rst!vefYearAvails
    tlVef.iPctSellout = rst!vefPctSellout
    tlVef.iMnfVehGp2 = rst!vefMnfVehGp2
    tlVef.iMnfVehGp3Mkt = rst!vefMnfVehGp3Mkt
    tlVef.iMnfVehGp4Fmt = rst!vefMnfVehGp4Fmt
    tlVef.iMnfVehGp5Rsch = rst!vefMnfVehGp5Rsch
    tlVef.iMnfVehGp6Sub = rst!vefMnfVehGp6Sub
    tlVef.iNrfCode = rst!vefNrfCode
    tlVef.iSSMnfCode = rst!vefSSMnfCode
    tlVef.sStdPrice = rst!vefStdPrice
    tlVef.sStdInvTime = rst!vefStdInvTime
    tlVef.sStdAlter = rst!vefStdAlter
    tlVef.iStdIndex = rst!vefStdIndex
    tlVef.sStdAlterName = rst!vefStdAlterName
    tlVef.iRemoteID = rst!vefRemoteID
    tlVef.iAutoCode = rst!vefAutoCode
    tlVef.sExtUpdateRvf(0) = rst!vefExtUpdateRvf1
    tlVef.sExtUpdateRvf(1) = rst!vefExtUpdateRvf2
    tlVef.sExtUpdateRvf(2) = rst!vefExtUpdateRvf3
    tlVef.sExtUpdateRvf(3) = rst!vefExtUpdateRvf4
    tlVef.sExtUpdateRvf(4) = rst!vefExtUpdateRvf5
    tlVef.sExtUpdateRvf(5) = rst!vefExtUpdateRvf6
    tlVef.sExtUpdateRvf(6) = rst!vefExtUpdateRvf7
    tlVef.sExtUpdateRvf(7) = rst!vefExtUpdateRvf8
    tlVef.sStdSelCriteria = rst!vefStdSelCriteria
    tlVef.sStdOverrideFlag = rst!vefStdOverrideFlag
    tlVef.sContact = rst!vefContact
    rst.Close
End Sub

Sub gAgfSQLtoUDT(ilAgfCode As Integer, tlAgf As AGF)
    Dim slStr As String
    Dim rst As ADODB.Recordset
    Dim slSQLQuery As String

    slSQLQuery = "Select * from AGF_Agencies where agfCode = " & ilAgfCode
    Set rst = cnn.Execute(slSQLQuery)
    If rst.EOF Then
        rst.Close
        tlAgf.iCode = 0
        Exit Sub
    End If

    tlAgf.iCode = rst!agfCode
    tlAgf.sName = rst!agfName
    tlAgf.sAbbr = rst!agfAbbr
    tlAgf.sCityID = rst!agfCity
    tlAgf.iComm = rst!agfComm
    tlAgf.iSlfCode = rst!agfslfCode
    tlAgf.sBuyer = rst!agfBuyer
    tlAgf.sCodeRep = rst!agfCodeRep
    tlAgf.sCodeStn = rst!agfCodeStn
    tlAgf.sCreditRestr = rst!agfCreditRestr
    tlAgf.lCreditLimit = rst!agfCreditLimit
    tlAgf.sPaymRating = rst!agfPaymRating
    tlAgf.sShowISCI = rst!agfISCI
    tlAgf.iMnfSort = rst!agfmnfSort
    tlAgf.sCntrAddr(0) = rst!agfCntrAddr1
    tlAgf.sCntrAddr(1) = rst!agfCntrAddr2
    tlAgf.sCntrAddr(2) = rst!agfCntrAddr3
    tlAgf.sBillAddr(0) = rst!agfBillAddr1
    tlAgf.sBillAddr(1) = rst!agfBillAddr2
    tlAgf.sBillAddr(2) = rst!agfBillAddr3
    tlAgf.iArfLkCode = rst!agfarfLkCode
    tlAgf.sPhone = rst!agfPhone
    tlAgf.sFax = rst!agfFax
    tlAgf.iArfCntrCode = rst!agfarfCntrCode
    tlAgf.iArfInvCode = rst!agfarfInvCode
    tlAgf.sCntrPrtSz = rst!agfCntrPrtSz
    tlAgf.iTrfCode = rst!agfTrfCode
    tlAgf.sCrdApp = rst!agfCrdApp
    tlAgf.sCrdRtg = rst!agfCrdRtg
    tlAgf.iPnfBuyer = rst!agfpnfBuyer
    tlAgf.iPnfPay = rst!agfpnfPay
    tlAgf.iPct90 = rst!agfPct90
    tlAgf.sCurrAR = rst!agfCurrAR
    tlAgf.sUnbilled = rst!agfUnbilled
    tlAgf.sHiCredit = rst!agfHiCredit
    tlAgf.sTotalGross = rst!agfTotalGross
    slStr = Format$(rst!agfDateEntrd, sgShowDateForm)
    gPackDate slStr, tlAgf.iDateEntrd(0), tlAgf.iDateEntrd(1)
    tlAgf.iNSFChks = rst!agfNSFChks
    slStr = Format$(rst!agfDateLstInv, sgShowDateForm)
    gPackDate slStr, tlAgf.iDateLstInv(0), tlAgf.iDateLstInv(1)
    slStr = Format$(rst!agfDateLstPaym, sgShowDateForm)
    gPackDate slStr, tlAgf.iDateLstPaym(0), tlAgf.iDateLstPaym(1)
    tlAgf.iAvgToPay = rst!agfAvgToPay
    tlAgf.iLstToPay = rst!agfLstToPay
    tlAgf.iNoInvPd = rst!agfNoInvPd
    tlAgf.iMerge = rst!agfMerge
    tlAgf.iUrfCode = rst!agfurfCode
    tlAgf.sState = rst!agfState
    slStr = Format$(rst!agfCrdAppDate, sgShowDateForm)
    gPackDate slStr, tlAgf.iCrdAppDate(0), tlAgf.iCrdAppDate(1)
    slStr = Format$(rst!agfCrdAppTime, "h:mm:ssAM/PM")
    gPackTime slStr, tlAgf.iCrdAppTime(0), tlAgf.iCrdAppTime(1)
    tlAgf.sPkInvShow = rst!agfPkInvShow
    tlAgf.iRemoteID = rst!agfRemoteID
    tlAgf.iAutoCode = rst!agfAutoCode
    slStr = Format$(rst!agfSyncDate, sgShowDateForm)
    gPackDate slStr, tlAgf.iSyncDate(0), tlAgf.iSyncDate(1)
    slStr = Format$(rst!agfSyncTime, "h:mm:ssAM/PM")
    gPackTime slStr, tlAgf.iSyncTime(0), tlAgf.iSyncTime(1)
    tlAgf.iSourceID = rst!agfSourceID
    tlAgf.sCntrExptForm = rst!agfCntrExptForm
    tlAgf.iMnfInvTerms = rst!agfMnfInvTerms
    tlAgf.sXMPProposalBand = rst!agfXMLProposalBand
    tlAgf.s1or2DigitRating = rst!agf1or2DigitRating
    tlAgf.sXMLCallLetters = rst!agfXMLCallLetters
    tlAgf.sXMLDates = rst!agfXMLDates
    tlAgf.sUnused = rst!agfUnused
    rst.Close
End Sub

