Attribute VB_Name = "ProposalSubs"
Public Function gDefaultCHF() As CHF
    
    Dim slDate As String
    Dim c As Integer
    Dim tlChf As CHF
    
    With tlChf
        .lCode = 0
        .lCntrNo = 0                                    '*set this from site
        .sAgyEstNo = ""                                 '*set this
        .lVefCode = 0                                   'might have to set this
        .iExtRevNo = 0
        .iXMLSentUrfCode = 0
        gPackDate Now, .iOHDDate(0), .iOHDDate(1)
        gPackTime Now, .iOHDTime(0), .iOHDTime(1)
        .iCntRevNo = 0
        gPackDate Now, .iPropDate(0), .iPropDate(1)
        gPackTime Now, .iPropTime(0), .iPropTime(1)
        .sType = "C"                                  'standard Contract
        .iAdfCode = 0                                   '*set this field
        .sProduct = ""                                  '*set this field
        .iAgfCode = 0                                'set this field
        For c = 1 To 9                               'Talk to Jim as to who he wants as the salesman
            .iSlfCode(c) = 0
        Next c
        For c = 1 To 9
            .lComm(c) = 0
        Next c
        .lComm(0) = 1000000
        .iMnfComp(0) = 0
        .iMnfComp(1) = 0
        .iMnfExcl(0) = 0
        .iMnfExcl(1) = 0
        .sBuyer = ""                                 '*set this field
        .sPhone = ""                                 'set buyer phone number - phone num is packed
        .iEDSSentUrfCode = 0
        .iPctBudget = 0
        For c = 0 To 4
            .iMnfRevSet(c) = 0
        Next c
        .sCppCpm = "N"                               'set based on tab picked
        .iMerchPct = 0
        .iPromoPct = 0
        For c = 1 To 9
            .iSlspCommPct(c) = 0
        Next c
        .iSlspCommPct(0) = 10000
        For c = 0 To 3
            .iMnfDemo(c) = 0                            '*set mnfDemo(0)  the demo chosen
        Next c
        For c = 0 To 3
            .lTarget(c) = 0
        Next c
        .iPctTrade = 0
        .iRcfCode = 0                                   '*set to rate card used in the computations
        .lInputGross = 0                            'set the gross, ttl $ of aired spots
        .sBillCycle = "S"
        .sInvGp = "A"
        '.iPctTag = 0
        .sAdServerDefined = "N"
        .sPodSpotDefined = "N"
        .lCxfCode = 0
        .lCxfChgR = 0
        .lCxfInt = 0                                'set this or cxfCode to the comment sent from the web
        .lCxfMerch = 0
        .lCxfProm = 0
        .lCxfCanc = 0
        .iPropVer = 1
        .sMGMiss = "G"
        .sStatus = "W"                               'save as W
        .sTitle = ""
        For c = 0 To 9
            .iMnfSubCmpy(c) = 0
        Next c
        gPackDate "", .iDtNeed(0), .iDtNeed(1)
        .iMnfBus = 0
        .lGuar = 0
        .iEDSSentExtRevNo = 0
        For c = 0 To 6                             'replace with Ubound and LBound for safety, all loops
            .iMnfCmpy(c) = 0
        Next c
        For c = 0 To 6
            .iCmpyPct(c) = 0
        Next c
        .sResvNew = "N"
        .lChfCode = 0
        .iMnfPotnType = 0
        .sPrint = "N"
        .sDiscrep = "N"
        .sNewBus = ""                              'not sure on this one
        .sSchStatus = ""                           'set to P,
        .sAgyCTrade = "N"
        gPackDate "", .iStartDate(0), .iStartDate(1)    '*set date range of all lines
        gPackDate "", .iEndDate(0), .iEndDate(1)        '*set date range of all lines
        .imnfSeg = 0
        .iUrfCode = 0                              'set this field
        .sCBSOrder = "N"
        .iHdChg = 0
        gPackDate "1/1/1970", .iEDSSentDate(0), .iEDSSentDate(1)
        gPackTime "12am", .iEDSSentTime(0), .iEDSSentTime(1)
        .sSource = "P"
        .iXMLSentExtRevNo = 0
        gPackDate "", .iXMLSentDate(0), .iXMLSentDate(1)
        .sDelete = "N"
        .lSifCode = 0
        gPackTime "12am", .iXMLSentTime(0), .iXMLSentTime(1)
        .lAirTimeGross = 0                         'must be set = ttl gross from above
        .sInstallDefined = "N"
        .sRepDBID = ""
        .sNRProcessed = "Y"
        .lEffCode = 0
        .sHideDemo = "N"
        .iPnfBuyer = 0                             'set to buyer record
        .lSpotChfCode = 0
        .sNoAssigned = "Y"
        .lGrImp = 0                                'maybe set this
        .lGRP = 0                                  'maybe set this
        .sNTRDefined = "N"
        .lNTRGross = 0
    End With
    gDefaultCHF = tlChf
End Function
Public Function gDefaultCLF() As CLF
    Dim tlClf As CLF
    Dim c As Integer
    Dim slDate As String
    
    slDate = gNow()
    With tlClf
        .lChfCode = 0                             'set this field
        .iLine = 0                                'set this field
        .iCntRevNo = 0
        .iVefCode = 0                             'set this field
        .iRpfCode = 0
        .sBB = "N"
        .sExtra = "N"
        .sPgmTime = ""
        .iBreak = 0
        .iPosition = 0
        gPackTime "", .iStartTime(0), .iStartTime(1)
        gPackTime "", .iEndTime(0), .iEndTime(1)
        .iNoGames = 0
        .iSpotsOrdered = 0
        .iSpotsBooked = 0
        .iSpotsWrite = 0
        .sCntPct = ""    ' set to 1 for weekly or 0 for daily based on the tab
        .iLen = 0                                'set this field
        .sPreempt = "P"
        gPackTime "", .iPrefStartTime(0), .iPrefStartTime(1)
        gPackTime "", .iPrefEndTime(0), .iPrefEndTime(1)
        For c = 0 To 6
            .sPrefDays(c) = ""
        Next c
        .iPctAllocation = 0
        .lAcquisitionCost = 0
        .sSoloAvail = "N"
        .sLiveCopy = ""
        .sSportsByWeek = ""
        .iEDSSentExtRevNo = 0
        .iXMLSentExtRevNo = 0
        .sLineChgd = "Y"
        .sHideCBS = "N"
        .sPriceType = ""
        .iPriority = -1
        .lCxfCode = 0
        .sSchStatus = "P"
        .iUrfCode = 0                            'set this field
        .sDelete = "N"
        gPackDate Now, .iEntryDate(0), .iEntryDate(1)
        gPackTime Now, .iEntryTime(0), .iEntryTime(1)
        .iMnfDemo = 0                            'set this field
        .lCPM = 0                                'set this field
        .lCPP = 0                                'set this field
        .lGrImp = 0                              'set this field
        .iDnfCode = 0                            'set this field
        .iRdfCode = 0                            'set this field
        .iPropVer = 1
        gPackDate "", .iStartDate(0), .iStartDate(1)  'set this for the line only
        gPackDate "", .iEndDate(0), .iEndDate(1)      'set this for the line only
        .iMnfSocEco = 0
        .sType = "S"
        .iPkLineNo = 0                           'set this for hidden lines
        .iAdvtSepFlag = 0
        .iBBOpenLen = 0
        .iBBCloseLen = 0
        .sOV2DefinedBits = Chr(0)
        .lCode = 0
        .lghfcode = 0
        .sGameLayout = ""
        .lRafCode = 0
        .sACT1LineupCode = ""
        .sACT1StoredTime = ""
        .sACT1StoredSpots = ""
        .sACT1StoreClearPct = ""
        .sACT1DaypartFilter = ""
        .sUnused = ""
    End With
    gDefaultCLF = tlClf
End Function
Public Function gDefaultCFF() As CFF
    Dim tlCff As CFF
    Dim c As Integer
    
    With tlCff
        .lChfCode = 0                            'set this field
        .iClfLine = 0                            'set this field
        .iCntRevNo = 0
        gPackDate "", .iStartDate(0), .iStartDate(1)   'set this field
        gPackDate "", .iEndDate(0), .iEndDate(1)       'set this field
        .iSpotsWk = 0                            'daily = 0, weekly = num spots per week; if price tab get from web
        For c = 0 To 6
            .iDay(c) = 0                         'set all of these for weekly 1 or daily use the num spots per day based
        Next c                                   'on the daypart days only valid days
        iXSpotsWk = 0
        For c = 0 To 6
            .sXDay(c) = ""   '
        Next c
        .sDelete = "N"
        .lActPrice = 0                           'set this field actual spot price
        .lPropPrice = 0                          'set this same as  above proposed ratecard price
        .lBBPrice = 0
        .sPriceType = "T"
        .iPropVer = 1
        .lAdjPrice = 0
        .lCode = 0
        .sUnused = ""
    End With
    gDefaultCFF = tlCff
End Function

Public Function gDefaultCXF() As CXF

    Dim tlCxf As CXF
    Dim c As Integer
    
    With tlCxf
        .lCode = 0                                                  'Internal code number for
                                                                    ' Comment-Contract
        .sComType = "O"                                                   ' Comment Type(O=Other; L=Contract
                                                                        ' Line; C=Cancellation; R=Change
                                                                        ' Reason; M=Merchandising;
                                                                        ' P=Promotion; I=Internal;
                                                                        ' N=Personnel; D=Invoice
                                                                        ' Disclaimer)
        .sShProp = "Y"                                                  ' Show on Proposal Y/N
        .sShOrder = "N"                                                 ' Show on Order Y/N
        .sShSpot = "N"                                                  ' Show on Log Y/N
        .sShInv = "N"                                                   ' Show on Invoices
        .iRemoteID = 0                                                 ' Unique ID = Remote ID + AutoCode
        .lAutoCode = 0                                              ' Unique ID = Remote ID + AutoCode
        '.iSyncDate(0 To 1)                                         ' Sync Date (from Central or
        gPackDate "", .iSyncDate(0), .iSyncDate(1)                  ' set this field
        'iSyncTime(0 To 1)                                          ' Sync Time (from Central or
        gPackTime "12:00:00AM", .iSyncTime(0), .iSyncTime(1)
                                                                    ' Remote)
        .iSourceID = 0                                              ' Remote User ID (avoid resend to
                                                                    ' sender)
        .sShInsertion = "N"                                              ' Show on Insertion Order Y/N
        .sUnused = ""
        .sComment = ""                                              ' Last bytes after the comment
                                                                    ' must be 0
    End With
    gDefaultCXF = tlCxf
End Function


Public Function gDefaultADF(Optional slCrditApproval As String = "A") As String

    Dim slSql As String
    
    gDefaultADF = ""
    slSql = "Insert into ADF_Advertisers("
    slSql = slSql & "adfCode, "
    slSql = slSql & "adfName, "
    slSql = slSql & "adfAbbr, "
    slSql = slSql & "adfProd, "
    slSql = slSql & "adfslfCode, "
    slSql = slSql & "adfagfCode, "
    slSql = slSql & "adfmnfComp1, "
    slSql = slSql & "adfmnfComp2, "
    slSql = slSql & "adfmnfExcl1, "
    slSql = slSql & "adfmnfExcl2, "
    slSql = slSql & "adfCppCpm, "
    slSql = slSql & "adfmnfDemo1, "
    slSql = slSql & "adfmnfDemo2, "
    slSql = slSql & "adfmnfDemo3, "
    slSql = slSql & "adfmnfDemo4, "
    slSql = slSql & "adfTarget1, "
    slSql = slSql & "adfTarget2, "
    slSql = slSql & "adfTarget3, "
    slSql = slSql & "adfTarget4, "
    slSql = slSql & "adfCreditRestr, "
    slSql = slSql & "adfCreditLimit, "
    slSql = slSql & "adfPaymRating, "
    slSql = slSql & "adfISCI, "
    slSql = slSql & "adfmnfSort, "
    slSql = slSql & "adfBilAgyDir, "
    slSql = slSql & "adfarfLkCode, "
    slSql = slSql & "adfarfContrCode, "
    slSql = slSql & "adfarfInvCode, "
    slSql = slSql & "adfCntrPrtSz, "
    slSql = slSql & "adfTrfCode, "
    slSql = slSql & "adfCrdApp, "
    slSql = slSql & "adfpnfBuyer, "
    slSql = slSql & "adfpnfPay, "
    slSql = slSql & "adfPct90, "
    slSql = slSql & "adfCurrAR, "
    slSql = slSql & "adfUnBilled, "
    slSql = slSql & "adfHiCredit, "
    slSql = slSql & "adfTotalGross, "
    slSql = slSql & "adfDateEntrd, "
    slSql = slSql & "adfNSFChks, "
    slSql = slSql & "adfAvgTopay, "
    slSql = slSql & "adfLstToPay, "
    slSql = slSql & "adfNoInvPd, "
    slSql = slSql & "adfNewBus, "
    slSql = slSql & "adfMerge, "
    slSql = slSql & "adfurfCode, "
    slSql = slSql & "adfState, "
    slSql = slSql & "adfCrdApptime, "
    slSql = slSql & "adfPkInvShow, "
    slSql = slSql & "adfGuar, "
    slSql = slSql & "adfLastYearNew, "
    slSql = slSql & "adfLastMonthNew, "
    slSql = slSql & "adfRateOnInv, "
    slSql = slSql & "adfmnfBus, "
    slSql = slSql & "adfUnused1, "
    slSql = slSql & "adfAllowRepMG, "
    slSql = slSql & "adfBonusOnInv, "
    slSql = slSql & "adfRepInvGen, "
    slSql = slSql & "adfMnfInvTerms, "
    slSql = slSql & "adfPolitical) "
    
    slSql = slSql & "Values ("
    
    slSql = slSql & "Replace" & ", "                                        'Code
    slSql = slSql & "'" & gFixQuote(sgAdfname) & "', "              'Name
    slSql = slSql & "'" & gFixQuote(sgAdfAbbrv) & "', "             'Abbrv
    slSql = slSql & "'" & gFixQuote(sgProdName) & "', "             'Prod
    slSql = slSql & igSlfCode & ", "                                'Salesperson
    slSql = slSql & igAgfCode & ", "                                'Agency Code
    slSql = slSql & 0 & ", "                                        'mnfComp1
    slSql = slSql & 0 & ", "                                        'mnfComp2
    slSql = slSql & 0 & ", "                                        'mnfExcl1
    slSql = slSql & 0 & ", "                                        'mnfExcl2
    slSql = slSql & "'" & "N" & "' ,"                               'CppCpm
    slSql = slSql & 0 & ", "                                        'mnfDemo1
    slSql = slSql & 0 & ", "                                        'mnfDemo2
    slSql = slSql & 0 & ", "                                        'mnfDemo3
    slSql = slSql & 0 & ", "                                        'mnfDemo4
    slSql = slSql & 0 & ", "                                        'mnfTarget1
    slSql = slSql & 0 & ", "                                        'mnfTarget2
    slSql = slSql & 0 & ", "                                        'mnfTarget3
    slSql = slSql & 0 & ", "                                        'mnfTarget4
    slSql = slSql & "'" & "N" & "' ,"                               'CreditRestr
    slSql = slSql & 0 & ", "                                        'CreditLimit
    slSql = slSql & "'" & "1" & "', "                               'PaymRating
    slSql = slSql & "'" & "Y" & "' ,"                               'ISCI
    slSql = slSql & 0 & ", "                                        'mnfSort
    slSql = slSql & "'" & "A" & "' ,"                               'BilAgyDir
    slSql = slSql & 0 & ", "                                        'arfLkCode
    slSql = slSql & 0 & ", "                                        'arfContrCode
    slSql = slSql & 0 & ", "                                        'arfInvCode
    slSql = slSql & "'" & "W" & "' ,"                               'CntrPrtSz
    slSql = slSql & 0 & ", "                                        'TrfCode
    slSql = slSql & "'" & slCrditApproval & "' ,"                               'CrdApp
    slSql = slSql & 0 & ", "                                        'pnfBuyer
    slSql = slSql & 0 & ", "                                        'pnfPay
    slSql = slSql & 0 & ", "                                        'Pct09
    slSql = slSql & "'" & "0.00" & "' ,"                            'CurrAR
    slSql = slSql & "'" & "0.00" & "' ,"                            'UnBilled           COBOL Money
    slSql = slSql & "'" & "0.00" & "' ,"                            'HiCredit           COBOL Money
    slSql = slSql & "'" & "0.00" & "' ,"                            'TotalGross         COBOL Money
    slSql = slSql & "'" & Format$(gNow, sgSQLDateForm) & "', "      'DateEntrd
    slSql = slSql & 0 & ", "                                        'NSFChks
    slSql = slSql & 0 & ", "                                        'AvgToPay
    slSql = slSql & 0 & ", "                                        'LstToPay
    slSql = slSql & 0 & ", "                                        'NoInvPpay
    slSql = slSql & "'" & "N" & "' ,"                               'NewBus
    slSql = slSql & 0 & ", "                                        'Merge
    slSql = slSql & 2 & ", "                                        'urfCode
    slSql = slSql & "'" & "A" & "' ,"                               'State
    slSql = slSql & "'" & Format$("12:00AM", sgSQLTimeForm) & "', " 'CrdAppTime
    slSql = slSql & "'" & "T" & "' ,"                               'PkInvShow
    slSql = slSql & 0 & ", "                                        'Guar
    slSql = slSql & 0 & ", "                                        'LastYearNew
    slSql = slSql & 0 & ", "                                        'LastMonthNew
    slSql = slSql & "'" & "Y" & "' ,"                               'RateOnInv
    slSql = slSql & 0 & ", "                                        'MnfBus
    slSql = slSql & 0 & ", "                                        'Unused1
    slSql = slSql & "'" & "Y" & "' ,"                               'AllowRepMG
    slSql = slSql & "'" & "Y" & "' ,"                               'BonusOnInv
    slSql = slSql & "'" & "I" & "' ,"                               'RepInvGen
    slSql = slSql & 0 & ", "                                        'mnfInvTerms
    slSql = slSql & "'" & "N" & "' )"                               'Political
    gDefaultADF = slSql
End Function

Public Function gDefaultPRF() As String

    Dim slSql As String
    
    gDefaultPRF = ""
    slSql = "Insert into PRF_Product_Names("
    slSql = slSql & "prfCode, "
    slSql = slSql & "prfadfCode, "
    slSql = slSql & "prfName, "
    slSql = slSql & "prfmnfComp1, "
    slSql = slSql & "prfmnfComp2, "
    slSql = slSql & "prfmnfExcl1, "
    slSql = slSql & "prfmnfExcl2, "
    slSql = slSql & "prfpnfBuyer, "
    slSql = slSql & "prfCppCpm, "
    slSql = slSql & "prfmnfDemo1, "
    slSql = slSql & "prfmnfDemo2, "
    slSql = slSql & "prfmnfDemo3, "
    slSql = slSql & "prfmnfDemo4, "
    slSql = slSql & "prfTarget1, "
    slSql = slSql & "prfTarget2, "
    slSql = slSql & "prfTarget3, "
    slSql = slSql & "prfTarget4, "
    slSql = slSql & "prfLastCPP1, "
    slSql = slSql & "prfLastCPP2, "
    slSql = slSql & "prfLastCPP3, "
    slSql = slSql & "prfLastCPP4, "
    slSql = slSql & "prfLastCPM1, "
    slSql = slSql & "prfLastCPM2, "
    slSql = slSql & "prfLastCPM3, "
    slSql = slSql & "prfLastCPM4, "
    slSql = slSql & "prfState, "
    slSql = slSql & "prfUrfCode, "
    slSql = slSql & "prfmnfBus, "
    slSql = slSql & "prfRemoteID, "
    slSql = slSql & "prfAutoCode, "
    slSql = slSql & "prfSyncTime, "
    slSql = slSql & "prfSourceID )"
    
    slSql = slSql & "Values ("
    
    slSql = slSql & "Replace" & ", "                                        'Code
    slSql = slSql & lgAdfCode & ", "                                'adfCode            ************
    slSql = slSql & "'" & gFixQuote(sgProdName) & "', "             'Name               *************
    slSql = slSql & 0 & ", "                                        'mnfComp1
    slSql = slSql & 0 & ", "                                        'mnfComp2
    slSql = slSql & 0 & ", "                                        'mnfExcl1
    slSql = slSql & 0 & ", "                                        'mnfExcl2
    slSql = slSql & 0 & ", "                                        'pnfBuyer
    
    'P = CPP M = CPM N = Unknown
    
    slSql = slSql & "'" & "P" & "' ,"                               'CppCpm
    slSql = slSql & 0 & ", "                                        'mnfDemo1    ********
    slSql = slSql & 0 & ", "                                        'mnfDem2
    slSql = slSql & 0 & ", "                                        'mnfDemo3
    slSql = slSql & 0 & ", "                                        'mnfDemo4
    slSql = slSql & 0 & ", "                                        'mnfTarget1
    slSql = slSql & 0 & ", "                                        'mnfTarget2
    slSql = slSql & 0 & ", "                                        'mnfTarget3
    slSql = slSql & 0 & ", "                                        'mnfTarget4
    slSql = slSql & 0 & ", "                                        'LastCPP1
    slSql = slSql & 0 & ", "                                        'LastCPP2
    slSql = slSql & 0 & ", "                                        'LastCPP3
    slSql = slSql & 0 & ", "                                        'LastCPP4
    slSql = slSql & 0 & ", "                                        'LastCPM1
    slSql = slSql & 0 & ", "                                        'LastCPM2
    slSql = slSql & 0 & ", "                                        'LastCPM3
    slSql = slSql & 0 & ", "                                        'LastCPM4
    slSql = slSql & "'" & "A" & "' ,"                               'State
    slSql = slSql & 2 & ", "                                        'UrfCode
    slSql = slSql & 0 & ", "                                        'mnfBus
    slSql = slSql & 0 & ", "                                        'remoteID
    slSql = slSql & 0 & ", "                                        'autoCode      ***********
    slSql = slSql & "'" & Format$("12:00AM", sgSQLTimeForm) & "', " 'syncTime
    slSql = slSql & 0 & ")"                                         'sourceID
    gDefaultPRF = slSql
End Function

Public Function gDefaultAgf() As String
    Dim slSQLQuery As String
    
    gDefaultAgf = ""
    slSQLQuery = "Insert Into AGF_Agencies ( "
    slSQLQuery = slSQLQuery & "agfCode, "
    slSQLQuery = slSQLQuery & "agfName, "
    slSQLQuery = slSQLQuery & "agfAbbr, "
    slSQLQuery = slSQLQuery & "agfCity, "
    slSQLQuery = slSQLQuery & "agfComm, "
    slSQLQuery = slSQLQuery & "agfslfCode, "
    slSQLQuery = slSQLQuery & "agfBuyer, "
    slSQLQuery = slSQLQuery & "agfCodeRep, "
    slSQLQuery = slSQLQuery & "agfCodeStn, "
    slSQLQuery = slSQLQuery & "agfCreditRestr, "
    slSQLQuery = slSQLQuery & "agfCreditLimit, "
    slSQLQuery = slSQLQuery & "agfPaymRating, "
    slSQLQuery = slSQLQuery & "agfISCI, "
    slSQLQuery = slSQLQuery & "agfmnfSort, "
    slSQLQuery = slSQLQuery & "agfCntrAddr1, "
    slSQLQuery = slSQLQuery & "agfCntrAddr2, "
    slSQLQuery = slSQLQuery & "agfCntrAddr3, "
    slSQLQuery = slSQLQuery & "agfBillAddr1, "
    slSQLQuery = slSQLQuery & "agfBillAddr2, "
    slSQLQuery = slSQLQuery & "agfBillAddr3, "
    slSQLQuery = slSQLQuery & "agfarfLkCode, "
    slSQLQuery = slSQLQuery & "agfPhone, "
    slSQLQuery = slSQLQuery & "agfFax, "
    slSQLQuery = slSQLQuery & "agfarfCntrCode, "
    slSQLQuery = slSQLQuery & "agfarfInvCode, "
    slSQLQuery = slSQLQuery & "agfCntrPrtSz, "
    slSQLQuery = slSQLQuery & "agfTrfCode, "
    slSQLQuery = slSQLQuery & "agfCrdApp, "
    slSQLQuery = slSQLQuery & "agfCrdRtg, "
    slSQLQuery = slSQLQuery & "agfpnfBuyer, "
    slSQLQuery = slSQLQuery & "agfpnfPay, "
    slSQLQuery = slSQLQuery & "agfPct90, "
    slSQLQuery = slSQLQuery & "agfCurrAR, "
    slSQLQuery = slSQLQuery & "agfUnbilled, "
    slSQLQuery = slSQLQuery & "agfHiCredit, "
    slSQLQuery = slSQLQuery & "agfTotalGross, "
    slSQLQuery = slSQLQuery & "agfDateEntrd, "
    slSQLQuery = slSQLQuery & "agfNSFChks, "
    slSQLQuery = slSQLQuery & "agfDateLstInv, "
    slSQLQuery = slSQLQuery & "agfDateLstPaym, "
    slSQLQuery = slSQLQuery & "agfAvgToPay, "
    slSQLQuery = slSQLQuery & "agfLstToPay, "
    slSQLQuery = slSQLQuery & "agfNoInvPd, "
    slSQLQuery = slSQLQuery & "agfMerge, "
    slSQLQuery = slSQLQuery & "agfurfCode, "
    slSQLQuery = slSQLQuery & "agfState, "
    slSQLQuery = slSQLQuery & "agfCrdAppDate, "
    slSQLQuery = slSQLQuery & "agfCrdAppTime, "
    slSQLQuery = slSQLQuery & "agfPkInvShow, "
    slSQLQuery = slSQLQuery & "agfRemoteID, "
    slSQLQuery = slSQLQuery & "agfAutoCode, "
    slSQLQuery = slSQLQuery & "agfSyncDate, "
    slSQLQuery = slSQLQuery & "agfSyncTime, "
    slSQLQuery = slSQLQuery & "agfSourceID, "
    slSQLQuery = slSQLQuery & "agfCntrExptForm, "
    slSQLQuery = slSQLQuery & "agfMnfInvTerms, "
    slSQLQuery = slSQLQuery & "agfXMLProposalBand, "
    slSQLQuery = slSQLQuery & "agf1or2DigitRating, "
    slSQLQuery = slSQLQuery & "agfXMLCallLetters, "
    slSQLQuery = slSQLQuery & "agfXMLDates, "
    slSQLQuery = slSQLQuery & "agfUnused "
    slSQLQuery = slSQLQuery & ") "
    slSQLQuery = slSQLQuery & "Values ( "
    slSQLQuery = slSQLQuery & "Replace" & ", "                              'Code
    slSQLQuery = slSQLQuery & "'" & gFixQuote(sgAgfname) & "', "    'Name
    slSQLQuery = slSQLQuery & "'" & gFixQuote(sgAgfAbbr) & "', "    'Abbrv
    slSQLQuery = slSQLQuery & "'" & gFixQuote(sgAgfCity) & "', "    'City
    slSQLQuery = slSQLQuery & 0 & ", "                              'Comm
    slSQLQuery = slSQLQuery & igSlfCode & ", "                      'Salesperson
    slSQLQuery = slSQLQuery & "'" & gFixQuote(sgAgfBuyer) & "', "   'Buyer
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "           'Code Rep
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "           'Code station
    slSQLQuery = slSQLQuery & "'" & gFixQuote("N") & "', "          'Credit Restrictions
    slSQLQuery = slSQLQuery & 0 & ", "                              'Credit Limit
    slSQLQuery = slSQLQuery & "'" & gFixQuote("1") & "', "          'PaymRating
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "           'ISCI
    slSQLQuery = slSQLQuery & 0 & ", "                              'mnfSort
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "           'Contract Address
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "           'Billing Address
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "
    slSQLQuery = slSQLQuery & 0 & ", "                              'arfLkCode
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "           'Phone
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "           'Fax
    slSQLQuery = slSQLQuery & 0 & ", "                              'Contract EDI service code: ArfCntrCode
    slSQLQuery = slSQLQuery & 0 & ", "                              'Invoice EDI Service code: ArfInvCode
    slSQLQuery = slSQLQuery & "'" & gFixQuote("W") & "', "          'Paper width: CntrPrtSz
    slSQLQuery = slSQLQuery & 0 & ", "                              'Tax rate: TrfCode
    slSQLQuery = slSQLQuery & "'" & gFixQuote("R") & "', "          'CrdApp
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "           'CrdRtg
    slSQLQuery = slSQLQuery & 0 & ", "                              'PnfBuyer
    slSQLQuery = slSQLQuery & 0 & ", "                              'PnfPay
    slSQLQuery = slSQLQuery & 0 & ", "                              'Pct90
    slSQLQuery = slSQLQuery & "'" & "0.00" & "', "                  'CurrAR
    slSQLQuery = slSQLQuery & "'" & "0.00" & "', "                  'Unbilled
    slSQLQuery = slSQLQuery & "'" & "0.00" & "', "                  'HiCredit
    slSQLQuery = slSQLQuery & "'" & "0.00" & "', "                  'TotalGross
    slSQLQuery = slSQLQuery & "'" & Format$(gNow, sgSQLDateForm) & "', "    'DateEntrd
    slSQLQuery = slSQLQuery & 0 & ", "                              'NSFChks
    slSQLQuery = slSQLQuery & "'" & Format$("1970-12-13", sgSQLDateForm) & "', "   'DateLstInv
    slSQLQuery = slSQLQuery & "'" & Format$("1970-12-31", sgSQLDateForm) & "', "  'DateLstPaym
    slSQLQuery = slSQLQuery & 0 & ", "                              'AvgToPay
    slSQLQuery = slSQLQuery & 0 & ", "                              'LstToPay
    slSQLQuery = slSQLQuery & 0 & ", "                              'NoInvPd
    slSQLQuery = slSQLQuery & 0 & ", "                              'Merge
    slSQLQuery = slSQLQuery & 2 & ", "                              'UrfCode
    slSQLQuery = slSQLQuery & "'" & gFixQuote("A") & "', "          'State
    slSQLQuery = slSQLQuery & "'" & Format$(tlagf.sCrdAppDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & "'" & Format$("12:00AM", sgSQLTimeForm) & "', "   'CrdAppTime
    slSQLQuery = slSQLQuery & "'" & gFixQuote("T") & "', "          'PkInvShow
    slSQLQuery = slSQLQuery & 0 & ", "                              'RemoteID
    slSQLQuery = slSQLQuery & 0 & ", "                              'AutoCode
    slSQLQuery = slSQLQuery & "'" & Format$("1970-12-31", sgSQLDateForm) & "', "    'SyncDate
    slSQLQuery = slSQLQuery & "'" & Format$("12:00AM", sgSQLTimeForm) & "', "       'SyncTime
    slSQLQuery = slSQLQuery & 0 & ", "                              'SourceID
    slSQLQuery = slSQLQuery & "'" & gFixQuote("C") & "', "          'CntrExptForm
    slSQLQuery = slSQLQuery & 0 & ", "                              'MnfInvTerms
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "           'sXMLProposalBand
    slSQLQuery = slSQLQuery & "'" & gFixQuote("1") & "', "          '1or2DigitRating
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "           'XMLCallLetters
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "           'XMLDates
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "' "            'Unused
    slSQLQuery = slSQLQuery & ") "
    gDefaultAgf = slSQLQuery

End Function

Public Function gDefaultSlf()

    gDefaultSlf = ""
    slSQLQuery = "Insert Into SLF_Salespeople ( "
    slSQLQuery = slSQLQuery & "slfCode, "
    slSQLQuery = slSQLQuery & "slfFirstName, "
    slSQLQuery = slSQLQuery & "slfLastName, "
    slSQLQuery = slSQLQuery & "slfsofCode, "
    slSQLQuery = slSQLQuery & "slfPhone, "
    slSQLQuery = slSQLQuery & "slfFax, "
    slSQLQuery = slSQLQuery & "slfmnfSlsTeam, "
    slSQLQuery = slSQLQuery & "slfMerge, "
    slSQLQuery = slSQLQuery & "slfJobTitle, "
    slSQLQuery = slSQLQuery & "slfState, "
    slSQLQuery = slSQLQuery & "slfurfCode, "
    slSQLQuery = slSQLQuery & "slfCodeStn, "
    slSQLQuery = slSQLQuery & "slfSalesGoal, "
    slSQLQuery = slSQLQuery & "slfUnderComm, "
    slSQLQuery = slSQLQuery & "slfOverComm, "
    slSQLQuery = slSQLQuery & "slfRemUnderComm, "
    slSQLQuery = slSQLQuery & "slfRemOverComm, "
    slSQLQuery = slSQLQuery & "slfStartCommPaid, "
    slSQLQuery = slSQLQuery & "SlfStartSales, "
    slSQLQuery = slSQLQuery & "slfStartCommDate, "
    slSQLQuery = slSQLQuery & "slfRemoteID, "
    slSQLQuery = slSQLQuery & "slfAutoCode, "
    slSQLQuery = slSQLQuery & "slfSyncDate, "
    slSQLQuery = slSQLQuery & "slfSyncTime, "
    slSQLQuery = slSQLQuery & "slfNewClientComm, "
    slSQLQuery = slSQLQuery & "slfIncClientComm, "
    slSQLQuery = slSQLQuery & "slfUnused "
    slSQLQuery = slSQLQuery & ") "
    slSQLQuery = slSQLQuery & "Values ( "
    slSQLQuery = slSQLQuery & "Replace" & ", "                              'Code
    slSQLQuery = slSQLQuery & "'" & gFixQuote(sgSlfFirstName) & "', "       'FirstName
    slSQLQuery = slSQLQuery & "'" & gFixQuote(sgSlfLastName) & "', "        'LastName
    slSQLQuery = slSQLQuery & 0 & ", "                                      'SofCode
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "                   'Phone
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "                   'Fax
    slSQLQuery = slSQLQuery & 0 & ", "                                      'MnfSlsTeam
    slSQLQuery = slSQLQuery & 0 & ", "                                      'Merge
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "                   'JobTitle
    slSQLQuery = slSQLQuery & "'" & gFixQuote("A") & "', "                  'State
    slSQLQuery = slSQLQuery & 0 & ", "                                      'UrfCode
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "', "                   'CodeStn
    slSQLQuery = slSQLQuery & 0 & ", "                                      'SalesGoal
    slSQLQuery = slSQLQuery & 0 & ", "                                      'UnderComm
    slSQLQuery = slSQLQuery & 0 & ", "                                      'OverComm
    slSQLQuery = slSQLQuery & 0 & ", "                                      'RemUnderComm
    slSQLQuery = slSQLQuery & 0 & ", "                                      'RemOverComm
    slSQLQuery = slSQLQuery & 0 & ", "                                      'StartCommPaid
    slSQLQuery = slSQLQuery & 0 & ", "                                      'StartSales
    slSQLQuery = slSQLQuery & "'" & Format$("1970-12-31", sgSQLDateForm) & "', "    'StartCommDate
    slSQLQuery = slSQLQuery & 0 & ", "                                      'RemoteID
    slSQLQuery = slSQLQuery & 0 & ", "                                      'AutoCode
    slSQLQuery = slSQLQuery & "'" & Format$("1979-12-31", sgSQLDateForm) & "', "    'SyncDate
    slSQLQuery = slSQLQuery & "'" & Format$("12:00AM", sgSQLTimeForm) & "', "       'SyncTime
    slSQLQuery = slSQLQuery & 0 & ", "                                      'NewClientComm
    slSQLQuery = slSQLQuery & 0 & ", "                                      'IncClientComm
    slSQLQuery = slSQLQuery & "'" & gFixQuote("") & "' "                    'Unused
    slSQLQuery = slSQLQuery & ") "
    
    gDefaultSlf = slSQLQuery
End Function
