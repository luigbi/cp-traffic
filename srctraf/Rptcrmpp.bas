Attribute VB_Name = "RPTCRMPP"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrmpp.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
'Public igYear As Integer                'budget year used for filtering
Dim imOwner As Integer                  'true if owner option
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF

Dim hmAgf As Integer            'Agency file handle
Dim imAgfRecLen As Integer      'AGF record length
Dim tmAgf As AGF
Dim hmSof As Integer            'Office file handle
Dim imSofRecLen As Integer      'SOF record length
Dim tmSof As SOF
Dim hmSlf As Integer            'Salesperson file handle
Dim imSlfRecLen As Integer      'SLF record length
Dim tmSlfSrchKey As INTKEY0     'SLF key image
Dim tmSlf As SLF
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim imTrade As Integer  'true = include trade contracts
Dim imCash As Integer
Dim imMerchant As Integer   'true = include merchandise transactions
Dim imPromotion As Integer      'true =include promotions transactions
'  Receivables File
Dim hmRvf As Integer        'receivables file handle
Dim tmRvf As RVF            'RVF record image
Dim imRvfRecLen As Integer  'RVF record length
'  Receivables Report File
Dim hmRvr As Integer        'receivables report file handle
Dim tmRvr As RVR            'RVR record image
Dim imRvrRecLen As Integer  'RVR record length
Dim tlSofList() As SOFLIST    'list of selling office codes and sales sourcecodes
Dim tmPifKey() As PIFKEY          'array of vehicle codes and start/end indices pointing to the participant percentages
                                        'i.e Vehicle XYZ has 2 sales sources, each with 3 participants.  That will be a total of
                                        '6 entries.  Vehicle XYZ points to lo index equal to 1, and a hi index equal to 6; the
                                        'next vehicle will be a lo index of 7, etc.
Dim tmPifPct() As PIFPCT          'all vehicles and all percentages from PIF

'
'
'
'                   Generate Participant Payment Status
'                   Gather History (PVF) and Receivables (RVF) for
'                   all transaction types and split their dollars based
'                   on the particpants share.
'
'                   This routine was copied from gCRRvrGen (rptsel),
'                   then modified.
'                   Created: 12/1/97    d.hosaka
'
'           7-7-06 add single contract selectivity for debugging
'           7-22-07 Immplement varying participant table
Sub gCRPPGen()
    Dim ilRet As Integer
    Dim ilLoop  As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoopOnFile As Integer             '2 passes, 1 for History, then Receivables
    Dim slStr As String
    Dim llAmt As Long
    Dim llDate As Long
    Dim ilValidDates As Integer             'true if trans date falls within requested start & end or
                                            'if distribution report- true if date entered falls within
                                            'start & end dates and trans date in past
    Dim ilTemp As Integer
    Dim llTransGross As Long                'total Gross of one transaction, used to balance splitting of owners
    Dim llTransNet As Long                  'total net of one trans, used to balance split owners
    Dim ilFoundOne As Integer
    Dim ilLoopSlsp As Integer
    Dim ilMatchCntr As Integer              'selectivity on holds, & contr types (remnants, PIs, etc)
    Dim ilHowMany As Integer                'times to loop - up to 10 records created per trans if by slsp with splits,
                                            'or just 1 per transaction if vehicle or advt option
    Dim ilHowManyDefined As Integer         '# of actual participants or slsp to process in splits
    Dim llProcessPct As Long                '% of slsp split or vehicle owner split (else 100%)
    Dim llEarliestDate As Long              'start date of data to retrieve from PRF or RVF
    Dim llLatestDate As Long                'end date of data to retrieve from PRF or RVF
    Dim ilTransFound As Integer
    Dim ilIncludeH As Integer
    Dim ilIncludeI As Integer
    Dim ilIncludeP As Integer
    Dim ilIncludeA As Integer
    Dim ilIncludeW As Integer
    Dim ilMatchSSCode As Integer            'matching sales source for participant option
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim slDate As String
    Dim llSingleCntr As Long
    'ReDim ilProdPct(1 To 1) As Integer            '5-22-07
    'ReDim ilMnfGroup(1 To 1) As Integer           '5-22-07
    'ReDim ilMnfSSCode(1 To 1) As Integer          '5-22-07
    'Index zero ignored in arrays below
    ReDim ilProdPct(0 To 1) As Integer            '5-22-07
    ReDim ilMnfGroup(0 To 1) As Integer           '5-22-07
    ReDim ilMnfSSCode(0 To 1) As Integer          '5-22-07
    Dim ilUse100pct As Integer                    '8-21-07 use 100% for participant share, do not search for the %


    hmRvr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRvr, "", sgDBPath & "Rvr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRvr)
        btrDestroy hmRvr
        Exit Sub
    End If
    imRvrRecLen = Len(tmRvr)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        'btrDestroy hmRvf
        btrDestroy hmRvr
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        btrDestroy hmCHF
        btrDestroy hmSlf
        'btrDestroy hmRvf
        btrDestroy hmRvr
        Exit Sub
    End If
    imSlfRecLen = Len(tmSlf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        btrDestroy hmCHF
        btrDestroy hmSlf
        'btrDestroy hmRvf
        btrDestroy hmRvr
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)

    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        btrDestroy hmVef
        btrDestroy hmCHF
        btrDestroy hmSlf
        'btrDestroy hmRvf
        btrDestroy hmRvr
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAgf)
        btrDestroy hmAgf
        btrDestroy hmMnf
        btrDestroy hmVef
        btrDestroy hmCHF
        btrDestroy hmSlf
        'btrDestroy hmRvf
        btrDestroy hmRvr
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)
    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSof)
        btrDestroy hmSof
        btrDestroy hmAgf
        btrDestroy hmMnf
        btrDestroy hmVef
        btrDestroy hmCHF
        btrDestroy hmSlf
        'btrDestroy hmRvf
        btrDestroy hmRvr
        Exit Sub
    End If
    imSofRecLen = Len(tmSof)
    imOwner = True
    imTrade = False                                         'assume trades should be included
    imCash = True
    imMerchant = False                                      'merchandising transactions
    imPromotion = False                                         'promotions transactions
    ilIncludeH = False                                     'include HI (inv history) transactions
    ilIncludeI = True                                     'include all I (invoice) transactions
    ilIncludeP = True                                     'include all P (payment) transactions
    ilIncludeA = True                                     'include all A (adjustment) transactions
    ilIncludeW = True                                     'include all W (Write off) transactions
    
    slStr = RptSelPP!edcDistributeTo.Text               'Latest date to retrieve from PRF or RVF
    llLatestDate = gDateValue(slStr)
    llSingleCntr = Val(RptSelPP!edcContract)        '7-7-06

    '11-2-10 changed to use entered date to start from (previously used start std date of year processing
    slStr = RptSelPP!edcEarliestDate.Text
    llEarliestDate = gDateValue(slStr)
    'obtain start of the standard broadcast year to begin extracting RVF & PHF
'    gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
'    slDate = "1/15/" & Trim$(slYear)
'    slDate = gObtainStartStd(slDate)
'    llEarliestDate = gDateValue(slDate)
    
    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        btrDestroy hmRvr
        Exit Sub
    End If
    imRvfRecLen = Len(tmRvf)

    'build array of selling office codes and their sales sources
    ilTemp = 0
    ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tlSofList(0 To ilTemp) As SOFLIST
        tlSofList(ilTemp).iSofCode = tmSof.iCode
        tlSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
        ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilTemp = ilTemp + 1
    Loop
    'Create table of the vehicles and their participants with share of split
    gCreatePIFForRpts llEarliestDate, tmPifKey(), tmPifPct(), RptSelPP

    For ilLoopOnFile = 1 To 2 Step 1                 '2 passes, first History, then Receivables
        'handles and buffers for PHF and RVF will be the same

        ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation

        Do While ilRet = BTRV_ERR_NONE
        If (tmRvf.lCode = 166859 And ilLoopOnFile = 1) Or ((tmRvf.lCode = 174422 Or tmRvf.lCode = 174423) And (ilLoopOnFile = 21)) Then
            ilRet = ilRet
        End If
            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slStr   'using tran date for filter unless its a payment (PI)
            llDate = gDateValue(slStr)                  'convert trans date to test if within requested limits
            'test for the valid transactions to include
            ilTransFound = False
            If (ilIncludeI) And (Left$(tmRvf.sTranType, 1) = "I") Then
                ilTransFound = True
            End If
            If (ilIncludeP) And (Left$(tmRvf.sTranType, 1) = "P") Then
                If tmRvf.sTranType = "PI" Then          'for payments, need to use entered date rather than tran date for filter
                    gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slStr   'using tran date for filter unless its a payment (PI)
                    llDate = gDateValue(slStr)                  'convert trans date to test if within requested limits
                End If
                ilTransFound = True
            End If
            If (ilIncludeA) And (Left$(tmRvf.sTranType, 1) = "A") Then
                ilTransFound = True
            End If
            If (ilIncludeW) And (Left$(tmRvf.sTranType, 1) = "W") Then
                ilTransFound = True
            End If
            If (ilIncludeH) And (Left$(tmRvf.sTranType, 1) = "H") Then
                ilTransFound = True
            End If
            If tmRvf.sCashTrade = "C" And Not (imCash) Then
                ilTransFound = False
            ElseIf tmRvf.sCashTrade = "T" And Not (imTrade) Then
                ilTransFound = False
            ElseIf tmRvf.sCashTrade = "M" And Not (imMerchant) Then
                ilTransFound = False
            ElseIf tmRvf.sCashTrade = "P" And Not (imPromotion) Then
                ilTransFound = False
            End If
            gPDNToLong tmRvf.sNet, llAmt

            ilValidDates = False
            'Filter on date entered for PO and trans date for everything else
            If (ilTransFound And llDate >= llEarliestDate And llDate <= llLatestDate) Then
                ilValidDates = True
            End If


            If ((ilValidDates) And (ilTransFound)) And ((llSingleCntr = 0) Or (llSingleCntr > 0 And llSingleCntr = tmRvf.lCntrNo)) Then '7-7-06 add selective contract test
                If (tmRvf.lCntrNo > 0) Then             'this trans is other than a PO (PO don'thave contract #s)
                    'Filter on type of transaction, date filter, and can't be an on account trans.
                    'get contract from history or rec file
                    tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'get matching contr recd

                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")
                         ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    ilMatchCntr = True

                    If tmChf.lCntrNo <> tmRvf.lCntrNo Or tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M" Then
                        ilMatchCntr = False
                    End If
                    If tmChf.iPctTrade = 100 And Not imTrade Then  'trades?
                        ilMatchCntr = False
                    End If
                Else                            'contract # not present
                    mFakeRvrSlsp                'setup slsp & comm from RVf
                    ilMatchCntr = True
                    ilRet = BTRV_ERR_NONE

                End If

                If ((ilRet = BTRV_ERR_NONE) And (ilMatchCntr)) Then
                    LSet tmRvr = tmRvf
                    tmRvr.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
                    tmRvr.iGenDate(1) = igNowDate(1)
                    'tmRvr.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
                    'tmRvr.iGenTime(1) = igNowTime(1)
                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                    tmRvr.lGenTime = lgNowTime
                    If ilLoopOnFile = 1 Then
                        tmRvr.sSource = "H"                    'let crystal know these records are histroy/receivables (vs contracts)
                    Else
                        tmRvr.sSource = "R"
                    End If
                                                                '(for cash distribution reports)
                    gPDNToLong tmRvf.sGross, llTransGross
                    gPDNToLong tmRvf.sNet, llTransNet
                    ilHowMany = 0
                    ilHowManyDefined = 0
                    If tmRvf.iSlfCode <> tmSlf.iCode Then   'only read if not already in mem
                        tmSlfSrchKey.iCode = tmRvf.iSlfCode
                        ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    End If
                    For ilLoop = LBound(tlSofList) To UBound(tlSofList)
                        If tlSofList(ilLoop).iSofCode = tmSlf.iSofCode Then
                            ilMatchSSCode = tlSofList(ilLoop).iMnfSSCode          'Sales source
                            Exit For
                        End If
                    Next ilLoop
                    'if PO need special coding since it doesnt have airing vehicle
                    'ReDim ilProdPct(1 To 1) As Integer            '5-22-07
                    'ReDim ilMnfGroup(1 To 1) As Integer           '5-22-07
                    'ReDim ilMnfSSCode(1 To 1) As Integer          '5-22-07
                    ReDim ilProdPct(0 To 1) As Integer            '5-22-07
                    ReDim ilMnfGroup(0 To 1) As Integer           '5-22-07
                    ReDim ilMnfSSCode(0 To 1) As Integer          '5-22-07


                    If tmRvf.sTranType = "PO" Then
                        tmRvr.imnfOwner = 0             'no owner
                        tmRvr.iProdPct = 10000
                        tmRvr.iMnfSSCode = ilMatchSSCode
                        'write out the PO recrd, all passed parms will not be used (use common point to do btrinsert)
                        mObtainPPShare llProcessPct, llTransGross, llTransNet, ilHowManyDefined, ilLoop
                    Else            'all tran types except POs
                        'get the vehicle for this transaction
                        tmVefSrchKey.iCode = tmRvf.iAirVefCode
                        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        'For ilLoop = 1 To 8 Step 1
                        '    If tmVef.iMnfSSCode(ilLoop) = ilMatchSSCode Then     'get count of how many actually
                                                                                'have to be processed.  Need to find
                                                                                'the last one because of extra pennies
                                                                                'goes to the last one processed
                         '       ilHowManyDefined = ilHowManyDefined + 1
                         '   End If
                        'Next ilLoop
                        'ilHowMany = 8               '8 max participants


                        'gInitPartGroupAndPcts tmRvf.iAirVefCode, ilMatchSSCode, tmRvf.iMnfGroup, ilMnfSSCode(), ilMnfGroup(), ilProdPct(), tmRvf.iTranDate(), tmPifKey(), tmPifPct()
                        ilUse100pct = False             '8-21-07 find the participant share, dont default to 100% share
                        gInitPartGroupAndPcts tmRvf.iAirVefCode, ilMatchSSCode, 0, ilMnfSSCode(), ilMnfGroup(), ilProdPct(), tmRvf.iInvDate(), tmPifKey(), tmPifPct(), ilUse100pct

                        'For ilLoop = 1 To UBound(ilMnfSSCode) - 1 Step 1
                            ilHowManyDefined = UBound(ilMnfSSCode)
                        'Next ilLoop
                        ilHowMany = UBound(ilMnfSSCode)

                        'For Owner , create as many as 3 records per transaction.  (up to 3 owners per vehicle)
                        For ilLoop = 0 To ilHowMany - 1 Step 1        'loop based on report option
                            slStr = ".00"
                            gStrToPDN slStr, 2, 6, tmRvr.sGross
                            gStrToPDN slStr, 2, 6, tmRvr.sNet
                             If ilMnfSSCode(ilLoop + 1) = ilMatchSSCode Then
                            'If tmVef.iMnfSSCode(ilLoop + 1) = ilMatchSSCode Then
                                'llProcessPct = tmVef.iProdPct(ilLoop + 1)
                                llProcessPct = ilProdPct(ilLoop + 1)
                                llProcessPct = llProcessPct * 100      'make it xxx.xxxx
                                If llProcessPct = 0 Then
                                    llProcessPct = 1000000
                                End If
                            Else
                                llProcessPct = 0                'this participant not used
                            End If

                            'If tmChf.lComm(ilLoop) > 0 Then
                            If llProcessPct > 0 Then
                                ilFoundOne = False
                                If (Not RptSelPP!ckcAll.Value = vbChecked) Then                               'slsp, check if any of the split slsp should be excluded
                                    For ilLoopSlsp = 0 To RptSelPP!lbcSelection(0).ListCount - 1 Step 1
                                        If RptSelPP!lbcSelection(0).Selected(ilLoopSlsp) Then              'selected slsp
                                            slNameCode = tgVehicle(ilLoopSlsp).sKey        'pick up slsp code
                                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                            'If Val(slCode) = tmVef.iMnfGroup(ilLoop + 1) Then
                                            If Val(slCode) = ilMnfGroup(ilLoop + 1) Then
                                                ilFoundOne = True
                                                Exit For
                                            End If
                                        End If
                                    Next ilLoopSlsp
                                Else                                                'all other options the record has already been filtered
                                    ilFoundOne = True
                                End If

                                If ilFoundOne Then
                                    'tmRvr.imnfOwner = tmVef.iMnfGroup(ilLoop + 1)
                                    'tmRvr.iProdPct = tmVef.iProdPct(ilLoop + 1)
                                    'tmRvr.iMnfSSCode = tmVef.iMnfSSCode(ilLoop + 1)
                                    '5-22-07 use the varying participant table
                                    tmRvr.imnfOwner = ilMnfGroup(ilLoop + 1)
                                    tmRvr.iProdPct = ilProdPct(ilLoop + 1)
                                    tmRvr.iMnfSSCode = ilMnfSSCode(ilLoop + 1)

                                    'get share for splits and write out pre-pass report record
                                    mObtainPPShare llProcessPct, llTransGross, llTransNet, ilHowManyDefined, ilLoop
                                End If                          'ilFoundOne
                            End If                              'llProcessPct > 0
                        Next ilLoop                             'loop for 10 slsp possible splits, or 3 possible owners, otherwise loop once
                    End If                                  'Trantype = PO
                End If                                      'ilmatchcnt
            End If                                          'ilvaliddates
            ilRet = btrGetNext(hmRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        ilRet = btrClose(hmRvf)

        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            btrDestroy hmCHF
            btrDestroy hmRvr
            btrDestroy hmSlf
            btrDestroy hmVef
            btrDestroy hmMnf
            btrDestroy hmAgf
            btrDestroy hmSof
            Exit Sub
        End If
        imRvfRecLen = Len(tmRvf)
        llEarliestDate = 0                              'include everything from beginning of time to user entered date from receivables
        '1-9-12 only get the requested transactions from receivables; if not, the payment is offset with the billing transactions in the "owed" column, resulting in $0
        'llLatestDate = gDateValue("12/30/2039")         'get everything in the Receivables
    Next ilLoopOnFile                                   '2 passes, first History, then Receivbles

    Erase tlSofList
    ilRet = btrClose(hmRvr)
    ilRet = btrClose(hmRvf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmSof)
End Sub
'
'*****************************************************************************************
'                   sub mFakeRvrSlsp - Contract # doesnt exist for
'                                        split slsp %.  Use the sales
'                                        person from Receivables file and
'                                        assume he gets 100% .
'                   Created 11/18/96 DH
'*****************************************************************************************
Sub mFakeRvrSlsp()
Dim ilLoop As Integer
    For ilLoop = 0 To 9 Step 1
        If ilLoop = 0 Then
            tmChf.iSlfCode(ilLoop) = tmRvf.iSlfCode
            tmChf.lComm(ilLoop) = 1000000
        Else
            tmChf.iSlfCode(ilLoop) = 0
            tmChf.lComm(ilLoop) = 0
        End If
    Next ilLoop
End Sub
'
'***************************************************************************************************************
'
'                mObtainPPShare : Calculate the Gross and Net portions of a transaction if
'                                  splitting the $ up by participant or slsp.  Update the
'                                  Report record's (RVR) gross and net $.
'                <input>    llProcesspct - % allocated to split slsp or participant
'                           llTransGross - orig. gross $ before splitting.  All monies
'                                           must be accounted for.  Extra pennies goes
'                                           to last one
'                           llTransNet - same as above except for net value.
'                           ilHowManyDefined - # of slsp or participants to process.
'                           ilLoop - # of times processed so far for participant or slsp
'                           (Compare illoop against ilhowmanydefined - for the last one
'                           give it the extra penny(s))
'                <output>   tmRvr.sGross
'                           tmRvr.sNet
'
'               Created 11/18/96 DH  (extracted & modified from rptsel gcrrvrgen)
'
'***************************************************************************************************************
Sub mObtainPPShare(llProcessPct As Long, llTransGross As Long, llTransNet As Long, ilHowManyDefined As Integer, ilLoop As Integer)
Dim slPct As String
Dim slAmount As String
Dim slDollar As String
Dim llGrossDollar As Long
Dim llNetDollar As Long
Dim ilRet As Integer
Dim slStr As String
        If tmRvr.sTranType <> "PO" Then                     'do splits for cash (excluding on account payments)
                                                            'and billing distribution
            slPct = gLongToStrDec(llProcessPct, 4)           'slsp split share in % or Owner pct.  If advt or vehicle
                                                                'options, slsp is force to100%

            gPDNToStr tmRvf.sGross, 2, slAmount
            slDollar = gMulStr(slPct, slAmount)                 'slsp gross portion of possible split
            llGrossDollar = Val(gRoundStr(slDollar, "01.", 0))
            llTransGross = llTransGross - llGrossDollar
            If ilLoop = ilHowManyDefined - 1 And RptSelPP!ckcAll.Value = vbChecked Then               'last slsp or participant processed? Handle extra pennies
                llGrossDollar = llGrossDollar + llTransGross  'last record written for splits, left over pennies goes to last owner
            End If
            slStr = gLongToStrDec(llGrossDollar, 2)
            gStrToPDN slStr, 2, 6, tmRvr.sGross
            gPDNToStr tmRvf.sNet, 2, slAmount
            slDollar = gMulStr(slPct, slAmount)                 'slsp net portion of possible split
            llNetDollar = Val(gRoundStr(slDollar, "01.", 0))
            llTransNet = llTransNet - llNetDollar
            If ilLoop = ilHowManyDefined - 1 And RptSelPP!ckcAll.Value = vbChecked Then               'last slsp or participant processed?  handle extra pennies
                llNetDollar = llNetDollar + llTransNet         'last record writen for splits, left over pennies goes to last owner
            End If
            slStr = gLongToStrDec(llNetDollar, 2)
            gStrToPDN slStr, 2, 6, tmRvr.sNet
    End If
    'If tmRvr.sNet <> ".00" And tmRvr.lDistAmt <> 0 Then         'dont show $0
    gPDNToStr tmRvr.sNet, 2, slAmount
    If slAmount <> "0.00" Then
        ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
    End If
End Sub
