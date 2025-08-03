Attribute VB_Name = "Rptcrcb"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptcrcb.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  imFirstActivate               lmCntrNo                      tmChfSrchKey1             *
'*  imTerminate                   imBypassFocus                 imFirstFocus              *
'*                                                                                        *
'******************************************************************************************

' File Name: Rptcrcb.bas
'
' Release: 5.5
'
' Description:
'
Option Explicit
Option Compare Text


Dim imMajorSet As Integer
Dim imMinorSet As Integer
Dim imInclVefCodes As Integer
Dim imUsevefcodes() As Integer     'vehicles codes to include or exclude
Dim imInclVGCodes As Integer
Dim imUseVGCodes() As Integer       'vehicle group codes from MNF
Dim imInclSSCodes As Integer
Dim imUseSSCodes() As Integer       'Sales source codes to include/exclude (mnf)

Dim tmSofList() As SOFLIST
Dim imFirstTime As Integer
Dim tlChfAdvtExt() As CHFADVTEXT
Dim hmAgf As Integer            'Agencyf ile handle
Dim tmAgfSrchKey As INTKEY0     'AGF key image
Dim imAgfRecLen As Integer        'AGF record length
Dim tmAgf As AGF

Dim hmCHF As Integer            'Contract header file handle
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF

Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF

Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF

Dim hmGrf As Integer            'Generic report file handle
Dim imGrfRecLen As Integer        'GRF record length
Dim tmGrf As GRF

Dim hmSof As Integer            'Office file handle
Dim imSofRecLen As Integer      'SOF record length
Dim tmSof As SOF

Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim tmMnfSS() As MNF                    'array of Sales Sources MNF
Dim tmNTRMNF() As MNF

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length

Dim hmSbf As Integer        'SBF file handle
Dim tmSbf As SBF            'SBF record image
Dim imSbfRecLen As Integer  'SBF record length



Type ACCDEFSELECT               'accrual/deferral selections
    lStart As Long             'input start date
    lEnd As Long               'input end date
    lContract As Long           'selective contract #
    iDays(0 To 6) As Integer   'days of the week selected
    lSingleCntr As Long         'selective contract #
    iSTd As Integer             'include std contracts
    iResv As Integer            'include reservatioins
    iRemnant As Integer         'include remnant
    iDR As Integer              'include DR
    iPI As Integer              'include Per Inquiry
    iPSA As Integer             'include PSA
    iPromo As Integer           'include Promo
    iAirTime As Integer         'include Air Time
    iRep As Integer             'include REP
    iNTR As Integer             'include NTR
    iHC As Integer              'include Hard Cost
    iPolit As Integer           'include politicals
    iNonPolit As Integer        'include non-politicals
    iSpotsByCount As Integer    'counts by spot count, false if do 30" unit counts
    iCalCycle As Integer
    iSTdCycle As Integer
    iWklyCycle As Integer
End Type


Dim tmPriceTypes As PRICETYPES
Dim tmSelect As ACCDEFSELECT





'
'                      Calculate Gross & Net $, and Split Cash/Trade $ from a schedule line
'                      mCalcTotalAmt - Loop and accumulate all days for one gross total.  This
'                       report combines all days for one total.  Then, calculate the net
'                       value from that; and split cash/trade (if applicable to contract).
'                       Total all days for more accuracy in $.
'
'                       <input> llGross() - array of gross $ for days requested from user input
'                               llSpots() - array of spots for the line
'                               ilCorT - 1 = Cash , 2 = Trade processing
'                               slPctTrade  - % of trade
'                               slCashAgyComm - agency comm %
'                               llGross = gross $ buckets for each day requested
'                       <output> llTotalGross - total gross value for all days requested
'                               llTotalNet - net values for total days requested
'                               llTotalSpots - total spots for the cash or trade portion
Private Sub mCalcTotalAmt(llGross() As Long, llSpots() As Long, llTotalGross As Long, llTotalNet As Long, llTotalSpots As Long, ilCorT As Integer, slPctTrade As String, slCashAgyComm As String)
Dim ilTemp As Integer
Dim slAmount As String
Dim slSharePct As String
Dim slStr As String
Dim slCode As String
Dim slDollar As String
Dim slNet As String
Dim slSpots As String
Dim slSpotShare As String

    llTotalGross = 0
    llTotalSpots = 0
    For ilTemp = 0 To UBound(llGross)
        llTotalGross = llTotalGross + llGross(ilTemp)
        llTotalSpots = llTotalSpots + llSpots(ilTemp)
    Next ilTemp

    slAmount = gLongToStrDec(llTotalGross, 2)
    slSpots = gLongToStrDec(llTotalSpots, 2)
    slSharePct = gLongToStrDec(10000, 2)
    slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
    slStr = gRoundStr(slStr, "1", 0)

    slSpotShare = gMulStr(slSharePct, slSpots)              '100% spots in string format
    slSpotShare = gRoundStr(slSpotShare, "1", 0)

    If ilCorT = 1 Then                 'all cash commissionable
        slCode = gSubStr("100.", slPctTrade)                'get the cash % (100-trade%)
        slDollar = gDivStr(gMulStr(slStr, slCode), "100")              'slsp gross
        slNet = gRoundStr(gDivStr(gMulStr(slDollar, gSubStr("100.00", slCashAgyComm)), "100.00"), ".01", 0)

        slSpotShare = gDivStr(gMulStr(slSpotShare, slCode), "100")

    Else
        If ilCorT = 2 Then                'at least cash is commissionable
            slCode = gIntToStrDec(tgChfCB.iPctTrade, 0)
            slDollar = gDivStr(gMulStr(slStr, slCode), "100")

            If tgChfCB.iAgfCode > 0 And tgChfCB.sAgyCTrade = "Y" Then
                slNet = gRoundStr(gDivStr(gMulStr(slDollar, gSubStr("100.00", slCashAgyComm)), "100.00"), "1", 0)
            Else
                slNet = slDollar    'no commission , net is same as gross
            End If

            slSpotShare = gDivStr(gMulStr(slSpotShare, slCode), "100")
        End If
    End If
    llTotalGross = Val(slDollar)
    llTotalNet = Val(slNet)
    llTotalSpots = Val(slSpotShare)

End Sub
'           mCreateDays
'           Obtain all contracts with an active date between the dates requested.
'           For each valid contract, build a record for each line saving the vehicle,
'           contract, gross & net $ with spot counts for cash & trade.
'           All stats are obtained from the contract; spots and $ are averaged across
'           the valid airing days of each week and distributed on those valid airing days.
'
' Weekly buy example:   9 spots M-F @ $75.
'                       total $ per week $675 (9 spots x $75)
'                       Number valid airing days: 5 (M-F)
'                       Avg $ distribution on each valid day of week:  $675/ 5 = $135/vallid airing day
'                       Avg spot distribtuion on each day of week:  9/5 = 1.8
'  Daily buys are on days where they air
'
'   GRF prepass record:
'   grfGenDate - generation date
'   grfGenTime - generation time
'   grfvefCode - vehicle code
'   grfchfcode - contract code (to retrieve advt code, s/s, sales office
'   grfBktType - C = C/T, Z = Hard Cost
'   grfCode 2 - 0 = direct, 1 = agy, 2 = NTR, 3 = political, 4 = Hardcost
'   grfrdfcode = mnfcode for NTR type
'   grfPerGenl(1) - mnf vehicle group item
'   grfPer1 - Cash gross
'   grfPer2 - cash net
'   grfPer3 - cash spots
'   grfPer4 - Trade gross
'   grfPer5 - trade net
'   grfPer6 - trade spots
'********************************************************************************************
Sub mCreateDays()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilTemp                        llStdStart                *
'*  llEnd                         ilWhichset                                              *
'******************************************************************************************

Dim ilRet As Integer
Dim slStart As String            'start date to gather
Dim slEnd As String              'end date to gather
Dim slCntrStatus As String          'list of contract status to gather (working, order, hold, etc)
Dim slCntrType As String            'list of contract types to gather (Per inq, Direct Response, Remnants, etc)
Dim ilHOState As Integer            'which type of HO cntr states to include (whether revisions should be included)
Dim llContrCode As Long
Dim ilCurrentRecd As Integer
Dim illoop As Integer
Dim ilClf As Integer                'loop count for lines
Dim llDate As Long
Dim ilCorT As Integer
Dim ilStartCorT As Integer
Dim ilEndCorT As Integer
Dim slCashAgyComm As String
Dim slPctTrade As String
Dim slNet As String
Dim llNet As Long
Dim slGross As String
Dim llGross As Long
Dim slGrossPct As String
Dim llAmt As Long
ReDim tlSbf(0 To 0) As SBF
'TTP 10855 - prevent overflow due to too many NTR items
'Dim ilSbf As Integer
Dim llSbf As Long
Dim slDate As String
Dim tlSBFTypes As SBFTypes
Dim ilIsItHardCost As Integer
Dim ilIsItPolitical As Integer
Dim ilVefIndex As Integer
Dim ilVefOK As Integer
Dim ilCntOK As Integer
Dim llTotalGross As Long
Dim llTotalNet As Long
Dim llTotalSpots As Long
Dim ilSlfCode As Integer
Dim ilmnfMinorCode As Integer
Dim ilMnfMajorCode As Integer
ReDim llCalSpots(0 To 0) As Long        'init buckets for daily calendar values
ReDim llCalAmt(0 To 0) As Long
ReDim llCalAcqAmt(0 To 0) As Long
Dim slNTRAgyComm As String


    'setup type statement as to which type of SBF records to retrieve (only NTR)
    tlSBFTypes.iNTR = True          'include NTR billing
    tlSBFTypes.iInstallment = False      'exclude Installment billing
    tlSBFTypes.iImport = False           'exclude rep import billing

    slCntrStatus = "HOGN"                 'statuses: hold, order, unsch hold, uns order
    slCntrType = ""
    If tmSelect.iSTd Then
        slCntrType = Trim$(slCntrType) & "C"        'std
    End If
    If tmSelect.iResv Then
        slCntrType = Trim$(slCntrType) & "V"        'reservation
    End If
    If tmSelect.iRemnant Then
        slCntrType = Trim$(slCntrType) & "T"        'remnant
    End If
    If tmSelect.iDR Then
        slCntrType = Trim$(slCntrType) & "R"        'DR
    End If
    If tmSelect.iPI Then
        slCntrType = Trim$(slCntrType) & "Q"        'PI
    End If
    If tmSelect.iPSA Then
        slCntrType = Trim$(slCntrType) & "S"        'PSA
    End If
    If tmSelect.iPromo Then
        slCntrType = Trim$(slCntrType) & "M"        'Promo
    End If
    'slCntrType = "CVTRQ"         'all types: PI, DR, etc.  except PSA(p) and Promo(m)
    ilHOState = 2                       'get latest orders & revisions  (HOGN plus any revised orders WCI)
    'build table (into tlchfadvtext) of all contracts that fall within the dates required

    slStart = Format$(tmSelect.lStart, "m/d/yy")
    slEnd = Format$(tmSelect.lEnd, "m/d/yy")
    ilRet = gObtainCntrForDate(RptSelCt, slStart, slEnd, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())

    tmGrf.lGenTime = lgNowTime
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1                                          'loop while llCurrentRecd < llRecsRemaining

        llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCB, tgClfCB(), tgCffCB())
        'test contract types
        ilCntOK = False
        ilIsItPolitical = gIsItPolitical(tgChfCB.iAdfCode)           'its a political, include this contract?
        'test for inclusion if its political adv and politicals requested, or
        'its not a political adv and politicals
        If (tmSelect.iPolit And ilIsItPolitical) Or ((tmSelect.iNonPolit) And (Not ilIsItPolitical)) Then
            'test for single contract
            If tmSelect.lContract = 0 Or tmSelect.lContract <> 0 And tmSelect.lContract = tgChfCB.lCntrNo Then    'single contract for debugging
                'test for Sales Source
                ilSlfCode = gBinarySearchSlf(tgChfCB.iSlfCode(0)) 'return index to salesp record to get selling office
                'test the maching selling office
                For illoop = 0 To UBound(tmSofList)
                    If tmSofList(illoop).iSofCode = tgMSlf(ilSlfCode).iSofCode Then
                        'get the associated S/S with this office
                        If gFilterLists(tmSofList(illoop).iMnfSSCode, imInclSSCodes, imUseSSCodes()) Then
                            ilCntOK = True
                        End If
                    End If
                    If ilCntOK = True Then
                        Exit For
                    End If
                Next illoop

            End If
        End If
        
        '8-16-12 test billing cycles
        If tgChfCB.sBillCycle = "C" And Not tmSelect.iCalCycle Then
            ilCntOK = False
        End If
        If tgChfCB.sBillCycle = "S" And Not tmSelect.iSTdCycle Then
            ilCntOK = False
        End If
        If tgChfCB.sBillCycle = "W" And Not tmSelect.iWklyCycle Then
            ilCntOK = False
        End If
        

        If ilCntOK Then
             'obtain agency for commission
             If tgChfCB.iAgfCode > 0 Then
                slCashAgyComm = ".00"
                tmAgfSrchKey.iCode = tgChfCB.iAgfCode
                ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slCashAgyComm = gIntToStrDec(tmAgf.iComm, 2)
                End If
             Else
                 slCashAgyComm = ".00"
             End If              'iagfcode > 0

             slPctTrade = gIntToStrDec(tgChfCB.iPctTrade, 0)
             If tgChfCB.iPctTrade = 0 Then                     'setup loop to do cash & trade
                 ilStartCorT = 1
                 ilEndCorT = 1
             ElseIf tgChfCB.iPctTrade = 100 Then
                 ilStartCorT = 2
                 ilEndCorT = 2
             Else
                 ilStartCorT = 1     'split cash/trade
                 ilEndCorT = 2
             End If

            If tgChfCB.sNTRDefined = "Y" And (tmSelect.iNTR = True Or tmSelect.iHC = True) Then        'this has NTR billing
                ilRet = gObtainSBF(RptSelCb, hmSbf, tgChfCB.lCode, slStart, slEnd, tlSBFTypes, tlSbf(), 0)   '11-28-06 add last parm to indicate which key to use

                For llSbf = LBound(tlSbf) To UBound(tlSbf) - 1
                    tmSbf = tlSbf(llSbf)
                    gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), slDate
                    llDate = gDateValue(slDate)
                    ilIsItHardCost = gIsItHardCost(tmSbf.iMnfItem, tmNTRMNF())

                    ilVefOK = False
                    'check for selective vehicle
                    ilVefIndex = gBinarySearchVef(tmSbf.iAirVefCode)      'find the vehicle record to see if its REP

                    'If RptSelCb!ckcAll.Value = vbChecked Then
                    '    ilVefOK = True
                    'Else
                    'If ilVefIndex > 0 Then      'vehicle must exist
                    If ilVefIndex >= 0 Then      'vehicle must exist
                        If gFilterLists(tmSbf.iAirVefCode, imInclVefCodes, imUsevefcodes()) Then
                            ilVefOK = mVGSelection(ilVefIndex)      'check vehicle group selection
                        End If
                    End If
                    'End If
                    'include the NTR if it falls within the dates requested, and its a regular NTR that should be included or
                    'a hard cost NTR that should be included
                    If (llDate >= tmSelect.lStart And llDate <= tmSelect.lEnd) And ((tmSelect.iNTR And Not ilIsItHardCost) Or (tmSelect.iHC And ilIsItHardCost)) And ilVefOK Then
                        'make sure original agy commission from airtime is retained
                        slNTRAgyComm = slCashAgyComm
                         If tmSbf.sAgyComm = "N" Then        'ntr comm flag overrides the contract
                            'slCashAgyComm = ".00"
                            slNTRAgyComm = ".00"
                        End If

                        For ilCorT = ilStartCorT To ilEndCorT
                            If ilCorT = 1 Then              'cash
                                slPctTrade = gSubStr("100.", gIntToStrDec(tgChfCB.iPctTrade, 0))
                            Else            'trade portion
                                slPctTrade = gIntToStrDec(tgChfCB.iPctTrade, 0)
                            End If
                            'convert the $ to gross & net strings
                            llAmt = tmSbf.lGross * tmSbf.iNoItems
                            slGross = gLongToStrDec(llAmt, 2)       'convert to xxxx.xx
                            'slGrossPct = gSubStr("100.00", slCashAgyComm)        'determine  % to client (normally 85%)
                            slGrossPct = gSubStr("100.00", slNTRAgyComm)        'determine  % to client (normally 85%)
                            slNet = gDivStr(gMulStr(slGrossPct, slGross), "100")    'net value

                            'calculate the new gross & net if split cash/trade
                            slNet = gDivStr(gMulStr(slNet, slPctTrade), "100")
                            llNet = gStrDecToLong(slNet, 2)
                            slGross = gDivStr(gMulStr(slGross, slPctTrade), "100")
                            llGross = gStrDecToLong(slNet, 2)

                            If llNet <> 0 Then           'bypass $0
                                tmGrf.iVefCode = tmSbf.iAirVefCode
                                tmGrf.lChfCode = tgChfCB.lCode
                                tmGrf.iRdfCode = tmSbf.iMnfItem         'mnf type
                                tmGrf.iCode2 = 2                        'ntr
                                tmGrf.sBktType = "C"                    'this is cash/trade combined
                                If ilIsItHardCost Then
                                    tmGrf.iCode2 = 4
                                    tmGrf.sBktType = "Z"                'force hard cost to end
                                End If

                                If ilCorT = 1 Then                   'cash portion
                                    'tmGrf.lDollars(1) = llAmt       'gross
                                    'tmGrf.lDollars(2) = llNet         'net
                                    'tmGrf.lDollars(3) = 0           'no spots on NTR
                                    'tmGrf.lDollars(4) = 0      'gross trade
                                    'tmGrf.lDollars(5) = 0       'trade net
                                    'tmGrf.lDollars(6) = 0          'trade- no spots on NTR
                                
                                    tmGrf.lDollars(0) = llAmt       'gross
                                    tmGrf.lDollars(1) = llNet         'net
                                    tmGrf.lDollars(2) = 0           'no spots on NTR
                                    tmGrf.lDollars(3) = 0      'gross trade
                                    tmGrf.lDollars(4) = 0       'trade net
                                    tmGrf.lDollars(5) = 0          'trade- no spots on NTR
                                Else
                                    'tmGrf.lDollars(4) = llAmt      'gross trade
                                    'tmGrf.lDollars(5) = llNet       'trade net
                                    'tmGrf.lDollars(6) = 0          'trade- no spots on NTR
                                    'tmGrf.lDollars(1) = 0       'gross
                                    'tmGrf.lDollars(2) = 0         'net
                                    'tmGrf.lDollars(3) = 0           'no spots on NTR

                                    tmGrf.lDollars(3) = llAmt      'gross trade
                                    tmGrf.lDollars(4) = llNet       'trade net
                                    tmGrf.lDollars(5) = 0          'trade- no spots on NTR
                                    tmGrf.lDollars(0) = 0       'gross
                                    tmGrf.lDollars(1) = 0         'net
                                    tmGrf.lDollars(2) = 0           'no spots on NTR

                                End If
                                gGetVehGrpSets tmSbf.iAirVefCode, imMinorSet, imMajorSet, ilmnfMinorCode, ilMnfMajorCode    '7-16-02 obtain vehicle group code, some options may not use it
                                'tmGrf.iPerGenl(1) = ilMnfMajorCode
                                tmGrf.iPerGenl(0) = ilMnfMajorCode

                                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                             End If
                        Next ilCorT
                    End If
                Next llSbf
            End If
            slPctTrade = gIntToStrDec(tgChfCB.iPctTrade, 0) 're-establish if went to NTR

            'process air time
            For ilClf = LBound(tgClfCB) To UBound(tgClfCB) - 1 Step 1
                tmClf = tgClfCB(ilClf).ClfRec

                ilVefIndex = gBinarySearchVef(tmClf.iVefCode)      'find the vehicle record to see if its REP
                ilVefOK = False
                If (tmSelect.iAirTime And (tgMVef(ilVefIndex).sType = "C" Or tgMVef(ilVefIndex).sType = "S" Or tgMVef(ilVefIndex).sType = "G")) Or (tmSelect.iRep And tgMVef(ilVefIndex).sType = "R") Then
                    If tmClf.sType = "H" Or tmClf.sType = "S" Then      'use only hidden and std lines (no packages)
                        'check for selective vehicle
                         If gFilterLists(tmClf.iVefCode, imInclVefCodes, imUsevefcodes()) Then
                            ilVefOK = mVGSelection(ilVefIndex)      'check vehicle group selection
                        End If
                    End If
                End If

                If ilVefOK Then
                    ReDim llCalSpots(0 To 0) As Long        'init buckets for daily calendar values
                    ReDim llCalAmt(0 To 0) As Long
                    ReDim llCalAcqAmt(0 To 0) As Long
                    gCalendarFlights tgClfCB(ilClf), tgCffCB(), tmSelect.lStart, tmSelect.lEnd, tmSelect.iDays(), tmSelect.iSpotsByCount, llCalAmt(), llCalSpots(), llCalAcqAmt(), tmPriceTypes
                    For ilCorT = ilStartCorT To ilEndCorT Step 1        '2 passes if split cash/trade
                        mCalcTotalAmt llCalAmt(), llCalSpots(), llTotalGross, llTotalNet, llTotalSpots, ilCorT, slPctTrade, slCashAgyComm
                        If ilCorT = 1 Then          'cash
                            'tmGrf.lDollars(1) = llTotalGross       'gross
                            'tmGrf.lDollars(2) = llTotalNet               'net
                            'tmGrf.lDollars(3) = llTotalSpots
                            'tmGrf.lDollars(4) = 0      'gross trade
                            'tmGrf.lDollars(5) = 0       'trade net
                            'tmGrf.lDollars(6) = 0
                        
                            tmGrf.lDollars(0) = llTotalGross       'gross
                            tmGrf.lDollars(1) = llTotalNet               'net
                            tmGrf.lDollars(2) = llTotalSpots
                            tmGrf.lDollars(3) = 0      'gross trade
                            tmGrf.lDollars(4) = 0       'trade net
                            tmGrf.lDollars(5) = 0
                        Else                        'trade
                            'tmGrf.lDollars(4) = llTotalGross      'gross trade
                            'tmGrf.lDollars(5) = llTotalNet       'trade net
                            'tmGrf.lDollars(6) = llTotalSpots
                            'tmGrf.lDollars(1) = 0       'gross
                            'tmGrf.lDollars(2) = 0               'net
                            'tmGrf.lDollars(3) = 0
                        
                            tmGrf.lDollars(3) = llTotalGross      'gross trade
                            tmGrf.lDollars(4) = llTotalNet       'trade net
                            tmGrf.lDollars(5) = llTotalSpots
                            tmGrf.lDollars(0) = 0       'gross
                            tmGrf.lDollars(1) = 0               'net
                            tmGrf.lDollars(2) = 0
                        End If

                    Next ilCorT                             'process cash or trade portion
                    tmGrf.iVefCode = tmClf.iVefCode
                    tmGrf.lChfCode = tgChfCB.lCode
                    tmGrf.iRdfCode = 0                  'not NTR (no mnf type)
                    tmGrf.iCode2 = 1                     'air time
                    tmGrf.sBktType = "C"                'cash/trade is shown on same line, separate from the hard costs
                    If tgChfCB.iAgfCode = 0 Then
                        tmGrf.iCode2 = 0                '4-6-07 direct
                    End If
                    If ilIsItPolitical Then                 'political
                        tmGrf.iCode2 = 3
                    End If
                    'If tmGrf.lDollars(1) + tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(6) <> 0 Then
                    If tmGrf.lDollars(0) + tmGrf.lDollars(2) + tmGrf.lDollars(3) + tmGrf.lDollars(5) <> 0 Then

                        gGetVehGrpSets tmClf.iVefCode, imMinorSet, imMajorSet, ilmnfMinorCode, ilMnfMajorCode    '7-16-02 obtain vehicle group code, some options may not use it
                        'tmGrf.iPerGenl(1) = ilMnfMajorCode
                        tmGrf.iPerGenl(0) = ilMnfMajorCode

                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    End If
                End If
            Next ilClf                          'next schedule line
        End If                                  'selective contract #
    Next ilCurrentRecd

    Exit Sub
End Sub

'
'
'
'           mClosefiles - Close all applicable files for
'
'
Sub mCloseFiles()
Dim ilRet As Integer
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmGrf)

   btrDestroy hmCHF
    btrDestroy hmVef
    btrDestroy hmMnf
    btrDestroy hmSof
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmAgf
    btrDestroy hmSbf
    btrDestroy hmAgf
    btrDestroy hmGrf
End Sub

'
'
'           mOpenFiles - open files applicable to Accrual/Defferal report
'                           This report takes all contracts within a requested date stpan
'                           and projects the $ and spot counts using an averaging method
'                           across the valid days of the week
'
'
Function mOpenFiles() As Integer
Dim ilRet As Integer
Dim ilTemp As Integer
Dim ilError As Integer
Dim slStamp As String
Dim tlSof As SOF

    ilError = False

    hmSbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenFilesErr
    gBtrvErrorMsg ilRet, "mOpenFiles (btrOpen SBF)", RptSelIA
    On Error GoTo 0
    imSbfRecLen = Len(tmSbf)

    hmAgf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenFilesErr
    gBtrvErrorMsg ilRet, "mOpenFiles (btrOpen AGF)", RptSelIA
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenFilesErr
    gBtrvErrorMsg ilRet, "mOpenFiles (btrOpen CHF)", RptSelIA
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)

    hmGrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenFilesErr
    gBtrvErrorMsg ilRet, "mOpenFiles (btrOpen GRF)", RptSelIA
    On Error GoTo 0
    imGrfRecLen = Len(tmGrf)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenFilesErr
    gBtrvErrorMsg ilRet, "mOpenFiles (btrOpen VEF)", RptSelIA
    On Error GoTo 0
    imVefRecLen = Len(tmVef)

    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenFilesErr
    gBtrvErrorMsg ilRet, "mOpenFiles (btrOpen MNF)", RptSelIA
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)

    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenFilesErr
    gBtrvErrorMsg ilRet, "mOpenFiles (btrOpen SOF)", RptSelIA
    On Error GoTo 0
    imSofRecLen = Len(tlSof)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenFilesErr
    gBtrvErrorMsg ilRet, "mOpenFiles (btrOpen CLF)", RptSelIA
    On Error GoTo 0
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenFilesErr
    gBtrvErrorMsg ilRet, "mOpenFiles (btrOpen CFF)", RptSelIA
    On Error GoTo 0
    imCffRecLen = Len(tmCff)

    'build array of selling office codes and their sales sources.  Used to test inclusion/exclusion of S/S
    ilTemp = 0
    ReDim tmSofList(0 To 0) As SOFLIST
    ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tmSofList(0 To ilTemp) As SOFLIST
        tmSofList(ilTemp).iSofCode = tmSof.iCode
        tmSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
        ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilTemp = ilTemp + 1
    Loop
    ilRet = gObtainMnfForType("S", slStamp, tmMnfSS())        'build sales sources

    imFirstTime = True              'set to create the header record in text file only once
    mOpenFiles = ilError
    Exit Function

mOpenFilesErr:
    ilError = True
    Return
End Function
'
'           gCreateAccruDefer - Gather all contracts for requested date span.  Normally used to get
'           get numbers for end of month accrual/deferral, the calculations used to obtain them
'           will average the $ across the valid days of the week.
'           For example:  9 spots M-F @ $50 ea = $450 total for the week
'           if only M, T and We requested, the averaging takes the total $ per week and divides
'           by the number of valid airing days.  In this case 450/5 = $90 on each valid day of week.
'           The result would be $270 for the 3 days requested.  Daily buys use the exact day & # spts
'           from the schedule line (no averaging)
'
'           Prepass file is built into GRF for output.
'
'
'       csi_calFrom repl. edcSelCFrom - from date
'       csi_calTo repl. edcSelCFrom1 - to date
'       ckcSelc8(0 to 6) - days of week
'       rbcSelCSelect(0 -2) sales source, sales origin, or vehicle
'       cbcSet1 - vehicle group
'       edcSelCto - contract #
'       ckcSelC10 - Summary Only
'       ckcSelC6 - contract types (std, reserved, etc)
'       ckcSelC3 - other types (i.e. line, ntr): air time, rep, ntr, hardcost, political, non-polit
'       ckcSelC5 - line rate types (charge, .00, adu, etc)
'       rbcSelC4(0 to 1): spot counts vs unit counts
'       lbcSelection(3) - vehicle list
'       lbcSelectin(4) Sales source list
'       lbcSelection(7) - vehicle group item list
Public Sub gGenAccrueDefer()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilTemp                                                                                *
'******************************************************************************************

Dim slDate As String
Dim illoop As Integer
Dim slStamp As String
Dim ilRet As Integer

        If mOpenFiles() = 0 Then
            ilRet = gObtainMnfForType("I", slStamp, tmNTRMNF())        'NTR Item types
            '9-111-19 use csi calendar control vs edit box
'            slDate = RptSelCb!edcSelCFrom.Text          'input start date
            slDate = RptSelCb!CSI_CalFrom.Text          'input start date
            tmSelect.lStart = gDateValue(slDate)
'            slDate = RptSelCb!edcSelCFrom1.Text         'input end date
            slDate = RptSelCb!CSI_CalTo.Text         'input end date
            tmSelect.lEnd = gDateValue(slDate)
            'option to use unit counts have been disabled until someone can think of a way to use it
            If RptSelCb!rbcSelC4(0).Value Then          'spot counts
                tmSelect.iSpotsByCount = True
            Else
                tmSelect.iSpotsByCount = False          'counts by 30"  units
            End If
            tmSelect.lContract = Val(RptSelCb!edcSelCTo.Text)     'selective contract #
            For illoop = 0 To 6                         'days of the week
                tmSelect.iDays(illoop) = gSetCheck(RptSelCb!ckcSelC8(illoop))
            Next illoop
            tmSelect.iSTd = gSetCheck(RptSelCb!ckcSelC6(0))     'std
            tmSelect.iResv = gSetCheck(RptSelCb!ckcSelC6(1))     'reserved
            tmSelect.iRemnant = gSetCheck(RptSelCb!ckcSelC6(2))     'remnant
            tmSelect.iDR = gSetCheck(RptSelCb!ckcSelC6(3))     'DR
            tmSelect.iPI = gSetCheck(RptSelCb!ckcSelC6(4))     'PI
            tmSelect.iPSA = gSetCheck(RptSelCb!ckcSelC6(5))     'PSA
            tmSelect.iPromo = gSetCheck(RptSelCb!ckcSelC6(6))     'Promo

            tmSelect.iAirTime = gSetCheck(RptSelCb!ckcSelC3(0))     'Air Time
            tmSelect.iRep = gSetCheck(RptSelCb!ckcSelC3(1))        'REP
            tmSelect.iNTR = gSetCheck(RptSelCb!ckcSelC3(2))     'NTR
            tmSelect.iHC = gSetCheck(RptSelCb!ckcSelC3(3))          'HardCost
            tmSelect.iPolit = gSetCheck(RptSelCb!ckcSelC3(4))          'Politicals
            tmSelect.iNonPolit = gSetCheck(RptSelCb!ckcSelC3(5))         'Non-politcals
            tmSelect.iCalCycle = gSetCheck(RptSelCb!ckcSelC13(0))            'include cal billed
            tmSelect.iSTdCycle = gSetCheck(RptSelCb!ckcSelC13(1))            'include std billed
            tmSelect.iWklyCycle = gSetCheck(RptSelCb!ckcSelC13(2))            'include wkly billed
            
            tmPriceTypes.iCharge = gSetCheck(RptSelCb!ckcSelC5(0))     'Chargeable lines
            tmPriceTypes.iZero = gSetCheck(RptSelCb!ckcSelC5(1))     '.00 lines
            tmPriceTypes.iADU = gSetCheck(RptSelCb!ckcSelC5(2))     'adu lines
            tmPriceTypes.iBonus = gSetCheck(RptSelCb!ckcSelC5(3))          'bonus lines
            tmPriceTypes.iNC = gSetCheck(RptSelCb!ckcSelC5(4))          'N/C lines
            tmPriceTypes.iRecap = gSetCheck(RptSelCb!ckcSelC5(5))         'recapturable
            tmPriceTypes.iSpinoff = gSetCheck(RptSelCb!ckcSelC5(6))       'spinoff
            tmPriceTypes.iMG = gSetCheck(RptSelCb!ckcSelC5(7))       'mg rates
            


            ReDim imUseVGCodes(0 To 0) As Integer
            ReDim imUsevefcodes(0 To 0) As Integer
            ReDim imUseSSCodes(0 To 0) As Integer
            'get the vehicles codes to include/exclude, use lbcselection(3)
            gObtainCodesForMultipleLists 3, tgVehicle(), imInclVefCodes, imUsevefcodes(), RptSelCb
            'get the sales source codes to include/exclude, use lbcselection(4)
            gObtainCodesForMultipleLists 4, tgMnfCodeCB(), imInclSSCodes, imUseSSCodes(), RptSelCb

            illoop = RptSelCb!cbcSet1.ListIndex
            imMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())
            If imMajorSet > 0 Then
                'get the vehicle group codes to include/exclude, use lbcselection(7)
                gObtainCodesForMultipleLists 7, tgMNFCodeRpt(), imInclVGCodes, imUseVGCodes(), RptSelCb
            End If
            mCreateDays

            mCloseFiles
            Screen.MousePointer = vbDefault
        End If
End Sub

Public Function mVGSelection(illoop As Integer) As Integer
Dim ilWhichset As Integer
Dim ilVefOK As Integer

        ilVefOK = False
        If imMajorSet > 0 Then
            'test for vehicle group
            If imMajorSet = 1 Then
                ilWhichset = tgMVef(illoop).iOwnerMnfCode
            ElseIf imMajorSet = 2 Then
                ilWhichset = tgMVef(illoop).iMnfVehGp2
            ElseIf imMajorSet = 3 Then
                ilWhichset = tgMVef(illoop).iMnfVehGp3Mkt
            ElseIf imMajorSet = 4 Then
                ilWhichset = tgMVef(illoop).iMnfVehGp4Fmt
            ElseIf imMajorSet = 5 Then
                ilWhichset = tgMVef(illoop).iMnfVehGp5Rsch
            ElseIf imMajorSet = 6 Then
                ilWhichset = tgMVef(illoop).iMnfVehGp6Sub
            End If
            If gFilterLists(ilWhichset, imInclVGCodes, imUseVGCodes()) Then
                ilVefOK = True
            End If
        Else
            ilVefOK = True
        End If
        mVGSelection = ilVefOK
End Function
