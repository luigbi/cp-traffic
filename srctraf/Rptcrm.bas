Attribute VB_Name = "RPTCRM"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrm.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
Dim imMajorSet As Integer               '7-16-02 vehicle group
Dim imMinorSet As Integer               '7-16-02
Dim imAdvt As Integer                  'true if advt option
Dim imSlsp As Integer                   'true if slsp option
Dim imVehicle As Integer                'true if vehicle option
Dim imAirVeh As Integer
Dim imBillVeh As Integer
Dim imOwner As Integer                  'true if owner option
Dim imAgency As Integer                 'true if agency option
Dim imInvoice As Integer                'true if invoice option
Dim imProducer As Integer               '2-10-00 true if producer option (no splitting of participants)
Dim imNTR As Integer                    '9-17-02 Item Billing Type (NTR)
Dim imHardCost As Integer               '3-17-05 hard cost option
Dim imSS As Integer                     '10-18-02 Sales Source option
Dim imSO As Integer                     '11-25-06 sales origin
Dim imGrossNeg As Integer
Dim imNetNeg As Integer
Dim imOffice As Integer
Dim imPolit As Integer                  '7-15-08 include politicals   inv reg
Dim imNonPolit As Integer               '7-15-08 incl non politicals  Inv reg

Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF

Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF

Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer        'CFF record length
Dim tmCff As CFF

Dim hmAdf As Integer            'Advertisr file handle
Dim imAdfRecLen As Integer      'ADF record length
Dim tmAdfSrchKey As INTKEY0     'ADF key image
Dim tmAdf As ADF

Dim hmAgf As Integer            'Agency file handle
Dim imAgfRecLen As Integer      'AGF record length
Dim tmAgfSrchKey As INTKEY0     'AGF key image
Dim tmAgf As AGF
Dim hmSof As Integer            'Office file handle
Dim imSofRecLen As Integer      'SOF record length
Dim tmSof As SOF
Dim hmSlf As Integer            'Salesperson file handle
Dim imSlfRecLen As Integer      'SLF record length
Dim tmSlfSrchKey As INTKEY0     'SLF key image
Dim tmSlf As SLF
'Dim hmUrf As Integer            'User file handle
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim tmNTRMNF() As MNF           'NTR Item types

Dim hmSbf As Integer
Dim imSbfRecLen As Integer
Dim tmSbf As SBF
Dim tmSbfSrchKey1 As LONGKEY0    'SBF key image
Dim tmSbfList() As SBF

'Dim imCbfRecLen As Integer      'CBF record length
'Dim tmCbfSrchKey As CBFKEY0     'Gen date and time
'Dim tmCbf As CBF
'Dim tmZeroCbf As CBF

Dim imIvrRecLen As Integer      'IVR record length
Dim tmIvr As IVR
Dim hmIvr As Integer

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length

Dim tmPnf As PNF                  'Personnel Projections
Dim hmPnf As Integer
Dim tmPnfSrchKey As INTKEY0       'personnel code
Dim imPnfRecLen As Integer        'PJF record length

'Quarterly Avails
Dim imPromo As Integer  'True=Include Promo
Dim imTrade As Integer  'true = include trade contracts
Dim imCash As Integer
Dim imMerchant As Integer   'true = include merchandise transactions
Dim imPromotion As Integer      'true =include promotions transactions
'Log Calendar
'Copy Report
Dim hmCpr As Integer            'Copy Report file handle
Dim tmCpr() As CPR                'CPR record image
Dim tmCprSrchKey As CPRKEY0            'CPR record image
Dim imCprRecLen As Integer        'CPR record length

Dim hmCxf As Integer            'Comment file handle
Dim imCxfRecLen As Integer      'CXF record length
Dim tmCxfSrchKey As LONGKEY0    'CXF key record image
Dim tmCxf As CXF

'Copy inventory
' Copy Combo Inventory File
'  Copy Product/Agency File
' Time Zone Copy FIle
'  Media code File
'  Rating Book File
'  Receivables File
Dim hmRvf As Integer        'receivables file handle
Dim tmRvf As RVF            'RVF record image
Dim imRvfRecLen As Integer  'RVF record length
Dim tmRvfList() As RVF
'  Receivables Report File
Dim hmRvr As Integer        'receivables report file handle
Dim tmRvr As RVR            'RVR record image
Dim tmRvrSrchKey As RVRKEY0   'RVR key record image
Dim imRvrRecLen As Integer  'RVR record length

Dim hmUor As Integer        'User report file handle
Dim tmUor As UOR            'User record image
Dim tmUorSrchKey As UORKEY0   'User key record image
Dim imUorRecLen As Integer  'URF record length
Dim tmUstUOR As UOR            'Date: 4/17/20202 User record image for ust.mkd

Dim hmUrf As Integer        'User table file handle
Dim tmUrf As URF           'User record image
Dim imUrfRecLen As Integer  'URF record length
Dim tmUrfSrchKey As INTKEY0

Dim hmUst As Integer        'User affiliate table file handle
Dim tmUst As UST           'User record image
Dim imUstRecLen As Integer  'UST record length
Dim tmUstSrchKey As INTKEY0

Dim hmUaf As Integer            'User Activity file
Dim imUafRecLen As Integer      'UAF record length
Dim tmUaf As UAF

Dim hmAfr As Integer            'Temp file for User Activity report
Dim imAfrRecLen As Integer      'AFR record length
Dim tmAfr As AFR
Dim hmVsf As Integer        'VEhicle Slsp table file handle
Dim tmVsf As VSF            'VSF record image
Dim imVsfRecLen As Integer  'VSF record length

Dim hmIihf As Integer               'invoice import file for barter stations
Dim tmIihf As IIHF
Dim tmIihfSrchKey0 As LONGKEY0
Dim tmIihfSrchKey1 As IIHFKEY1
Dim tmIihfSrchKey2 As IIHFKEY2
Dim tmIihfSrchKey3 As IIHFKEY3
Dim imIihfRecLen As Integer

Dim hmTxr As Integer            'text temporary file
Dim imTxrRecLen As Integer
Dim tmTxr As TXR

Dim tlSofList() As SOFLIST    'list of selling office codes and sales sourcecodes

Dim tmPifKey() As PIFKEY          'array of vehicle codes and start/end indices pointing to the participant percentages
                                        'i.e Vehicle XYZ has 2 sales sources, each with 3 participants.  That will be a total of
                                        '6 entries.  Vehicle XYZ points to lo index equal to 1, and a hi index equal to 6; the
                                        'next vehicle will be a lo index of 7, etc.
Dim tmPifPct() As PIFPCT          'all vehicles and all percentages from PIF
Dim tmChfAdvtExt() As CHFADVTEXT

Type INSTALLINFO
    iMonthIndex As Integer          '1 - 12
    lOrdered As Long                'ordered $ from lines & NTR
    lInstallment As Long            'installment $ entered per month by user
    lBilling As Long                '$ billed from RVF/phf (rvftype/phftype = "I")
    lRevenue As Long                'revenue (earned) from phf/rvf (rvftype/phftype = "A" or blank)
    sTranType As String * 2         'tran type from rvf/phf
    lStartDate As Long              'start date of the month
    lComment As Long                'AN comment code
    sCommentFlag As String * 1      'if multiple AN with more than 1 comment, set flag to indicate it on report
End Type
Type INSTALLDISCREP
    'sTranType(1 To 36) As String * 2
    'lInstallment(1 To 36) As Long
    'lBilling(1 To 36) As Long
    'lStartDate(1 To 36) As Long
    'Index zero ignored in the arrays below
    sTranType(0 To 36) As String * 2
    lInstallment(0 To 36) As Long
    lBilling(0 To 36) As Long
    lStartDate(0 To 36) As Long
End Type
'TTP 10117 - Reports: Cash Distribution report - add export option to export to CSV
Dim tmMnfSrchKey As INTKEY0
Dim tmMnfList() As MNFLIST      'array of mnf codes for Missed reasons and billing rules
Dim hmPrf As Integer            'Product file handle
Dim tmPrfSrchKey As LONGKEY0     'PrF record image
Dim imPrfRecLen As Integer       'PrF record length
Dim tmPrf As PRF
Dim lmExportCount As Long ' TTP 10252 - Ageing Summary by Month: overflow error when exporting
Dim hmExport As Integer
Dim smExportStatus As String
Dim smClientName As String
'TTP 10144 - Ageing Summary by Month - Export option
Type AGEINGSUMMARY
    iYear As Integer
    iMonth As Integer
    iAgencyCode As Integer
    iAdvertiserCode As Integer
    lProductCode As Long
    lContractNumber As Long
    lInvoiceNumber As Long
    iInvoiceDate(1) As Integer 'rvfTranDate
    dBalance As Double
    iDaysBehind As Integer
    iSalesPersonCode As Integer
End Type
Dim tmAgingSummary() As AGEINGSUMMARY

'TTP 10118 -Billing Distribution Export to CSV
Type INVDISTSUMMARY
    iParticipant As Integer
    iAiringVehicle As Integer
    sCashTrade As String
    iSalesSource As Integer
    iVehicle As Integer
    dGross As Double
    dNet As Double
    lDistDue As Long
    sMissingSSFlag As String * 1
End Type
Dim tmInvDistSummary() As INVDISTSUMMARY
'TTP 10902
Dim imLastAdfCode As Integer
Dim smLastAdfName As String
Dim imLastAgyCode As Integer
Dim smLastAgyName As String
Dim imLastSofCode As Integer
Dim smLastSofName As String

'*******************************************************
'*                                                     *
'*      Procedure Name:mZeroBalance                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer all zero balanced     *
'*                      invoices and trades to revenue *
'*                      history                        *
'*            5/7/97 dh: Change to zero balance all    *
'*                      transactions as of previous    *
'*                      closing period                 *
'*            2/18/20   lmzbrvfcode() is the output;   *
'*                      if include/not include record  *
'*                      in prepass, bypass if exists   *
'*******************************************************
Private Sub mZeroBalance(lmZBRvfCode() As Long)
    Dim tlRvf As RVF            'RVF or PHF record image
    Dim ilRvfRecLen As Integer  'RVF record length
    Dim hlRvf As Integer        'RVF or PHF file handle
    Dim hlPhf As Integer        'RVF or PHF file handle
    Dim hlChf As Integer
    Dim hlSbf As Integer
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim llRvfRecPos As Long
    Dim slDate As String
    Dim llEndPrevPeriod As Long
    Dim slEndPrevPeriod As String
    Dim slNet As String
    Dim ilUpper As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim illoop As Integer
    Dim llCashCount As Long  '2-8-02 chg from int. to long
    Dim ilTradeCount  As Integer
    Dim ilMerchCount As Integer
    Dim ilPromoCount  As Integer
    Dim slNowDate As String
    Dim slMsg As String
    Dim ilPass As Integer
    Dim slStr As String
    Dim ilBypassRvf As Integer
    Dim llNextLk As Long
    ReDim lmZBRvfCode(0 To 0) As Long       'build this array for RVF records to by bypassed
    Dim llLatestCashDate As Long
    Dim ilValidDates As Integer
    
    'Date: 03/18/2020 get latest cash date to include
    slStr = RptSel!CSI_CalTo.Text                   'Latest cash date to include
    llLatestCashDate = gDateValue(slStr)
    If llLatestCashDate = 0 Then                    'if end date not entered, use all
        llLatestCashDate = gDateValue("12/29/2069")
    End If
    
    slStr = RptSel!CSI_CalFrom.Text                 'Latest Billing date to include
    llEndPrevPeriod = gDateValue(slStr)
    If llEndPrevPeriod = 0 Then                     'if end date not entered, use all
        llEndPrevPeriod = gDateValue("12/29/2069")
    End If
    
    slNowDate = Format$(gNow(), "m/d/yy")
    hlRvf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlRvf, "", sgDBPath & "RVF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    hlPhf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlPhf, "", sgDBPath & "PHF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    
    hlChf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlChf, "", sgDBPath & "CHF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If

    hlSbf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlSbf, "", sgDBPath & "SBF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    
    
    'Passes:
    '        1 = Match on Invoice Numbers, Agency, Advertiser and Cash/Trade
    '        2 = Match on Contract Number, Agency, Advertiser and Cash/Trade
    '        3 = Zero in InvNo, CntrNo, slfCode, Air/Bill Vehicle and Match on Agency, Advertiser and Cash/Trade
    '        4 = Zero in InvNo, CntrNo, slfCode, Air/Bill Vehicle and Match on Aging, Agency, Advertiser and Cash/Trade
    '  2-16-12 5 = Zero in InvNo, CntrNo, slfCode, Air/Bill Vehicle and Match on Check#, Agency, Advertiser and Cash/Trade
    llCashCount = 0         '2-8-02 chg from int. to long
    ilTradeCount = 0
    ilMerchCount = 0
    ilPromoCount = 0
    'For ilPass = 1 To 4 Step 1
    For ilPass = 1 To 5 Step 1

        Screen.MousePointer = vbHourglass
        ReDim tmZP(0 To 0) As ZEROPURGE     'it's a global declaration  --> RECDEFAL.bas
        ReDim tmZPLink(0 To 0) As ZPLINK
        tmZP(0).lInvNo = 0
        tmZP(0).iAgfCode = -1
        tmZP(0).iAdfCode = -1
        tmZP(0).sAmount = ".00"
        tmZP(0).sType = ""
        tmZP(0).sCheckNo = ""
        tmZP(0).lFirstLk = -1
        ilUpper = 0
        ilRvfRecLen = Len(tlRvf)
        'gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slEndPrevPeriod
        'gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slEndPrevPeriod
        'llEndPrevPeriod = gDateValue(slEndPrevPeriod)
        ilRet = btrGetFirst(hlRvf, tlRvf, ilRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE)    'gathering all the passes
            If (tlRvf.sCashTrade <> "T") Or (tgSpf.sRUseTrade = "Y") Then
                'Ignore trade
                gUnpackDate tlRvf.iTranDate(0), tlRvf.iTranDate(1), slDate
                
                'Date: 3/12/2020 commented out test for transaction date; no need filter out payments after closing period
                'If gDateValue(slDate) <= llEndPrevPeriod Then
'If tlRvf.lCode = 1487 Then Stop
                'Date: 3/18/2020 filter cash transactions on/before the last cash date to include; billing transactions on/before billing date to include
                ilValidDates = False
                If (tlRvf.sTranType = "AN" Or tlRvf.sTranType = "IN") Then      'Billing
                    If (gDateValue(slDate) <= llEndPrevPeriod) Then
                        ilValidDates = True
                    End If
                Else        'PO, PI, JE                                         'Cash
                    If (gDateValue(slDate) <= llLatestCashDate) Then
                        ilValidDates = True
                    End If
                End If
                
                If ilValidDates Then
                    ilFound = False
                    ilBypassRvf = False
                    'Check if RVF code PASS 2-5, determine if need to filter out tlRvf as previously zero balanced
                    'Add a binary search on lmZBRvfCode, to find out if record is being used or not (zero balanced) --> mBinarySearchRVF
                    'if record is already zero balanced, then set ilBypassRvf to "TRUE"
                    ilBypassRvf = IIF(mBinarySearchRvf(tlRvf.lCode, lmZBRvfCode()) > 0, True, False)
                    If ilPass = 1 And ilBypassRvf = False Then
                        'Remove salesperson = 0 test as salesperson is now set sometimes
                        'If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                        If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                            ilBypassRvf = True
                        Else
                            For illoop = LBound(tmZP) To UBound(tmZP) - 1 Step 1
                                If (tmZP(illoop).lInvNo = tlRvf.lInvNo) And (tmZP(illoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(illoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(illoop).sType = tlRvf.sCashTrade) Then
                                    ilIndex = illoop
                                    ilFound = True
                                    Exit For
                                End If
                            Next illoop
                        End If
                    ElseIf ilPass = 2 And ilBypassRvf = False Then
                        'Remove salesperson = 0 test as salesperson is now set sometimes
                        'If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                        If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                            ilBypassRvf = True
                        Else
                            For illoop = LBound(tmZP) To UBound(tmZP) - 1 Step 1
                                If (tmZP(illoop).lInvNo = tlRvf.lCntrNo) And (tmZP(illoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(illoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(illoop).sType = tlRvf.sCashTrade) Then
                                    ilIndex = illoop
                                    ilFound = True
                                    Exit For
                                End If
                            Next illoop
                        End If
                    ElseIf ilPass = 3 And ilBypassRvf = False Then
                        'Remove salesperson = 0 test as salesperson is now set sometimes
                        'If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                        If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                            For illoop = LBound(tmZP) To UBound(tmZP) - 1 Step 1
                                'If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) And (tmZP(ilLoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(ilLoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(ilLoop).sType = tlRvf.sCashTrade) Then
                                If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) And (tmZP(illoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(illoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(illoop).sType = tlRvf.sCashTrade) Then
                                    ilIndex = illoop
                                    ilFound = True
                                    Exit For
                                End If
                            Next illoop
                        Else
                            ilBypassRvf = True
                        End If
                    ElseIf ilPass = 4 And ilBypassRvf = False Then
                        'Remove salesperson = 0 test as salesperson is now set sometimes
                        'If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                        If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                            For illoop = LBound(tmZP) To UBound(tmZP) - 1 Step 1
                                'If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) And (tmZP(ilLoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(ilLoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(ilLoop).sType = tlRvf.sCashTrade) And (tmZP(ilLoop).lInvNo = (CLng(tlRvf.iAgePeriod) * 10000 + tlRvf.iAgingYear)) Then
                                If (tmZP(illoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(illoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(illoop).sType = tlRvf.sCashTrade) And (tmZP(illoop).lInvNo = (CLng(tlRvf.iAgePeriod) * 10000 + tlRvf.iAgingYear)) Then
                                    ilIndex = illoop
                                    ilFound = True
                                    Exit For
                                End If
                            Next illoop
                        Else
                            ilBypassRvf = True
                        End If
                    ElseIf ilPass = 5 And ilBypassRvf = False Then          '2-17-12 add 5th pass to match on check #
                        'Remove salesperson = 0 test as salesperson is now set sometimes
                        'If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                        If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                            For illoop = LBound(tmZP) To UBound(tmZP) - 1 Step 1
                                '6/7/15: Check number changed to string
                                ''If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) And (tmZP(ilLoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(ilLoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(ilLoop).sType = tlRvf.sCashTrade) And (tmZP(ilLoop).lInvNo = (CLng(tlRvf.iAgePeriod) * 10000 + tlRvf.iAgingYear)) Then
                                'If (tmZP(ilLoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(ilLoop).iadfCode = tlRvf.iadfCode) And (tmZP(ilLoop).sType = tlRvf.sCashTrade) And (tmZP(ilLoop).lInvNo = tlRvf.lCheckNo) Then
                                If (tmZP(illoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(illoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(illoop).sType = tlRvf.sCashTrade) And (UCase$(Trim$(tmZP(illoop).sCheckNo)) = UCase$(Trim$(tlRvf.sCheckNo))) Then
                                    ilIndex = illoop
                                    ilFound = True
                                    Exit For
                                End If
                            Next illoop
                        Else
                            ilBypassRvf = True
                        End If
                    End If
                    If Not ilBypassRvf Then
                        If Not mRepBilled(tlRvf, hlChf, hlSbf) Then
                            ilBypassRvf = True
                        End If
                    End If
                    If Not ilBypassRvf Then
                        gPDNToStr tlRvf.sNet, 2, slNet
                        If Not ilFound Then
                            ilIndex = UBound(tmZP)
                            ReDim Preserve tmZP(0 To ilIndex + 1) As ZEROPURGE
                            If ilPass = 1 Then
                                tmZP(ilIndex).lInvNo = tlRvf.lInvNo
                            ElseIf ilPass = 2 Then
                                tmZP(ilIndex).lInvNo = tlRvf.lCntrNo
                            ElseIf ilPass = 3 Then
                                tmZP(ilIndex).lInvNo = -1
                            ElseIf ilPass = 4 Then
                                tmZP(ilIndex).lInvNo = CLng(tlRvf.iAgePeriod) * 10000 + tlRvf.iAgingYear
                            ElseIf ilPass = 5 Then          '2-17-12 add 5th pass to match on check#
                                tmZP(ilIndex).sCheckNo = tlRvf.sCheckNo
                            End If
                            tmZP(ilIndex).iAgfCode = tlRvf.iAgfCode
                            tmZP(ilIndex).iAdfCode = tlRvf.iAdfCode
                            tmZP(ilIndex).sAmount = slNet
                            tmZP(ilIndex).sAmount = gAddStr(Trim$(tmZP(ilIndex).sAmount), gLongToStrDec(tlRvf.lTax1, 2))
                            tmZP(ilIndex).sAmount = gAddStr(Trim$(tmZP(ilIndex).sAmount), gLongToStrDec(tlRvf.lTax2, 2))
                            tmZP(ilIndex).sType = tlRvf.sCashTrade
                            tmZP(ilIndex).lFirstLk = UBound(tmZPLink)
                            tmZPLink(UBound(tmZPLink)).lNextLk = -1
                            tmZPLink(UBound(tmZPLink)).lRvfCode = tlRvf.lCode
                            ReDim Preserve tmZPLink(0 To UBound(tmZPLink) + 1) As ZPLINK
                            tmZP(ilIndex + 1).lInvNo = 0
                            tmZP(ilIndex + 1).iAdfCode = -1
                            tmZP(ilIndex + 1).iAdfCode = -1
                            tmZP(ilIndex + 1).sAmount = ".00"
                            tmZP(ilIndex + 1).sType = ""
                            tmZP(ilIndex + 1).sCheckNo = ""
                            tmZP(ilIndex + 1).lFirstLk = -1
                        Else
                            tmZP(ilIndex).sAmount = gAddStr(Trim$(tmZP(ilIndex).sAmount), slNet)
                            tmZP(ilIndex).sAmount = gAddStr(Trim$(tmZP(ilIndex).sAmount), gLongToStrDec(tlRvf.lTax1, 2))
                            tmZP(ilIndex).sAmount = gAddStr(Trim$(tmZP(ilIndex).sAmount), gLongToStrDec(tlRvf.lTax2, 2))
                            llNextLk = tmZP(ilIndex).lFirstLk
                            Do
                                If tmZPLink(llNextLk).lNextLk = -1 Then
                                    tmZPLink(llNextLk).lNextLk = UBound(tmZPLink)
                                    tmZPLink(UBound(tmZPLink)).lNextLk = -1
                                    tmZPLink(UBound(tmZPLink)).lRvfCode = tlRvf.lCode
                                    ReDim Preserve tmZPLink(0 To UBound(tmZPLink) + 1) As ZPLINK
                                    Exit Do
                                End If
                                llNextLk = tmZPLink(llNextLk).lNextLk
                            Loop 'While llNextLk <> -1
                        End If
                    End If
                End If
            End If
            ilRet = btrGetNext(hlRvf, tlRvf, ilRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        For illoop = LBound(tmZP) To UBound(tmZP) - 1 Step 1
            'tmZP holds the dollars sum of the records that matched during the pass; if greater zero
            'Transfer zero balance
            If (gCompNumberStr(Trim$(tmZP(illoop).sAmount), ".00") = 0) Then '--> record is zero purged

                'lmZBRvfCode array holds the rvf code of all records that are zero balanced [purged]
                llNextLk = tmZP(illoop).lFirstLk
                Do  'building RVF zero purged records
                    lmZBRvfCode(UBound(lmZBRvfCode)) = tmZPLink(llNextLk).lRvfCode
                    ReDim Preserve lmZBRvfCode(0 To UBound(lmZBRvfCode) + 1) As Long
                    If tmZPLink(llNextLk).lNextLk = -1 Then
                        Exit Do
                    End If
                    llNextLk = tmZPLink(llNextLk).lNextLk
                Loop
             End If
        Next illoop
        'sort lmZBRvfCode descending
        If UBound(lmZBRvfCode) > 1 Then
            ArraySortTyp fnAV(lmZBRvfCode(), 0), UBound(lmZBRvfCode), 0, LenB(lmZBRvfCode(1)), 0, -2, 0
        End If
    Next ilPass '1-5 passes
    ilRet = btrClose(hlRvf)
    btrDestroy hlRvf
    ilRet = btrClose(hlPhf)
    btrDestroy hlPhf
    ilRet = btrClose(hlChf)
    btrDestroy hlChf
    ilRet = btrClose(hlSbf)
    btrDestroy hlSbf
End Sub

Public Function mBinarySearchRvf(ByVal lRvfCode As Long, llRvfCodes() As Long) As Long
    '6/16/21 - JW - TTP 10210 - changed to use Long, to prevent silent failure of Overflow error 6, and a -1 result
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim llRet As Long
    On Error GoTo mBinarySearchRvfErr
    llMin = LBound(llRvfCodes)
    llMax = UBound(llRvfCodes)
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If lRvfCode = llRvfCodes(llMiddle) Then
            'found the match
            mBinarySearchRvf = llMiddle
            Exit Function
        ElseIf lRvfCode < llRvfCodes(llMiddle) Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchRvf = -1
    Exit Function
mBinarySearchRvfErr:
    mBinarySearchRvf = -1
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:m RepBilled                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if Rep Billed        *
'*******************************************************
Private Function mRepBilled(tlRvf As RVF, hlChf As Integer, hlSbf As Integer) As Integer
    Dim ilBillVefCode As Integer
    Dim ilVef As Integer
    Dim ilRet As Integer
    Dim ilCount As Integer
    Dim tlChf As CHF
    Dim tlSbf As SBF
    Dim ilChfRecLen As Integer
    Dim ilSbfRecLen As Integer
    Dim tlChfSrchKey1 As CHFKEY1    'CHF key record image
    Dim tlSbfSrchKey0 As SBFKEY0    'SBF key record image

    ilChfRecLen = Len(tlChf)
    ilSbfRecLen = Len(tlSbf)
    
    mRepBilled = True
    ilCount = 0
    'Only check NTR transactions
    If (sgRepDef = "Y") And (tlRvf.sTranType = "IN") Then
        If (tgSpf.sPostCalAff = "N") And ((Asc(tgSpf.sUsingFeatures8) And REPBYDT) <> REPBYDT) Then
            Exit Function
        End If
        'Bypass Rep vehicles that are associated with NTR transaction
        If (tgSpf.sUsingNTR <> "Y") Or (tlRvf.iMnfItem <= 0) Then
            ilBillVefCode = tlRvf.iBillVefCode
            ilVef = gBinarySearchVef(ilBillVefCode)
            If ilVef <> -1 Then
                If tgMVef(ilVef).sType = "R" Then
                    tlChfSrchKey1.lCntrNo = tlRvf.lCntrNo
                    tlChfSrchKey1.iCntRevNo = 32000
                    tlChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hlChf, tlChf, ilChfRecLen, tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                    'If contract missing, assume that transaction entered via Backlog, 5/8/04
                    If (ilRet <> BTRV_ERR_NONE) Or (tlChf.lCntrNo <> tlRvf.lCntrNo) Then
                        Exit Function
                    End If
                    Do While (ilRet = BTRV_ERR_NONE) And (tlChf.lCntrNo = tlRvf.lCntrNo)
                        If tlChf.sSchStatus = "F" Then
                            Exit Do
                        End If
                        ilRet = btrGetNext(hlChf, tlChf, ilChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (tlChf.lCntrNo = tlRvf.lCntrNo) And (tlChf.sSchStatus = "F") Then
                        'Billing
                        tlSbfSrchKey0.lChfCode = tlChf.lCode
                        tlSbfSrchKey0.iDate(0) = tlRvf.iTranDate(0)
                        tlSbfSrchKey0.iDate(1) = tlRvf.iTranDate(1)
                        tlSbfSrchKey0.sTranType = "T"
                        ilRet = btrGetEqual(hlSbf, tlSbf, ilSbfRecLen, tlSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tlChf.lCode = tlSbf.lChfCode)
                            If (tlSbf.iDate(0) <> tlRvf.iTranDate(0)) Or (tlSbf.iDate(1) <> tlRvf.iTranDate(1)) Then
                                Exit Do
                            End If
                            If tlSbf.sTranType = "T" Then
                                If tlSbf.sBilled <> "Y" Then
                                    mRepBilled = False
                                    Exit Function
                                End If
                                ilCount = ilCount + 1
                            End If
                            ilRet = btrGetNext(hlSbf, tlSbf, ilSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                    If ilCount <= 0 Then
                        mRepBilled = False
                    End If
                End If
            End If
        End If
    End If
    Exit Function
End Function


''*******************************************************
''*                                                     *
''*      Procedure Name:gCRPlayListClear                *
''*                                                     *
''*             Created:10/09/93      By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments:Clear Play List                 *
''*                     for Crystal report              *
''*                                                     *
'       7-6-15 move to module rptextra.bas              *
''*******************************************************
'Sub gCRPlayListClear()
'    Dim ilRet As Integer
'    hmCpr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmCpr, "", sgDBPath & "Cpr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmCpr)
'        btrDestroy hmCpr
'        Exit Sub
'    End If
'    ReDim tmCpr(0 To 0) As CPR
'    imCprRecLen = Len(tmCpr(0))
'    tmCprSrchKey.iGenDate(0) = igNowDate(0)
'    tmCprSrchKey.iGenDate(1) = igNowDate(1)
'    'tmCprSrchKey.iGenTime(0) = igNowTime(0)
'    'tmCprSrchKey.iGenTime(1) = igNowTime(1)
'    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
'    tmCprSrchKey.lGenTime = lgNowTime
'    ilRet = btrGetGreaterOrEqual(hmCpr, tmCpr(0), imCprRecLen, tmCprSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
'    Do While (ilRet = BTRV_ERR_NONE) And (tmCpr(0).iGenDate(0) = igNowDate(0)) And (tmCpr(0).iGenDate(1) = igNowDate(1)) And (tmCpr(0).lGenTime = lgNowTime)
'        ilRet = btrDelete(hmCpr)
'        ilRet = btrGetNext(hmCpr, tmCpr(0), imCprRecLen, BTRV_LOCK_NONE, SETFORWRITE)
'    Loop
'    Erase tmCpr
'    ilRet = btrClose(hmCpr)
'    btrDestroy hmCpr
'End Sub
'Sub gCrRvrClear()
'                   Name changed to gRvrClear and moved to common module (rptextra.bas)
''*******************************************************
''*                                                     *
''*      Procedure Name:gCRRvrClear                     *
''*                                                     *
''*             Created:10/21/96      By:D. Hosaka      *
''*                                                     *
''*            Comments:Clear Receivables Report file   *
''*                     for Crystal report              *
''*                                                     *
''*******************************************************
'    Dim ilRet As Integer
'    hmRvr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmRvr, "", sgDBPath & "Rvr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmRvr)
'        btrDestroy hmRvr
'        Exit Sub
'    End If
'    imRvrRecLen = Len(tmRvr)
'    tmRvrSrchKey.iGenDate(0) = igNowDate(0)
'    tmRvrSrchKey.iGenDate(1) = igNowDate(1)
'    'tmRvrSrchKey.iGenTime(0) = igNowTime(0)
'    'tmRvrSrchKey.iGenTime(1) = igNowTime(1)
'    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
'    tmRvrSrchKey.lGenTime = lgNowTime
'    ilRet = btrGetGreaterOrEqual(hmRvr, tmRvr, imRvrRecLen, tmRvrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
'    Do While (ilRet = BTRV_ERR_NONE) And (tmRvr.iGenDate(0) = igNowDate(0)) And (tmRvr.iGenDate(1) = igNowDate(1)) And (tmRvr.lGenTime = lgNowTime)
'        ilRet = btrDelete(hmRvr)
'        ilRet = btrGetNext(hmRvr, tmRvr, imRvrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
'    Loop
'    ilRet = btrClose(hmRvr)
'    btrDestroy hmRvr
'End Sub
'
'
'                   Pre-pass file for Invoice REgisters and Distribution reports
'
'               2/8/98 DH When a participant exists that is zero % the rounding
'               causes .01 cents off.  Don't know why the participant is defined
'               with 0%.
'
'               4/12/99 CReate separate array of selective advt, agy or vehicles
'               instead of parsing the SORTCODE array for every transaction.
'               Also, create array of either entries selected, or entries not
'               selected, whichever is less (to hopefully speed up processing)
'               8/23/99 Inv Reg: When more than half of the items selected from the list box,
'               the selections were not showing.
'               10-29-99 New sort office in Invoice Register (office/vehicle)
'               12-20-99 Add option Ageing by Sales Source
'               2-10-00 Add option to Ageing by Producer (doesn't split participants,
'                       new invoicing option)
'               4-20-00 Use new commission structure for Salesperson Inv. register
'                   (match airing vehicle with the contract header slsp sub-company)
'               8-4-00 Fix formula error when selecting participant on Ageing by Producer
'               12-26-01 Change Inv register by office/Vehicle to use split slsp'
'               6-4-02 Make Cash REceipts 2 pass, for RVF & PHF
'               8-13-02 Changed ilSeqNo from integer to long.  Uses rvrurfcode as temporaray sort
'                   field for Crystal report.  Changed to use rvrrefinvno
'               9-17-02 Implement Invoice Register by NTR
'               10-18-02 Inv Reg- Sales Source option; vehicle group option
'               11-2-02 Handle extra pennies when splitting by slsp or participant.  Do all math as positive, not negative
'               11-4-02 Include transactions without Sales Source or Participant (transactions without airing vehicle defined,
'                       or a package vehicle stored as the airing vehicle and it doesnt contain a Sales Source) on the Participant report Aging
'               2-28-03 Test Sales Source, if Ask dont split revenue by participant for ageing by participant
'               3-19-03 Option to include Receivables, History or Both in Invoice Register
'               3-27-03 Fix Selective participants for Cash & Bill Distr, Ageing by Participant, SS, & Prod
'               7-11-03 Remove Cash Receipts code and combine with Cash summary (gCashGen)
'               11-11-03 Look for direct advertisrs that have been changed to regular agencies (in addition to testing for /direct,
'                       test for /non- (payee).  Treat them the same.
'               5-10-04 Add option to show sales source/participant on separate line to Acct History report
'               3-17-05 add option to include hard cost only in Invoice Register
'               1-20-06 Add option to Ageing by Participant, Sales Source and Producer to select Sales Sources
'               1-30-06 If the vehicle doesnt have a sale source defined matching he contract slsp sales source, show the
'                       transaction and alert user
'               7-3-06 Ageing by Payee- allow any trans to be printed when vehicle gropu selected and no vehicle group exists, they will sort to the top
'               11-25-06 Inv Register - Sales origin sort
'               1-30-07 Tax Register
'               5-18-07 Implement varying participant % by date
'               6-09-08 Hard Cost can be selected along with NTR and Airtime on ageing reports.
'TTP 10117 - Reports: Cash Distribution report - add export option to export to CSV, TTP 10164 - Ageing summary by month - export option for Audacy A/R dump
Sub gCRRvrGen(Optional blExport As Boolean = False)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilHowMany                     ilmnfParticipant                                        *
'******************************************************************************************

    Dim llSeqNo As Long                  '8-13-02 chged to long, running seq # of trans created to keep like transactions
                                            'apart in Cash Distribution report
    Dim ilListIndex As Integer              'report option
    Dim ilRet As Integer
    Dim illoop  As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoopOnFile As Integer             '2 passes, 1 for History, then Receivables
    Dim slStr As String
    Dim llCntrNo As Long                     'User entered contract #
    Dim llInvNo As Long                      '7-7-17 Selective Inv # Acct Hist
    Dim llAmt As Long
    Dim llDate As Long
    Dim llDateEntrd As Long
    Dim ilValidDates As Integer             'true if trans date falls within requested start & end or
                                            'if distribution report- true if date entered falls within
                                            'start & end dates and trans date in past
    Dim ilWhichMonth As Integer             'flag for level of control in Cash Disribution report:
                                            '1=current month cash, 2= prior month distributed cash,
                                            '3=prior month undistributed cash (PO)
    Dim ilTemp As Integer
    Dim llUndistCash As Long                'running total of all PO thrown away (trns date in past)
                                            'formula to be sent to Cash Distr by participant
    Dim llTransGross As Long                'total Gross of one transaction, used to balance splitting of owners
    Dim llTransNet As Long                  'total net of one trans, used to balance split owners
    Dim ilFoundOne As Integer
    Dim ilListBoxInx As Integer
    Dim ilFoundOption As Integer
    Dim ilLoopSlsp As Integer
    Dim ilMatchCntr As Integer              'selectivity on holds, & contr types (remnants, PIs, etc)
                                            'or just 1 per transaction if vehicle or advt option
    Dim ilHowManyDefined As Integer         '# of actual participants or slsp to process in splits
    Dim llProcessPct As Long                '% of slsp split or vehicle owner split (else 100%)
    Dim llEarliestDate As Long              'start date of data to retrieve Billing from PRF or RVF
    Dim llLatestDate As Long                'end date of data to retrieve Billing from PRF or RVF
    Dim slEarliestDate As String            '5-15-19 start date of data to retrieve Billing from PRF or RVF
    Dim slLatestDate As String              '5-15-19 end date of data to retrieve Billing from PRF or RVF
    Dim llEarliestCashDate As Long          'start date of data to retrieve Cash from PRF or RVF
    Dim llLatestCashDate As Long            'end date of data to retrieve Cash from PRF or RVF
    Dim ilTransFound As Integer
    Dim ilIncludeH As Integer
    Dim ilIncludeI As Integer
    Dim ilIncludeP As Integer
    Dim ilIncludeA As Integer
    Dim ilIncludeW As Integer
    Dim ilIncludeNTR As Integer
    Dim ilIncludeBilling As Integer         'true to include rvf/phf type "I" billing
    Dim ilIncludeEarned As Integer          'true to include earned revenue type "A"
    Dim ilStartFile As Integer              '1 = file PVF, 2 = file RVF
    Dim ilEndFile As Integer                '1 = file PVF, 2 = file RVF
    Dim ilMatchSSCode As Integer            'matching sales source for participant option
    Dim ilIncludeCodes As Integer               'true = include codes stored in ilusecode array,
                                            'false = exclude codes store din ilusecode array
    'ReDim ilUseCodes(1 To 1) As Integer       'valid advt, agency or vehicles codes to process--
    ReDim ilUseCodes(0 To 0) As Integer       'valid advt, agency or vehicles codes to process--
                                            'or advt, agy or vehicles codes not to process
    Dim ilMnfSubCo As Integer               '4-20-00
    Dim ilWhichset As Integer
    Dim ilCkcAll As Integer
    ReDim ilMnfcodes(0 To 0) As Integer     'array of valid vehicle groups to gather
    ReDim ilMnfSSCodes(0 To 0) As Integer   '1-20-06 Sales source codes for AGeing
    Dim ilEarliestAgeMM As Integer          'earliest MM/YY to process - user entred
    Dim ilEarliestAgeYY As Integer
    Dim ilLatestAgeMM As Integer             'latest MM/YY to process - user entred
    Dim ilLatestAgeYY As Integer
    ReDim ilExternalAdvList(0 To 0) As Integer
    Dim ilExternalAdv As Integer
    Dim ilUpper As Integer
    Dim slStamp As String
    ReDim tlMnf(0 To 0) As MNF
    Dim ilAskforUpdate As Integer
    Dim ilMissingSS As Integer         '1-30-06
    Dim ilValidSelect As Integer
    Dim ilExitFor As Integer
    Dim ilDistribution As Integer               'true if Billing or Cash Distribution
    Dim lmZBRvfCode() As Long                   'array of zero balanced RVF records
    Dim bZeroBalanced As Boolean                'flag for zero balanced RVF record
    Dim slRvfTranDate As String
    
    'TTP 10130 - Ageing report: if the include ageing MM/YY earliest/latest dates go across a year, it fails to find any data
    Dim sltmpEarliestDate As String
    Dim sltmpLatestDate As String
    Dim sltmpAgingDate As String
    
    'ReDim ilProdPct(1 To 1) As Integer            '5-1-07
    'ReDim ilMnfGroup(1 To 1) As Integer           '5-1-07
    'ReDim ilMnfSSCode(1 To 1) As Integer          '5-1-07
    'Index zero ignored in arrays below
    ReDim ilProdPct(0 To 1) As Integer            '5-1-07
    ReDim ilMnfGroup(0 To 1) As Integer           '5-1-07
    ReDim ilMnfSSCode(0 To 1) As Integer          '5-1-07
    Dim ilUse100pct As Integer                  '8-21-07 use 100% for participant share if rvfmnfgroup exists
    Dim llParticipantAdjustedDate As Long        '5-15-08 requested date minus 2 year to make sure the transaction has reference to participants
    ReDim ilParticipantDate(0 To 1) As Integer
    Dim ilIsItPolitical As Integer                  '7-15-08
    Dim ilAirNtrHard As Integer                 '7-21-08 flag for multiple choice in ageing reports Dan M
    Dim ilOKtoSeeVeh As Integer                 '4-15-11
    Dim blIsItAnyAgeing As Boolean              '8-27-15 is the report request any form of an ageing
    Dim llTotalSplits As Long                   '3-23-18 for slsp splits if exceeding over 100%, do not track running totals of splits which gives last slsp extra pennies
    Dim tlTranType As TRANTYPES                 '5-15-19 speed up processing anduse generalized routine which requires this array
    ReDim tlRvf(0 To 0) As RVF                  '5-15-19
    Dim llLoopOnTrans  As Long                  '5-15-19
    Dim ilWhichDate As Integer                  '5-15-19    0 = tran date, 1 = ageing date
    Const FLAG_NTR = 1
    Const FLAG_HARD_COST = 2
    Dim slFileName As String
    Dim slRepeat As String
    ReDim tmAgingSummary(0) As AGEINGSUMMARY
    ReDim tmInvDistSummary(0) As INVDISTSUMMARY
    
    ilListIndex = RptSel!lbcRptType.ListIndex               'report option from  Invoicing or Collections
    
    'TTP 10117 - Cash Distribution Export - Generate filename and Header
    If blExport = True Then
        RptSel.lacExport.Caption = "Exporting..."
        RptSel.lacExport.Refresh
        lmExportCount = 0
        slRepeat = "A"
        smClientName = Trim$(tgSpf.sGClient)
        If tgSpf.iMnfClientAbbr > 0 Then
            tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                smClientName = Trim$(tmMnf.sName)
            End If
        End If
        'Generate Export Filename
        Do
            ilRet = 0
            If ilListIndex = COLL_DISTRIBUTE Then
                'Invoice
                If RptSel.rbcSelCSelect(0).Value = True Then slFileName = "DistInv-"
                'Check#
                If RptSel.rbcSelCSelect(1).Value = True Then slFileName = "DistChk-"
                'Part
                If RptSel.rbcSelCSelect(2).Value = True Then slFileName = "DistPart-"
                
                'DateRange
                slFileName = slFileName & Format(RptSel.CSI_CalFrom.Text, "mmddyy")
                slFileName = slFileName & "To"
                slFileName = slFileName & Format(RptSel.CSI_CalTo.Text, "mmddyy")
                slFileName = slFileName & " - "
            End If
            'TTP 10164 - Ageing summary by month - export option for Audacy A/R dump
            If ilListIndex = COLL_AGEMONTH Then
                slFileName = "AgeMonth-"
            End If
            'TTP 10118 - Billing Distribution.
            If ilListIndex = INV_DISTRIBUTE Then
                If RptSel!rbcSelCInclude(0).Value Then          'detail
                    slFileName = "Invowndt-"
                Else
                    slFileName = "Invownsm-"
                End If
            End If
            
            'Todays Date
            slFileName = slFileName & Format(gNow, "mmddyy")
            slFileName = slFileName & slRepeat & " " & gFileNameFilter2(Trim$(smClientName))
            slFileName = slFileName & ".csv"
            'Check if exists, make new character
            ilRet = gFileExist(sgExportPath & slFileName)
            If ilRet = 0 Then       'if went to mOpenErr , there was a filename that existed with same name. Increment the letter
                slRepeat = Chr(Asc(slRepeat) + 1)
            End If
        Loop While ilRet = 0
        'Create File
        ilRet = gFileOpen(sgExportPath & slFileName, "OUTPUT", hmExport)
        If ilRet <> 0 Then
            MsgBox "Error writing file:" & sgExportPath & slFileName & vbCrLf & "Error:" & ilRet & " - " & Error(ilRet)
            Close #hmExport
            Exit Sub
        End If
        
        'Generate Header
        If ilListIndex = COLL_DISTRIBUTE Then
            'by Participant
            If RptSel.rbcSelCSelect(2).Value = True Then slStr = "Participant,Air Vehicle,Bill Vehicle,Source,Agency,Advertiser,Product,Contract,Inv#,Invoice Date,Paid Amount,Distribution %,Distribution Amount"
        End If
        'TTP 10164 - Ageing summary by month - export option for Audacy A/R dump
        If ilListIndex = COLL_AGEMONTH Then
            slStr = "Year/Month,Agency,Advertiser,AE,Product,Order Number,Invoice Number,Invoice Date,Balance"
        End If
        'TTP 10118 - Billing Distribution.
        If ilListIndex = INV_DISTRIBUTE Then
            If RptSel!rbcSelCInclude(0).Value Then          'detail
                'Invowndt
                'TTP 10459 - Bill Dist. participant shows sales source: Added Sales Source Column
                'slStr = "Participant,Sales Source,Airing Vehicle,Advertiser,Agency,Contract,Invoice Date,Inv #,Type,Gross Billed,Net Billed,Distribution %,Distribution Due,MissingSSFlag"
                'TTP 10902 - Billing Dist Export Change
                slStr = "Account ID,Participant,Sales Source,Airing Vehicle,Advertiser,Advertiser Reference ID,Political,Agency,Agency Reference ID,Contract,Product,Sales Office,Salesperson ID,Salesperson,Invoice Date,Inv #,Type,Gross Billed,Net Billed,Distribution %,Distribution Due,MissingSSFlag,Cash/Trade Flag"
            Else
                'Invownsm
                slStr = "Participant,Sales Source,Airing Vehicle,Gross,Net,Distribution Due,MissingSSFlag"
            End If
        End If
        'Write Header
        Print #hmExport, slStr
    End If
    
    hmRvr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRvr, "", sgDBPath & "Rvr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRvr)
        btrDestroy hmRvr
        Exit Sub
    End If
    imRvrRecLen = Len(tmRvr)
    'hmRvf = CBtrvTable() 'CBtrvObj()            'read History files using RVF handles and buffers
    'ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet <> BTRV_ERR_NONE Then
    '    ilRet = btrClose(hmRvf)
    '    btrDestroy hmRvf
    '    btrDestroy hmRvr
    '    Exit Sub
    'End If
    'imRvfRecLen = Len(tmRvf)
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

    hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()       'SBF required for Tax REgister
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSbf)
        btrDestroy hmSbf
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
    imSbfRecLen = Len(tmSbf)
    
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()       'VSF required for slsp to see only their stuff
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVsf)
        btrDestroy hmVsf
        btrDestroy hmSbf
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
    imVsfRecLen = Len(tmVsf)


    ilListIndex = RptSel!lbcRptType.ListIndex               'report option from  Invoicing or Collections

    If igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_ACCTHIST Then          'acct history needs to create another prepass to contain decrypted user name on report
        hmTxr = CBtrvTable(ONEHANDLE) 'CBtrvObj()       'VSF required for slsp to see only their stuff
        ilRet = btrOpen(hmTxr, "", sgDBPath & "txr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmTxr)
            btrDestroy hmTxr
            btrDestroy hmVsf
            btrDestroy hmSbf
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
        imTxrRecLen = Len(tmTxr)
    End If
   
    'TTP 10117 - Cash Distribution Export - for Product lookup
    If igRptCallType = COLLECTIONSJOB And (ilListIndex = COLL_DISTRIBUTE Or ilListIndex = COLL_AGEMONTH) Then 'TTP 10117 & 10164
        hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmTxr)
            btrDestroy hmTxr
            btrDestroy hmVsf
            btrDestroy hmSbf
            btrDestroy hmSof
            btrDestroy hmAgf
            btrDestroy hmMnf
            btrDestroy hmVef
            btrDestroy hmCHF
            btrDestroy hmSlf
            'btrDestroy hmRvf
            btrDestroy hmRvr
            btrDestroy hmPrf
            Exit Sub
        End If
        imPrfRecLen = Len(tmPrf)
    End If
    
    ilDistribution = False              'not cash or billing distribution
    blIsItAnyAgeing = False             '8-27-15
    'If ((igRptCallType = COLLECTIONSJOB) And (ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGEVEHICLE Or ilListIndex = COLL_CASH)) Then      'these 2 agings have vehicle group selectivity
    If ((igRptCallType = COLLECTIONSJOB) And (ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGEVEHICLE)) Then       'these 2 agings have vehicle group selectivity
        illoop = RptSel!cbcSet1.ListIndex
        imMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())
    ElseIf (igRptCallType = INVOICESJOB And ilListIndex = INV_REGISTER) Then     'Inv Register by airing vehicle & sales source have market selectivity
        If (RptSel!rbcSelCSelect(5).Value Or RptSel!rbcSelCSelect(8).Value) Or (RptSel!rbcSelCSelect(1).Value) Then    '5-13-11 add vg subsort to advt
            illoop = RptSel!cbcSet1.ListIndex
            imMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())
        End If
    Else
        illoop = 0
        imMajorSet = 0
    End If
    If illoop = 0 Then              'no vehicle group selected
        ilCkcAll = True             'force everything included
    Else
        If RptSel!ckcAllGroups.Value = vbChecked Then
            ilCkcAll = True
        Else
            ilCkcAll = False
        End If
    End If

    If Not ilCkcAll Or ((igRptCallType = COLLECTIONSJOB) And (ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGEVEHICLE)) Or (igRptCallType = INVOICESJOB And ilListIndex = INV_REGISTER) Then             'build array of selected vehicle groups
        For ilTemp = 0 To RptSel!lbcSelection(7).ListCount - 1 Step 1
            If RptSel!lbcSelection(7).Selected(ilTemp) Then
                slNameCode = tgSOCode(ilTemp).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilMnfcodes(UBound(ilMnfcodes)) = Val(slCode)
                ReDim Preserve ilMnfcodes(0 To UBound(ilMnfcodes) + 1)
            End If
        Next ilTemp
    End If

    '1-20-06 build array of sales sources selected
    For ilTemp = 0 To RptSel!lbcSelection(9).ListCount - 1 Step 1
        If RptSel!lbcSelection(9).Selected(ilTemp) Then
            slNameCode = tgMNFCodeRpt(ilTemp).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilMnfSSCodes(UBound(ilMnfSSCodes)) = Val(slCode)
            ReDim Preserve ilMnfSSCodes(0 To UBound(ilMnfSSCodes) + 1)
        End If
    Next ilTemp
    ilMnfSSCodes(UBound(ilMnfSSCodes)) = 0      'set one as zero to get POs and other loose ends
    ReDim Preserve ilMnfSSCodes(0 To UBound(ilMnfSSCodes) + 1)

    imPolit = True              'assume all reports include politicals / non-politicals
    imNonPolit = True           '7-15-08 inv register only report with option
    imAdvt = False                                          '
    imSlsp = False
    imVehicle = False
    imAirVeh = False
    imBillVeh = False
    imOwner = False
    imInvoice = False
    imAgency = False
    imProducer = False                                      '2-10-00
    imTrade = True                                         'assume trades should be included
    imCash = True
    imMerchant = False                                      'merchandising transactions
    imPromotion = False                                         'promotions transactions
    imNTR = False                                           '9-17-02
    imHardCost = False                                      '3-17-05
    imSS = False                                            '10-18-02 Sales Source option
    imSO = False                                            '11-25-06 sales origin
    ilIncludeH = False                                     'include HI (inv history) transactions
    ilIncludeI = False                                     'include all I (invoice) transactions
    ilIncludeP = False                                     'include all P (payment) transactions
    ilIncludeA = False                                     'include all A (adjustment) transactions
    ilIncludeW = False                                     'include all W (Write off) transactions
    ilIncludeNTR = 0                                  'Assume both air time & NTR included
    ilIncludeBilling = False                         'assume exclude the billing installment records until the report is determined,
                                                     'most reports will ignore type "I"
    ilIncludeEarned = True                          'assume include the earned installment records until thereport is determmined

    ilStartFile = 1                 'assume reading both phf & rvf
    ilEndFile = 2                   'assume reading both phf & rvf
    ilWhichDate = 0                 '5-15-19 default to use tran date in search
    '5-15-19 speed up processing by using extended reads in generalized routine , which needs this array of valid trans to include.
    'setup transaction types to retrieve from history and receivables
    tlTranType.iAdj = False              'adjustments
    tlTranType.iInv = False              'invoices
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iCash = True
    tlTranType.iTrade = True
    tlTranType.iMerch = False
    tlTranType.iPromo = False
    tlTranType.iNTR = True
    tlTranType.iAirTime = True


    'setup table of advertisers that are externally invoiced (if applicable)
    ilRet = gObtainAdvt()
    If Not ilRet Then
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmRvr)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmSof)
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

    '2-28-03 Gather the sales sources to determine how to update the vehicle (RVF/PHF).
    'Required to determine whether to automatically split transactions by revenue share in some reports
    ilRet = gObtainMnfForType("S", slStamp, tlMnf())
    ilRet = gObtainMnfForType("I", slStamp, tmNTRMNF())        'NTR Item types

    gObtainSalesperson                                          '4-17-11 maintain slsp in memory at all times
    
    ilUpper = 0
    For illoop = LBound(tgCommAdf) To UBound(tgCommAdf) - 1
        If tgCommAdf(illoop).sRepInvGen = "E" Then
            ilExternalAdvList(ilUpper) = tgCommAdf(illoop).iCode
            ReDim Preserve ilExternalAdvList(0 To ilUpper + 1)
            ilUpper = ilUpper + 1
        End If
    Next illoop

    llInvNo = 0                                     '7-7-17 selective inv # in Acct History
    If igRptCallType = INVOICESJOB Then
'        slStr = RptSel!edcSelCFrom.Text                'Earliest date to retrieve from PRF or RVF
        slStr = RptSel!CSI_CalFrom.Text                '8-15-19 Earliest date to retrieve from PRF or RVF
        llEarliestDate = gDateValue(slStr)
'        slStr = RptSel!edcSelCTo.Text                'Latest date to retrieve from PRF or RVF
        slStr = RptSel!CSI_CalTo.Text                 '8-15-19 Latest date to retrieve from PRF or RVF
        llLatestDate = gDateValue(slStr)
        slEarliestDate = Format(llEarliestDate, "ddddd")    '5-15-19
        slLatestDate = Format(llLatestDate, "ddddd")         '5-15-19
        If ilListIndex = INV_REGISTER Then
            imPolit = gSetCheck(RptSel!ckcSelC11(0).Value)      '7-15-08 new feature to include politicals/non-politicals
            imNonPolit = gSetCheck(RptSel!ckcSelC11(1).Value)

            If RptSel!rbcSelC12(0).Value Then     'billing
                ilIncludeEarned = False
                ilIncludeBilling = True
            Else
                ilIncludeEarned = True                  'include earned installment revenue (type = "A")
                ilIncludeBilling = False                'exclude installment billing (type = "I")
            End If
            '2-2-03 determine if airtime, NTR or both
            If RptSel!rbcSelC6(0).Value Then           'air time only
                ilIncludeNTR = 1
                tlTranType.iNTR = False               '5-15-19
            ElseIf RptSel!rbcSelC6(1).Value Then
                ilIncludeNTR = 2                        'include ntr only
                tlTranType.iAirTime = False             '5-15-19
            End If
            
            If RptSel!rbcSelC4(0).Value = True Then      'include cash only
                imTrade = False
                tlTranType.iTrade = False               '5-15-19
            End If
             If RptSel!rbcSelC4(1).Value = True Then      'include trade only
                imCash = False
                tlTranType.iCash = False                    '5-15-19
            End If
           
            'always go thru both files
             '3-19-03 determine if receivables, history, or both
            'If RptSel!rbcSelC8(0).Value Then            'receivables only?
            '    ilStartFile = 2                                                '
            'ElseIf RptSel!rbcSelC8(1).Value Then
           '     ilEndFile = 1                        'history only?
            'End If

            If RptSel!rbcSelCSelect(0).Value Then               'invoice option
                imInvoice = True
            ElseIf RptSel!rbcSelCSelect(1).Value Then           'advt option
                imAdvt = True
                mObtainCodes 5, tgAdvertiser(), ilIncludeCodes, ilUseCodes()
            ElseIf RptSel!rbcSelCSelect(2).Value Then           'agency
                imAgency = True
                mObtainCodes 1, tgAgency(), ilIncludeCodes, ilUseCodes()
            ElseIf RptSel!rbcSelCSelect(3).Value Then           'slsp
                imSlsp = True
            ElseIf (RptSel!rbcSelCSelect(4).Value) Then             'billing vehicle
                imVehicle = True
                imBillVeh = True
                mObtainCodes 6, tgCSVNameCode(), ilIncludeCodes, ilUseCodes()
            ElseIf (RptSel!rbcSelCSelect(5).Value) Then            'airing vehicle
                imVehicle = True
                imAirVeh = True
                mObtainCodes 6, tgCSVNameCode(), ilIncludeCodes, ilUseCodes()
            ElseIf RptSel!rbcSelCSelect(6).Value Then           'airing vehicle for office/vehicle option
                imVehicle = True
                imAirVeh = True
                mObtainCodes 6, tgCSVNameCode(), ilIncludeCodes, ilUseCodes()
            ElseIf RptSel!rbcSelCSelect(7).Value Then           'Item Billing Types (NTR)
                imNTR = True
                ilIncludeNTR = 2                        'include ntr only
                tlTranType.iAirTime = False             '5-15-19
                tlTranType.iNTR = True                  '5-15-19
                mObtainCodes 8, tgMnfCodeCT(), ilIncludeCodes, ilUseCodes()
            ElseIf RptSel!rbcSelCSelect(8).Value Then           'Sales Source
                imSS = True
                mObtainCodes 9, tgMNFCodeRpt(), ilIncludeCodes, ilUseCodes()
            ElseIf RptSel!rbcSelCSelect(9).Value Then            'sales origin
                imSO = True

                If RptSel!rbcSelC8(0).Value Then                  'no vehicle totals
                    RptSel!ckcAll.Value = vbChecked                 'fake out all vehicles included
                Else                                              'major totals by vehicle, see if any selected
                    mObtainCodes 6, tgCSVNameCode(), ilIncludeCodes, ilUseCodes()       'bill or airing vehicle major totals
                    If RptSel!rbcSelC8(1).Value Then                'bill vehicle
                        imBillVeh = True
                    Else
                        imAirVeh = True
                    End If
                End If
            End If
            If RptSel!ckcSelC3(0).Value = vbChecked Then                    'include invoices
                ilIncludeI = True
                tlTranType.iInv = True                          '5-15-19
                'ilIncludeH = True                           'history invoice trans.
            End If
            If RptSel!ckcSelC3(1).Value = vbChecked Then               'include adjustmetns
                ilIncludeA = True
                tlTranType.iAdj = True                          '5-15-19
            End If
            '3-20-03 option to include/exclude History transactions
            If RptSel!ckcSelC3(2).Value = vbChecked Then              'include history
                 ilIncludeH = True                           'history invoice trans.
                 tlTranType.iInv = True                     '5-15-19
            End If

            '3-17-05 see if NTR Hard cost only
            If RptSel!ckcSelC7.Value = vbChecked Then
                imHardCost = True
                tlTranType.iHardCost = True                 '5-15-19
            End If
        ElseIf ilListIndex = INV_DISTRIBUTE Then
            ilIncludeBilling = True                 'include type "I" for installment billing
            ilIncludeEarned = False                 'exclude type "A" installment revenue
            ilDistribution = True
            imOwner = True
            If RptSel!ckcSelC3(0).Value = vbChecked Then                    'include invoices
                ilIncludeI = True
                tlTranType.iInv = True                      '5-15-19
                'ilIncludeH = True                           'history invoice trans.
            End If
            If RptSel!ckcSelC3(1).Value = vbChecked Then               'include adjustmetns
                ilIncludeA = True
                tlTranType.iAdj = True                      '5-15-19
            End If
            '3-20-03 option to include/exclude History transactions
            If RptSel!ckcSelC3(2).Value = vbChecked Then              'include history
                 ilIncludeH = True                           'history invoice trans.
                 tlTranType.iInv = True                     '5-15-19
            End If
            mObtainCodes 3, tgVehicle(), ilIncludeCodes, ilUseCodes()
            slStr = RptSel!edcCheck.Text
            llCntrNo = Val(slStr)

        ElseIf ilListIndex = INV_TAXREGISTER Then
            ilIncludeBilling = False                'exclude "I" installment billing
            ilIncludeEarned = True                  'include "A" installment revenue
            ilRet = gObtainTrf()
            imAirVeh = True
            imVehicle = True
            mObtainCodes 6, tgCSVNameCode(), ilIncludeCodes, ilUseCodes()
            ilIncludeI = gSetCheck(RptSel!ckcSelC3(0).Value)   'incl invoices
            ilIncludeH = gSetCheck(RptSel!ckcSelC3(0).Value)   'incl invoices
            ilIncludeA = gSetCheck(RptSel!ckcSelC3(1).Value)   'incl adjustments
            ilIncludeP = gSetCheck(RptSel!ckcSelC3(2).Value)   'incl payments
            ilIncludeW = gSetCheck(RptSel!ckcSelC3(3).Value)   'incl journal entries
            slStr = RptSel!edcSelCTo1.Text
            llCntrNo = Val(slStr)
            'imMerchant = True                                      'merchandising transactions
            'imPromotion = True                                         'promotions transactions
            'imHardCost = True                                   'include air time/NTRs
            ilIncludeNTR = 0                                    'air time and ntr
            tlTranType.iAirTime = True                          '5-15-19
            tlTranType.iNTR = True                              '5-15-19
            tlTranType.iInv = ilIncludeI                        '5-15-19
            tlTranType.iAdj = ilIncludeA                        '5-15-19
            tlTranType.iPymt = ilIncludeP                       '5-15-19
            tlTranType.iWriteOff = ilIncludeW                   '5-15-19
        End If
    'Cash Distribution, Account History , All Ageings come thru here
    ElseIf igRptCallType = COLLECTIONSJOB Then
        '8-28-19 use csi calendar control vs edit box
'        slStr = RptSel!edcSelCFrom.Text                ' Earliest date to retrieve  from PRF or RVF
        slStr = RptSel!CSI_CalFrom.Text                ' Earliest date to retrieve  from PRF or RVF
        llEarliestDate = gDateValue(slStr)
'        slStr = RptSel!edcSelCTo.Text               'Latest date to retrieve  from PRF or RVF
        slStr = RptSel!CSI_CalTo.Text               'Latest date to retrieve  from PRF or RVF
        llLatestDate = gDateValue(slStr)
        If llLatestDate = 0 Then                    'if end date not entered, use all
            llLatestDate = gDateValue("12/29/2069")
        End If
        slEarliestDate = Format(llEarliestDate, "ddddd")    '5-15-19
        slLatestDate = Format(llLatestDate, "ddddd")         '5-15-19

        '7-11-03 remove cash receipts & payment history code and combine in gCashGen
        'If ilListIndex = COLL_PAYHISTORY Then
       '     If RptSel!rbcSelCSelect(0).Value Then           'agy
        '        imAgency = True
        '
        '        mObtainCodes 1, tgAgency(), ilIncludeCodes, ilUseCodes()
        '
        '    ElseIf RptSel!rbcSelCSelect(1).Value Then       'advt
        '        imAdvt = True
        '        mObtainCodes 0, tgAdvertiser(), ilIncludeCodes, ilUseCodes()
        '
        '    Else                                            'vehicle
        '        imVehicle = True
        '        imAirVeh = True
        '        mObtainCodes 6, tgCSVNameCode(), ilIncludeCodes, ilUseCodes()
        '    End If
        '    ilIncludeP = True
        '    ilIncludeW = True
        '    mGetCashTradeOption         'check user selections for cash/trade/merch/promo
        '    'determine if airtime, NTR or both
        '    If RptSel!rbcSelC6(0).Value Then           'air time only
        '        ilIncludeNTR = 1                        '
        '    ElseIf RptSel!rbcSelC6(1).Value Then
        '        ilIncludeNTR = 2                        'include ntr only
        '    End If
        'ElseIf ilListIndex = COLL_CASH Then
        '    slStr = RptSel!edcSelCFrom1.Text
        '    llCheckNo = Val(slStr)
        '    If RptSel!ckcSelC3(0).Value = vbChecked Then       'payments selected
        '        ilIncludeP = True
        '    End If
        '    If RptSel!ckcSelC3(1).Value = vbChecked Then       'journal entries selected
        '        ilIncludeW = True
        '    End If
        '    mGetCashTradeOption         'check user selections for cash/trade/merch/promo

            'determine if airtime, NTR or both
        '    If RptSel!rbcSelC6(0).Value Then           'Air time only
        '        ilIncludeNTR = 1                        'include ntr only
        '    ElseIf RptSel!rbcSelC6(1).Value Then
        '        ilIncludeNTR = 2                       'include ntr only
        '    End If

        '    If RptSel!rbcSelC4(0).Value Then   'date option
        '        RptSel!ckcAll.Value = vbChecked  'fake out selectivity
        '    Else                            'slsp option
        '        imSlsp = True
        '        mObtainCodes 5, tgSalesperson(), ilIncludeCodes, ilUseCodes()
        '    End If

        If ilListIndex = COLL_DISTRIBUTE Then           'cash distribution
            ilDistribution = True
            ilIncludeP = True
            ilIncludeW = True                           '5-9-13 include WU (ActionB) and WD (ActionD) only, for bounced and redeposited checks
            imOwner = True                                  'split up cash amonst participants
            mObtainCodes 3, tgVehicle(), ilIncludeCodes, ilUseCodes()
            ilAirNtrHard = 3                                '1-13-11 always include all types of cash (airtime, ntr, hard cost)
            tlTranType.iAirTime = True                      '5-15-19
            tlTranType.iNTR = True                          '5-15-19
            tlTranType.iHardCost = True                     '5-15-19
            tlTranType.iPymt = ilIncludeP                   '5-15-19
            tlTranType.iWriteOff = ilIncludeW               '5-15-19
            ilWhichDate = 0                                 '5-15-19 use tran date for search; doesnt matter using tran or ageing date since entire file read for past payments undistributed (search from slearliestdate)
            slEarliestDate = "1/1/1970"                     '5-15-19 need to read the entire file to find past payments undistributed
        ElseIf ilListIndex = COLL_ACCTHIST Then                   'Account History
            If RptSel!ckcSelC3(0).Value = vbChecked Then                    'include invoices
                ilIncludeI = True
                ilIncludeH = True                           'history invoice trans.
                slStr = RptSel!edcLatestCashDate.Text       '7-7-17 selective inv #
                llInvNo = Val(slStr)
            End If
            If RptSel!ckcSelC3(1).Value = vbChecked Then               'include adjustmetns
                ilIncludeA = True
                tlTranType.iAdj = True                                  '5-15-19
            End If
            If RptSel!ckcSelC3(2).Value = vbChecked Then               'include Payments
                ilIncludeP = True
                tlTranType.iPymt = True                                 '5-15-19
            End If
            If RptSel!ckcSelC3(3).Value = vbChecked Then               'include Write-offs
                ilIncludeW = True
                tlTranType.iWriteOff = True                             '5-15-19
            End If
            imMerchant = True                                      'merchandising transactions
            tlTranType.iMerch = True                                '5-15-19
            imPromotion = True                                         'promotions transactions
            tlTranType.iPromo = True                                    '5-15-19
            imInvoice = True                                        'default to sort by invoice #
            tlTranType.iInv = True                                  '5-15-19
            imHardCost = True
            tlTranType.iHardCost = True                             '5-15-19
            slStr = RptSel!edcSelCTo1.Text
            llCntrNo = Val(slStr)

            If RptSel!rbcSelCInclude(0).Value Then              'include Billing Installment records?
                ilIncludeBilling = True
                ilIncludeEarned = False
            ElseIf RptSel!rbcSelCInclude(1).Value Then        'include earned
                ilIncludeEarned = True
                ilIncludeBilling = False

            Else                                            'include both
                ilIncludeEarned = True
                ilIncludeBilling = True
            End If

            If RptSel!rbcSelCSelect(0).Value Then               'advertiser
                imAdvt = True
                mObtainCodes 0, tgAdvertiser(), ilIncludeCodes, ilUseCodes()

            ElseIf RptSel!rbcSelCSelect(1).Value Then           'agy option
                imAgency = True
                mObtainAgyAdvCodes ilIncludeCodes, ilUseCodes(), 2    'assume using lbcselection(2) for list box of direct & agencies
            End If

            'always go thru both files
            '5-18-09 determine if receivables, history, or both
            If RptSel!rbcSelC12(0).Value Then            'history only?
                ilEndFile = 1                                                '
            ElseIf RptSel!rbcSelC12(1).Value Then            'receivables only?
                ilStartFile = 2
            End If
            ilAirNtrHard = 3                                '1-13-11 always include all types of cash (airtime, ntr, hard cost)
            tlTranType.iAirTime = True                      '5-15-19
            tlTranType.iNTR = True                          '5-15-19
            tlTranType.iHardCost = True                     '5-15-19
        ElseIf ilListIndex = COLL_AGEOWNER Or ilListIndex = COLL_AGESS Then             'ageing by owner
            blIsItAnyAgeing = True
            ilIncludeBilling = True              'include any installment billing records
            ilIncludeEarned = False             'exclude any installment revenue records
            imOwner = True
            'everything gets included from RVF file for ageing
            '2-10-00 following code moved to common place below, used also with coll_ageproducer
            'ilIncludeI = True                                     'include all I (invoice) transactions
            'ilIncludeP = True                                     'include all P (payment) transactions
            'ilIncludeA = True                                     'include all A (adjustment) transactions
            'ilIncludeW = True                                     'include all W (Write off) transactions
            'mGetCashTradeOption                         'determine cash/trade/merch or promo inclusion
            'ilStartFile = 2                             'bypass PHF
            'slStr = RptSel!edcSelCFrom.Text               'Latest date to retrieve from  RVF
            'llLatestDate = gDateValue(slStr)
            'If llLatestDate = 0 Then                    'if end date not entered, use all
            '    llLatestDate = gDateValue("12/29/2069")
            'End If
            'llEarliestDate = gDateValue("1/1/1970")
            mObtainCodes 3, tgVehicle(), ilIncludeCodes, ilUseCodes()

        ElseIf ilListIndex = COLL_AGEPRODUCER Then            '2-10-00 no participant splitting of transactions
            blIsItAnyAgeing = True
            imProducer = True
            mObtainCodes 3, tgVehicle(), ilIncludeCodes, ilUseCodes()
            '7-8-02
        ElseIf ilListIndex = COLL_AGEPAYEE Then
            blIsItAnyAgeing = True
            ilIncludeBilling = True              'include any installment billing records
            ilIncludeEarned = False             'exclude any installment revenue records

            imAgency = True
            '9-16-03 use common routine to build valid codes
            mObtainAgyAdvCodes ilIncludeCodes, ilUseCodes(), 2    'assume using lbcselection(2) for list box of direct & agencies

            '7-16-02
            illoop = RptSel!cbcSet1.ListIndex
            imMajorSet = tgVehicleSets1(illoop).iCode
            imMinorSet = 0                      'unused
        ElseIf ilListIndex = COLL_AGEVEHICLE Then          '7-17-02
            blIsItAnyAgeing = True
            ilIncludeBilling = True              'include any installment billing records
            ilIncludeEarned = False             'exclude any installment revenue records

            imAirVeh = True
            imVehicle = True
            mObtainCodes 6, tgCSVNameCode(), ilIncludeCodes, ilUseCodes()

            '7-16-02
            illoop = RptSel!cbcSet1.ListIndex
            imMajorSet = tgVehicleSets1(illoop).iCode
            imMinorSet = 0                      'unused
        ElseIf ilListIndex = COLL_AGESLSP Then                              '4-17-11 make ageing by slsp a prepass
            blIsItAnyAgeing = True
            ilIncludeBilling = True              'include any installment billing records
            ilIncludeEarned = False             'exclude any installment revenue records

            imSlsp = True
            mObtainCodes 5, tgSalesperson(), ilIncludeCodes, ilUseCodes()
        End If
        'If ilListIndex = COLL_AGEOWNER Or ilListIndex = COLL_AGESS Or ilListIndex = COLL_AGEPRODUCER Or ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGEVEHICLE Or ilListIndex = COLL_AGESLSP Then       '4-17-11 make ageing by slsp prepass
        If blIsItAnyAgeing Then                 '8-27-15
            ilIncludeBilling = True              'include any installment billing records
            ilIncludeEarned = False             'exclude any installment revenue records

'            '3-21-05 see if NTR Hard cost only  7-21-08   moved to new location, after test for airtime/ntr. Dan M
'            If RptSel!ckcSelC6Add(HardCost).Value = 1 Then
'           ' If RptSel!rbcSelCInclude(4).Value = True Then
'                imHardCost = True
'            Else
'                imHardCost = False
'            End If

            'Setup all filters for the Ageing report
            'Determine the earliest MM/YY ageing dates to use: if blank assume 01/1970
            slStr = RptSel!edcSelCTo.Text
            If Trim$(slStr) = "" Then
                ilEarliestAgeMM = 1
                ilEarliestAgeYY = 1970
            Else
                ilRet = gParseItem(slStr, 1, "/", slCode)
                ilEarliestAgeMM = Val(slCode)

                ilRet = gParseItem(slStr, 2, "/", slCode)
                ilEarliestAgeYY = Val(slCode)
                If (ilEarliestAgeYY >= 0) And (ilEarliestAgeYY <= 69) Then
                    ilEarliestAgeYY = 2000 + ilEarliestAgeYY
                ElseIf (ilEarliestAgeYY >= 70) And (ilEarliestAgeYY <= 99) Then
                    ilEarliestAgeYY = 1900 + ilEarliestAgeYY
                End If

            End If

             'Determine the latest MM/YY ageing dates to use: if blank assume 12/2069
            slStr = RptSel!edcSelCTo1.Text
            If Trim$(slStr) = "" Then
                ilLatestAgeMM = 12
                ilLatestAgeYY = 2069
            Else
                ilRet = gParseItem(slStr, 1, "/", slCode)
                ilLatestAgeMM = Val(slCode)

                ilRet = gParseItem(slStr, 2, "/", slCode)
                ilLatestAgeYY = Val(slCode)
                If (ilLatestAgeYY >= 0) And (ilLatestAgeYY <= 69) Then
                    ilLatestAgeYY = 2000 + ilLatestAgeYY
                ElseIf (ilLatestAgeYY >= 70) And (ilLatestAgeYY <= 99) Then
                    ilLatestAgeYY = 1900 + ilLatestAgeYY
                End If
            End If

            'everything gets included from RVF file for ageing
            ilIncludeI = True                                     'include all I (invoice) transactions
            ilIncludeP = True                                     'include all P (payment) transactions
            ilIncludeA = True                                     'include all A (adjustment) transactions
            ilIncludeW = True                                     'include all W (Write off) transactions
            tlTranType.iInv = True                      '5-15-19
            tlTranType.iPymt = True                     '5-15-19
            tlTranType.iAdj = True                      '5-15-19
            tlTranType.iWriteOff = True                 '5-15-19
            mGetCashTradeOption tlTranType                         'determine cash/trade/merch or promo inclusion

            '7-21-08 allow hard cost/airtime/ntr combinations. Dan M
            If RptSel!ckcSelC6Add(Airtime).Value = 1 Then
                ilIncludeNTR = 1            'ntr/airtime flag
                tlTranType.iAirTime = True     '5-15-19
            End If
           If RptSel!ckcSelC6Add(NTR).Value = 1 Then
                ilAirNtrHard = FLAG_NTR     'ntr/hardcost flag
                ilIncludeNTR = ilIncludeNTR + 2
                tlTranType.iNTR = True          '5-15-19
            End If
            If RptSel!ckcSelC6Add(HardCost).Value = 1 Then
                ilAirNtrHard = ilAirNtrHard + FLAG_HARD_COST
                If ilIncludeNTR = 1 Then
                    ilIncludeNTR = 3
                    tlTranType.iHardCost = True     '5-15-19
                ElseIf ilIncludeNTR = 0 Then
                    ilIncludeNTR = 2
                    tlTranType.iHardCost = True          '5-15-19
                End If
            End If


            ilStartFile = 2                             'bypass PHF
            '8-27-15 Cash (PO, PI, JE) and Billing (IN, AN) dates are now separate
            '8-28-19 use csi calendar control vs edit box
'           slStr = RptSel!edcSelCFrom.Text               'Latest date to retrieve from  RVF
            slStr = RptSel!CSI_CalFrom.Text               'Latest date to retrieve from  RVF
            llLatestDate = gDateValue(slStr)
            If llLatestDate = 0 Then                    'if end date not entered, use all
                llLatestDate = gDateValue("12/29/2069")
            End If
            llEarliestDate = gDateValue("1/1/1970")

'            slStr = RptSel!edcLatestCashDate.Text       'Latest cash date to retrieve from  RVF
            slStr = RptSel!CSI_CalTo.Text                'Latest cash date to retrieve from  RVF
            llLatestCashDate = gDateValue(slStr)
            If llLatestCashDate = 0 Then                    'if end date not entered, use all
                llLatestCashDate = gDateValue("12/29/2069")
            End If
            llEarliestCashDate = gDateValue("1/1/1970")
            
            slEarliestDate = Format(llEarliestDate, "ddddd")    '5-15-19
            '11/15/19 determine which is later, latest bill date or latest cash date.  use that for retrieval, further filtering done after all recds retrieved from rvf
            If llLatestDate > llLatestCashDate Then
                slLatestDate = Format(llLatestDate, "ddddd")
            Else
                slLatestDate = Format(llLatestCashDate, "ddddd")
            End If
'            slLatestDate = Format(llLatestDate, "ddddd")        '5-15-19

        ElseIf ilListIndex = COLL_AGEMONTH Then     'set the defaults since no selectivity
            blIsItAnyAgeing = True
            ilIncludeBilling = True              'include any installment billing records
            ilIncludeEarned = False             'exclude any installment revenue records

            'Setup all filters for the Ageing report
            'Determine the earliest MM/YY ageing dates to use: if blank assume 01/1970
           
             ilEarliestAgeMM = 1
             ilEarliestAgeYY = 1970
        
             'Determine the latest MM/YY ageing dates to use: if blank assume 12/2069
            ilLatestAgeMM = 12
            ilLatestAgeYY = 2069

            'everything gets included from RVF file for ageing
            ilIncludeI = True                                     'include all I (invoice) transactions
            ilIncludeP = True                                     'include all P (payment) transactions
            ilIncludeA = True                                     'include all A (adjustment) transactions
            ilIncludeW = True                                     'include all W (Write off) transactions
            imTrade = False
            imCash = True
            tlTranType.iInv = True                      '5-15-19
            tlTranType.iPymt = True                     '5-15-19
            tlTranType.iAdj = True                      '5-15-19
            tlTranType.iWriteOff = True                 '5-15-19
            tlTranType.iTrade = False                   '5-15-19
            tlTranType.iCash = True                     '5-15-19
            
            ilIncludeNTR = 1            'ntr/airtime flag
            tlTranType.iNTR = True          '5-15-19
            ilAirNtrHard = FLAG_NTR     'ntrflag
            If RptSel!ckcSelC6Add(0).Value = vbChecked Then            'hardcost only option for selection
                ilIncludeNTR = ilIncludeNTR + 2
                tlTranType.iHardCost = True             '5-15-19
            End If
            
            If RptSel!ckcSelC6Add(HardCost).Value = 1 Then
                ilAirNtrHard = ilAirNtrHard + FLAG_HARD_COST
                If ilIncludeNTR = 1 Then
                    ilIncludeNTR = 3
                    tlTranType.iHardCost = True         '5-15-19
                ElseIf ilIncludeNTR = 0 Then
                    ilIncludeNTR = 2
                    tlTranType.iHardCost = True
                End If
            End If
            
            'This ageing doesnt use any date input, all is gathered thru the last reconciled date in site
            gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slStr
            llLatestDate = gDateValue(slStr)
            If llLatestDate = 0 Then                    'if end date not entered, use all
                llLatestDate = gDateValue("12/29/2069")
            End If
            llEarliestDate = gDateValue("1/1/1970")
            
            'slstr contains the last date reconciled, use as default
            llLatestCashDate = gDateValue(slStr)
            If llLatestCashDate = 0 Then                    'if end date not entered, use all
                llLatestCashDate = gDateValue("12/29/2069")
            End If
            llEarliestCashDate = gDateValue("1/1/1970")

            slEarliestDate = Format(llEarliestDate, "ddddd")     '5-15-19
            slLatestDate = Format(llLatestDate, "ddddd")         '5-15-19
            ilStartFile = 2                             'bypass PHF

        End If
    End If
    'if user entered start date prior to the prev reconciling period, go thru the history file (PVF)
    'Otherwise, just go thru RVF
    gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slStr       'end date of prior reconciling period
    llDate = gDateValue(slStr)
    If ilStartFile = 1 Then
        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            btrDestroy hmRvr
            Exit Sub
        End If
        imRvfRecLen = Len(tmRvf)
    Else
        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read only RVF file (not PHF)
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            btrDestroy hmRvr
            Exit Sub
        End If
        imRvfRecLen = Len(tmRvf)
    End If

    'If imOwner Then    '2-10-00 comment out testing for owner, always gather the sales sources
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
    'End If
    'For ilLoopOnFile = ilStartFile To ilEndFile Step 1                 '2 passes, first History, then Receivables
    'changed to go always go thru PVF and RVF.  Code above still represents testing to see if
    'there is history (or receivables) based on the last date invoiced.  Retain in case
    'needed later
    llSeqNo = 0                                     'init seq # that is incremented every time a new trans. is processed
                                                    'when distribution $ for participants, each set created for the same trans
                                                    'is given the same seq #.
    'depending on report, may need to get participant splits prior to the requested report.
    'i.e. For cash distribution could be trying to process an old transaction whose
    'participant pcentages wont be in memory.  Then an erroneous line would be printed.
    'get participant percentages at least 2 year back to ensure that all transactions printed have
    'a participant reference
    llParticipantAdjustedDate = llEarliestDate - 730
    'gCreatePIFForRpts llEarliestDate, tmPifKey(), tmPifPct(), RptSel
    gCreatePIFForRpts llParticipantAdjustedDate, tmPifKey(), tmPifPct(), RptSel

    tmRvr.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmRvr.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmRvr.lGenTime = lgNowTime
    
    '5-22-19 change manner of obtaining the username to show on Acct History because it is encrypted.  Previously using
    'subreports but caused severe slowness.  Based on the users in memory that are decrypted, create
    'parallel records in txr.  For each transaction to print, point the transaction to the txr record
    'The pointer from rvr to be used had to be an extra field, not used in the Acct HIstory (rvrprodpct) since
    'rvrurfcode was overlayed and used for other features
    If igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_ACCTHIST Then
        For illoop = LBound(tgPopUrf) To UBound(tgPopUrf) - 1
            tmTxr.lGenTime = tmRvr.lGenTime
            tmTxr.iGenDate(0) = tmRvr.iGenDate(0)
            tmTxr.iGenDate(1) = tmRvr.iGenDate(1)
            tmTxr.iGeneric1 = tgPopUrf(illoop).iCode
            tmTxr.sText = Trim$((tgPopUrf(illoop).sRept))        'name to show on report, decrypted
            ilRet = btrInsert(hmTxr, tmTxr, imTxrRecLen, INDEXKEY0)
        Next illoop
    End If
    
    'For ilLoopOnFile = 1 To 2 Step 1                 '2 passes, first History, then Receivables
    For ilLoopOnFile = ilStartFile To ilEndFile Step 1 '8-31-99
        'handles and buffers for PHF and RVF will be the same

        '5-15-19 use generalized routine with extended reads for speed
        If ilLoopOnFile = 1 Then     'only history selected
            ilRet = gObtainPhfOrRvf(RptSel, slEarliestDate, slLatestDate, tlTranType, tlRvf(), 1, ilWhichDate)      'history, use tran date or ageing date
        Else                                        'only receivables selected
            ilRet = gObtainPhfOrRvf(RptSel, slEarliestDate, slLatestDate, tlTranType, tlRvf(), 2, ilWhichDate) 'receivables, use tran date or ageing date
        End If
        
        'Date: 02/20/2020 'option to suppress zero balance; create array of zero balanced RVF records
        ReDim lmZBRvfCode(0 To 0)
        If RptSel!ckcSuppressZB.Value = 1 Then
            mZeroBalance lmZBRvfCode()      'create array of RVF codes to exclude
        End If
        
        '--------------------------------------------------------
'        ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        For llLoopOnTrans = LBound(tlRvf) To UBound(tlRvf) - 1          '5-15-19 loop on the array of transactions returned
            tmRvf = tlRvf(llLoopOnTrans)
'        Do While ilRet = BTRV_ERR_NONE
            
            'Date: 2020-02-20 suppress RVF records that are zero balanced
            bZeroBalanced = False
            If UBound(lmZBRvfCode) > 0 Then
                'check if RVF code is on the list; if so, exclude from the report
                bZeroBalanced = IIF(mBinarySearchRvf(tmRvf.lCode, lmZBRvfCode()) < 0, False, True)
            End If

            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slRvfTranDate

            If Not bZeroBalanced Then
                ilFoundOption = False
                If RptSel!ckcAll.Value = vbChecked And (igRptCallType = INVOICESJOB And ilListIndex = INV_REGISTER And Not imNTR) Then                            'report all
                    ilFoundOption = True
                ElseIf imAdvt Then                              'advt option
                    If igRptCallType = INVOICESJOB Then
                        ilListBoxInx = 5
                    End If
                    ilValidSelect = mFindMatchingItem(ilIncludeCodes, ilUseCodes())
                    If ilValidSelect Then                     'valid vehicle
                        ilFoundOption = True
                    End If
    
                ElseIf imAgency Then                              'agency option
                    If igRptCallType = INVOICESJOB Then
                        ilListBoxInx = 1
                    ElseIf igRptCallType = COLLECTIONSJOB Then
                        'If ilListIndex = COLL_PAYHISTORY Then    '7-8-02
                        '    ilListBoxInx = 1    '2
                        If ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_ACCTHIST Then     '7-8-02
                            ilListBoxInx = 2                        'agy & direct advt list box
                        ElseIf ilListBoxInx = COLL_AGEVEHICLE Then      '7-17-02
                            ilListBoxInx = 6
                        End If
                    End If
                    If ilIncludeCodes Then
                        For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                            If ilUseCodes(ilTemp) < 0 Then      '11-11-03 do not test tmrvf.iagfcode, because there are some transactions that do have agencies
                                                                'as well as the advt having been changed to non-direct
                            'If tmRvf.iAgfCode = 0 Then              'direct, codes has been negated to know to test advertisr code
                                If ilUseCodes(ilTemp) = -tmRvf.iAdfCode Then
                                    ilFoundOption = True
                                    Exit For
                                End If
                            Else
                                If ilUseCodes(ilTemp) = tmRvf.iAgfCode Then
                                    ilFoundOption = True
                                    Exit For
                                End If
                            End If
                        Next ilTemp
                    Else
                        ilFoundOption = True        '8/23/99 when more than half selected, selection fixed
                        For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                            If ilUseCodes(ilTemp) < 0 Then          '11-11-03 do not test tmrvf.iagfcode, because there are some transactions that do have agencies
                                                                    'as well as the advt having been changed to non-direct
                            'If tmRvf.iAgfCode = 0 Then              'direct, codes has been negated to know to test advertisr code
                                If ilUseCodes(ilTemp) = -tmRvf.iAdfCode Then
                                    ilFoundOption = False
                                    Exit For
                                End If
                            Else
                                If ilUseCodes(ilTemp) = tmRvf.iAgfCode Then
                                    ilFoundOption = False
                                    Exit For
                                End If
                            End If
                        Next ilTemp
                    End If
                ElseIf imVehicle Then      'vehicle  option
                    ilValidSelect = mFindMatchingItem(ilIncludeCodes, ilUseCodes())
                    If ilValidSelect Then                     'valid vehicle
                        ilFoundOption = True
                    End If
    
                ElseIf imNTR Then                              'NTR Item Billing option
                    ilFoundOption = False
                    If tmRvf.iMnfItem > 0 Then              'all ntr must have an mnf item reference
                        If igRptCallType = INVOICESJOB Then
                            ilValidSelect = mFindMatchingItem(ilIncludeCodes, ilUseCodes())        'test if NTR item matches selections
                            If ilValidSelect Then                     'valid vehicle
                                ilFoundOption = True
                            End If
                        End If
                    End If
                ElseIf imSS Then                    'Sales source
                    'Determine what this transactions sales Source is
                    '4-17-11 slsp table in memory
    '                If tmRvf.iSlfCode <> tmSlf.iCode And tmRvf.iSlfCode <> 0 Then      'only read if not already in mem
    '                    tmSlfSrchKey.iCode = tmRvf.iSlfCode
    '                    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '                End If
                    
                    ilRet = gBinarySearchSlf(tmRvf.iSlfCode)
                    If ilRet = -1 Then
                        ilFoundOption = False
                    Else
                        tmSlf = tgMSlf(ilRet)
    
                        ilMatchSSCode = mFindMatchSSCode(tmSlf.iSofCode, tlSofList())   'get the sales source for this sof trans.
                        If ilIncludeCodes Then
                            For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                                If ilUseCodes(ilTemp) = ilMatchSSCode Then
                                    ilFoundOption = True
                                    Exit For
                                End If
                            Next ilTemp
                        Else
                            ilFoundOption = True        '10-18-02 when more than half selected, selection fixed
                            For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                                If ilUseCodes(ilTemp) = ilMatchSSCode Then
                                    ilFoundOption = False
                                    Exit For
                                End If
                            Next ilTemp
                        End If
                    End If
                ElseIf imSO Then                    '11-25-06 sales origin
                    If Not RptSel!rbcSelC8(0).Value Then        'totals by vehicle?
                        ilValidSelect = mFindMatchingItem(ilIncludeCodes, ilUseCodes())
                        If ilValidSelect Then                     'valid vehicle, now check for sales origins
                            ilFoundOption = True
                        End If
    
                    End If
    
                    '  Retain the following code in case need to implement option to select sales origins
                    '  currently no selection for sales regions.  The screen has the option hidden (plcSelC10)
                    '  if option is opened up, need to read in the Sales Source record to get the sales origin
                    'If tmRvf.iSlfCode <> tmSlf.iCode And tmRvf.iSlfCode <> 0 Then      'only read if not already in mem
                    '    tmSlfSrchKey.iCode = tmRvf.iSlfCode
                    '    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    'End If
                    'ilMatchSSCode = mFindMatchSSCode(tmSlf.iSofCode, tlSofList())   'get the sales source for this sof trans.
                    'For ilTemp = 0 To UBound(ilMnfSSCodes) - 1
                    '    If ilMnfSSCodes(ilTemp) = ilMatchSSCode Then
                    '        'check to see if local, natl, regional is included
                    '        'if sales origin opened up as option, test here for selection.  need to get the MNF Sales Source record
                    '        'to test the origin.  Currently, only the Sales Source code has been built into memory
                    '        Exit For
                    '    End If
                    'Next ilTemp
                Else
    
                    ilFoundOption = True                    'slsp & owners are filtered later
                    '1-20-06 filter any selected sales sources
                    If ((igRptCallType = COLLECTIONSJOB) And (ilListIndex = COLL_AGESS Or ilListIndex = COLL_AGEOWNER Or ilListIndex = COLL_AGEPRODUCER)) Then
                        'Determine what this transactions sales Source is
                        If RptSel!ckcAllGroups.Value = vbUnchecked Then   'if all sales sources, always include other
    
                            ilFoundOption = mTestValidSalesSource(tmRvf.iSlfCode, tlSofList(), ilMnfSSCodes())
                            
    '                        ilRet = gBinarySearchSlf(tmRvf.iSlfCode)
    '                        If ilRet = -1 Then
    '                            ilMatchSSCode = 0
    '                        Else
    '                            tmSlf = tgMSlf(ilRet)
    '                            ilMatchSSCode = mFindMatchSSCode(tmSlf.iSofCode, tlSofList())   'get the sales source for this sof trans.
    '                        End If
    '                        ilFoundOption = False
    '                        For ilTemp = 0 To UBound(ilMnfSSCodes) - 1
    '                            If ilMnfSSCodes(ilTemp) = ilMatchSSCode Then
    '                                ilFoundOption = True
    '                                Exit For
    '                            End If
    '                        Next ilTemp
    '                        If ilFoundOption Then
    '                        End If
                        End If
                    End If
                End If
                
                If (igRptCallType = INVOICESJOB And ilListIndex = INV_REGISTER) And (imAgency Or imSlsp) Then       'invoice register, either by agy or slsp has additional filter by sales source
                    If ilFoundOption Then
                        ilFoundOption = mTestValidSalesSource(tmRvf.iSlfCode, tlSofList(), ilMnfSSCodes())
                    End If
                End If
    
                'test for the valid transactions to include
                ilTransFound = False
                If (ilIncludeI) And (Left$(tmRvf.sTranType, 1) = "I") Then
                    ilTransFound = True
                End If
                If (ilIncludeP) And (Left$(tmRvf.sTranType, 1) = "P") Then
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
                'determine if installment billing transactions and what types to include (billingor revenue)
                If Trim$(tmRvf.sType) = "" Or (tmRvf.sType = "I" And ilIncludeBilling = True) Or (tmRvf.sType = "A" And ilIncludeEarned = True) Then
                    'ilTransFound = True
                    'if transaction type to include is already false, dont change the flag
                Else
                    ilTransFound = False
                End If
    
                'test for selective contracts
                If (igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_ACCTHIST) Or (igRptCallType = INVOICESJOB And ilListIndex = INV_TAXREGISTER) Or (igRptCallType = INVOICESJOB And ilListIndex = INV_DISTRIBUTE) Then
                    If ilTransFound And (llCntrNo = 0 Or tmRvf.lCntrNo = llCntrNo) Then
                        ilTransFound = True
                        If igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_ACCTHIST Then      '7-7-17 filter on inv #
                            If (llInvNo = 0 Or tmRvf.lInvNo = llInvNo) Then
                            Else
                                ilTransFound = False
                            End If
                        End If
                    Else
                        ilTransFound = False
                    End If
                End If
                'Else
                '3-12-09 always to test cash/trade/merchandising/promotions
                If (tmRvf.sCashTrade = "C" And Not imCash) Then 'Or (tmRvf.sCashTrade = "C" And ilIncludeNTR = 1 And tmRvf.imnfItem = 0) Or (tmRvf.sCashTrade = "C" And ilIncludeNTR = 0 And tmRvf.imnfItem <> 0) Then   'if cash transaction and cash not requested, ignore. Or, if
                                                                            'cash and ntr only requested which doesnt have an Item bill pointer, ignore it
                                                                            'or cash & air time only requested and its an NTR
                    ilTransFound = False
                ElseIf tmRvf.sCashTrade = "T" And Not (imTrade) Then
                    ilTransFound = False
                ElseIf tmRvf.sCashTrade = "M" And Not (imMerchant) Then
                    ilTransFound = False
                ElseIf tmRvf.sCashTrade = "P" And Not (imPromotion) Then
                    ilTransFound = False
                End If
                'End If
    
    
                If ilIncludeNTR = 1 And tmRvf.iMnfItem > 0 Then     'if air time only and this is an NTR trans, ignore it
                    ilTransFound = False
                End If
                If ilIncludeNTR = 2 And tmRvf.iMnfItem = 0 Then     'if NTR only and this is not an NTR trans, ignore it
                    ilTransFound = False
                End If
    
                '3-17-05 exlude hard cost if not hard cost only    6-9-08 allowed multiple choice hard cost, airtime, ntr
                If (tmRvf.iMnfItem > 0) Then
                    ilRet = gIsItHardCost(tmRvf.iMnfItem, tmNTRMNF())
                    If (igRptCallType = COLLECTIONSJOB) Then
                        'If (ilListIndex <> COLL_ACCTHIST) Then 'if acct history report,  dont test it
                            If (ilAirNtrHard And FLAG_HARD_COST) = FLAG_HARD_COST Then     'want to see hardcost.  See ntr?
                                If ((ilAirNtrHard And FLAG_NTR) <> FLAG_NTR) And Not ilRet Then     'Don't want to see ntr
                                    ilTransFound = False
                                End If
                            Else
                                If ilRet = True Then
                                    ilTransFound = False
                                End If
                            End If
                        'End If
                    '5-11-09 test for hard cost only when ntr are included.  I think it was by design that Hard cost was included with NTR,
                    'and then user could get Hard Cost Only on a separate report.
                    ElseIf igRptCallType = INVOICESJOB Then
                        If ilListIndex = INV_REGISTER Then      'Invoice register has Hard cost only option
                            If imHardCost Then          'hard cost only?
                                If Not ilRet Then
                                    ilTransFound = False
                                End If
                            Else            'user has not requested hard cost
                                If ilRet Then           'its a hard cost, need to exclude
                                    ilTransFound = False
                                End If
                            End If
                        End If
                    End If
                End If
                If ilFoundOption = True Then
    
                    If ((igRptCallType = COLLECTIONSJOB) And (ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGEVEHICLE)) Or (igRptCallType = INVOICESJOB And ilListIndex = INV_REGISTER) Then       'Ageing by Payee or Vehicle are only options with vehicle group sorts
                        'check selectivity of vehicle groups
                        If Not ilCkcAll Or (imMajorSet > 0 And Not ilCkcAll) Then
                            ilFoundOption = False
                            'For ilLoop = LBound(tgMVef) To UBound(tgMVef)
                            '    If tgMVef(ilLoop).iCode = tmRvf.iAirVefCode Then
                                illoop = gBinarySearchVef(tmRvf.iAirVefCode)
                                If illoop <> -1 Then
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
    
                                    If Not ilCkcAll Then
                                        'If ((igRptCallType = COLLECTIONSJOB) And (ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGEVEHICLE)) Then
                                        '    ilFoundOption = True        'always allow tran without vehicle group defintion to be reported on the ageing by payee
                                        'End If
                                        For ilTemp = 0 To UBound(ilMnfcodes) - 1
                                            If ilMnfcodes(ilTemp) = ilWhichset Then
                                                ilFoundOption = True
                                                Exit For
                                            End If
                                        Next ilTemp
                                        If ilFoundOption Then
                            '                Exit For
                                        Else                        'default no vehicle group to always include if major sort is by vehicle group
                                            If ilWhichset = 0 Then
                                                ilFoundOption = True
                                            End If
                                        End If
                                    Else
                                        ilFoundOption = True
                                    End If
                                Else
                                    If (igRptCallType = COLLECTIONSJOB) And (ilListIndex = COLL_AGEPAYEE) Then
                                        ilFoundOption = True        'always allow a tran without an airing vehicle to be  reported on the ageing by payee
                                    End If
    
                                End If
                            'Next ilLoop
                        Else
                            ilFoundOption = True
                        End If
    
    
                        '7-15-08 if the record is valid for inclusion/exclusion of politicals and non-politicals
                        ilIsItPolitical = gIsItPolitical(tmRvf.iAdfCode)           'its a political, include this contract?
                        If ilIsItPolitical Then                 'its a political
                            If Not imPolit Then                 'Include politicals?
                                ilFoundOption = False           'no, exclude them
                            End If
                        Else                                    'not a plitical
                            If Not imNonPolit Then              'include non politicals?
                                ilFoundOption = False           'no, exclude them
                            End If
                        End If
                    End If
                End If
    
                gPDNToLong tmRvf.sNet, llAmt
                gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slStr
                llDate = gDateValue(slStr)                  'convert trans date to test if within requested limits
    
                'If ilListIndex = COLL_CASH And igRptCallType = COLLECTIONSJOB Then     'for Cash Receipts, have the option to select on deposit date or date entered
                '    If RptSel!rbcSelCSelect(1).Value Then   'date entered, put the date to test in common field
                '        gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slStr 'convert date entered to test if within requested limits
                '        llDate = gDateValue(slStr)
                '    End If
                '    If llCheckNo <> 0 And llCheckNo <> tmRvf.lCheckNo Then  'see if user wants a specific check #
                '        ilTransFound = False
                '    End If
                'End If
    
                gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slStr 'convert date entered to test if within requested limits
                                                            'for prior month distributed in Cash Distribution report
    
                llDateEntrd = gDateValue(slStr)                  'convert trans date to test if within requested limits
                'valid record must be an "Invoice" type, non-zero amount, and transaction date within the start date of the
                'cal year and end date of the current cal month requested
                'If (ilTransFound And llAmt <> 0 And llDate >= llEarliestDate And llDate <= llLatestDate And ilFoundOption) Then         'looking for Invoice types only
                'show $0 so there's an audit trail of all invoices, etc created
    
    
                ilValidDates = False
                
                If blIsItAnyAgeing Then
                    If (tmRvf.sTranType = "AN" Or tmRvf.sTranType = "IN") Then
                        If (ilTransFound) And (llDate >= llEarliestDate And llDate <= llLatestDate) Then
                            ilValidDates = True
                            ilWhichMonth = 1            'current month date
                        End If
                    Else        'PO, PI, JE
                        If (ilTransFound) And (llDate >= llEarliestCashDate And llDate <= llLatestCashDate) Then
                            ilValidDates = True
                            ilWhichMonth = 1            'current month date
                        End If
                    End If
                'non-ageing report
                Else
                    If (ilTransFound) And (llDate >= llEarliestDate And llDate <= llLatestDate) Then
                        ilValidDates = True
                        ilWhichMonth = 1            'current month date
                    End If
                End If
    
                If igRptCallType = COLLECTIONSJOB Then
                    If ilListIndex = COLL_DISTRIBUTE Then       'Check for PO trans that havent been distributed yet
                        ilValidDates = False
                        
                        '5-9-13 More special testing for Cash Distribution to adjust for bounced and redeposited checks, which are journal entries
                        'WU with action type = B are bounced checked, WD with action type "D" are redeposited checks
                        If (tmRvf.sTranType = "WU" And tmRvf.sAction = "B") Or (tmRvf.sTranType = "WD" And tmRvf.sAction = "D") Or (Left$(tmRvf.sTranType, 1) = "P") Then
                            'Use date entered instead of transaction date for Distribution report
                            gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slStr
                            llDate = gDateValue(slStr)                  'convert entry date to test if within requested limits
                            If (ilTransFound And llDate >= llEarliestDate And llDate <= llLatestDate) Then
                                ilValidDates = True
                                ilWhichMonth = 1            'current month date
                            End If
                            If (ilTransFound And llDateEntrd >= llEarliestDate And llDateEntrd <= llLatestDate And llDate < llEarliestDate) Then
                                If tmRvf.sTranType = "PO" And llDate < llEarliestDate Then          'on account or in advance payment, accumulate
                                                                        'for total not distributed yet if trans date in prior month
                                    gPDNToLong tmRvf.sNet, llTransNet
                                    llUndistCash = llUndistCash + llTransNet
                                Else                                'PI FROM PO applied
                                    ilValidDates = True
                                    ilWhichMonth = 2        'prior month distributed
                                End If
                            Else                            'date entered are not with the requested
                                'is this a PO with date entered and trans date all in past, accum $ not distributed
                                If (ilTransFound And llDateEntrd < llEarliestDate And llDate < llEarliestDate And tmRvf.sTranType = "PO") Then
                                    gPDNToLong tmRvf.sNet, llTransNet
                                    llUndistCash = llUndistCash + llTransNet
                                End If
                            End If
                        End If
                    'any ageing needs to test for earliest/latest ageing MM/YY filter
                    'ElseIf ilListIndex = COLL_AGEOWNER Or ilListIndex = COLL_AGESS Or ilListIndex = COLL_AGEPRODUCER Or ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGEVEHICLE Or ilListIndex = COLL_AGESLSP Or ilListIndex = COLL_AGEMONTH Then
                     ElseIf blIsItAnyAgeing Then                   '8-27-15
                        If ilValidDates Then
                            'TTP 10130 - Ageing report: if the include ageing MM/YY earliest/latest dates go across a year, it fails to find any data
'                            If tmRvf.iAgingYear >= ilEarliestAgeYY And tmRvf.iAgingYear <= ilLatestAgeYY Then
'                                   'valid year, find if valid ageing month
'                                'If tmRvf.iAgePeriod < ilEarliestAgeMM Or tmRvf.iAgePeriod > ilLatestAgeYY Then
'                                If tmRvf.iAgePeriod < ilEarliestAgeMM Or tmRvf.iAgePeriod > ilLatestAgeMM Then
'                                    ilValidDates = False
'                                End If
'                            Else
'                                ilValidDates = False
'                            End If
                            sltmpEarliestDate = ilEarliestAgeMM & "/15/" & ilEarliestAgeYY
                            sltmpLatestDate = ilLatestAgeMM & "/15/" & ilLatestAgeYY
                            sltmpAgingDate = tmRvf.iAgePeriod & "/15/" & tmRvf.iAgingYear
                            ilValidDates = True
                            If DateValue(sltmpAgingDate) < DateValue(sltmpEarliestDate) Then
                                ilValidDates = False
                            End If
                            If DateValue(sltmpAgingDate) > DateValue(sltmpLatestDate) Then
                                ilValidDates = False
                            End If
                        End If
                    End If
                End If
    
                'dates have been determined.  Valid transaction to process?
                If ((ilValidDates) And (ilTransFound) And (ilFoundOption)) Then
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
                        mFakeChf
                        If tmChf.lCntrNo <> tmRvf.lCntrNo Or tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M" Then
                            ilMatchCntr = False
                        End If
                        'test for trade later, found that if contract changed from cash to trade and the contract was generated as trade (prior to changing)
                        'the transaction wasnt printed, causing out of balance with ageing and last reconciled $
                        'filter by rvfCashTrade flag only
                        'If tmChf.iPctTrade = 100 And Not imTrade Then  'trades?
                        '    ilMatchCntr = False
                        'End If
                        ilRet = BTRV_ERR_NONE           '5-17-03 reset error return- fake contract may have been created if contract # wasnt found
                    Else                            'contract # not present
                        mFakeRvrSlsp                'setup slsp & comm from RVf
                        ilMatchCntr = True
                        ilRet = BTRV_ERR_NONE
                    End If
                    'If ((ilRet = BTRV_ERR_NONE) And (ilMatchCntr)) Then
    
                    '4-15-11  see if user allowed to see contract
                    ilOKtoSeeVeh = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode())
    
                     If (ilMatchCntr) And (ilOKtoSeeVeh) Then
    
                        'set flag if its an Externally billed contract
                        ilExternalAdv = False
                        For illoop = 0 To UBound(ilExternalAdvList) - 1
                            If tmRvf.iAdfCode = ilExternalAdvList(illoop) Then
                                ilExternalAdv = True
                                Exit For
                            End If
                        Next illoop
    
                        LSet tmRvr = tmRvf
    
    '                    tmRvr.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    '                    tmRvr.iGenDate(1) = igNowDate(1)
    '                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    '                    tmRvr.lGenTime = lgNowTime
                        If ilLoopOnFile = 1 Then
                            tmRvr.sSource = "H"                    'let crystal know these records are histroy/receivables (vs contracts)
                        Else
                            tmRvr.sSource = "R"
                        End If
                        tmRvr.iMnfGroup = ilWhichMonth              '1=currnt month, 2=prior month distri, 3=priormonth undistr
                                                                    '(for cash distribution reports)
                        gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slCode
                        llDate = gDateValue(slCode)
                        gPDNToLong tmRvf.sNet, tmRvr.lDistAmt       'orig amt paid
                        imGrossNeg = False
                        imNetNeg = False
    
                        gPDNToLong tmRvf.sGross, llTransGross
                        gPDNToLong tmRvf.sNet, llTransNet
                        If llTransGross < 0 Then
                            llTransGross = -llTransGross
                            imGrossNeg = True
                        End If
                        If llTransNet < 0 Then
                            llTransNet = -llTransNet
                            imNetNeg = True
                        End If
    
                        'ilHowMany = 0
                        ilHowManyDefined = 0
                        '12-26-01 if Inv Reg is by office/vehicle, variable imSlsp is set to false & imvehicle is set to true.  Need to
                        'go thru some of the slsp code for split revenue
                        '4-17-11 make AGeing by Slsp a prepass and test for slsp selectivity
    '                    If (imSlsp And igRptCallType <> COLLECTIONSJOB) Or (igRptCallType = INVOICESJOB And ilListIndex = INV_REGISTER And RptSel!rbcSelCSelect(6).Value) Then
                        If (imSlsp) Or (igRptCallType = INVOICESJOB And ilListIndex = INV_REGISTER And RptSel!rbcSelCSelect(6).Value) Then       '2-25-19 slsp option to split
                        'slsp option, see if slsp splits and how many
                            '3-20-18 see if total splits exceed 100%, which will cause the total invoice register to be out of balance from the invoices.
                            'if exceeding 100%, each slsp gets their total share, and there will be no extra pennies to give the last slsp
                            llTotalSplits = 0
                            If ((imSlsp) And (igRptCallType = COLLECTIONSJOB) And RptSel!ckcOption.Value = vbUnchecked) Then
                                llTotalSplits = 1000000
                                ilHowManyDefined = 1
                            Else
                                For illoop = 0 To 9 Step 1
                                    If tmChf.iSlfCode(illoop) > 0 Then
                                        ilHowManyDefined = ilHowManyDefined + 1
                                        llTotalSplits = llTotalSplits + tmChf.lComm(illoop)
                                    End If
                                Next illoop
                                If ilHowManyDefined = 1 And tmChf.lComm(0) = 0 Then                       'theres only 1 slsp
                                    tmChf.lComm(0) = 1000000                'xxx.xxxx
                                Else
                                    If ilHowManyDefined > 1 Then
                                        ilHowManyDefined = 10                      'more than 1 slsp, process all 10 because some in the
                                                                            'middle may not be used
                                    End If
                                End If
                                'ilHowManyDefined = ilHowMany            'these % do not have to add up to 100, so the actual #
                                                                        'of slsp to process can be the same (its differnt from
                                                                        'the owners, which have to add to 100%)
                            End If
                        ElseIf imOwner Or imProducer Then                         'owner, get the vehicle groups associated with the vehicle
                            '4-17-11 slsp array in memory, no reading of record
    '                        If tmRvf.iSlfCode <> tmSlf.iCode And tmRvf.iSlfCode <> 0 Then      'only read if not already in mem
    '                            tmSlfSrchKey.iCode = tmRvf.iSlfCode
    '                            ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '                        End If
    '                        If tmRvf.iSlfCode = 0 Or ilRet <> BTRV_ERR_NONE Then              '2-15-01 slsp isnt indicated in rvf record, init the fields used so
    '                                                                'that incorrect data isnt used
    '                            tmSlf.iSofCode = 0
    '                            tmSlf.iCode = 0
    '                            ilMatchSSCode = 0
    '                        End If
                            
                            ilRet = gBinarySearchSlf(tmRvf.iSlfCode)
                            If ilRet = -1 Or tmRvf.iSlfCode = 0 Then     'not found or no slsp defined in record
                                tmSlf.iSofCode = 0
                                tmSlf.iCode = 0
                                ilMatchSSCode = 0
                            Else
                                tmSlf = tgMSlf(ilRet)
                                ilMatchSSCode = mFindMatchSSCode(tmSlf.iSofCode, tlSofList())
    
                            End If
    
                            '2-28-03 determine update method
                            ilAskforUpdate = False      'assume to split the revenue share, everything goes into RVF
                            For illoop = LBound(tlMnf) To UBound(tlMnf) - 1
                                If tlMnf(illoop).iCode = ilMatchSSCode Then
                                    If Trim$(tlMnf(illoop).sUnitType) = "A" Then
                                        ilAskforUpdate = True       'dont split any revenue by participants, invoicing has already done that
                                    End If
                                    Exit For
                                End If
                            Next illoop
    
                            'get the vehicle for this transaction
                            tmVefSrchKey.iCode = tmRvf.iAirVefCode
                            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet <> BTRV_ERR_NONE Then       '2-14-01 , vehicle isnt defined with transactions, init the fields that is accessed so
                                                                 'that incorrect data isnt used
                                'For ilLoop = 1 To 8
                                    'tmVef.iProdPct(ilLoop) = 0
                                    'If ilLoop = 1 Then
                                     '   tmVef.iProdPct(1) = 10000  'force to thinking a vehicle exists with 100% revenue
                                    'End If
                                    'tmVef.iMnfGroup(ilLoop) = 0
                                    'tmVef.iMnfSSCode(ilLoop) = 0
                                'Next ilLoop
                                'ilHowManyDefined = 1
                                tmVef.sType = "C"               'default to conventional for lack of anything else
                                ilMnfSSCode(1) = ilMatchSSCode
                                ilMnfGroup(1) = tmRvf.iMnfGroup
                                ilProdPct(1) = 10000
                                ilHowManyDefined = 1
                            Else
                                ilUse100pct = False
                                '6-27-08 Cash distributions needs to determine participant percentage based on the ageing date, not tran date
                                ilParticipantDate(0) = tmRvf.iTranDate(0)
                                ilParticipantDate(1) = tmRvf.iTranDate(1)
                                If ilListIndex = COLL_DISTRIBUTE And igRptCallType = COLLECTIONSJOB Then
                                    slStr = Trim$(str(tmRvf.iAgePeriod)) & "/15/" & Trim$(str(tmRvf.iAgingYear))
                                    slStr = gObtainEndStd(slStr)        'use std end date of ageing period since billing by std
                                    gPackDate slStr, ilParticipantDate(0), ilParticipantDate(1)
                                End If
    
                                gInitPartGroupAndPcts tmRvf.iAirVefCode, ilMatchSSCode, 0, ilMnfSSCode(), ilMnfGroup(), ilProdPct(), ilParticipantDate(), tmPifKey(), tmPifPct(), ilUse100pct
                                For illoop = 1 To UBound(ilMnfSSCode) Step 1
                                    ilHowManyDefined = UBound(ilMnfSSCode)
                                Next illoop
                                'ilHowMany = UBound(ilMnfSSCode)
                            End If
    
                            'For ilLoop = 1 To 8
                            '    If ilMatchSSCode = tmVef.iMnfSSCode(ilLoop) Then
                            '        ilHowManyDefined = ilLoop           'find the last participant of the matching S/S
                            '    End If
                            'Next ilLoop
    
    
    '                    ElseIf imProducer Then          '8-4-00
    '                        ilHowManyDefined = 8
    '                        'get the vehicle for this transaction
    '                        tmVefSrchKey.iCode = tmRvf.iAirVefCode
    '                        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        Else
                            ilHowManyDefined = 1
                        End If
    
                        'Go thru and split $ if necessary (i.e. splitting participants or splitting slsp revenue)
                        llSeqNo = llSeqNo + 1          'seq # that is incremented every time a new trans. is processed
                                                        'when distribution $ for participants, each set created for the same trans
                                                        'is given the same seq #.
                        'tmRvr.iUrfCode = ilSeqNo        'replace the user code with running seq #.   (need an unused field)
                        tmRvr.lRefInvNo = llSeqNo       '8-13-02 chged to use long (from integer)
                        'For Advt & Vehicle, create 1 record per transaction
                        'For Owner , create as many as 3 records per transaction.  (up to 3 owners per vehicle)
                        'For slsp, create as many as 10 records per trans (up to 10 split slsp) per trans.
                        ilMissingSS = True                              'assume no matching SS defined in this vehicle
                        For illoop = 0 To ilHowManyDefined - 1 Step 1        'loop based on report option
                            slStr = ".00"
                            gStrToPDN slStr, 2, 6, tmRvr.sGross
                            gStrToPDN slStr, 2, 6, tmRvr.sNet
                            '12-26-01 if Inv Reg is by office/vehicle, variable imSlsp is set to false & imvehicle is set to true.  Need to
                            'go thru some of the slsp code for split revenue
                            '4-17-11 make ageing by slsp a prepass; test for slsp
    '                        If (imSlsp And igRptCallType <> COLLECTIONSJOB) Or (igRptCallType = INVOICESJOB And ilListIndex = INV_REGISTER And RptSel!rbcSelCSelect(6).Value) Then
                            If (imSlsp) Or (igRptCallType = INVOICESJOB And ilListIndex = INV_REGISTER And RptSel!rbcSelCSelect(6).Value) Then      '2-25-19 implement slsp split on ageing
                                ReDim llSlfSplit(0 To 9) As Long           '4-20-00 slsp slsp share %
                                ReDim ilSlfCode(0 To 9) As Integer             '4-20-00
                                ReDim llSlfSplitRev(0 To 9) As Long         '2-2-04 unused in this report (need for subroutine)
                                ilMnfSubCo = gGetSubCmpy(tmChf, ilSlfCode(), llSlfSplit(), tmRvf.iAirVefCode, False, llSlfSplitRev())
                                'use this slsp if the airing vehicles associated sub-co matches the slsp sub-co, or the slsp sub-co is not defined
                                '4-20-00 If tmChf.islfCode(ilLoop) > 0 And tmChf.lComm(ilLoop) > 0 And (ilMnfSubCo = tmChf.iMnfSubCmpy(ilLoop) Or tmChf.iMnfSubCmpy(ilLoop) = 0) Then
                                If ilSlfCode(illoop) > 0 And llSlfSplit(illoop) > 0 Then
                                    '4-20-00 llProcessPct = tmChf.lComm(ilLoop)          'slsp split % or 100%
                                    llProcessPct = llSlfSplit(illoop)          'slsp split % or 100%
                                Else
                                    llProcessPct = 0
                                End If
                            ElseIf imOwner Or imProducer Then
                                '7-21-14 this has already called to get participant/pcts
                                'gInitPartGroupAndPcts tmRvf.iAirVefCode, ilMatchSSCode, 0, ilMnfSSCode(), ilMnfGroup(), ilProdPct(), ilParticipantDate(), tmPifKey(), tmPifPct(), ilUse100pct
                                ilExitFor = False                       '6-17-14 flag to indicate after processing owner to exit the loop since
                                                                        'the participant does not have to split revenue.  False indicates to split the participants
                                If ((ilListIndex = COLL_AGEOWNER Or ilListIndex = COLL_AGESS Or ilListIndex = COLL_AGEPRODUCER) And (igRptCallType = COLLECTIONSJOB)) And tmRvf.sTranType = "PO" Then         'PO dont have any of this info
                                    'PO for the ageing need to show the 100% amount
                                    If ilMatchSSCode = 0 Then
                                        tmRvr.imnfOwner = 0
                                        tmRvr.iProdPct = 10000
                                        tmRvr.iMnfSSCode = ilMatchSSCode  '0
                                        llProcessPct = 10000
                                    Else
                                        tmRvr.imnfOwner = ilMnfGroup(1) 'ilMnfGroup(LBound(ilMnfGroup))
                                        tmRvr.iMnfSSCode = ilMatchSSCode
                                    End If
                                    tmRvr.iProdPct = 10000
                                    llProcessPct = 10000
                                    ilExitFor = True                '6-17-14, dont split the participant, process @ 100%
                                Else                            'billing or cash distribution, ageing by owner or producer or ss
                                    'If (ilLoop + 1 = 1 And tmVef.iMnfGroup(ilLoop + 1) = 0 And tmVef.iMnfSSCode(ilLoop + 1) = 0) Then
                                     'if first time thru, and theres not a participant or sales source defined
                                     If (illoop + 1 = 1 And ilMnfGroup(illoop + 1) = 0) Then    'And ilMnfSSCode(ilLoop + 1) = 0) Then
                                        tmRvr.imnfOwner = 0
                                        tmRvr.iProdPct = 10000
                                        tmRvr.iMnfSSCode = ilMnfSSCode(illoop + 1)
                                    Else        'process the participant split
                                        If (tmVef.sType = "R" And ilExternalAdv) Or ilAskforUpdate Then           '1-14-03 for REP vehicles, dont split the externally billed contracts
                                            'ASK if Update Sales source should already be split by participants
                                            'override the Source of where transaction came from (PHF or RVF),
                                            'if from History, change code to "X" to indicate its an external billed transaction for the Billing & Cash Distribution
                                            If tmVef.sType = "R" And ilExternalAdv Then
                                                If tmRvf.sTranType = "HI" Then        'billed externally
                                                    tmRvr.sSource = "X"         'indicate billed externally for Bill/Cash Distribution
                                                End If
                                            End If
    
    
                                            'the transaction should have a participant defined since its sales source is "Ask"
                                            'if it doesnt have a participant, force to show the total $ on the ageing.
                                            'but if Cash or Billing Distribution, no distribution should be given but it
                                            'needs to be shown on the report and also highlighted that a participant or sales
                                            'source is missing
                                            tmRvr.imnfOwner = tmRvf.iMnfGroup
                                            'find correct group from vehicle table to get the proper % distribution
                                            tmRvr.iProdPct = -1
                                            'need to update the prepass record to show the correct % of split (i.e. 20%);
                                            'but will process at 100% because the transaction has already been split
    
                                            If (tmRvf.iMnfGroup = 0) And (Not ilDistribution) Then                'the sales source is Ask, but there is no designated participant reference in trasnaction
                                                                                    'get the first one from the vehicle.  If distribution report, need to be able to flag
                                                                                    'that entry as not having a valid participant and/or sales source reference.  Those entries will
                                                                                    'be flagged with 0% participation.
                                                tmRvr.imnfOwner = 0                 'unknown participant
                                                tmRvr.iMnfSSCode = ilMnfSSCode(1)
                                                tmRvr.iProdPct = 0
                                                llProcessPct = 1000000
                                                ilExitFor = True                    'exit after 1st participant, process @ 100%
                                            Else
                                                For ilTemp = illoop To ilHowManyDefined - 1     'get the designated participants info.  On the Billing
                                                                                                'and Cash Distribution, need to show the participant share in terms
                                                                                                'of the %; but give the participant the entire revenue of the trans.
                                                                                                'because it has already been split.
                                                    'If tmRvf.iMnfGroup = tmVef.iMnfGroup(ilTemp + 1) And tmVef.iMnfSSCode(ilTemp + 1) = ilMatchSSCode Then
                                                     If tmRvf.iMnfGroup = ilMnfGroup(ilTemp + 1) And ilMnfSSCode(ilTemp + 1) = ilMatchSSCode Then
                                                        'tmRvr.iProdPct = tmVef.iProdPct(ilTemp + 1)
                                                        tmRvr.iProdPct = ilProdPct(ilTemp + 1)
                                                        tmRvr.iMnfSSCode = ilMnfSSCode(illoop + 1)
                                                        llProcessPct = 1000000
                                                        ilExitFor = True            'exit after 1st participant, process @ 100%
                                                        Exit For
                                                    End If
                                                Next ilTemp
                                            End If
    
    
                                            If tmRvr.iProdPct < 0 Then          'didnt find a group and sales source because it wasnt defined or
                                                                                'it doesnt exist, dont know which participant to give it to
                                                'tmRvr.iMnfSSCode = tmVef.iMnfSSCode(ilLoop + 1)
                                                tmRvr.iMnfSSCode = ilMnfSSCode(illoop + 1)
                                                tmRvr.iProdPct = 0
                                                ilHowManyDefined = 8
                                                llProcessPct = 0
                                                illoop = 7
                                            End If
                                        Else            'do the splits
                                            'tmRvr.imnfOwner = tmVef.iMnfGroup(ilLoop + 1)
                                            tmRvr.imnfOwner = ilMnfGroup(illoop + 1)
                                            'tmRvr.iProdPct = tmVef.iProdPct(ilLoop + 1)
                                            tmRvr.iProdPct = ilProdPct(illoop + 1)
                                            'tmRvr.iMnfSSCode = tmVef.iMnfSSCode(ilLoop + 1)
                                            tmRvr.iMnfSSCode = ilMnfSSCode(illoop + 1)
                                            llProcessPct = ilProdPct(illoop + 1)         'make xxx.xxxx (100.0000)
                                            llProcessPct = llProcessPct * 100                    'make xxx.xxxx (100.0000)
                                        End If
                                    End If
                                End If
    
                                'for the ageing, everything needs to be included and shown.  If the share split is 0%,
                                'no sales source or participant but need to include it and show all $
                                '7-21-14 transaction all have a sales source.  if not, its set up for 100%.
                                'the participant table may be set up with 0% share, process all of them  to equal 100% for all participants
    '                            If igRptCallType = COLLECTIONSJOB And (ilListIndex = COLL_AGEOWNER Or ilListIndex = COLL_AGESS Or ilListIndex = COLL_AGEPRODUCER) Then
    '                                If llProcessPct = 0 Then
    '                                    llProcessPct = 1000000
    '                                End If
    '                            End If
                            Else                'remaining options (like advt, agency dont split revenue)
                                llProcessPct = 1000000              'if advt or vehicle, just use slsp since there will always be one
                                                                    'and only one pass will be processed
                            End If
    
                            If llProcessPct > 0 Then        'only continue if % of revenue needs processing; if zero ignore the trans.
                                ilMissingSS = False         '1-30-06 found a matching Sales Source
                                ilFoundOne = False
                                '8-4-00
                                If (imSlsp Or imOwner Or imProducer) Then   'And (Not RptSel!ckcAll.Value = vbChecked) Then                              'slsp, check if any of the split slsp should be excluded
                                    If (imOwner Or imProducer) Then
                                        If (RptSel!ckcAll.Value = vbChecked) Then
                                            ilFoundOne = True
                                        Else
                                            If ilIncludeCodes Then
                                                For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                                                    If (ilAskforUpdate) Or (tmVef.sType = "R" And ilExternalAdv) Or imProducer Then
                                                        If ilUseCodes(ilTemp) = tmRvf.iMnfGroup Then
                                                            ilFoundOne = True
                                                            Exit For
                                                        End If
                                                    Else
                                                        'If ilUseCodes(ilTemp) = tmVef.iMnfGroup(ilLoop + 1) Then
                                                         If ilUseCodes(ilTemp) = ilMnfGroup(illoop + 1) Then
                                                            ilFoundOne = True
                                                            Exit For
                                                        End If
                                                    End If
                                                Next ilTemp
                                            Else
                                                ilFoundOne = True        'selections to exclude
                                                For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                                                    If (ilAskforUpdate) Or (tmVef.sType = "R" And ilExternalAdv) Or imProducer Then
                                                        If ilUseCodes(ilTemp) = tmRvf.iMnfGroup Then
                                                            ilFoundOne = False
                                                            Exit For
                                                        End If
                                                    Else
                                                        'If ilUseCodes(ilTemp) = tmVef.iMnfGroup(ilLoop + 1) Then
                                                          If ilUseCodes(ilTemp) = ilMnfGroup(illoop + 1) Then
                                                                ilFoundOne = False
                                                                Exit For
                                                            End If
                                                        End If
                                                Next ilTemp
                                            End If
                                        End If
    
                                    ElseIf imSlsp Then                         'slsp option
                                        If igRptCallType = INVOICESJOB Then
                                            'ilListBoxInx = 2
                                            If RptSel!ckcAll.Value = vbUnchecked Then
                                                For ilLoopSlsp = 0 To RptSel!lbcSelection(2).ListCount - 1 Step 1
                                                    If RptSel!lbcSelection(2).Selected(ilLoopSlsp) Then              'selected slsp
                                                        slNameCode = tgSalesperson(ilLoopSlsp).sKey    'Traffic!lbcSalesperson.List(ilLoopSlsp)         'pick up slsp code
                                                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                                        If Val(slCode) = tmChf.iSlfCode(illoop) Then
                                                            ilFoundOne = True
                                                            
                                                            Exit For
                                                        End If
                                                    End If
                                                Next ilLoopSlsp
                                            Else
                                                ilFoundOne = True           'all selected
                                            End If
                                        ElseIf igRptCallType = COLLECTIONSJOB Then       'payment history
                                            'ilListBoxInx = 5                        'slsp selection list box
                                            ilRet = gBinarySearchSlf(tmRvf.iSlfCode)
                                            If ilRet <> -1 Then         'slsp found
                                                tmSlf = tgMSlf(ilRet)
                                            End If
                                            'Transaction like PO and JE may not have salesperson reference, if ALL salesperson report,
                                            'need to see them on report
                                            If Not RptSel!ckcAllGroups.Value = vbChecked Then
                                                'selection on Sales Offices
                                                For ilLoopSlsp = 0 To RptSel!lbcSelection(7).ListCount - 1 Step 1
                                                    If RptSel!lbcSelection(7).Selected(ilLoopSlsp) Then
                                                        slNameCode = tgSOCode(ilLoopSlsp).sKey
                                                        ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                                        If Val(slCode) = tmSlf.iSofCode Then
                                                            ilFoundOne = True
                                                            Exit For
                                                        End If
                                                    End If
                                                Next ilLoopSlsp
                                            Else
                                                ilFoundOne = True           'all offices selected
                                            End If
                                            If ilFoundOne Then              'valid sales office, check slsp
                                                If RptSel!ckcAll.Value = vbUnchecked Then           '8-25-11 was testing wrong check box for slsp selection
                                                    ilFoundOne = False
                                                    For ilLoopSlsp = 0 To RptSel!lbcSelection(5).ListCount - 1 Step 1
                                                        If RptSel!lbcSelection(5).Selected(ilLoopSlsp) Then              'selected slsp
                                                            slNameCode = tgSalesperson(ilLoopSlsp).sKey    'Traffic!lbcSalesperson.List(ilLoopSlsp)         'pick up slsp code
                                                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                                            If Val(slCode) = tmChf.iSlfCode(illoop) Then
                                                                ilFoundOne = True
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next ilLoopSlsp
                                                End If
                                            End If
                                            'End If
                                    
                                            'If ilListIndex = COLL_PAYHISTORY Then
                                            '    ilListBoxInx = 5
                                            'End If
                                            'If ilListIndex = COLL_CASH Then
                                           '     ilListBoxInx = 5
                                           '     If tmRvf.sTranType = "PO " Then
                                           '         ilFoundOne = True       'always include the PO since its cash by salesperson
                                           '     End If
                                           '  End If
                                        End If
    
    '                                    For ilLoopSlsp = 0 To RptSel!lbcSelection(ilListBoxInx).ListCount - 1 Step 1
    '                                        If RptSel!lbcSelection(ilListBoxInx).Selected(ilLoopSlsp) Then              'selected slsp
    '                                            slNameCode = tgSalesperson(ilLoopSlsp).sKey    'Traffic!lbcSalesperson.List(ilLoopSlsp)         'pick up slsp code
    '                                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '                                            If Val(slCode) = tmChf.iSlfCode(ilLoop) Then
    '                                                ilFoundOne = True
    '                                                Exit For
    '                                            End If
    '                                        End If
    '                                    Next ilLoopSlsp
                                        
                                    End If
                                Else                                                'all other options the record has already been filtered
                                    ilFoundOne = True
                                End If
                                If ilFoundOne Then
                                    '12-26-01 if Inv Reg is by office/vehicle, variable imSlsp is set to false & imvehicle is set to true.  Need to
                                    'go thru some of the slsp code for split revenue
                                    '4-17-11 make ageing by slsp a prepass
    '                                If (imSlsp And igRptCallType <> COLLECTIONSJOB) Or (igRptCallType = INVOICESJOB And ilListIndex = INV_REGISTER And RptSel!rbcSelCSelect(6).Value) Then
                                    If (imSlsp) Or (igRptCallType = INVOICESJOB And ilListIndex = INV_REGISTER And RptSel!rbcSelCSelect(6).Value) Then       '2-25-19 slsp split on ageing implemented
                                        '4-20-00 tmrvr.islfCode = tmChf.islfCode(ilLoop) 'slsp code
                                         tmRvr.iSlfCode = ilSlfCode(illoop) 'slsp code
                                        'get share for splits and write out pre-pass report record
                                        '4-20-00mObtainRvrShare llProcessPct, llTransGross, llTransNet, ilHowManyDefined, ilLoop, ilListIndex
                                        
                                        '3-23-18 Total slsp split percentages exceed 100 %.  Calc from orig gross & net values rather than from the running total to give
                                        'pennies left over to last slsp.  this is so those splits will balance to the invoice.  But if over 100%, this report will not
                                        'balance to the invoices.
                                        If llTotalSplits > 1000000 Then
                                            gPDNToLong tmRvf.sGross, llTransGross
                                            gPDNToLong tmRvf.sNet, llTransNet
                                        End If
                                        mObtainRvrShare llProcessPct, llTransGross, llTransNet, ilHowManyDefined, illoop, ilListIndex
                                    ElseIf imOwner Or imProducer Then
    
                                        '4-20-00 get share for splits and write out pre-pass report record
                                        '5-9-00
                                        mObtainRvrShare llProcessPct, llTransGross, llTransNet, ilHowManyDefined + 1, illoop + 1, ilListIndex
                                        If ilExitFor Then                           'force exit out of loop
                                            'ilLoop = ilHowManyDefined            'end the loop
                                            Exit For                               '6-17-14
                                        End If
    '                                ElseIf imProducer Then          '2-10-00
    '                                    tmRvr.imnfOwner = 0         '1-20-06 init fields inc ase owner not found
    '                                    tmRvr.iProdPct = 10000
    '                                    tmRvr.iMnfSSCode = 0
    '
    '                                    If tmRvf.iSlfCode <> tmSlf.iCode Then   'only read if not already in mem
    '                                        tmSlfSrchKey.iCode = tmRvf.iSlfCode
    '                                        ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '                                    End If
    '                                    tmRvr.iMnfSSCode = mFindMatchSSCode(tmSlf.iSofCode, tlSofList())
    '
    '                                    tmRvr.imnfOwner = tmRvf.iMnfGroup
    '                                    ilHowManyDefined = 0
    '                                    If tmRvr.imnfOwner = 0 And tmRvf.iAirVefCode > 0 Then   '1-20-06 if there isnt a participant reference, see if only one participant exists for this sales source
    '                                        'get the vehicle for this transaction
    '                                        tmVefSrchKey.iCode = tmRvf.iAirVefCode
    '                                        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '
    '                                        gInitPartGroupAndPcts tmRvf.iAirVefCode, ilMatchSSCode, tmRvf.iMnfGroup, ilMnfSSCode(), ilMnfGroup(), ilProdPct(), tmRvf.iTranDate(), tmPifKey(), tmPifPct()
    '
    '                                        For ilLoopSlsp = 1 To 8 Step 1
    '                                            'If tmVef.iMnfSSCode(ilLoopSlsp) = tmRvr.iMnfSSCode And ((ilLoopSlsp = 1) Or (tmVef.iProdPct(ilLoopSlsp) > 0 And ilLoopSlsp > 1)) Then        'get count of how many actually
    '                                            If ilMnfSSCode(ilLoopSlsp) = tmRvr.iMnfSSCode And ((ilLoopSlsp = 1) Or (ilProdPct(ilLoopSlsp) > 0 And ilLoopSlsp > 1)) Then        'get count of how many actually
    '
    '                                                                                                'have to be processed.  Need to find
    '                                                                                                'the last one because of extra pennies
    '                                                                                                'goes to the last one processed
    '                                                ilHowManyDefined = ilHowManyDefined + 1
    '                                                If ilHowManyDefined = 1 Then                          'if the first participant of the matching salessource, save that one
    '                                                    'ilmnfParticipant = tmVef.iMnfGroup(ilLoopSlsp)
    '                                                     ilmnfParticipant = ilMnfGroup(ilLoopSlsp)
    '                                                End If
    '                                            End If
    '                                        Next ilLoopSlsp
    '                                        If ilHowManyDefined = 1 Then            'only one participant for this sales source, insure that producer report can sort it
    '                                            tmRvr.imnfOwner = ilmnfParticipant
    '                                        End If
    '                                    End If
    '                                    '5-9-00
    '                                    mObtainRvrShare llProcessPct, llTransGross, llTransNet, ilHowManyDefined, ilLoop, ilListIndex
    '                                    ilLoop = 8
                                    Else
                                        '5-10-04 for Acct History to show Sales Source/participant on separate line
                                        '4-17-11 slsp record in memory, no reading record
    '                                    If tmRvf.iSlfCode <> tmSlf.iCode Then   'only read if not already in mem
    '                                        tmSlfSrchKey.iCode = tmRvf.iSlfCode
    '                                        ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '                                    End If
                                        
                                        ilRet = gBinarySearchSlf(tmRvf.iSlfCode)
                                        If ilRet = -1 Then          'not found
                                            tmRvr.iMnfSSCode = 0
                                            tmRvr.imnfOwner = 0
                                            tmRvr.iProdPct = 0
                                        Else
                                            tmSlf = tgMSlf(ilRet)
                                            
                                            tmRvr.iMnfSSCode = mFindMatchSSCode(tmSlf.iSofCode, tlSofList())
                                            tmRvr.imnfOwner = tmRvf.iMnfGroup
                                        End If
                                         '1-30-09 if vehicle and showing owners share, need to get the % share to report
                                         'can calculate it
                                         If imVehicle And RptSel!ckcSelC3(0).Value = vbChecked Then
                                             ilUse100pct = False
                                             ilParticipantDate(0) = tmRvf.iTranDate(0)
                                             ilParticipantDate(1) = tmRvf.iTranDate(1)
                                             gInitPartGroupAndPcts tmRvf.iAirVefCode, tmRvr.iMnfSSCode, 0, ilMnfSSCode(), ilMnfGroup(), ilProdPct(), ilParticipantDate(), tmPifKey(), tmPifPct(), ilUse100pct
                                             tmRvr.iProdPct = ilProdPct(1)       'get the owner pct share
                                             tmRvr.imnfVefGroup = ilMnfGroup(1)     '9-20-11 get the participant based on PIF table. will be overridden if another vehicle group selected to sort
                                        End If
                                        mObtainRvrShare llProcessPct, llTransGross, llTransNet, ilHowManyDefined, illoop, ilListIndex
                                        'End If
                                        'get share for splits and write out pre-pass report record 2-10-00 Move to above
                                        '5-9-00
                                        'mObtainRvrShare llProcessPct, llTransGross, llTransNet, ilHowManyDefined, ilLoop, ilListIndex
                                    End If
                                End If                          'ilFoundOne
                            End If                              'llProcessPct > 0
                        Next illoop                             'loop for 10 slsp possible splits, or 3 possible owners, otherwise loop once
    
                        If ilMissingSS Then            '1-30-06  was sales souce found?
                            If (igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_DISTRIBUTE) Or (igRptCallType = INVOICESJOB And ilListIndex = INV_DISTRIBUTE) Then
                                'highlight in report so user knows no distribution for this vehicle due to missing SS definition
                                'gross and net amts remain unsplit since there are no matching paricipants
                                tmRvr.sGross = tmRvf.sGross
                                tmRvr.sNet = tmRvf.sNet
                                tmRvr.lDistAmt = 0
                                tmRvr.imnfOwner = 0
                                tmRvr.iProdPct = 0  '10000
                                tmRvr.iMnfSSCode = ilMatchSSCode
                                tmRvr.sSource = "#"             'flag as missing participant
                                If (igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_DISTRIBUTE) Then
                                    gPDNToLong tmRvf.sNet, tmRvr.lDistAmt
                                    tmRvr.sNet = ".00"
                                End If
                                
                                'Export
                                If blExport = True Then
                                    If ilListIndex = COLL_DISTRIBUTE Then smExportStatus = mExportCashDist(tmRvr) 'TTP 10117 - Cash Distribution Export
                                    If ilListIndex = COLL_AGEMONTH Then smExportStatus = mSaveAgeSummary(tmRvr) 'TTP 10164 - Ageing summary by month - export option for Audacy A/R dump
                                    If ilListIndex = INV_DISTRIBUTE Then
                                        If RptSel!rbcSelCInclude(0).Value Then          'detail
                                            smExportStatus = mExportInvDist(tmRvr) 'TTP 10118 -Billing Distribution Export to CSV
                                        Else
                                            smExportStatus = mSaveInvDistSummary(tmRvr)
                                        End If
                                    End If
                                Else
                                    ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
                                End If
                            End If
                        End If
    
                    End If                                      'ilmatchcnt
                End If                                          'ilvaliddates
            End If      'mBinarySearchRVF
            
            ilRet = btrGetNext(hmRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            'ReDim ilMnfGroup(1 To 1) As Integer
            'ReDim ilMnfSSCode(1 To 1) As Integer
            'ReDim ilProdPct(1 To 1) As Integer
            ReDim ilMnfGroup(0 To 1) As Integer
            ReDim ilMnfSSCode(0 To 1) As Integer
            ReDim ilProdPct(0 To 1) As Integer
        'Loop
        Next llLoopOnTrans                      '5-15-19 for llLoopOnTrans
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
        llSeqNo = 0
    Next ilLoopOnFile                                   '2 passes, first History, then Receivbles



    'show total records exported
    If blExport = True Then
        lmExportCount = 0
        If ilListIndex = COLL_AGEMONTH Then mExportAgeMonthSummary 'TTP 10164 - Ageing summary by month - export option for Audacy A/R dump
        If ilListIndex = INV_DISTRIBUTE Then mExportInvDistSummary 'TTP 10118 -Billing Distribution Export to CSV
        
        Close #hmExport
        If InStr(1, smExportStatus, "Error") > 0 Then
            RptSel.lacExport.Caption = "Export Failed:" & smExportStatus
        Else
            RptSel.lacExport.Caption = "Export Stored in- " & sgExportPath & slFileName
        End If
    Else
        'send formula to rptsel of all undistributed cash from prior month
        If igRptCallType = COLLECTIONSJOB Then
            If ilListIndex = COLL_DISTRIBUTE Then
                If RptSel!rbcSelCSelect(0).Value Then       'invoice option cash distribution
                    ilRet = gSetFormula("CashUndistr", llUndistCash)
                End If
            End If
        End If
    End If
    
    Erase tlSofList, ilUseCodes, llSlfSplit, ilSlfCode
    Erase tlRvf                             '5-15-19
    ilRet = btrClose(hmRvr)
    ilRet = btrClose(hmRvf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmVsf)
    
    
    If igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_ACCTHIST Then          'acct history needs to create another prepass to contain decrypted user name on report
        ilRet = btrClose(hmTxr)
        btrDestroy hmTxr
    End If
        
    btrDestroy hmSbf
    btrDestroy hmSof
    btrDestroy hmAgf
    btrDestroy hmMnf
    btrDestroy hmVef
    btrDestroy hmCHF
    btrDestroy hmSlf
    btrDestroy hmRvr
    btrDestroy hmRvf
    btrDestroy hmVsf
    
    If igRptCallType = COLLECTIONSJOB And (ilListIndex = COLL_DISTRIBUTE Or ilListIndex = COLL_AGEMONTH) Then 'TTP 10117 & 10164
        ilRet = btrClose(hmPrf)
        btrDestroy hmPrf
    End If
    
    
End Sub
'
'*********************************************************************************
'                  ilOrAnd = gGetCountSelected lblSelection
'
'                     <input> - list box to test selection
'                     <output> - ilOrAnd 0 = Do Or equal testing filtering on selection
'                                        1 = Do And  not equal filtering on selection
'
'**********************************************************************************
'
Function gGetCountSelected(ilIndex As Integer) As Integer
Dim illoop As Integer
Dim ilCount As Integer

    ilCount = 0
    For illoop = 0 To RptSel!lbcSelection(ilIndex).ListCount - 1 Step 1
        If RptSel!lbcSelection(ilIndex).Selected(illoop) Then
            ilCount = ilCount + 1
        End If
    Next illoop
    If ilCount < (RptSel!lbcSelection(ilIndex).ListCount / 2) Then  'selected less than half the total in the list box
        gGetCountSelected = 0                   'assume to do the "or" condition for filtering
    Else
        gGetCountSelected = 1
    End If
End Function
'
'
'       gMerchGen - Generate prepass file for
'                   Merchandising (Promotions) report.
'                   Go thru PHF & RVF and find all
'                   Merchandise & or Promotions transactions
'                   (I & A types) that fall within the 12months from
'                   the requested year/quarter.  Build
'                   GRF records containing contract #, all
'                   split slsp with their %.
'
'       4/2/99 change merch/promo % from 3 to 2 decimal places (100.00% max)
'       6/14/06 option for selective contract #
'
Sub gMerchGen()
Dim ilRet As Integer
Dim illoop As Integer
Dim llDate As Long
Dim slStr As String
'ReDim llStdStartDates(1 To 13) As Long
ReDim llStdStartDates(0 To 13) As Long      'Index zero ignored
Dim ilLoopOnFile As Integer
Dim ilError As Integer
Dim ilFoundOption As Integer
Dim ilTemp As Integer
Dim slNameCode As String
Dim slCode As String
Dim llAmt As Long
Dim ilFoundMonth As Integer
Dim ilMonthNo As Integer
Dim ilMatchSSCode As Integer
Dim slPctFrom As String            'user input range of percent to include from (0 assumes all)
                                    'pct may be .5% (5000) thru 1.5% (15000)
Dim slPctTo As String              'user input range of percnt to include to (0 assumes all)
Dim ilPctFrom As Integer
Dim ilPctTo As Integer
Dim ilPct As Integer
Dim ilSlfCode As Integer            'code used to detrmine what data user is allowed to see
Dim ilListBoxIndex As Integer       '0 for advt, 6 for vehicles
Dim llSingleCntr As Long            'selective contract # input

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imGrfRecLen = Len(tmGrf)
    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imRvfRecLen = Len(tmRvf)

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imCHFRecLen = Len(tmChf)

    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSlfRecLen = Len(tmSlf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imVefRecLen = Len(tmVef)

    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imMnfRecLen = Len(tmMnf)
    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSofRecLen = Len(tmSof)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmRvf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmSof)
        btrDestroy hmGrf
        btrDestroy hmRvf
        btrDestroy hmCHF
        btrDestroy hmSlf
        btrDestroy hmVef
        btrDestroy hmMnf
        btrDestroy hmSof
        Exit Sub
    End If
    imPromotion = False
    imMerchant = False
    If RptSel!rbcSelC4(0).Value Then            'select merchandising
        imMerchant = True
    Else
        imPromotion = True
    End If

    llSingleCntr = Val(RptSel!edcCheck.Text)        'selective contract #

    ReDim ilSelected(0 To 0) As Integer
    If RptSel!rbcSelCSelect(0).Value Then            'select vehicles
        ilListBoxIndex = 6
        If Not RptSel!ckcAll.Value = vbChecked Then
            For ilTemp = 0 To RptSel!lbcSelection(6).ListCount - 1 Step 1
                If RptSel!lbcSelection(6).Selected(ilTemp) Then              'selected slsp
                    slNameCode = tgCSVNameCode(ilTemp).sKey    'RptSelCt!lbcCSVNameCode.List(ilTemp)         'pick up slsp code
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilSelected(UBound(ilSelected)) = Val(slCode)
                    ReDim Preserve ilSelected(0 To UBound(ilSelected) + 1)
                End If
            Next ilTemp
        End If
    Else
        ilListBoxIndex = 0                          'select advt
        If Not RptSel!ckcAll.Value = vbChecked Then
            For ilTemp = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
                If RptSel!lbcSelection(0).Selected(ilTemp) Then              'selected slsp
                    slNameCode = tgAdvertiser(ilTemp).sKey    'RptSelCt!lbcCSVNameCode.List(ilTemp)         'pick up slsp code
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilSelected(UBound(ilSelected)) = Val(slCode)
                    ReDim Preserve ilSelected(0 To UBound(ilSelected) + 1)
                End If
            Next ilTemp
        End If
    End If
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

    'User file is always open.  Find out what this user is allowed to see
    If tgUrf(0).iCode = 1 Or tgUrf(0).iCode = 2 Then    'guide or counterpoint password
        ilSlfCode = 0                   'allow guide & CSI to get all stuff
    Else
        ilSlfCode = tgUrf(0).iSlfCode   'slsp gets to see only his own stuff
    End If
    slStr = RptSel!edcSelCTo.Text      'user input Pct from & to
    'slPctFrom = gRoundStr(slStr, ".001", 3)
    'ilPctFrom = gStrDecToInt(slPctFrom, 3)
    slPctFrom = gRoundStr(slStr, ".01", 2)     '4/2/99 chg from 3 to 2 dec places
    ilPctFrom = gStrDecToInt(slPctFrom, 2)
    slStr = RptSel!edcSelCTo1.Text      'user input Pct from & to
    'slPctTo = gRoundStr(slStr, ".001", 3)
    'ilPctTo = gStrDecToInt(slPctTo, 3)
    slPctTo = gRoundStr(slStr, ".01", 2)
    ilPctTo = gStrDecToInt(slPctTo, 2)

    'slPctFrom = gRoundStr(slStr, ".0001", 4)
    'ilPctFrom = gStrDecToInt(slPctFrom, 4)
    'slStr = RptSel!edcSelCTo1.Text      'user input Pct from & to
    'slPctTo = gRoundStr(slStr, ".0001", 4)
    'ilPctTo = gStrDecToInt(slPctTo, 4)

    If slPctTo = "0.00" Then
        'slPctTo = "30.000"                 'assume highest allowed
        'slPctTo = "3.0000"
        slPctTo = "100.00"                  'assume highest allowed
        ilPctTo = gStrDecToInt(slPctTo, 2)
    End If
    illoop = (Val(RptSel!edcSelCFrom1.Text) - 1) * 3 + 1     'convert qtr to month index
    slStr = Trim$(str$(illoop)) & "/15/" & Trim$(str$(igYear))      'format xx/xx/xxxx
    'build array of 13 start standard dates - Everything is obtained from PHF & RVF
    For illoop = 1 To 13 Step 1
        slStr = gObtainStartStd(slStr)
        llStdStartDates(illoop) = gDateValue(slStr)
        slStr = gObtainEndStd(slStr)
        llDate = gDateValue(slStr) + 1                      'increment for next month
        slStr = Format$(llDate, "m/d/yy")
    Next illoop

    For ilLoopOnFile = 1 To 2 Step 1
        'handles and buffers for PHF and RVF will be the same
        ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            ilFoundOption = False
            If RptSel!ckcAll.Value Then                           'report all
                ilFoundOption = True
            Else                             'check selective vehicles or adv
                If ilListBoxIndex = 6 Then      'veh
                    For ilTemp = 0 To UBound(ilSelected) - 1 Step 1
                        If tmRvf.iAirVefCode = ilSelected(ilTemp) Then
                            ilFoundOption = True
                            Exit For
                        End If
                    Next ilTemp
                Else                            'advt
                    For ilTemp = 0 To UBound(ilSelected) - 1 Step 1
                        If tmRvf.iAdfCode = ilSelected(ilTemp) Then
                            ilFoundOption = True
                            Exit For
                        End If
                    Next ilTemp
                End If
            End If

            If ilFoundOption = True Then
                If llSingleCntr > 0 And llSingleCntr <> tmRvf.lCntrNo Then
                    ilFoundOption = False
                End If
            End If

            gPDNToLong tmRvf.sNet, llAmt
            slCode = Trim$(str$(tmRvf.iAgePeriod) & "/15/" & Trim$(str$(tmRvf.iAgingYear)))
            slStr = gObtainEndStd(slCode)

            llDate = gDateValue(slStr)                  'convert trans date to test if within requested limits
            'valid record must be an "Invoice", History Invoices, or Adjsutment types, non-zero amount, and transaction date within the start date of the
            'cal year and end date of the current cal month requested
            If ((Left$(tmRvf.sTranType, 1) = "I" Or tmRvf.sTranType = "HI" Or Left$(tmRvf.sTranType, 1) = "A") And (llAmt <> 0) And (llDate >= llStdStartDates(1) And llDate < llStdStartDates(13)) And ((tmRvf.sCashTrade = "P" And imPromotion) Or (tmRvf.sCashTrade = "M" And imMerchant)) And (ilFoundOption) And (ilSlfCode = 0 Or ilSlfCode = tmRvf.iSlfCode)) Then         'looking for Invoice types only
                'get contract from history or rec file

                tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
                   'dont check for altered status since merch/promo is updated at the times its changed, not after scheduling
                'Do While (ilret = BTRV_ERR_NONE) And (tmChf.lCntrno = tmRvf.lCntrno)    'And (tmChf.sSchStatus = "A")
                '     ilret = btrGetNext(hmChf, tmChf, imChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                'Loop

                'ilValidStatus = False
                '12-28-12  if theres no contract #, create a fake entry to get something printed
                gFakeChf tmRvf, tmChf
                'only look for HOGN statuses, look for the valid header belonging to the transaction
                Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sStatus <> "H" And tmChf.sStatus <> "O" And tmChf.sStatus <> "G" And tmChf.sStatus <> "N")
                    ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    'If tmChf.sStatus = "H" Or tmChf.sStatus = "O" Or tmChf.sStatus = "G" Or tmChf.sStatus = "N" Then    'remove 10-26-99 Or (tmChf.sStatus = "W" And tmChf.iCntRevNo > 0) Then
                    '    ilValidStatus = True
                    'End If
                Loop
                If (imMerchant) Then
                    ilPct = tmChf.iMerchPct
                Else
                    ilPct = tmChf.iPromoPct
                End If
                If (tmChf.lCntrNo = tmRvf.lCntrNo And ilRet = BTRV_ERR_NONE) And (ilPct >= ilPctFrom And ilPct <= ilPctTo) And (tmChf.sStatus = "H" Or tmChf.sStatus = "O" Or tmChf.sStatus = "G" Or tmChf.sStatus = "N") Then
                    'format remainder of record
                    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
                    tmGrf.iGenDate(1) = igNowDate(1)
                    'tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
                    'tmGrf.iGenTime(1) = igNowTime(1)
                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                    tmGrf.lGenTime = lgNowTime
                    tmGrf.lChfCode = tmChf.lCode           'contr internal code
                    'tmGrf.iDateGenl(0, 1) = tmRvf.iTranDate(0)    'date billed or paid
                    'tmGrf.iDateGenl(1, 1) = tmRvf.iTranDate(1)
                    tmGrf.iDateGenl(0, 0) = tmRvf.iTranDate(0)    'date billed or paid
                    tmGrf.iDateGenl(1, 0) = tmRvf.iTranDate(1)
                    tmGrf.iVefCode = tmRvf.iAirVefCode
                    tmGrf.iAdfCode = tmRvf.iAdfCode
                    'tmGrf.iPerGenl(1) = tmRvf.iAgfCode
                    tmGrf.iPerGenl(0) = tmRvf.iAgfCode
                    tmGrf.sDateType = tmRvf.sCashTrade          'Type field used for C = Cash, T = Trade
                    tmGrf.iCode2 = ilPct
                    If ilLoopOnFile = 1 Then
                        tmGrf.sBktType = "H"                    'let crystal know these records are histroy/receivables (vs contracts)
                    Else
                        tmGrf.sBktType = "R"
                    End If
                    'determine the month that this transaction falls within
                    ilFoundMonth = False
                    gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slCode
                    llDate = gDateValue(slCode)
                    For ilMonthNo = 1 To 12 Step 1         'loop thru months to find the match
                        If llDate >= llStdStartDates(ilMonthNo) And llDate < llStdStartDates(ilMonthNo + 1) Then
                            ilFoundMonth = True
                            Exit For
                        End If
                    Next ilMonthNo
                    If ilFoundMonth Then
                        'obtain the sales source for major sort of business booked reports
                        If tmRvf.iSlfCode <> tmSlf.iCode Then        'only read slsp recd if not in mem already
                            tmSlfSrchKey.iCode = tmRvf.iSlfCode         'find the slsp to obtain the sales source code
                            ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            For illoop = LBound(tlSofList) To UBound(tlSofList)
                                If tlSofList(illoop).iSofCode = tmSlf.iSofCode Then
                                    ilMatchSSCode = tlSofList(illoop).iMnfSSCode          'Sales source
                                    Exit For
                                End If
                            Next illoop
                        End If
                        'For ilLoop = 1 To 12 Step 1
                        For illoop = 0 To 11 Step 1
                            tmGrf.lDollars(illoop) = 0          'init $ fields
                        Next illoop
                        tmGrf.iSofCode = ilMatchSSCode          'sales source code
                        'Crystal program needs to retrieve max 10 slsp from Contract recd for slsp splits
                        '*****  temporary test until item name defined
                        'ilTemp = 50
                        'slStr = gIntToStrDec(ilTemp, 4)
                        'gPDNToStr tmRvf.sNet, 2, slCode
                        'slDollar = gMulStr(slStr, slCode)
                        'tmGrf.lDollars(ilMonthNo) = Val(gRoundStr(slDollar, "01.", 0))
                        gPDNToLong tmRvf.sNet, tmGrf.lDollars(ilMonthNo - 1)
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    End If
                End If                          'btrv_err_none & contract matches
            End If                              'all conditions meet filtering
            ilRet = btrGetNext(hmRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop

        're-open PHF for next time thru
        ilRet = btrClose(hmRvf)                         'Close Receivables file

        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmSlf)
            ilRet = btrClose(hmVef)
            ilRet = btrClose(hmMnf)
            btrDestroy hmGrf
            btrDestroy hmRvf
            btrDestroy hmCHF
            btrDestroy hmSlf
            btrDestroy hmVef
            btrDestroy hmMnf
            Exit Sub
        End If
        imRvfRecLen = Len(tmRvf)                    'prepare History file for Comparison dates needed for
    Next ilLoopOnFile                               '2 passes, first History, then Receivbles
    Erase tlSofList, ilSelected
End Sub
Sub gMerchRecapGen()
Dim ilRet As Integer
Dim slStartDate As String               'temporary date for strings
Dim slEndDate As String                 'temporary date for strings
Dim llActiveStartDate As Long           'earliest contract header start date to use
Dim llActiveEndDate As Long             'latest contract header end date to use
Dim llDate As Long                      'temporary  for serial dates
Dim llDate2 As Long                     'temporary for serial dates
Dim ilLoopOnFile As Integer             'loop variable to go thru PHF, then RVF
Dim llAmt As Long                       'net amount of transaction
Dim ilNoWeeks As Integer                'total span of airing weeks to calc avg per week
Dim ilNoMonths As Integer               'total span of airing months to calc avg per month
Dim llAnotherDate As Long
Dim slStr As String                     'temp string
Dim ilSlfCode As Integer                'Code to detrmine what user is allowed to see
Dim ilValidStatus As Integer
    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imRvfRecLen = Len(tmRvf)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmRvf)
        btrDestroy hmAdf
        btrDestroy hmRvf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAdfRecLen = Len(tmAdf)
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmRvf)
        btrDestroy hmGrf
        btrDestroy hmAdf
        btrDestroy hmRvf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    hmCHF = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmRvf)
        btrDestroy hmCHF
        btrDestroy hmGrf
        btrDestroy hmAdf
        btrDestroy hmRvf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)
    'User file is always open.  Find out what this user is allowed to see
    If tgUrf(0).iCode = 1 Or tgUrf(0).iCode = 2 Then    'guide or counterpoint password
        ilSlfCode = 0                   'allow guide & CSI to get all stuff
    Else
        ilSlfCode = tgUrf(0).iSlfCode   'slsp gets to see only his own stuff
    End If
    'setup earliest and latest active to test transactions to be included
    slStartDate = RptSel!edcSelCFrom.Text   'Earliest Active Start date
    If slStartDate = "" Then
        slStartDate = "1/5/1970" 'Monday
    End If
    llActiveStartDate = gDateValue(slStartDate)
    slStartDate = RptSel!edcSelCFrom1.Text   'End date
    If (StrComp(slStartDate, "TFN", 1) = 0) Or (Len(slStartDate) = 0) Then
        llActiveEndDate = gDateValue("12/29/2069")    'Sunday
    Else
        llActiveEndDate = gDateValue(slStartDate)
    End If

    For ilLoopOnFile = 1 To 2 Step 1
        'handles and buffers for PHF and RVF will be the same
        ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            'Filter transactions on ageing date, transaction type (include "I", "A", "H") and cash/trade M or P
            'slStartDate = Trim$(Str$(tmRvf.iAgePeriod) & "/15/" & Trim$(Str$(tmRvf.iAgingYear)))
            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slStartDate
            llDate = gDateValue(gObtainEndStd(slStartDate))
            gPDNToLong tmRvf.sNet, llAmt

            'valid record must be an "Invoice", History Invoices, or Adjustment types, non-zero amount, and ageing month-year date within the start date of the
            ' dates requested
            If ((Left$(tmRvf.sTranType, 1) = "I" Or tmRvf.sTranType = "HI" Or Left$(tmRvf.sTranType, 1) = "A") And (llAmt <> 0) And (llDate >= llActiveStartDate And llDate <= llActiveEndDate)) And (tmRvf.sCashTrade = "P" Or tmRvf.sCashTrade = "M") And (ilSlfCode = 0 Or ilSlfCode = tmRvf.iSlfCode) Then
                'get contract from history or rec file
                tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd


                'ignored altered contract headers
                'Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrno = tmRvf.lCntrno) And (tmChf.sSchStatus = "A")
                '     ilRet = btrGetNext(hmChf, tmChf, imChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                'Loop
                ilValidStatus = False
                If tmChf.sStatus = "H" Or tmChf.sStatus = "O" Or tmChf.sStatus = "G" Or tmChf.sStatus = "N" Or (tmChf.sStatus = "W" And tmChf.iCntRevNo > 0) Then
                    ilValidStatus = True
                End If

                If ilRet = BTRV_ERR_NONE And ilValidStatus Then
                    'determine # of weeks for this order
                    gUnpackDate tmChf.iStartDate(0), tmChf.iStartDate(1), slStartDate
                    llDate = gDateValue(slStartDate)
                    gUnpackDate tmChf.iEndDate(0), tmChf.iEndDate(1), slEndDate
                    llDate2 = gDateValue(slEndDate)
                    ilNoWeeks = (llDate2 - llDate) / 7 + 1
                    'determine # std bdcst months from start to end
                    llAnotherDate = llDate
                    ilNoMonths = 0
                    Do While llAnotherDate < llDate2
                        slStr = gObtainEndStd(Format$(llAnotherDate, "m/d/yy"))
                        llAnotherDate = gDateValue(slStr) + 1
                        ilNoMonths = ilNoMonths + 1
                    Loop

                    'tmGrf.iGenTime(0) = igNowTime(0)        'generation time
                    'tmGrf.iGenTime(1) = igNowTime(1)
                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                    tmGrf.lGenTime = lgNowTime
                    tmGrf.iGenDate(0) = igNowDate(0)        'generation date
                    tmGrf.iGenDate(1) = igNowDate(1)
                    tmGrf.iVefCode = tmRvf.iAirVefCode      'airing vehicle code
                    tmGrf.iAdfCode = tmRvf.iAdfCode         ' advertiser
                    'tmGrf.iPerGenl(1) = tmRvf.iAgfCode         'agency code
                    tmGrf.iPerGenl(0) = tmRvf.iAgfCode         'agency code
                    'tmGrf.iCode2 = tmRvf.lPrfCode            'product name code
                    tmGrf.lCode4 = tmRvf.lPrfCode           '9-23-09 product code is long, not integer
                    tmGrf.sBktType = tmRvf.sCashTrade       'only M or P
                    tmGrf.lChfCode = tmRvf.lCntrNo         'contract ID
                    'tmGrf.iDateGenl(0, 1) = tmChf.iStartDate(0)
                    'tmGrf.iDateGenl(1, 1) = tmChf.iStartDate(1)
                    'tmGrf.iDateGenl(0, 2) = tmChf.iEndDate(0)
                    'tmGrf.iDateGenl(1, 2) = tmChf.iEndDate(1)
                    tmGrf.iDateGenl(0, 0) = tmChf.iStartDate(0)
                    tmGrf.iDateGenl(1, 0) = tmChf.iStartDate(1)
                    tmGrf.iDateGenl(0, 1) = tmChf.iEndDate(0)
                    tmGrf.iDateGenl(1, 1) = tmChf.iEndDate(1)
                    If ilLoopOnFile = 1 Then                'History file
                        tmGrf.sDateType = "H"
                    Else
                        tmGrf.sDateType = "R"               'receivables
                    End If
                    'tmGrf.iDateGenl(0, 1) = tmChf.iStartDate(0)    'contract start date
                    'tmGrf.iDateGenl(1, 1) = tmChf.iStartDate(1)
                    'tmGrf.iDateGenl(0, 2) = tmChf.iEndDate(0)   'contract start date
                    'tmGrf.iDateGenl(1, 2) = tmChf.iEndDate(1)
                    tmGrf.iDateGenl(0, 0) = tmChf.iStartDate(0)    'contract start date
                    tmGrf.iDateGenl(1, 0) = tmChf.iStartDate(1)
                    tmGrf.iDateGenl(0, 1) = tmChf.iEndDate(0)   'contract start date
                    tmGrf.iDateGenl(1, 1) = tmChf.iEndDate(1)
                    'overall merchandising amount
                    'tmGrf.lDollars(3) = llAmt
                    'tmGrf.lDollars(1) = 0
                    'tmGrf.lDollars(2) = 0
                    tmGrf.lDollars(2) = llAmt
                    tmGrf.lDollars(0) = 0
                    tmGrf.lDollars(1) = 0
                    If ilNoWeeks <> 0 Then
                        'calc avg price/week
                        'tmGrf.lDollars(1) = llAmt / ilNoWeeks
                        tmGrf.lDollars(0) = llAmt / ilNoWeeks
                    End If
                    If ilNoMonths <> 0 Then
                        'calc avg price/month
                        'tmGrf.lDollars(2) = llAmt / ilNoMonths
                        tmGrf.lDollars(1) = llAmt / ilNoMonths
                    End If
                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                End If                  'ilValidStatus
            End If                      'Rvf tests
            'TTP 10264 - JW - 8/3/21 - Merchandising/Promotions recap report doesn't complete running - the Following Line was Not Leftovers and Should NOT be removed
            ilRet = btrGetNext(hmRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)      '1-14-21 this is left over code from 5/17/19 that should have been removed
        Loop
        're-open PHF for next time thru
        ilRet = btrClose(hmRvf)                         'Close Receivables file

        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmVef)
            btrDestroy hmGrf
            btrDestroy hmRvf
            btrDestroy hmCHF
            btrDestroy hmVef
            Exit Sub
        End If
        imRvfRecLen = Len(tmRvf)                    'prepare History file for Comparison dates needed for
    Next ilLoopOnFile                               '2 passes, first History, then Receivbles
    ilRet = btrClose(hmUrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmRvf)
End Sub
'
'
'                   mFakeChf: Billed & booked subroutine to
'                   create a header with fields from the
'                   Receivables record so that the transaction
'                   is included in report.  Header is not found
'                   when transactions are added without references
'                   to the contract (no contract entered) or
'                   the contract # is zero
'
'                   7/14/99
'
'
Sub mFakeChf()
Dim illoop As Integer
    If tmRvf.lCntrNo <> tmChf.lCntrNo Then
        tmChf.lCntrNo = tmRvf.lCntrNo
        tmChf.lCode = 0
        tmChf.sSchStatus = "F"              'assume fully scheduled
        tmChf.sStatus = "O"                 'assume order
        tmChf.sType = "C"                   'standard contract
        If tmRvf.sCashTrade = "T" Then
            tmChf.iPctTrade = 100
        Else                                'could be merchndising, promotions or cash
            tmChf.iPctTrade = 0
        End If
        For illoop = 0 To 9 Step 1
        If illoop = 0 Then
            tmChf.iSlfCode(illoop) = tmRvf.iSlfCode
            tmChf.lComm(illoop) = 1000000
        Else
            tmChf.iSlfCode(illoop) = 0
            tmChf.lComm(illoop) = 0
        End If
        Next illoop
        tmChf.iAgfCode = tmRvf.iAgfCode
        tmChf.iAdfCode = tmRvf.iAdfCode

    End If
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
Dim illoop As Integer
    For illoop = 0 To 9 Step 1
        If illoop = 0 Then
            tmChf.iSlfCode(illoop) = tmRvf.iSlfCode
            tmChf.lComm(illoop) = 1000000
        Else
            tmChf.iSlfCode(illoop) = 0
            tmChf.lComm(illoop) = 0
        End If
    Next illoop
End Sub
'
'
'
'           mGetCashTradeOption - setup which types of
'           transactions to process:  Cash, tRade, Merchandising
'           or Promotions
'
'           8/31/99
'
Sub mGetCashTradeOption(tlTranType As TRANTYPES)
    If RptSel!rbcSelCInclude(0).Value Then             'cash only
        imTrade = False
        imCash = True
        tlTranType.iTrade = False                       '5-15-19
        tlTranType.iCash = True                         '5-15-19
    ElseIf RptSel!rbcSelCInclude(1).Value Then          'trade only
        imTrade = True
        imCash = False
        tlTranType.iTrade = True                       '5-15-19
        tlTranType.iCash = False                         '5-15-19
    ElseIf RptSel!rbcSelCInclude(2).Value Then          'merchandising only
        imMerchant = True
        imTrade = False
        imCash = False
        tlTranType.iTrade = False                       '5-15-19
        tlTranType.iCash = False                         '5-15-19
        tlTranType.iMerch = True                         '5-15-19
    ElseIf RptSel!rbcSelCInclude(3).Value Then          'promotions only
        imPromotion = True
        imTrade = False
        imCash = False
        tlTranType.iTrade = False                       '5-15-19
        tlTranType.iCash = False                         '5-15-19
        tlTranType.iPromo = True                         '5-15-19
    End If
End Sub
'
'
'               Payment or Usage History report subroutine
'               mObtainCodes - get all codes to process or exclude
'               When selecting advt, agy or vehicles--make testing
'               of selection more efficient.  If more than half of
'               the entries are selected, create an array with entries
'               to exclude.  If less than half of entries are selected,
'               create an array with entries to include.
'               <input> ilListIndex - list box to test
'                       lbcListbox - array containing sort codes
'               <output> ilIncludeCodes - true if test to include the codes in array
'                                          false if test to exclude the codes in array
'                        ilUseCodes - array of advt, agy or vehicles codes to include/exclude
Sub mObtainCodes(ilListIndex As Integer, lbcListBox() As SORTCODE, ilIncludeCodes, ilUseCodes() As Integer)
Dim ilHowManyDefined As Integer
Dim ilHowMany As Integer
Dim slNameCode As String
Dim illoop As Integer
Dim slCode As String
Dim ilRet As Integer
    ilHowManyDefined = RptSel!lbcSelection(ilListIndex).ListCount
    'ilHowMany = RptSel!lbcSelection(ilListIndex).SelectCount
    ilHowMany = RptSel!lbcSelection(ilListIndex).SelCount
    If ilHowMany > (ilHowManyDefined / 2) + 1 Then  'more than half selected
        ilIncludeCodes = False
    Else
        ilIncludeCodes = True
    End If
    For illoop = 0 To RptSel!lbcSelection(ilListIndex).ListCount - 1 Step 1
        slNameCode = lbcListBox(illoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If RptSel!lbcSelection(ilListIndex).Selected(illoop) And ilIncludeCodes Then               'selected ?
            ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
            ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
        Else        'exclude these
            If (Not RptSel!lbcSelection(ilListIndex).Selected(illoop)) And (Not ilIncludeCodes) Then
                ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
                ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
            End If
        End If
    Next illoop
End Sub
'
'***************************************************************************************************************
'
'                mObtainRvrShare : Calculate the Gross and Net portions of a transaction if
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
'                           ilListIndex = index to report type
'                <output>   tmRvr.sGross
'                           tmRvr.sNet
'
'               Created 11/18/96 DH
'               11-2-02  If original value is negative, make positive to do match and take care
'                        of extra pennies when splitting by slsp or participant
'               2-10-04 Show $0 transactions on Acct History
'               3-8-04 PO trans erroneously showing as $0 on Acct History
'***************************************************************************************************************
Sub mObtainRvrShare(llProcessPct As Long, llTransGross As Long, llTransNet As Long, ilHowManyDefined As Integer, illoop As Integer, ilListIndex As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilTrfAgyAdvt                                                                          *
'******************************************************************************************

Dim slPct As String
Dim slAmount As String
Dim slDollar As String
Dim llGrossDollar As Long
Dim llNetDollar As Long
Dim ilRet As Integer
Dim slStr As String
Dim ilMnfMajorCode As Integer       '7-16-02
Dim ilmnfMinorCode As Integer       '7-16-02
Dim ilLoopOnUsers As Integer
Dim slYear As String
Dim slMonth As String
Dim slDay As String

        If tmRvr.sTranType <> "PO" Then                     'do splits for cash (excluding on account payments)
                                                            'and billing distribution
            slPct = gLongToStrDec(llProcessPct, 4)           'slsp split share in % or Owner pct.  If advt or vehicle
                                                                'options, slsp is force to100%

            gPDNToStr tmRvf.sGross, 2, slAmount
            slDollar = gMulStr(slPct, slAmount)                 'slsp gross portion of possible split
            llGrossDollar = Val(gRoundStr(slDollar, "01.", 0))
            If imGrossNeg Then                              'if the original gross was negative, do all math as positive to handle extra pennies
                llGrossDollar = -llGrossDollar
            End If
            llTransGross = llTransGross - llGrossDollar
            If (illoop = ilHowManyDefined - 1 And RptSel!ckcAll.Value = vbChecked) Or llTransGross < 0 Then               'last slsp or participant processed? Handle extra pennies
                llGrossDollar = llGrossDollar + llTransGross  'last record written for splits, left over pennies goes to last owner
            End If
            If imGrossNeg Then                  'make the value back to negative if thats what the orig amount was, to store in RVR
                llGrossDollar = -llGrossDollar
            End If
            slStr = gLongToStrDec(llGrossDollar, 2)
            gStrToPDN slStr, 2, 6, tmRvr.sGross
            gPDNToStr tmRvf.sNet, 2, slAmount
            slDollar = gMulStr(slPct, slAmount)                 'slsp net portion of possible split
            llNetDollar = Val(gRoundStr(slDollar, "01.", 0))
            If imNetNeg Then                            'if the original gross was negative, do all math as positive to handle extra pennies
                llNetDollar = -llNetDollar
            End If
            llTransNet = llTransNet - llNetDollar
            If (illoop = ilHowManyDefined - 1 And RptSel!ckcAll.Value = vbChecked) Or llTransNet < 0 Then               'last slsp or participant processed?  handle extra pennies
                llNetDollar = llNetDollar + llTransNet         'last record writen for splits, left over pennies goes to last owner
            End If
            If imNetNeg Then                        'make the value back to negative if thats what the orig amount was, to store it back in RVR
                llNetDollar = -llNetDollar
            End If
            slStr = gLongToStrDec(llNetDollar, 2)
            gStrToPDN slStr, 2, 6, tmRvr.sNet
            'For the billing distribution, show the airing vehicles original billing amounts gross & net,
            'not the participants share of the airings gross & net.
            'Show the participants share stored in tmrvf.ldistamt
            If igRptCallType = INVOICESJOB And ilListIndex = INV_DISTRIBUTE Then
                tmRvr.lDistAmt = llNetDollar
                tmRvr.sGross = tmRvf.sGross
                tmRvr.sNet = tmRvf.sNet
            End If
            
            If igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_AGEMONTH Then     '3-3-16need to determine end date of calendar month to store in prepass
                slStr = Trim$(str(tmRvr.iAgePeriod)) & "/01/" & Trim$(str(tmRvr.iAgingYear))
                slStr = gObtainEndCal(slStr)
                gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                tmRvr.lDistAmt = Val(slDay)     'last day of calendar month
            End If
        'End If'PO
    Else
        If igRptCallType = COLLECTIONSJOB Then              'leave net value in its orginal state if Payment history
            '2-15-01 add ageing by Sales source to create entry to print
            'If ilListIndex = COLL_PAYHISTORY Or ilListIndex = COLL_CASH Or ilListIndex = COLL_AGEOWNER Or ilListIndex = COLL_AGEPRODUCER Or ilListIndex = COLL_AGESS Or ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGEVEHICLE Then   '7-19-00 (test for producer to include PO) if ageing by owner and PO, there are not splits yet
            If ilListIndex = COLL_AGEOWNER Or ilListIndex = COLL_AGEPRODUCER Or ilListIndex = COLL_AGESS Or ilListIndex = COLL_AGEPAYEE Or ilListIndex = COLL_AGEVEHICLE Or ilListIndex = COLL_ACCTHIST Or ilListIndex = COLL_AGEMONTH Then     '7-19-00 (test for producer to include PO) if ageing by owner and PO, there are not splits yet

                If imNetNeg Then                  'if orig amt was negative, make it back to positive to store in RVR
                    llTransNet = -llTransNet
                End If
                slStr = gLongToStrDec(llTransNet, 2)
                gStrToPDN slStr, 2, 6, tmRvr.sNet
                    
                If ilListIndex = COLL_AGEMONTH Then     '3-3-16 need to determine end date of calendar month to store in prepass
                    slStr = Trim$(str(tmRvr.iAgePeriod)) & "/01/" & Trim$(str(tmRvr.iAgingYear))
                    slStr = gObtainEndCal(slStr)
                    gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
                    tmRvr.lDistAmt = Val(slDay)     'last day of calendar month
                End If
            End If
        End If
    End If
    'Always write out the record for Invoice register or Acct History (2-10-04), even if $0
    If (igRptCallType = INVOICESJOB And ilListIndex = INV_REGISTER) Or (igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_ACCTHIST) Then
        tmRvr.iUrfCode = 0              'overlapped with the vehicle group set that has been selected (if applicable)
                                        '0 implies no vehicle group selected
        tmRvr.imnfVefGroup = 0
        If imMajorSet > 0 Then      'using vehicle groups?
            gGetVehGrpSets tmRvr.iAirVefCode, imMinorSet, imMajorSet, ilmnfMinorCode, ilMnfMajorCode    '7-16-02 obtain vehicle group code, some options may not use it
            tmRvr.imnfVefGroup = ilMnfMajorCode
            tmRvr.iUrfCode = imMajorSet         'used for report heading only
        End If
        'use rvrurfcode to point to txr
        If (igRptCallType = COLLECTIONSJOB And ilListIndex = COLL_ACCTHIST) Then        '10-13-15 acct history needs to show the user name which is encrypted. there is no
'                                                                                          'space in the temp record to store the name, link to another temporary file
'            tmTxr.lGenTime = tmRvr.lGenTime
'            tmTxr.iGenDate(0) = tmRvr.iGenDate(0)
'            tmTxr.iGenDate(1) = tmRvr.iGenDate(1)
'            tmTxr.lCsfCode = tmRvr.lCode                    'used to link to the temp record for decrypted user name that is stored in it
'            tmTxr.sText = ""
'            For ilLoopOnUsers = LBound(tgPopUrf) To UBound(tgPopUrf) - 1
'                If tgPopUrf(ilLoopOnUsers).iCode = tmRvf.iUrfCode Then          'rvfurfcode is used for another field for reporting (crystal)
'                    tmTxr.sText = Trim$((tgPopUrf(ilLoopOnUsers).sRept))        'name to show on report
'                    Exit For
'                End If
'            Next ilLoopOnUsers
'            ilRet = btrInsert(hmTxr, tmTxr, imTxrRecLen, INDEXKEY0)
            tmRvr.iProdPct = tmRvf.iUrfCode
        End If
        ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
    ElseIf (igRptCallType = INVOICESJOB And ilListIndex = INV_TAXREGISTER) Then
        'tmrvr.iprodpct = place the type of tax record, 1= tax1, 2 = tax2, 3 = no taxes applicable
        'tmrvf.iBackLogTrfCode - tax reference that transaction uses.  used to show the tax description names and sorting in crystal
        'Trades never have taxes, therefore falls into No tax applicable category even if the advt is taxable
        'for cash/trade splits, only the cash portion is shown as a taxable item

        If tmRvr.iBacklogTrfCode = 0 Then       'if theres a tax pointer, dont change it
            If tmRvr.iMnfItem > 0 Then
                'need to get the SBF for the tax reference code
                tmSbfSrchKey1.lCode = tmRvr.lSbfCode
                ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching advt recd
                If ilRet = BTRV_ERR_NONE Then
                    tmRvr.iBacklogTrfCode = gGetNTRTrfCode(tmSbf.iTrfCode)
                Else
                    tmRvr.iBacklogTrfCode = 0
                End If
            Else
                tmRvr.iBacklogTrfCode = gGetAirTimeTrfCode(tmRvr.iAdfCode, tmRvr.iAgfCode, tmRvr.iAirVefCode)
            End If
         End If

        If tmRvr.iBacklogTrfCode > 0 Then         'some tax should be applicable
            If tmRvr.lTax1 = 0 And tmRvr.lTax2 = 0 Then 'theres a tax reference, but the calculated taxes are 0
                'put out a record to show under the applicable tax reference defined
                tmRvr.iProdPct = 1                'flag for sorting, tax1
                If tmRvr.sCashTrade = "T" Then
                    tmRvr.iProdPct = 3              'not cash, no tax on trades
                    If RptSel!ckcSelC5(0).Value = vbChecked Then        'this is a non-taxable trans, include it?
                        tmRvr.iBacklogTrfCode = 0       'make sure this transaction doesnt get sorted with taxable items since its part of a cash/trade NTR
                        ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
                    End If
                Else
                    ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
                End If

            Else
                If tmRvr.lTax1 <> 0 Then
                    tmRvr.iProdPct = 1                'flag for sorting, tax1
                    ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
                End If
                If tmRvr.lTax2 <> 0 Then
                    tmRvr.iProdPct = 2                'flag for sorting, tax2
                    ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
                End If
            End If
        Else
            'create a record for non-taxable trasactions if requested
            If RptSel!ckcSelC5(0).Value = vbChecked Then        'include non-taxable transactions
                tmRvr.iProdPct = 3                      'invoices without taxes sort to the end
                ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
            End If
        End If

    Else
        If tmRvr.sNet <> ".00" And tmRvr.lDistAmt <> 0 Then         'dont show $0
            If imVehicle And RptSel!ckcSelC3(0).Value = vbChecked And imMajorSet = 1 Then      '9-20-11 if by vehicle with participant share by vehicle, do not
                                                                    'overwrite the vehicle group code.  It has been obtained based on the PIF table
                                                                    'analyzing the date changes
                tmRvr.imnfVefGroup = tmRvr.imnfVefGroup
            Else
                gGetVehGrpSets tmRvr.iAirVefCode, imMinorSet, imMajorSet, ilmnfMinorCode, ilMnfMajorCode    '7-16-02 obtain vehicle group code, some options may not use it
                tmRvr.imnfVefGroup = ilMnfMajorCode
            End If
            'TTP 10117 - Cash Distribution Export
            If RptSel!rbcOutput(3).Value = True Then
                'smExportStatus = mExportCashDist(tmRvr)
                If ilListIndex = COLL_DISTRIBUTE Then smExportStatus = mExportCashDist(tmRvr)
                If ilListIndex = COLL_AGEMONTH Then smExportStatus = mSaveAgeSummary(tmRvr) 'TTP 10164 - Ageing summary by month - export option for Audacy A/R dump
                If ilListIndex = INV_DISTRIBUTE Then
                    If RptSel!rbcSelCInclude(0).Value Then          'detail
                        smExportStatus = mExportInvDist(tmRvr) 'TTP 10118 -Billing Distribution Export to CSV
                    Else
                        smExportStatus = mSaveInvDistSummary(tmRvr)
                    End If
                End If
            Else
                ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
            End If
        End If
    End If
End Sub
'
'       mFindMatchSSCode - find the Sales Source based on the Selling Office
'
'       <input> - selling office code to match on
'       <return> - sales source code for matching selling office
'
'       Search thru array that contains listof selling offices and their
'       associated sales source
'

Public Function mFindMatchSSCode(ilSofCode As Integer, tlSofList() As SOFLIST) As Integer
Dim illoop As Integer
Dim ilMatchSSCode As Integer

    ilMatchSSCode = 0                   'assume no matching selling office found yet
    For illoop = LBound(tlSofList) To UBound(tlSofList)
        If tlSofList(illoop).iSofCode = ilSofCode Then
            ilMatchSSCode = tlSofList(illoop).iMnfSSCode          'Sales source
            Exit For
        End If
    Next illoop

    mFindMatchSSCode = ilMatchSSCode
End Function
'
'   Generate Cash  report for Cash Receipts, Cash Summary and
'               Payment History reports.  Remove the prepass
'               for Cash Receipts and Payment History and
'               combine into this subroutine
'
'   Go thru PHF & RVF for selected date span, & create rvr
'
'   12-1-03 Request to ignore transactions with $0
'   2-11-05 Chg array of PHF/RVF array from integer to long (prevent overflow)

Public Sub gCashGen(ilListIndex As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llLoopTran                                                                            *
'******************************************************************************************

Dim slStr As String
Dim slCode As String
Dim llEarliestDate As Long
Dim llLatestDate As Long
Dim illoop As Integer
Dim slStart As String
Dim slEnd As String
Dim ilRet As Integer
Dim ilError As Integer
'Dim ilLoopTran As Integer
'ReDim ilUseCodes(1 To 1) As Integer
ReDim ilUseCodes(0 To 0) As Integer
Dim ilIncludeCodes As Integer
'ReDim ilUseSOFCodes(1 To 1) As Integer
ReDim ilUseSOFCodes(0 To 0) As Integer
Dim ilIncludeSOFCodes As Integer
ReDim ilUseNTRCodes(0 To 0) As Integer
Dim ilIncludeNTRCodes As Integer
Dim ilmnfMinorCode As Integer
Dim ilMnfMajorCode As Integer
Dim ilIncludeP As Integer
Dim ilIncludeW As Integer
'6/7/15: Check number changed to string
'Dim llCheckNo As Long
Dim slCheckNo As String
Dim slStamp As String
Dim llAmt As Long           '12-01-03
Dim tlTranType As TRANTYPES
ReDim tlRvf(0 To 0) As RVF
ReDim tlMnf(0 To 0) As MNF
Dim llRvfLoop As Long
Dim ilWhichDate As Integer      '0=use tran date, 1 = use date entered
Dim ilEarliestAgeMM As Integer      '9-21-17 implement earliest mm/yy to include for cash by ageing sort.  For all other cash reports, default to 01/1970
Dim ilEarliestAgeYY As Integer
Dim ilOKtoSeeVeh As Integer
Dim ilLoopOnSlsp As Integer
Dim ilHowManyDefined As Integer
Dim slAmount As String
Dim slDollar As String
Dim llProcessPct As Long
Dim slPct As String
Dim llNetDollar As Long
Dim llComm As Long
Dim slOrigNetAmount As String
Dim llContract As Long          'Date: 9/5/2018 added filters: Contract number FYM
Dim llInvoice As Long           'Date: 9/5/2018 added filters: Invoice number FYM


    'RVF & PHF opened in general gObtainPHFRVF routine

    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSofRecLen = Len(tmSof)

    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSlfRecLen = Len(tmSlf)

    hmRvr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRvr, "", sgDBPath & "Rvr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imRvrRecLen = Len(tmRvr)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imVefRecLen = Len(tmVef)

    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imMnfRecLen = Len(tmMnf)
    
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imCHFRecLen = Len(tmChf)

    hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSbfRecLen = Len(tmSbf)
    

    tmRvr.iGenDate(0) = igNowDate(0)        'todays date used for retrieval/removal of records
    tmRvr.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmRvr.lGenTime = lgNowTime

    'setup transaction types to retrieve from history and receivables
    tlTranType.iAdj = False              'adjustments
    tlTranType.iInv = False              'invoices
    tlTranType.iWriteOff = True
    tlTranType.iPymt = True
    tlTranType.iCash = False
    tlTranType.iTrade = False
    tlTranType.iMerch = False
    tlTranType.iPromo = False
    tlTranType.iNTR = True
    tlTranType.iAirTime = True

    ilIncludeP = False
    ilIncludeW = False
    
    If ilListIndex = COLL_PAYHISTORY Or ilListIndex = COLL_CASH Or ilListIndex = COLL_SALESCOMM_COLL Or ilListIndex = COLL_CASHSUM Then
        '8-28-19 use csi cal control vs edit box
        slStr = RptSel!CSI_CalFrom.Text                'Earliest date to retrieve from PRF or RVF
        llEarliestDate = gDateValue(slStr)
        slStr = RptSel!CSI_CalTo.Text               'Latest date to retrieve from PRF or RVF
        llLatestDate = gDateValue(slStr)
        If llLatestDate = 0 Then                    'if end date not entered, use all
            llLatestDate = gDateValue("12/29/2069")
        End If
        slStart = Format$(llEarliestDate, "m/d/yy")
        slEnd = Format$(llLatestDate, "m/d/yy")
    Else
        slStr = RptSel!edcSelCFrom.Text                'Earliest date to retrieve from PRF or RVF
        llEarliestDate = gDateValue(slStr)
        slStr = RptSel!edcSelCTo.Text               'Latest date to retrieve from PRF or RVF
        llLatestDate = gDateValue(slStr)
        If llLatestDate = 0 Then                    'if end date not entered, use all
            llLatestDate = gDateValue("12/29/2069")
        End If
        slStart = Format$(llEarliestDate, "m/d/yy")
        slEnd = Format$(llLatestDate, "m/d/yy")
    End If

    ilWhichDate = 0                     'default to use tran date vs date entered
    If ilListIndex = COLL_PAYHISTORY Then      'always include all payments & journal entries for paymenthistory
        ilIncludeP = True
        ilIncludeW = True

    ElseIf ilListIndex = COLL_CASHSUM Then                  'cash receipts/cash summary
        If RptSel!ckcSelC3(0).Value = vbChecked Then       'payments selected
            ilIncludeP = True
        End If
        If RptSel!ckcSelC3(1).Value = vbChecked Then       'journal entries selected
            ilIncludeW = True
        End If
    ElseIf ilListIndex = COLL_CASH Or ilListIndex = COLL_SALESCOMM_COLL Then        'cash receipts or 1-19-18 sales comm on collections
        If RptSel!ckcSelC3(0).Value = vbChecked Or RptSel!ckcSelC3(1).Value = vbChecked Then
            ilIncludeP = True
        End If
        If RptSel!ckcSelC3(2).Value = vbChecked Then
            ilIncludeW = True
        End If
        If RptSel!rbcSelCSelect(1).Value Then       'use entry date (vs deposit date, tran date)
            ilWhichDate = 1
        End If
        
    End If

    If RptSel!rbcSelC6(0).Value Then           'Air time only
        tlTranType.iNTR = False
    ElseIf RptSel!rbcSelC6(1).Value Then
        tlTranType.iAirTime = False            'include ntr only
    End If
    mGetCashTradeOption tlTranType         ' check user selections for cash/trade/merch/promo
    If imCash Then
        tlTranType.iCash = True
    End If
    If imTrade Then
        tlTranType.iTrade = True
    End If
    If imMerchant Then
        tlTranType.iMerch = True
    End If
    If imPromo Then
        tlTranType.iPromo = True
    End If


    ilRet = gObtainMnfForType("Y", slStamp, tlMnf())        'obtain tran types

    imVehicle = False
    imSlsp = False
    imOffice = False
    imAdvt = False
    imAgency = False
    
    slStr = RptSel!edcText1.Text
    If Trim$(slStr) = "" Then
        ilEarliestAgeMM = 1
        ilEarliestAgeYY = 1970
    Else
        ilRet = gParseItem(slStr, 1, "/", slCode)
        ilEarliestAgeMM = Val(slCode)

        ilRet = gParseItem(slStr, 2, "/", slCode)
        ilEarliestAgeYY = Val(slCode)
        If (ilEarliestAgeYY >= 0) And (ilEarliestAgeYY <= 69) Then
            ilEarliestAgeYY = 2000 + ilEarliestAgeYY
        ElseIf (ilEarliestAgeYY >= 70) And (ilEarliestAgeYY <= 99) Then
            ilEarliestAgeYY = 1900 + ilEarliestAgeYY
        End If

    End If

    If ilListIndex = COLL_CASHSUM Then
        '6/7/15: Changed check number to string
        'llCheckNo = 0                               'include all check #s
        slCheckNo = ""
        If RptSel!rbcSelCSelect(0).Value Then       'vehicle
            imVehicle = True
            mObtainCodes 6, tgCSVNameCode(), ilIncludeCodes, ilUseCodes()
        Else
            imOffice = True
            mObtainCodes 0, tgSOCode, ilIncludeCodes, ilUseCodes()
        End If
        illoop = RptSel!cbcSet1.ListIndex           'Determine vehicle group selected
        imMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())
    ElseIf ilListIndex = COLL_CASH Then
        slStr = RptSel!edcCheck.Text           'selective check #
        '6/7/15: Changed check number to string
        'llCheckNo = Val(slStr)
        slCheckNo = UCase(Trim$(slStr))
        imSlsp = True
        If RptSel!rbcSelC4(1).Value Then       'slsp
            imSlsp = True
            mObtainCodes 5, tgSalesperson(), ilIncludeCodes, ilUseCodes()
            mObtainCodes 7, tgSOCode(), ilIncludeSOFCodes, ilUseSOFCodes()
            mObtainCodes 8, tgMnfCodeCT(), ilIncludeNTRCodes, ilUseNTRCodes()
        End If
        illoop = RptSel!cbcSet1.ListIndex           'Determine vehicle group selected
        imMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())
        
        'Date: 9/5/2018 added Contract and Invoice number filters   FYM
        llContract = 0: llInvoice = 0
        If RptSel!edcContract <> "" Then
            llContract = CLng(RptSel!edcContract)
        End If
        If RptSel!edcInvoice <> "" Then
            llInvoice = CLng(RptSel!edcInvoice)
        End If
    ElseIf ilListIndex = COLL_PAYHISTORY Then
        '6/7/15: Changed check number to string
        'llCheckNo = 0
        slCheckNo = ""
        If RptSel!rbcSelCSelect(0).Value Then           'agy
            imAgency = True
            'setup common area for the sorted list of agencies/direct adv
            ReDim tlSortCode(0 To 0) As SORTCODE

            For illoop = 0 To RptSel!lbcAgyAdvtCode.ListCount - 1 Step 1
                tlSortCode(illoop).sKey = RptSel!lbcAgyAdvtCode.List(illoop)
                ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 1) As SORTCODE
            Next illoop
            mObtainAgyAdvCodes ilIncludeCodes, ilUseCodes(), 2    'assume using lbcselection(2) for list box of direct & agencies
            'mObtainCodes 2, tlSortCode(), ilIncludeCodes, ilUseCodes()
        ElseIf RptSel!rbcSelCSelect(1).Value Then       'advt
            imAdvt = True
            mObtainCodes 0, tgAdvertiser(), ilIncludeCodes, ilUseCodes()
        End If
    'Cash Receipts is the foundation for Sales Commission on Collections.  Same selection controls have been used.
    ElseIf ilListIndex = COLL_SALESCOMM_COLL Then              '1-19-18 Sales Comm on Collections
        ilWhichDate = 0                                         'always use deposit date (trans date)
        slCheckNo = ""
        imSlsp = True
        mObtainCodes 5, tgSalesperson(), ilIncludeCodes, ilUseCodes()
        mObtainCodes 7, tgSOCode(), ilIncludeSOFCodes, ilUseSOFCodes()
        imMajorSet = 0                          'no vehicle groups are used
    End If

    'retrieve only payments and journal entries.  Cash Receipts has an option to include either PO/PI.
    'Generalized rtn will return all "P"
    ' coll_cash may only want phf or rvf  (rbcselc8(0) = phf, rbcSelc8(1) = rvf ) 6-04-08 Dan M
    If (ilListIndex = COLL_CASH) And (RptSel!rbcSelC8(0).Value = True Or RptSel!rbcSelC8(1).Value = True) Then
        'set which file wanted 1 is phf only, 2 is rvf only.
        If RptSel!rbcSelC8(0).Value = True Then     'only history selected
            ilRet = gObtainPhfOrRvf(RptSel, slStart, slEnd, tlTranType, tlRvf(), 1, ilWhichDate)
        Else                                        'only receivables selected
            ilRet = gObtainPhfOrRvf(RptSel, slStart, slEnd, tlTranType, tlRvf(), 2, ilWhichDate)
        End If
    Else                    '1-23-18 Sales Comm by Collections have option to get history, receivables or both files
        If (ilListIndex = COLL_SALESCOMM_COLL) Then
            ReDim tlRvf(0 To 0) As RVF
            'set which file wanted 1 is phf only, 2 is rvf only.
            If RptSel!rbcSelC8(2).Value Then        ' Both
                ilRet = mObtainPhfOrRvf(RptSel, slStart, slEnd, tlTranType, tlRvf(), 3)
            Else                                       'only receivables selected
                If RptSel!rbcSelC8(0).Value Then         'history
                    ilRet = mObtainPhfOrRvf(RptSel, slStart, slEnd, tlTranType, tlRvf(), 1)
                Else
                    ilRet = mObtainPhfOrRvf(RptSel, slStart, slEnd, tlTranType, tlRvf(), 2)
                End If
            End If
        Else        'all other options get both files
            ilRet = gObtainPhfRvf(RptSel, slStart, slEnd, tlTranType, tlRvf(), ilWhichDate) '12-14-06 add parm to indicate to use tran date or entry date
            If ilRet = 0 Then
                Exit Sub
            End If
        End If
    End If
    'transactions types other than paymnts & journal entries have been filtered out thru the call to RVF/PHF
    
    For llRvfLoop = LBound(tlRvf) To UBound(tlRvf) - 1
        tmRvf = tlRvf(llRvfLoop)
        ilRet = mFilterCashTrans(ilListIndex, tlTranType, llEarliestDate, llLatestDate, slCheckNo, ilIncludeP, ilIncludeW, ilEarliestAgeMM, ilEarliestAgeYY, llContract, llInvoice) 'determine date selectivity, cash/trade, airtime NTR
        If ilRet Then       'ok to include, check for selective vheicles vs sales office
            ilRet = mFilterLists(ilIncludeCodes, ilUseCodes())
            If ilRet = True Then        '11-12-03 continue other filters
'           3-15-18 Implementation of Sales Commission by Collection caused Sales Summary and Cash Pymt History to break.  No data was generated for these 2 reports.
'                If (ilListIndex = COLL_CASH) Or (ilListIndex = COLL_SALESCOMM_COLL) Then      'Cash Receipts or 1-19-18 Sales Comm on Collections needs to distinguish between PO & PI
                If (ilListIndex = COLL_SALESCOMM_COLL) Then
'                    If ilIncludeP = True Then       'include some form of the payments
'                        If InStr(tmRvf.sTranType, "P") <> 0 Then        'this is a Payment
'                            If (tmRvf.sTranType = "PI" And Not RptSel!ckcSelC3(0).Value = vbChecked) Or (tmRvf.sTranType = "PO" And Not RptSel!ckcSelC3(1).Value = vbChecked) Then
'                                ilRet = False
'                            End If
'                        End If
'                    End If
'                    If ilRet = True Then    'still a valid transaction, check the slsp office selection
                        'For Sales Comm by collections, need to also check that the
                        '8-5-10 add selectivity by sales office
'                        If ilListIndex = COLL_SALESCOMM_COLL Then
                            'read contract to see if slsp on order
                    ilHowManyDefined = 0
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
                        mFakeChf
                    Else                            'contract # not present
                        mFakeRvrSlsp                'setup slsp & comm from RVf
                    End If
                    For illoop = 0 To 9 Step 1
                        If tmChf.iSlfCode(illoop) > 0 Then
                            ilHowManyDefined = ilHowManyDefined + 1
                        End If
                    Next illoop
                    If ilHowManyDefined = 1 And tmChf.lComm(0) = 0 Then                       'theres only 1 slsp
                        tmChf.lComm(0) = 1000000                'xxx.xxxx
                    Else
                        If ilHowManyDefined > 1 Then
                            ilHowManyDefined = 10                      'more than 1 slsp, process all 10 because some in the
                                                                'middle may not be used
                        End If
                    End If
                    'ilHowManyDefined = 1
            
                    gPDNToStr tmRvf.sNet, 2, slOrigNetAmount
                    For ilLoopOnSlsp = 0 To ilHowManyDefined - 1
                        ilRet = gBinarySearchSlf(tmChf.iSlfCode(ilLoopOnSlsp))
                        If ilRet < 0 Then
                            tmSlf.iSofCode = 0
                        Else
                            tmSlf = tgMSlf(ilRet)
                        End If
                        
                        If (gFilterLists(tmSlf.iCode, ilIncludeCodes, ilUseCodes())) Then
                            If (gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, CLng(tmRvf.iBillVefCode), tmChf.iSlfCode())) Then
                                If gFilterLists(tmSlf.iSofCode, ilIncludeSOFCodes, ilUseSOFCodes()) Then        'office selected?
                                    'obtain the slsp rev split
                                    llProcessPct = tmChf.lComm(ilLoopOnSlsp)          'slsp split % or 100%
                                    slPct = gLongToStrDec(llProcessPct, 4)
                                    'gPDNToStr tmRvf.sNet, 2, slAmount
                                    'tmrvf.snet field is verwritten for each slsp split amount , that is written to prepass
                                    'slDollar = gMulStr(slPct, slAmount)                 'slsp net portion of rev split
                                    slDollar = gMulStr(slPct, slOrigNetAmount)
                                    llNetDollar = Val(gRoundStr(slDollar, ".01", 0))
                                    slDollar = gLongToStrDec(llNetDollar, 2)
                                    gStrToPDN slDollar, 2, 6, tmRvf.sNet
                                    'now get the slsp comm due.  Retrieve slsp comm from slsp file or the contract
                                    If tgSpf.sSubCompany = "Y" Then     'slsp by subcompany not applicable for this report
                                        llProcessPct = 0                'disabled for this report
                                    Else
                                        If tmRvf.sTranType = "WV" Or tmRvf.sTranType = "WB" Then      'penny variance & bad debt do not get comm
                                            llProcessPct = 0
                                        Else
                                            If tmRvf.sTranType = "PI" Or tmRvf.sTranType = "WU" Or tmRvf.sTranType = "WD" Then
                                                If tmRvf.iMnfItem > 0 Then      'NTR transaction
                                                    tmSbfSrchKey1.lCode = tmRvf.lSbfCode
                                                    ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching advt recd
                                                    If ilRet = BTRV_ERR_NONE Then
                                                        llProcessPct = CLng(tmSbf.iCommPct)
                                                    Else
                                                        llProcessPct = 0
                                                    End If
                                                Else
                                                    If tgSpf.sCommByCntr = "Y" Then     'retrieve slsp comm pct from cntr if in use.  Otherwise, take from Slsp profile
                                                        llProcessPct = tmChf.iSlspCommPct(ilLoopOnSlsp)
                                                        If llProcessPct = 10000 Then          'test for 100%, comm wouldnt be the entire rev share
                                                            llProcessPct = 0
                                                        End If
                                                    Else
                                                        llProcessPct = tmSlf.iUnderComm
                                                        If llProcessPct = 1000000 Then            'test for 100%, comm wouldnt be the entire rev share
                                                            llProcessPct = 0
                                                        End If
                                                    End If
                                                End If
                                            Else                    'PO or under defined Wx do not get commission
                                                llProcessPct = 0
                                            End If
                                        End If
                                    End If
                                    slPct = gLongToStrDec(llProcessPct, 2)
                                    slAmount = gLongToStrDec(llNetDollar, 2)
                                    slDollar = gMulStr(slPct, slAmount)             'slsp comm
                                    llComm = Val(gRoundStr(slDollar, ".01", 0))
                                    LSet tmRvr = tmRvf
                                    tmRvr.iSlfCode = tmSlf.iCode        'with splits, the slsp isnt the one stored in receivables (thats the primary one)
                                    tmRvr.lDistAmt = llComm
                                    tmRvr.iProdPct = CSng(llProcessPct)
                                    tmRvr.sSource = "R"
                                    If tmRvr.iPurgeDate(0) > 0 Or tmRvr.sTranType = "HI" Then  'if trans has purge date or its an HI, trans came from history
                                        tmRvr.sSource = "H"
                                    End If

                                    If llNetDollar <> 0 Then
                                        ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
                                    End If
                                  End If
                            End If
                        End If                  'gfilterlists (slsp)

                    Next ilLoopOnSlsp
'                End If
                        
                Else                            'Cash Summary, cash receipts or payment history reports
                    If ilListIndex = COLL_CASH Then     'cash summary and pymt history do not have transaction type option
                        If ilIncludeP = True Then       'include some form of the payments
                             If InStr(tmRvf.sTranType, "P") <> 0 Then        'this is a Payment
                                 If (tmRvf.sTranType = "PI" And Not RptSel!ckcSelC3(0).Value = vbChecked) Or (tmRvf.sTranType = "PO" And Not RptSel!ckcSelC3(1).Value = vbChecked) Then
                                     ilRet = False
                                 End If
                             End If
                         End If
                    End If
                    If tmRvf.iSlfCode <> tmSlf.iCode And tmRvf.iSlfCode <> 0 Then      'only read if not already in mem
                        tmSlfSrchKey.iCode = tmRvf.iSlfCode
                        ilError = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilError <> BTRV_ERR_NONE Then
                            tmSlf.iSofCode = 0
                        End If
                    End If
                    'ilRet is still true (OK to include)
                    If Not gFilterLists(tmSlf.iSofCode, ilIncludeSOFCodes, ilUseSOFCodes()) Then
                        ilRet = False
                    End If
                    
                    If (tmRvf.iMnfItem > 0) And (ilRet) And (tlTranType.iNTR) Then         '11-8-19 if NTR item, and passed previous filters and NTR is included, continue to test item type
                        If Not gFilterLists(tmRvf.iMnfItem, ilIncludeNTRCodes, ilUseNTRCodes()) Then
                            ilRet = False
                        End If
                    End If

                   
                    If (ilRet) Then           'include the entry from list box selectivity
                        LSet tmRvr = tmRvf
                        gGetVehGrpSets tmRvr.iAirVefCode, imMinorSet, imMajorSet, ilmnfMinorCode, ilMnfMajorCode    '7-16-02 obtain vehicle group code, some options may not use it
                        tmRvr.imnfVefGroup = ilMnfMajorCode
                        tmRvr.iProdPct = 0                  'used for mnf tran type code
                        tmRvr.sSource = "R"
                        If tmRvr.iPurgeDate(0) <> 0 Then  'if no purge date, tran not in history
                            tmRvr.sSource = "H"
                        End If
                        For illoop = LBound(tlMnf) To UBound(tlMnf) - 1
                            If Trim$(tlMnf(illoop).sUnitType) = tmRvr.sTranType Then
                                tmRvr.iProdPct = tlMnf(illoop).iCode            'for cash summary, the tran type desription is shown
                                Exit For
                            End If
                        Next illoop
                        gPDNToLong tmRvf.sNet, llAmt
                        If llAmt <> 0 Then '12-1-03
                            ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
                        End If
                    End If      'ilret
'                       End If           '3-15-18 If ilListIndex = COLL_SALESCOMM_COLL
'                    End If      ' ilRet = True
                End If          ' If (ilListIndex = COLL_SALESCOMM_COLL)
            End If              'if ilret after mFilterLists
                     
        End If                  'if ilRet aftermFilterCashTrans
    Next llRvfLoop              'for llRvfloop

    Erase ilUseCodes, tlMnf, tlRvf
    ilRet = btrClose(hmRvr)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSbf)
    btrDestroy hmRvr
    btrDestroy hmVef
    btrDestroy hmMnf
    btrDestroy hmSof
    btrDestroy hmCHF
    btrDestroy hmSbf
    Exit Sub
End Sub

'
'       mFilterCashTrans - determine if transaction should be included
'       for selectivity based on Transaction date, cash/trade/merch/promo,
'       Air Time or NTR
'
'       <input> ilListIndex - report (cash, cash summary)
'               tlTranType - structure containing inclusion/exclusion parms
'               llEarlestDate - User entered start date
'               llLatestDate - user entered end date
'               llCheckNo - selective check # (zero indicates all)
'               ilIncludeP - include Payments (vs journal entries)
'               ilIncludeW - include journal entries
'               Earliest ageing mm/yy used for Cash by Ageing report, All other cash reports default to 01/1970 and question is not asked
'               ilEarliestAgeMM - earliest ageing period to include
'               ilEArliestAgeYY - earliest ageing year to include
'       <return> true -include transaction, else false to exclude
'6/7/15: Changed Check Number from Long to string
'Public Function mFilterCashTrans(ilListIndex As Integer, tlTranType As TRANTYPES, llEarliestDate As Long, llLatestDate As Long, llCheckNo As Long, ilIncludeP As Integer, ilIncludeW As Integer) As Integer
Public Function mFilterCashTrans(ilListIndex As Integer, tlTranType As TRANTYPES, llEarliestDate As Long, llLatestDate As Long, slCheckNo As String, ilIncludeP As Integer, ilIncludeW As Integer, ilEarliestAgeMM As Integer, ilEarliestAgeYY As Integer, llContract As Long, llInvoice As Long) As Integer
Dim slStr As String
Dim llTranDate As Long
Dim ilOk  As Integer

    ilOk = True             'assume to include this trans

    If ilListIndex = COLL_CASH Then
        If RptSel!rbcSelCSelect(1).Value = True Then
            gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slStr
        Else
            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slStr
        End If
    ElseIf ilListIndex = COLL_CASHSUM Or ilListIndex = COLL_PAYHISTORY Then
        gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slStr
    End If
    llTranDate = gDateValue(slStr)                  'convert trans date to test if within requested limits
    'Sales Commissions on Collections have been date tested in mObtainPHForRVF routine because the entire file had to be searched
    'using both entry dates and tran dates
    If (llTranDate >= llEarliestDate And llTranDate <= llLatestDate) Or (ilListIndex = COLL_SALESCOMM_COLL) Then
        'transaction with date span is ok,
        'test for inclusion of Cash/Trade/Merchandising/Promotions
        If tmRvf.sCashTrade = "C" And Not tlTranType.iCash Then
            ilOk = False
        End If
        If tmRvf.sCashTrade = "T" And Not tlTranType.iTrade Then
            ilOk = False
        End If
        If tmRvf.sCashTrade = "M" And Not tlTranType.iMerch Then
            ilOk = False
        End If
        If tmRvf.sCashTrade = "P" And Not tlTranType.iPromo Then
            ilOk = False
        End If
If tmRvf.lCntrNo = 2 Then
ilOk = ilOk
End If
        'test for inclusion/Exclusion of AirTime vs NTR
        If tmRvf.iMnfItem = 0 And Not tlTranType.iAirTime Then          'air time
            ilOk = False
        End If
        If tmRvf.iMnfItem > 0 And Not tlTranType.iNTR Then          'NTR
            ilOk = False
        End If
        '6/7/15: Changed check number to string
        'If llCheckNo <> 0 And llCheckNo <> tmRvf.lCheckNo Then  'see if user wants a specific check #
        If Trim$(slCheckNo) <> "" And Trim$(slCheckNo) <> "0" And UCase$(Trim$(slCheckNo)) <> UCase$(Trim$(tmRvf.sCheckNo)) Then  'see if user wants a specific check #
            ilOk = False
        End If

        If (Left$(tmRvf.sTranType, 1) = "P") And Not ilIncludeP Then
            ilOk = False
        End If
        If (Left$(tmRvf.sTranType, 1) = "W") And Not ilIncludeW Then
            ilOk = False
        End If
        
        If ilListIndex = COLL_CASH Then
            '9-21-17 Implemented for Cash by Ageing , all other cash has it defaulted to 01/1970
            If tmRvf.iAgingYear >= ilEarliestAgeYY Then
                'valid year, find if valid ageing month
                If tmRvf.iAgePeriod < ilEarliestAgeMM Then
                    ilOk = False
                End If
            Else
                ilOk = False
            End If
            'Date: 9/5/2018 added filters: Contract and Invoice number  FYM
            If llContract > 0 Then
                If (tmRvf.lCntrNo <> llContract) Then
                    ilOk = False
                End If
            End If
            If llInvoice > 0 Then
                If (tmRvf.lInvNo <> llInvoice) Then
                    ilOk = False
                End If
            End If
        End If
    Else
        ilOk = False
    End If
    mFilterCashTrans = ilOk
End Function
'
'       mFilterLists - check the option and which list boxes to test
'       for inclusion/exclusion
'
'       <input>
'               ilIncludeCodes = true to include codes in array;
'                                false to exclude codes in array
'               ilUseCodes()- array of codes to include/exclude
'       <return> true = include transaction, else false to exclude
'
Public Function mFilterLists(ilIncludeCodes As Integer, ilUseCodes() As Integer) As Integer
Dim ilCompare As Integer
Dim ilTemp As Integer
Dim ilFoundOption As Integer
Dim ilRet As Integer
Dim ilListIndex As Integer

    ilListIndex = RptSel!lbcRptType.ListIndex               'report option from  Invoicing or Collections
    
    If ilListIndex = COLL_SALESCOMM_COLL Then               'if sales commission on collections, do not test for slsp here
                                                            'need to test all split slsp
        mFilterLists = True
        Exit Function
    End If
    
    ilFoundOption = False
    If imVehicle Then
        ilCompare = tmRvf.iAirVefCode
        If tmRvf.sTranType = "PO" And ilCompare = 0 Then        'no vehicle reference but include if its a PO since
                                                                'theres nothing to compare against but need to see the payment
            ilFoundOption = True
        End If
    ElseIf imOffice Then
        ilRet = gBinarySearchSlf(tmRvf.iSlfCode)
        If ilRet = -1 Then
            ilCompare = 0
            ilFoundOption = True                'no slsp reference, include since nothing to compare (might be PO)
        Else
            ilCompare = tgMSlf(ilRet).iSofCode
        End If
    ElseIf imSlsp Then
        ilRet = gBinarySearchSlf(tmRvf.iSlfCode)
        If ilRet = -1 Then
            ilCompare = tmRvf.iSlfCode
            ilFoundOption = True                'no slsp reference, include since nothing to compare (might be PO)
        Else
            ilCompare = tgMSlf(ilRet).iCode
        End If
    ElseIf imAdvt Then
        ilRet = gBinarySearchAdf(tmRvf.iAdfCode)
        If ilRet = -1 Then
           'ilCompare = tmRvf.iAdfCode
            'ilFoundOption = True                'no advt reference, include since nothing to compare (might be PO)
        Else
            ilCompare = tmRvf.iAdfCode
        End If
    ElseIf imAgency Then
        If tmRvf.iAgfCode = 0 Then          'direct
            ilRet = gBinarySearchAdf(tmRvf.iAdfCode)
            If ilRet = -1 Then
               ilCompare = tmRvf.iAdfCode
            Else
                ilCompare = tmRvf.iAdfCode
            End If
        Else
            ilRet = gBinarySearchAgf(tmRvf.iAgfCode)
            If ilRet = -1 Then
               ilCompare = tmRvf.iAgfCode
            Else
                ilCompare = tmRvf.iAgfCode
            End If
        End If
    End If


    If imVehicle Or imSlsp Or imOffice Or imAdvt Or imAgency Then
        If ilIncludeCodes Then
            If imAgency Then
                If tmRvf.iAgfCode = 0 Then      'direct
                    ilCompare = -ilCompare
                End If
            End If
            For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                If ilUseCodes(ilTemp) = ilCompare Then
                    ilFoundOption = True
                    Exit For
                End If
            Next ilTemp
        Else
            ilFoundOption = True        '8/23/99 when more than half selected, selection fixed
            If imAgency Then
                If tmRvf.iAgfCode = 0 Then      'direct
                    ilCompare = -ilCompare
                End If
            End If
            For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                If ilUseCodes(ilTemp) = ilCompare Then
                    ilFoundOption = False
                    Exit For
                End If
            Next ilTemp
        End If
    End If
    mFilterLists = ilFoundOption
End Function

'
'
'          Generate all payments generated from PO (OnAccount) payments.
'           These payments are flagged with an Action Code = "A"
'
'       Selectivity by transaction date and date entered.  Null dates assume all
'       Obtain from RVF only
'
'       9-11-03
'
'       3-19-04 Change to obtain from PHF as well as RVF
'       2-11-05 Chg array of PHF/RVF array from integer to long (prevent overflow)

Public Sub gGenPOApply()
Dim slStr As String
Dim llEarliestDate As Long
Dim llLatestDate As Long
Dim illoop As Integer
Dim slStart As String
Dim slEnd As String
Dim ilRet As Integer
Dim ilError As Integer
Dim llLoopTran As Long                  '2-11-05 chg to long
'ReDim ilUseCodes(1 To 1) As Integer
ReDim ilUseCodes(0 To 0) As Integer
Dim ilIncludeCodes As Integer
Dim ilmnfMinorCode As Integer
Dim ilMnfMajorCode As Integer
Dim llDate As Long

Dim tlTranType As TRANTYPES
ReDim tlRvf(0 To 0) As RVF



    hmRvr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRvr, "", sgDBPath & "Rvr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imRvrRecLen = Len(tmRvr)
    'RVF is opened in general READ routine

    'Setup transaction date selectivity
    '8-30-19 use csi cal control vs edit box
'    slStr = RptSel!edcSelCTo.Text                'Earliest date to retrieve from PRF or RVF
    slStr = RptSel!CSI_CalFrom2.Text                'Earliest date to retrieve from PRF or RVF
    llEarliestDate = gDateValue(slStr)
'    slStr = RptSel!edcSelCTo1.Text               'Latest date to retrieve from PRF or RVF
    slStr = RptSel!CSI_CalTo2.Text               'Latest date to retrieve from PRF or RVF
    llLatestDate = gDateValue(slStr)
    If llLatestDate = 0 Then                    'if end date not entered, use all
        llLatestDate = gDateValue("12/29/2069")
    End If
    slStart = Format$(llEarliestDate, "m/d/yy")
    slEnd = Format$(llLatestDate, "m/d/yy")

    tmRvr.iGenDate(0) = igNowDate(0)        'todays date used for retrieval/removal of records
    tmRvr.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmRvr.lGenTime = lgNowTime

    'setup transaction types to retrieve from history and receivables
    tlTranType.iAdj = False              'adjustments
    tlTranType.iInv = False              'invoices
    tlTranType.iWriteOff = False
    tlTranType.iPymt = True
    tlTranType.iCash = True
    tlTranType.iTrade = True
    tlTranType.iMerch = False
    tlTranType.iPromo = False
    tlTranType.iNTR = False
    tlTranType.iAirTime = True

    imVehicle = False
    imSlsp = False
    imOffice = False
    imAdvt = False
    imAgency = True



    mObtainAgyAdvCodes ilIncludeCodes, ilUseCodes(), 2    'assume using lbcselection(2) for list box of direct & agencies

    illoop = RptSel!cbcSet1.ListIndex           'Determine vehicle group selected
    imMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())

    'retrieve only payments and journal entries
    ilRet = gObtainPhfOrRvf(RptSel, slStart, slEnd, tlTranType, tlRvf(), 3, 0) '3-19-04 obtain Payments from RVF & phf (previously only rvf)
    If ilRet = 0 Then
        Exit Sub
    End If

    'setup filter of creation (entered) date  selectivity
    '8-30-19 use csi calendar control vs edit box
'    slStr = RptSel!edcSelCFrom.Text                'Earliest date to retrieve from PRF or RVF
    slStr = RptSel!CSI_CalFrom.Text                'Earliest date to retrieve from PRF or RVF
    llEarliestDate = gDateValue(slStr)
'    slStr = RptSel!edcSelCFrom1.Text               'Latest date to retrieve from PRF or RVF
    slStr = RptSel!CSI_CalTo.Text               'Latest date to retrieve from PRF or RVF
    llLatestDate = gDateValue(slStr)
    If llLatestDate = 0 Then                    'if end date not entered, use all
        llLatestDate = gDateValue("12/29/2069")
    End If
    slStart = Format$(llEarliestDate, "m/d/yy")
    slEnd = Format$(llLatestDate, "m/d/yy")


    'transactions types other than paymnts & journal entries have been filtered out thru the call to RVF/PHF
    For llLoopTran = LBound(tlRvf) To UBound(tlRvf) - 1
        tmRvf = tlRvf(llLoopTran)

        'transaction date was filter out from general rvf read routine
        gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slStr     'must be within trans date & date entered selectivity
        llDate = gDateValue(slStr)                  'convert trans date to test if within requested limits
        ilRet = mFilterLists(ilIncludeCodes, ilUseCodes())
        If ilRet Then           'include the entry from list box selectivity
            LSet tmRvr = tmRvf
            If tmRvr.sAction = "A" And llDate >= llEarliestDate And llDate <= llLatestDate Then
                gGetVehGrpSets tmRvr.iAirVefCode, imMinorSet, imMajorSet, ilmnfMinorCode, ilMnfMajorCode    '7-16-02 obtain vehicle group code, some options may not use it
                tmRvr.imnfVefGroup = ilMnfMajorCode

                ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
                On Error GoTo gGenPOApplyErr
                gBtrvErrorMsg ilRet, "btrInsert, gGenPOApply ", RptSel
                On Error GoTo 0
            End If
        End If
    Next llLoopTran

    Erase ilUseCodes, tlRvf
    ilRet = btrClose(hmRvr)
    btrDestroy hmRvr
    Exit Sub

gGenPOApplyErr:
    Exit Sub
End Sub
'
'
'           For agency selection, use the list of agencies & direct advertisers.
'           Need special code to handle the direct advertisers ( vs just agencies)
'           9-12-03
'
'           <input> ilincludecodes - true if use the array of codes to include
'                                    false if use the array of codes to exclude
'                   ilusecodes() - array of codes to include or exclude
'                   ilIndex - index of which lbcselection box (10-9-03)
Public Sub mObtainAgyAdvCodes(ilIncludeCodes As Integer, ilUseCodes() As Integer, ilIndex As Integer)
Dim ilHowManyDefined As Integer
Dim ilHowMany As Integer
Dim illoop As Integer
Dim slNameCode As String
Dim ilRet As Integer
Dim slCode As String

    ilHowManyDefined = RptSel!lbcSelection(ilIndex).ListCount
    ilHowMany = RptSel!lbcSelection(ilIndex).SelCount
    If ilHowMany > ilHowManyDefined / 2 Then    'more than half selected
        ilIncludeCodes = False
    Else
        ilIncludeCodes = True
    End If
    For illoop = 0 To RptSel!lbcSelection(ilIndex).ListCount - 1 Step 1
        slNameCode = RptSel!lbcAgyAdvtCode.List(illoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If RptSel!lbcSelection(ilIndex).Selected(illoop) And ilIncludeCodes Then               'selected ?
            If InStr(slNameCode, "/Direct") = 0 And InStr(slNameCode, "/Non-") = 0 Then         'not a direct, and not a direct that was changed to reg agency, this is reg agency
                ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
            Else
                ilUseCodes(UBound(ilUseCodes)) = -Val(slCode)
            End If
            ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
        Else        'exclude these
            If (Not RptSel!lbcSelection(ilIndex).Selected(illoop)) And (Not ilIncludeCodes) Then
                If InStr(slNameCode, "/Direct") = 0 And InStr(slNameCode, "/Non-") = 0 Then     'not a direct, and not a direct that was changed to reg agency, this is reg agency
                    ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
                Else
                    ilUseCodes(UBound(ilUseCodes)) = -Val(slCode)
                End If
                ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
            End If
        End If
    Next illoop
End Sub
'
'
'           Generate a credit/debit invoice created from receivables
'           adjustments only (AN transactions).  All AN are gathered
'           for a date span (date entered or invoice date or both).
'
'           11-8-03
'
'       2-11-05 Chg array of PHF/RVF array from integer to long (prevent overflow)

Public Sub gGenCreditMemo()
Dim slStr As String
Dim llEarliestDate As Long
Dim llLatestDate As Long
Dim slStart As String
Dim slEnd As String
Dim ilRet As Integer
Dim ilError As Integer
Dim llLoopTran As Long                  '2-11-05 chg to long
'ReDim ilUseCodes(1 To 1) As Integer
ReDim ilUseCodes(0 To 0) As Integer
Dim ilIncludeCodes As Integer
Dim llDate As Long
Dim llContrNo As Long
Dim llNet As Long

Dim tlTranType As TRANTYPES
ReDim tlRvf(0 To 0) As RVF


    hmRvr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRvr, "", sgDBPath & "Rvr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imRvrRecLen = Len(tmRvr)

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imCHFRecLen = Len(tmChf)

    'RVF is opened in general READ routine

    'Setup transaction date selectivity
'    slStr = RptSel!edcSelCTo.Text                'Earliest date to retrieve from RVF
    slStr = RptSel!CSI_CalFrom2.Text                'Earliest date to retrieve from RVF
    llEarliestDate = gDateValue(slStr)
'    slStr = RptSel!edcSelCTo1.Text               'Latest date to retrieve from RVF
    slStr = RptSel!CSI_CalTo2.Text               'Latest date to retrieve from RVF
    llLatestDate = gDateValue(slStr)
    If llLatestDate = 0 Then                    'if end date not entered, use all
        llLatestDate = gDateValue("12/29/2069")
    End If
    '2-12-04 If nothing in earliest date, default to an early date
    If llEarliestDate = 0 Then
        llEarliestDate = gDateValue("1/01/1980")
    End If
    slStart = Format$(llEarliestDate, "m/d/yy")
    slEnd = Format$(llLatestDate, "m/d/yy")

    slStr = RptSel!edcCheck.Text           'selective check #
    llContrNo = Val(slStr)

    tmRvr.iGenDate(0) = igNowDate(0)        'todays date used for retrieval/removal of records
    tmRvr.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmRvr.lGenTime = lgNowTime

    'setup transaction types to retrieve from history and receivables
    tlTranType.iAdj = True              'adjustments
    tlTranType.iInv = False              'invoices
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iCash = True
    tlTranType.iTrade = True
    tlTranType.iMerch = False
    tlTranType.iPromo = False
    tlTranType.iNTR = True
    tlTranType.iAirTime = True

    imVehicle = False
    imSlsp = False
    imOffice = False
    imAdvt = False
    imAgency = True



    mObtainAgyAdvCodes ilIncludeCodes, ilUseCodes(), 7    'use lbcselection(7) for list box of direct & agencies

    'retrieve only billing adjustment (AN) entries
    ilRet = gObtainPhfOrRvf(RptSel, slStart, slEnd, tlTranType, tlRvf(), 2, 0) 'obtain billing adjustments from RVF only
    If ilRet = 0 Then
        Exit Sub
    End If

    'setup filter of creation (entered) date  selectivity
'    slStr = RptSel!edcSelCFrom.Text                'Earliest date to retrieve from  RVF
    slStr = RptSel!CSI_CalFrom.Text                'Earliest date to retrieve from  RVF
    llEarliestDate = gDateValue(slStr)
'    slStr = RptSel!edcSelCFrom1.Text               'Latest date to retrieve from  RVF
    slStr = RptSel!CSI_CalTo.Text               'Latest date to retrieve from  RVF
    llLatestDate = gDateValue(slStr)
    If llLatestDate = 0 Then                    'if end date not entered, use all
        llLatestDate = gDateValue("12/29/2069")
    End If
    slStart = Format$(llEarliestDate, "m/d/yy")
    slEnd = Format$(llLatestDate, "m/d/yy")


    'transactions types other than billing adjustment entries have been filtered out thru the call to RVF/PHF
    For llLoopTran = LBound(tlRvf) To UBound(tlRvf) - 1
        tmRvf = tlRvf(llLoopTran)

        'transaction date was filter out from general rvf read routine
        gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slStr     'must be within trans date & date entered selectivity
        llDate = gDateValue(slStr)                  'convert trans date to test if within requested limits
        ilRet = mFilterLists(ilIncludeCodes, ilUseCodes())
        If ilRet Then           'include the entry from list box selectivity
            LSet tmRvr = tmRvf
            If llDate >= llEarliestDate And llDate <= llLatestDate And ((llContrNo = 0) Or (llContrNo > 0 And tmRvf.lCntrNo = llContrNo)) Then 'filter out by creation date selectivity
                'get the contract code and place in RVR so that the buyer and billing type can be
                'extracted and printed in header
                tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'get matching contr recd

                Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")
                     ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                mFakeChf        'default some values in the chf buffer if no contract exists

                tmRvr.lDistAmt = tmChf.lCode    'use another field to designate contract code for crystal reports
                gPDNToLong tmRvf.sNet, llNet
                If llNet <> 0 Then         'dont write out zero records
                    ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
                    On Error GoTo gGenMemoInvErr
                    gBtrvErrorMsg ilRet, "btrInsert, gGenMemoInv ", RptSel
                    On Error GoTo 0
                End If
            End If
        End If
    Next llLoopTran

    Erase ilUseCodes, tlRvf
    ilRet = btrClose(hmRvr)
    btrDestroy hmRvr
    Exit Sub

gGenMemoInvErr:
    Exit Sub
End Sub
'
'
'       Generate Mailing Labels from Agencies and Direct Advertisers
'       Place the requested list in IVR formatted with the full mailing
'       address.  The report will dump 2 or 3 columns across
'       10-13-03
'
'       2-19-05 implement vehicle labels
'
'       tmIvr.sAddr(1) = Payee
'       tmIvr.sAddr(2) = Attn:  Buyer or Paybles Name .  If none, line 2 of address
'       tmIvr.sAddr(3) = If   buyer/payables name, line 2 of addr, else line 3 of addr
'       tmIvr.sAddr(4) = if buyer/payables name, line 3 of addr, else line 4 of addr
'       tmIvr.sAddr(5) = if buyer/paybles name, line 4 of addr, else line 5 of addr
'
Public Sub gGenMailLabels()
                                          'or agy codes not to process
Dim ilRet As Integer
Dim illoop As Integer
Dim slNameCode As String
Dim slCode As String
Dim ilVef As Integer
Dim ilLineInx As Integer
Dim ilAddrLoop As Integer

    If RptSel!rbcSelC8(0).Value = True Then     'labels by payee

        hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmAgf)
            btrDestroy hmAgf
            Exit Sub
        End If
        imAgfRecLen = Len(tmAgf)

        hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmAdf)
            ilRet = btrClose(hmAgf)
            btrDestroy hmAdf
            btrDestroy hmAgf
            Exit Sub
        End If
        imAdfRecLen = Len(tmAdf)

        hmIvr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmIvr, "", sgDBPath & "Ivr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmIvr)
            ilRet = btrClose(hmAdf)
            ilRet = btrClose(hmAgf)
            btrDestroy hmIvr
            btrDestroy hmAdf
            btrDestroy hmAgf
            Exit Sub
        End If
        imIvrRecLen = Len(tmIvr)

        hmPnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmPnf, "", sgDBPath & "Pnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmPnf)
            ilRet = btrClose(hmIvr)
            ilRet = btrClose(hmAdf)
            ilRet = btrClose(hmAgf)
            btrDestroy hmPnf
            btrDestroy hmIvr
            btrDestroy hmAdf
            btrDestroy hmAgf
            Exit Sub
        End If
        imPnfRecLen = Len(tmPnf)

        imVehicle = False
        imSlsp = False
        imOffice = False
        imAdvt = False
        imAgency = True

        For illoop = 0 To RptSel!lbcSelection(2).ListCount - 1 Step 1
            slNameCode = RptSel!lbcAgyAdvtCode.List(illoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If RptSel!lbcSelection(2).Selected(illoop) Then                'selected ?
                If InStr(slNameCode, "/Direct") <> 0 Or InStr(slNameCode, "/Non-") <> 0 Then             'not a direct or not a direct that has been chged to a regular agency, this is reg agency
                    tmAdfSrchKey.iCode = Val(slCode)
                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching advt recd
                    If ilRet <> BTRV_ERR_NONE Then
                        tmAdf.sName = "Missing"
                    End If
                    'tmIvr.sAddr(1) = tmAdf.sName
                    tmIvr.sAddr(0) = tmAdf.sName
                    mLabelContact tmAdf.iPnfBuyer, tmAdf.iPnfPay        'determine contact
                    mLabelAddr tmAdf.sCntrAddr(), tmAdf.sBillAddr()
                Else
                    'regular agency
                    tmAgfSrchKey.iCode = Val(slCode)
                    ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching agency recd
                    If ilRet <> BTRV_ERR_NONE Then
                        tmAgf.sName = "Missing"
                    End If
                    'tmIvr.sAddr(1) = tmAgf.sName
                    tmIvr.sAddr(0) = tmAgf.sName
                    mLabelContact tmAgf.iPnfBuyer, tmAgf.iPnfPay        'determine contact
                    mLabelAddr tmAgf.sCntrAddr(), tmAgf.sBillAddr()
                End If
                tmIvr.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
                tmIvr.iGenDate(1) = igNowDate(1)
                gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                tmIvr.lGenTime = lgNowTime
                tmIvr.lCode = 0
                ilRet = btrInsert(hmIvr, tmIvr, imIvrRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    'error in writing prepass
                End If
            End If
        Next illoop

        ilRet = btrClose(hmPnf)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmAgf)
        btrDestroy hmPnf
        btrDestroy hmIvr
        btrDestroy hmAdf
        btrDestroy hmAgf

    Else                    'labels by vehicle
        hmIvr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmIvr, "", sgDBPath & "Ivr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmIvr)
            btrDestroy hmIvr
            Exit Sub
        End If
        imIvrRecLen = Len(tmIvr)


        For illoop = 0 To RptSel!lbcSelection(0).ListCount - 1 Step 1
            If RptSel!lbcSelection(0).Selected(illoop) Then                'selected vehicle
                slNameCode = tgAirNameCode(illoop).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVef = gBinarySearchVef(Val(slCode))
                If ilVef <> -1 Then
                    ilLineInx = 2                               'line 1 of address is always vehiclename
                    'tmIvr.sAddr(2) = ""
                    'tmIvr.sAddr(3) = ""
                    'tmIvr.sAddr(4) = ""
                    'tmIvr.sAddr(5) = ""
                    tmIvr.sAddr(1) = ""
                    tmIvr.sAddr(2) = ""
                    tmIvr.sAddr(3) = ""
                    tmIvr.sAddr(4) = ""
                    If Trim$(tgMVef(ilVef).sAddr(0)) <> "" Then         'must have an address to print this vehicles label
                        'tmIvr.sAddr(1) = Trim$(tgMVef(ilVef).sName)
                        tmIvr.sAddr(0) = Trim$(tgMVef(ilVef).sName)
                        If Trim$(tgMVef(ilVef).sContact) <> "" Then
                            'tmIvr.sAddr(2) = "Attn: " + tgMVef(ilVef).sContact
                            tmIvr.sAddr(1) = "Attn: " + tgMVef(ilVef).sContact
                            ilLineInx = ilLineInx + 1
                        End If
                         For ilAddrLoop = 0 To 2            'format the address, if blank, dont leave a blank line on the labels
                            If Trim$(tgMVef(ilVef).sAddr(ilAddrLoop)) <> "" Then
                                'tmIvr.sAddr(ilLineInx) = Trim$(tgMVef(ilVef).sAddr(ilAddrLoop))
                                tmIvr.sAddr(ilLineInx - 1) = Trim$(tgMVef(ilVef).sAddr(ilAddrLoop))
                                ilLineInx = ilLineInx + 1
                            End If
                         Next ilAddrLoop

                        tmIvr.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
                        tmIvr.iGenDate(1) = igNowDate(1)
                        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                        tmIvr.lGenTime = lgNowTime
                        tmIvr.lCode = 0
                        ilRet = btrInsert(hmIvr, tmIvr, imIvrRecLen, INDEXKEY0)
                        If ilRet <> BTRV_ERR_NONE Then
                            'error in writing prepass
                        End If

                    End If
                End If
            End If
        Next illoop
        ilRet = btrClose(hmIvr)
        btrDestroy hmIvr

    End If



End Sub
'
'
'           mLabelContact - retrieve the buyer or payables contact name from
'           Personnel file if applicable
'           <input> ilBuyer - pnf code for buyer name
'                   ilPayables - pnf code for payables
'          Test rbcSelC1Include(0-2) to show Buyer, Payables or None
'          <return> tmIvr.sname = Contact name or blank
Public Sub mLabelContact(ilBuyer As Integer, ilPayables As Integer)
Dim ilRet As Integer

    'tmIvr.sAddr(2) = ""
    tmIvr.sAddr(1) = ""
    If Not RptSel!rbcSelC4(2).Value = True Then       'dont show anything in Attn:
        'tmIvr.sAddr(2) = "Attn: "
        tmIvr.sAddr(1) = "Attn: "
        If RptSel!rbcSelC4(0).Value = True Then
            tmPnfSrchKey.iCode = ilBuyer
        Else
            tmPnfSrchKey.iCode = ilPayables
        End If
        ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching advt recd
        If ilRet <> BTRV_ERR_NONE Then
            tmPnf.sName = ""
            'tmIvr.sAddr(2) = ""
            tmIvr.sAddr(1) = ""
        End If
        'tmIvr.sAddr(2) = Trim$(tmIvr.sAddr(2)) & tmPnf.sName
        tmIvr.sAddr(1) = Trim$(tmIvr.sAddr(1)) & tmPnf.sName
    End If
End Sub
'
'
'           mLabelADdr - build the contract or billing address into IVR prepass
'           <input> slcontrAddr - contract address from either ADF or AGF
'                   slBillAddr - billing addr fro either ADF or AGF
'           Test Rptsel!rbcSelCInclude(0) for contract addr
Public Sub mLabelAddr(slContrAddr() As String * 40, slBillAddr() As String * 40)
Dim slAddress(0 To 2) As String
Dim illoop As Integer
Dim ilIndex As Integer

    If RptSel!rbcSelCInclude(0).Value Then          'use contract address
        For illoop = 0 To 2
            slAddress(illoop) = Trim(slContrAddr(illoop))
        Next illoop
    Else                                            'use billing address
        For illoop = 0 To 2
            slAddress(illoop) = Trim(slBillAddr(illoop))
        Next illoop
        If Trim$(slBillAddr(0)) = "" Then              'billing address is blank, default to contract addr
            For illoop = 0 To 2
                slAddress(illoop) = Trim(slContrAddr(illoop))
            Next illoop
        End If
    End If
    'If tmIvr.sAddr(2) = "" Then     'is there something in Attn: line?
    If tmIvr.sAddr(1) = "" Then     'is there something in Attn: line?
        ilIndex = 2
    Else
        ilIndex = 3
    End If
    For illoop = 0 To 2
        tmIvr.sAddr(ilIndex - 1) = slAddress(illoop)
        ilIndex = ilIndex + 1
    Next illoop
End Sub

'
'           gGenInvSummary - generate an invoice from the receivables that
'           looks like Invoice Form #1, and only gives a bottom lines total
'           of all types of transactions for the contract.  That is, NTR,
'           air time time, etc will be combined on the same invoice, without
'           showing what makes up the total (unless detail requested for internal
'           verification purposes).  History and Receivable transactins will be
'           included for IN & AN (by option).  If both air time & NTR exists on the
'           same contract, the air time invoice # is the number shown on the inv.
'           If no air time, whatever inv # it finds first is shown on the inv.
'
'            6-28-05
Public Sub gGenInvSummary()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slEnd                         ilError                                                 *
'******************************************************************************************

Dim slStr As String
Dim llEarliestDate As Long
Dim llLatestDate As Long
Dim slStartStd As String
Dim slEndStd As String
Dim slStart As String
Dim ilRet As Integer
Dim llLoopTran As Long                  '2-11-05 chg to long
Dim llDate As Long
Dim llContrNo As Long
Dim llNet As Long
Dim ilIncludeCodes As Integer               'true = include codes stored in ilusecode array,
                                        'false = exclude codes store din ilusecode array
'ReDim ilUseCodes(1 To 1) As Integer       'valid advt, agency or vehicles codes to process--
ReDim ilUseCodes(0 To 0) As Integer       'valid advt, agency or vehicles codes to process--
                                            'or advt, agy or vehicles codes not to process


Dim tlTranType As TRANTYPES
ReDim tlRvf(0 To 0) As RVF


    hmRvr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRvr, "", sgDBPath & "Rvr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        MsgBox "Cannot open RVF - Invoice Summary aborted", vbCritical
        Exit Sub
    End If
    imRvrRecLen = Len(tmRvr)

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        MsgBox "Cannot open CHF - Invoice Summary aborted", vbCritical
        btrClose hmRvf
        btrDestroy hmRvf
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmCxf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        MsgBox "Cannot open CXF - Invoice Summary aborted", vbCritical
        btrClose hmRvf
        btrClose hmCHF
        btrDestroy hmRvf
        btrDestroy hmCHF
        Exit Sub
    End If
    imCxfRecLen = Len(tmCxf)

    'RVF is opened in general READ routine

    'Setup transaction date selectivity based on month and year entered;
    'get std bdcst dates
    slStr = RptSel!edcSelCTo.Text   'Month user input
    'slStart = mVerifyMonth(slStr)
    slStart = gVerifyMonth(slStr)           '8-15-14 change to a global routine
    slStart = slStart & "/15/"
    slStr = RptSel!edcSelCTo1.Text  'Year user input
    slStart = slStart & Trim$(str$(gVerifyYear(slStr)))

    'Get standard broadcast month start and end dates based off user input month and year
    slStartStd = gObtainStartStd(slStart)
    llEarliestDate = gDateValue(slStartStd)
    slEndStd = gObtainEndStd(slStart)
    llLatestDate = gDateValue(slEndStd)

    slStr = slStartStd & "-" & slEndStd
    If Not gSetFormula("InvMonth", "'" & slStr & "'") Then      'pass crystal the billng period start & end dates to print
        MsgBox "Error calling InvMonth Formula, Invoice Summary aborted", vbCritical
        Exit Sub
    End If

    slStr = RptSel!edcCheck.Text           'selective check #
    llContrNo = Val(slStr)

    tmRvr.iGenDate(0) = igNowDate(0)        'todays date used for retrieval/removal of records
    tmRvr.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmRvr.lGenTime = lgNowTime

    'setup transaction types to retrieve from history and receivables
    tlTranType.iAdj = True              'adjustments
    If RptSel!ckcSelC7.Value <> vbChecked Then       'dont include if not checked
        tlTranType.iAdj = False
    End If
    tlTranType.iInv = True              'invoices
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iCash = True
    tlTranType.iTrade = True
    tlTranType.iMerch = False
    tlTranType.iPromo = False
    tlTranType.iNTR = True
    tlTranType.iAirTime = True

    imVehicle = False
    imSlsp = False
    imOffice = False
    imAdvt = False
    imAgency = False


    If RptSel!rbcSelCSelect(0).Value Then           'advt option
        imAdvt = True
        mObtainCodes 5, tgAdvertiser(), ilIncludeCodes, ilUseCodes()
    ElseIf RptSel!rbcSelCSelect(1).Value Then           'agency
        imAgency = True
        mObtainCodes 1, tgAgency(), ilIncludeCodes, ilUseCodes()
    ElseIf RptSel!rbcSelCSelect(2).Value Then           'slsp
        mObtainCodes 2, tgSalesperson(), ilIncludeCodes, ilUseCodes()
        imSlsp = True
    End If

    'retrieve only billing adjustment (AN) entries
    ilRet = gObtainPhfOrRvf(RptSel, slStartStd, slEndStd, tlTranType, tlRvf(), 3, 0) 'obtain billing adjustments from RVF only
    If ilRet = 0 Then
        Exit Sub
    End If


    'transactions types other than billing adjustment entries have been filtered out thru the call to RVF/PHF
    For llLoopTran = LBound(tlRvf) To UBound(tlRvf) - 1
        tmRvf = tlRvf(llLoopTran)

        'transaction date was filter out from general rvf read routine
        gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slStr     'must be within trans date & date entered selectivity
        llDate = gDateValue(slStr)                  'convert trans date to test if within requested limits
        ilRet = mFilterLists(ilIncludeCodes, ilUseCodes())
        If ilRet Then           'include the entry from list box selectivity
            LSet tmRvr = tmRvf
            '5-1-08 History records for installment rvftype = "A" should not be shown
            If ((llContrNo = 0) Or (llContrNo > 0 And tmRvf.lCntrNo = llContrNo)) And (Trim$(tmRvf.sType) = "" Or tmRvf.sType = "I") Then
                'get the contract code and place in RVR so that the buyer and billing type can be
                'extracted and printed in header
                tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'get matching contr recd

                Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")
                     ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                mFakeChf        'default some values in the chf buffer if no contract exists

                tmRvr.lCntrNo = tmChf.lCode    'Replace actual contr # with contract code to read header


                'Setup comment pointers only if show = yes for each applicable comment
                'Replace tmrvf.ldistamt with the "Other" comments pointer
                tmRvr.lDistAmt = 0                    'assume no "comment
                tmCxfSrchKey.lCode = tmChf.lCxfCode      'comment  code
                imCxfRecLen = Len(tmCxf)
                ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching comment recd
                If ilRet = BTRV_ERR_NONE Then
                                              'order/hold
                    If tmCxf.sShInv = "Y" Then        'show it on Invoice
                        tmRvr.lDistAmt = tmCxf.lCode
                    End If

                End If
                gPDNToLong tmRvf.sNet, llNet
                tmRvr.lAcquisitionCost = tgSpf.lBCxfDisclaimer      '9-30-11  ACQUISITION field is overlapped with the Invoice Disclaimer code for this report
                                                                    'if Acquisition field is required, find new field to place disclaimer code into, or add new field to record and change Invsummary.rpt
                If llNet <> 0 Then         'dont write out zero records
                    ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
                    On Error GoTo gGenInvSummaryErr
                    gBtrvErrorMsg ilRet, "btrInsert, gGenInvSummary ", RptSel
                    On Error GoTo 0
                End If
            End If
        End If
    Next llLoopTran

    Erase ilUseCodes, tlRvf
    ilRet = btrClose(hmRvr)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmCxf)
    btrDestroy hmRvr
    btrDestroy hmCHF
    btrDestroy hmCxf
    Exit Sub

gGenInvSummaryErr:
    Exit Sub
End Sub
'           mFindMatchingItem
'           Check to see if the airing or billing vehicle has been selected
'       <input> ilIncludesCodes = true if include codes, false if exclude codes
'               ilUseCodes() array of codes to include or exclude
'       <return> True if valid selection
Public Function mFindMatchingItem(ilIncludeCodes As Integer, ilUseCodes() As Integer)
Dim ilTemp As Integer
    mFindMatchingItem = False
    If ilIncludeCodes Then
        For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
            If imAirVeh Then
                If ilUseCodes(ilTemp) = tmRvf.iAirVefCode Then
                    mFindMatchingItem = True
                    Exit For
                End If
            ElseIf imBillVeh Then
                If ilUseCodes(ilTemp) = tmRvf.iBillVefCode Then
                    mFindMatchingItem = True
                    Exit For
                End If
            ElseIf imNTR Then
                If ilUseCodes(ilTemp) = tmRvf.iMnfItem Then
                    mFindMatchingItem = True
                    Exit For
                End If
            ElseIf imAdvt Then
                If ilUseCodes(ilTemp) = tmRvf.iAdfCode Then
                    mFindMatchingItem = True
                    Exit For
                End If
            End If
        Next ilTemp
    Else
        mFindMatchingItem = True        ' when more than half selected, selection fixed
        For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
            If imAirVeh Then
                If ilUseCodes(ilTemp) = tmRvf.iAirVefCode Then
                    mFindMatchingItem = False
                    Exit For
                End If
            ElseIf imBillVeh Then
                If ilUseCodes(ilTemp) = tmRvf.iBillVefCode Then
                    mFindMatchingItem = False
                    Exit For
                End If
            ElseIf imNTR Then
                If ilUseCodes(ilTemp) = tmRvf.iMnfItem Then
                    mFindMatchingItem = False
                    Exit For
                End If
            ElseIf imAdvt Then
                If ilUseCodes(ilTemp) = tmRvf.iAdfCode Then
                    mFindMatchingItem = False
                    Exit For
                End If
            End If
        Next ilTemp
    End If
End Function
'
'
'           Report to balance an installment contract
'           Active dates are entered by user.  Any installment contract spanning
'           any portion of those dates will be printed in its entirety.
'           For example:  Enter 10/1 - 10/30/07.  The order goes from 1/1/ - 10/5.
'                       All months jan - oct will be printed.
'           Each contract will show the month, its ordered $ (from contract lines/ntr),
'           the installment $ (from SBF, $ entered on installment screen), billing $ (rvftype = I)
'           and revenue $ (rvftype = "" and rvftype = "A")
'           Along with the billing and earned transactions, all invoice adjustments
'           will also be shown.  Bottom lines totals for all 4 monthly should balance
'
Public Sub gGenInstallmentReconcile()
Dim ilRet As Integer

'user entered active start/end date to retreive.  Variables for gObtainCntrForDate
Dim slStartDate As String
Dim llStartDate As Long
Dim slEndDate As String
Dim llEndDate As Long
Dim slCntrTypes As String                   'contr types to retrieve
Dim slCntrStatus As String                  'contr statuses to retrieve
Dim ilHOState As Integer                    'Hold/order contract states to retrieve
'
'
Dim slStr As String
Dim llSingleCntr As Long                    'contr code of user entered selective contr
Dim ilCurrentRecd As Integer                'loop index for contract to process from tlChfAdvtExt array
Dim llContrCode As Long                     'contract header internal code
Dim slChfStartDate As String                'start date of order
Dim slChfEndDate As String                  'end date of order
Dim llChfStartDate As Long                  'start date of order
Dim llChfEndDate As Long                    'end date of order
Dim llChfStartQtr As Long                   'start qtr (date) of order
Dim llChfEndQtr As Long                     'end qtr (date) of order
Dim ilTotalMonths As Integer                'totalmonths of order
'Dim llStdStartDates(1 To 37) As Long        'max 3 years of std dates
ReDim llStdStartDates(0 To 37) As Long        'max 3 years of std dates. Index zero ignored
'Dim ilMonthIndex(1 To 37) As Integer        'month index (1-12) for each of the std start dates in llStdStartDate array
ReDim ilMonthIndex(0 To 37) As Integer        'month index (1-12) for each of the std start dates in llStdStartDate array. Index zero ignored
Dim ilShowStdQtr As Integer                 'always show the months in std months: changed to 0 = std, 1 = cal, 2 = corp
Dim ilCorpStdYear As Integer                'std year that contract belongs in
Dim ilStartMonth As Integer                 'Start month of corp or std year that this contracts starts in (this report only  uses std)
Dim ilCurrStartQtr(0 To 1) As Integer
Dim tlSBFType As SBFTypes                   'Types to filter form SBF
Dim tlTranType As TRANTYPES                'transaction types to filter from RVF/PHF
'ReDim tlInstallInfo(1 To 1) As INSTALLINFO  'structure to maintain monthly $ info for the installment cntr
ReDim tlInstallInfo(0 To 1) As INSTALLINFO  'structure to maintain monthly $ info for the installment cntr. Index zero ignored
'ReDim tlInstallInfoAN(1 To 1) As INSTALLINFO   'strcture to maintain monthly $ info for ANs
ReDim tlInstallInfoAN(0 To 1) As INSTALLINFO   'strcture to maintain monthly $ info for ANs. Index zero ignored
Dim tlInstallDiscrep As INSTALLDISCREP      'this array accumulate totals IN & AN in the same monthly buckets to determine if there is a billing discrepancy
                                            'if the billed $ doesnt match the installment $ after invoicing, show it on Discrep Only version
Dim ilLoopOnRecd As Integer
Dim ilLoopOnDates As Integer
Dim llDate As Long
Dim slYear As String
Dim slMonth As String
Dim slDay As String
Dim llAmt As Long
'Dim llTempProject(1 To 36) As Long             'projection $ for sch lines, max 3 year
ReDim llTempProject(0 To 36) As Long             'projection $ for sch lines, max 3 year. Index zero ignored
Dim ilLoopOnIN As Integer
Dim ilLoopForDiscrep As Integer
Dim ilInclude As Integer
Dim ilDiscrepOnly As Integer
Dim ilIndex As Integer
Dim blFirstNTRFound As Boolean
Dim blGrossOption As Boolean
Dim slAmount As String
Dim ilAgyCommPct As Integer
Dim slAgyCommPct As String


        hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmCHF)
            btrDestroy hmCHF
            Exit Sub
        End If
        imCHFRecLen = Len(tmChf)

        hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            btrDestroy hmClf
            btrDestroy hmCHF
            Exit Sub
        End If
        imClfRecLen = Len(tmClf)

        hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            btrDestroy hmCff
            btrDestroy hmClf
            btrDestroy hmCHF
            Exit Sub
        End If
        imCffRecLen = Len(tmCff)

        hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmSbf)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            btrDestroy hmSbf
            btrDestroy hmCff
            btrDestroy hmClf
            btrDestroy hmCHF
            Exit Sub
        End If
        imSbfRecLen = Len(tmSbf)

        hmGrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmSbf)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            btrDestroy hmGrf
            btrDestroy hmSbf
            btrDestroy hmCff
            btrDestroy hmClf
            btrDestroy hmCHF
            Exit Sub
        End If
        imGrfRecLen = Len(tmGrf)

        slStr = RptSel!edcCheck.Text
        llSingleCntr = Val(slStr)      'selective contract #

'        slStartDate = RptSel!edcSelCFrom.Text
        slStartDate = RptSel!CSI_CalFrom.Text           '8-15-19 use csical control vs edit text control

        llStartDate = gDateValue(slStartDate)
        slStartDate = Format(llStartDate, "m/d/yy")   'make sure string start date has a year appended in case not entered with input

'        slEndDate = RptSel!edcSelCTo.Text
        slEndDate = RptSel!CSI_CalTo.Text               '8-15-19 use csical control vs edit text control
        llEndDate = gDateValue(slEndDate)
        slEndDate = Format(llEndDate, "m/d/yy")    'make sure string end date has a year appended in case not entered with input

        ilDiscrepOnly = True                '7-10-08 test user input for Discrepancies only.  Show only those months
                                            'that dont balance between installment $ and billed $
        If RptSel!ckcSelC5(0).Value = vbUnchecked Then
            ilDiscrepOnly = False
        End If
        
        blGrossOption = True
        If RptSel!rbcSelC4(1).Value = True Then         'net selected
            blGrossOption = False
        End If

'        ilShowStdQtr = True                         'always show the months of output in std month (vs corp)
        ilShowStdQtr = 0
        tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
        tmGrf.iGenDate(1) = igNowDate(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrf.lGenTime = lgNowTime

        'setup the SBF Types to include:  gather NTRs and the installments
        tlSBFType.iNTR = True
        tlSBFType.iInstallment = True
        tlSBFType.iImport = False

        'setup the transaction types to include: gather from PHF & RVF
        tlTranType.iAdj = True              'look only for adjustments in the History & Rec files
        tlTranType.iInv = True
        tlTranType.iWriteOff = False
        tlTranType.iPymt = False
        tlTranType.iCash = True
        tlTranType.iTrade = True
        tlTranType.iMerch = False               'always exclude Merchandise & promotions
        tlTranType.iPromo = False
        tlTranType.iNTR = True

        slCntrTypes = gBuildCntTypes()      'Setup valid types of contracts to obtain based on us
        slCntrStatus = ""
        slCntrStatus = "HOGN"             'sch/unsch holds & uns holds
        ilHOState = 2                      'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)
        sgCntrForDateStamp = ""            'init the time stamp to read in contracts upon re-entry
        'Build array of possible contracts that fall into last year or this years quarter and build into array tmChfAdvtExt
        If llSingleCntr > 0 Then
            ReDim tmChfAdvtExt(0 To 1) As CHFADVTEXT
            tmChfSrchKey1.lCntrNo = llSingleCntr
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                tmChfAdvtExt(0).lCode = tmChf.lCode
            End If
        Else
            ilRet = gObtainCntrForDate(RptSel, slStartDate, slEndDate, slCntrStatus, slCntrTypes, ilHOState, tmChfAdvtExt())
        End If

        For ilCurrentRecd = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1
            'only process contracts that are installments
            llContrCode = tmChfAdvtExt(ilCurrentRecd).lCode
            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChf, tgClf(), tgCff())

            If tgChf.sInstallDefined = "Y" Then
                'ReDim tlInstallInfo(1 To 37) As INSTALLINFO      'contract/install info by month, max 3 years
                ReDim tlInstallInfo(0 To 37) As INSTALLINFO      'contract/install info by month, max 3 years. Index zero ignored
                'ReDim tlInstallInfoAN(1 To 37) As INSTALLINFO
                ReDim tlInstallInfoAN(0 To 37) As INSTALLINFO       'Index zero ignored

                For ilLoopOnDates = 1 To 36                      'init the tran type field
                    tlInstallInfo(ilLoopOnDates).sTranType = ""
                    tlInstallInfoAN(ilLoopOnDates).sTranType = ""
                Next ilLoopOnDates
                For ilLoopOnDates = 1 To 36
                    tlInstallDiscrep.sTranType(ilLoopOnDates) = ""
                    tlInstallDiscrep.lBilling(ilLoopOnDates) = 0
                    tlInstallDiscrep.lInstallment(ilLoopOnDates) = 0
                    tlInstallDiscrep.lStartDate(ilLoopOnDates) = 0
                Next ilLoopOnDates
                
                '10-14-16 retrieve the agency to see if agency commissionable; if NTR exists, commission based on first NTR comm flag found.  Contract function has disallowed mixture of NTR and Airtime having some commision and some not
                'If direct, all must be non-commissionable.
                'If Agency, whatever the first NTR commission flag found will determine if commissionable or not.
                'if agency and no NTR, its commissionable.
                
                ilAgyCommPct = 0      'direct, no comm
                If tgChf.iAgfCode > 0 Then
                    ilIndex = gBinarySearchAgf(tgChf.iAgfCode)
                    If ilIndex >= 0 Then
                         ilAgyCommPct = tgCommAgf(ilIndex).iCommPct
                     End If
                End If
                
                'determine # of std months
                'gather contract $ from schedule lines & NTR
                gChfDatesToLong tgChf.iStartDate(), tgChf.iEndDate(), llChfStartDate, llChfEndDate, slChfStartDate, slChfEndDate  'convert contract header start & end dates to long
                'need the standard bdcst end date because billing will have a date at the end of the month and wont
                'include the last months billing/revenue
                slChfEndDate = gObtainEndStd(slChfEndDate)
                llChfEndDate = gDateValue(slChfEndDate)

                'Set up earliest/latest dates of contr, set to std dates.  Set array of starting bdcst months for summary page
                gFindMaxDates slChfStartDate, slChfEndDate, llChfStartQtr, llChfEndQtr, ilCurrStartQtr(), ilTotalMonths, llStdStartDates(), ilShowStdQtr, ilCorpStdYear, ilStartMonth, False   '1-6-21
                'setup the month indices for each of the std start dates
                For ilLoopOnRecd = 1 To 36
                    slStr = Format$(llStdStartDates(ilLoopOnRecd + 1) - 1, "m/d/yy")
                    gObtainYearMonthDayStr slStr, False, slYear, slMonth, slDay
                    ilMonthIndex(ilLoopOnRecd) = Val(slMonth)
                Next ilLoopOnRecd

                'gather installment $ (sbftrantype = "F")
                blFirstNTRFound = False                 '10-14-16 look for first NTR found, determine commision based on that flag (follow same rules as Invoicing function)
                                                        'contract function has been changed to restrict making all ntr and air time consistent , no mixtures of comm vs no comm
                                                        
                ReDim tmSbfList(0 To 0) As SBF
                ilRet = gObtainSBF(RptSel, hmSbf, tgChf.lCode, slChfStartDate, slChfEndDate, tlSBFType, tmSbfList(), 0)
                For ilLoopOnRecd = LBound(tmSbfList) To UBound(tmSbfList) - 1
                    tmSbf = tmSbfList(ilLoopOnRecd)
                    gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate        'end date of prior reconciling period
                    For ilLoopOnDates = 1 To 37
                        If llDate >= llStdStartDates(ilLoopOnDates) And llDate < llStdStartDates(ilLoopOnDates + 1) Then
                            If tmSbf.sTranType = "F" Then               'installment defined $
                                'accumulate the $ for this month
                                tlInstallInfo(ilLoopOnDates).lInstallment = tlInstallInfo(ilLoopOnDates).lInstallment + tmSbf.lGross
                                tlInstallInfo(ilLoopOnDates).iMonthIndex = ilMonthIndex(ilLoopOnDates)
                                tlInstallInfo(ilLoopOnDates).lStartDate = llStdStartDates(ilLoopOnDates)
                                tlInstallDiscrep.sTranType(ilLoopOnDates) = ""
                                tlInstallDiscrep.lInstallment(ilLoopOnDates) = tlInstallDiscrep.lInstallment(ilLoopOnDates) + tmSbf.lGross
                                tlInstallDiscrep.lStartDate(ilLoopOnDates) = llStdStartDates(ilLoopOnDates)
                                Exit For
                            ElseIf tmSbf.sTranType = "I" Then       'NTR
                                If Not blFirstNTRFound Then
                                    blFirstNTRFound = True              'found one NTR, use its commission used flag
                                    If tmSbf.sAgyComm = "N" Or blGrossOption Then        'no agy comm or running gross (dont take out comm)
                                        ilAgyCommPct = 0
                                    End If
                                End If
                                tlInstallInfo(ilLoopOnDates).lOrdered = tlInstallInfo(ilLoopOnDates).lOrdered + tmSbf.lGross * tmSbf.iNoItems       'Rate * # items
                                tlInstallInfo(ilLoopOnDates).iMonthIndex = ilMonthIndex(ilLoopOnDates)
                                tlInstallInfo(ilLoopOnDates).lStartDate = llStdStartDates(ilLoopOnDates)
                                tlInstallDiscrep.sTranType(ilLoopOnDates) = ""
                                tlInstallDiscrep.lStartDate(ilLoopOnDates) = llStdStartDates(ilLoopOnDates)
                                Exit For
                            End If
                        End If
                    Next ilLoopOnDates
                Next ilLoopOnRecd
                
                If (Not blFirstNTRFound) And (blGrossOption) Then         'no NTR found to change the commission structure and user requested gross; take no commission
                    ilAgyCommPct = 0
                End If

                'gather billing $, (rvftyp/phftype = "I"),
                'gather revenue (earned $ from receivables/history, rvftype/phftype = "A" or blank)
                ReDim tmRvfList(0 To 0) As RVF
                ilRet = gObtainPhfRvfbyCntr(RptSel, tgChf.lCntrNo, slChfStartDate, slChfEndDate, tlTranType, tmRvfList())
                If ilRet <> 0 Then          'error in rvf/phf
                    For ilLoopOnRecd = LBound(tmRvfList) To UBound(tmRvfList) - 1
                        tmRvf = tmRvfList(ilLoopOnRecd)
                        gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llDate        'end date of prior reconciling period
                        For ilLoopOnDates = 1 To 36
                            If llDate >= llStdStartDates(ilLoopOnDates) And llDate < llStdStartDates(ilLoopOnDates + 1) Then
                                'accumulate the $ for this month
                                If blGrossOption Then
                                    gPDNToLong tmRvf.sGross, llAmt
                                Else
                                    gPDNToLong tmRvf.sNet, llAmt
                                End If
                                If tmRvf.sType = "I" Then           'billing
                                    If tmRvf.sTranType = "AN" Then
                                        tlInstallInfoAN(ilLoopOnDates).lBilling = tlInstallInfoAN(ilLoopOnDates).lBilling + llAmt
                                        tlInstallInfoAN(ilLoopOnDates).sTranType = tmRvf.sTranType
                                        tlInstallInfoAN(ilLoopOnDates).iMonthIndex = ilMonthIndex(ilLoopOnDates)
                                        tlInstallInfoAN(ilLoopOnDates).lStartDate = llStdStartDates(ilLoopOnDates)

                                        tlInstallDiscrep.lBilling(ilLoopOnDates) = tlInstallDiscrep.lBilling(ilLoopOnDates) + llAmt
                                        tlInstallDiscrep.sTranType(ilLoopOnDates) = tmRvf.sTranType
                                        tlInstallDiscrep.lStartDate(ilLoopOnDates) = llStdStartDates(ilLoopOnDates)

                                    Else
                                        tlInstallInfo(ilLoopOnDates).lBilling = tlInstallInfo(ilLoopOnDates).lBilling + llAmt
                                        tlInstallInfo(ilLoopOnDates).sTranType = tmRvf.sTranType
                                        tlInstallInfo(ilLoopOnDates).iMonthIndex = ilMonthIndex(ilLoopOnDates)
                                        tlInstallInfo(ilLoopOnDates).lStartDate = llStdStartDates(ilLoopOnDates)

                                        tlInstallDiscrep.lBilling(ilLoopOnDates) = tlInstallDiscrep.lBilling(ilLoopOnDates) + llAmt
                                        tlInstallDiscrep.sTranType(ilLoopOnDates) = tmRvf.sTranType
                                        tlInstallDiscrep.lStartDate(ilLoopOnDates) = llStdStartDates(ilLoopOnDates)
                                   End If
                                    Exit For
                                ElseIf (Trim$(tmRvf.sType) = "" Or tmRvf.sType = "A") Then     'revenue
                                    If (Asc(tgSpf.sUsingFeatures6) And INSTALLMENTREVENUEEARNED) = INSTALLMENTREVENUEEARNED Then        'separate revenue and billing
                                        If tmRvf.sTranType = "AN" Then
                                            tlInstallInfoAN(ilLoopOnDates).lRevenue = tlInstallInfoAN(ilLoopOnDates).lRevenue + llAmt
                                            tlInstallInfoAN(ilLoopOnDates).sTranType = tmRvf.sTranType
                                            tlInstallInfoAN(ilLoopOnDates).iMonthIndex = ilMonthIndex(ilLoopOnDates)
                                            tlInstallInfoAN(ilLoopOnDates).lStartDate = llStdStartDates(ilLoopOnDates)

                                            tlInstallDiscrep.lBilling(ilLoopOnDates) = tlInstallDiscrep.lBilling(ilLoopOnDates) + llAmt
                                            tlInstallDiscrep.sTranType(ilLoopOnDates) = tmRvf.sTranType
                                            tlInstallDiscrep.lStartDate(ilLoopOnDates) = llStdStartDates(ilLoopOnDates)
                                        Else
                                            tlInstallInfo(ilLoopOnDates).lRevenue = tlInstallInfo(ilLoopOnDates).lRevenue + llAmt
                                            tlInstallInfo(ilLoopOnDates).sTranType = tmRvf.sTranType
                                            tlInstallInfo(ilLoopOnDates).iMonthIndex = ilMonthIndex(ilLoopOnDates)
                                            tlInstallInfo(ilLoopOnDates).lStartDate = llStdStartDates(ilLoopOnDates)

                                            tlInstallDiscrep.lBilling(ilLoopOnDates) = tlInstallDiscrep.lBilling(ilLoopOnDates) + llAmt
                                            tlInstallDiscrep.sTranType(ilLoopOnDates) = tmRvf.sTranType
                                            tlInstallDiscrep.lStartDate(ilLoopOnDates) = llStdStartDates(ilLoopOnDates)
                                        End If
                                        If tmRvf.lCefCode > 0 Then
                                            If tlInstallInfoAN(ilLoopOnDates).lComment > 0 Then
                                                tlInstallInfoAN(ilLoopOnDates).sCommentFlag = "*"
                                            Else
                                                tlInstallInfoAN(ilLoopOnDates).lComment = tmRvf.lCefCode
                                            End If
                                        End If
                                        Exit For
                                    Else                    'inv is revenue
                                        If tmRvf.sTranType = "AN" Then
                                            tlInstallInfoAN(ilLoopOnDates).lRevenue = tlInstallInfoAN(ilLoopOnDates).lRevenue + llAmt
                                            tlInstallInfoAN(ilLoopOnDates).lBilling = tlInstallInfoAN(ilLoopOnDates).lBilling + llAmt   '2-28-08 wrong field accumulated
                                            tlInstallInfoAN(ilLoopOnDates).sTranType = tmRvf.sTranType
                                            tlInstallInfoAN(ilLoopOnDates).iMonthIndex = ilMonthIndex(ilLoopOnDates)
                                            tlInstallInfoAN(ilLoopOnDates).lStartDate = llStdStartDates(ilLoopOnDates)

                                            tlInstallDiscrep.lBilling(ilLoopOnDates) = tlInstallDiscrep.lBilling(ilLoopOnDates) + llAmt
                                            tlInstallDiscrep.sTranType(ilLoopOnDates) = tmRvf.sTranType
                                            tlInstallDiscrep.lStartDate(ilLoopOnDates) = llStdStartDates(ilLoopOnDates)
                                        Else
                                            tlInstallInfo(ilLoopOnDates).lRevenue = tlInstallInfo(ilLoopOnDates).lRevenue + llAmt
                                            tlInstallInfo(ilLoopOnDates).lBilling = tlInstallInfo(ilLoopOnDates).lBilling + llAmt       '2-28-08 wrong field accumulated
                                            tlInstallInfo(ilLoopOnDates).sTranType = tmRvf.sTranType
                                            tlInstallInfo(ilLoopOnDates).iMonthIndex = ilMonthIndex(ilLoopOnDates)
                                            tlInstallInfo(ilLoopOnDates).lStartDate = llStdStartDates(ilLoopOnDates)

                                            tlInstallDiscrep.lBilling(ilLoopOnDates) = tlInstallDiscrep.lBilling(ilLoopOnDates) + llAmt
                                            tlInstallDiscrep.sTranType(ilLoopOnDates) = tmRvf.sTranType
                                            tlInstallDiscrep.lStartDate(ilLoopOnDates) = llStdStartDates(ilLoopOnDates)
                                        End If
                                        If tmRvf.lCefCode > 0 Then
                                            If tlInstallInfoAN(ilLoopOnDates).lComment > 0 Then
                                                tlInstallInfoAN(ilLoopOnDates).sCommentFlag = "*"
                                            Else
                                                tlInstallInfoAN(ilLoopOnDates).lComment = tmRvf.lCefCode
                                            End If
                                        End If
                                        Exit For
                                    End If
                                End If
                            End If
                        Next ilLoopOnDates
                    Next ilLoopOnRecd
                Else
                    MsgBox "Error retrieving rvf/phf for " & str$(tgChf.lCntrNo)
                End If
                'gather ordered (from schedule lines , ntr has already been gathered)
                For ilLoopOnRecd = LBound(tgClf) To UBound(tgClf) - 1
                    tmClf = tgClf(ilLoopOnRecd).ClfRec
                    If tmClf.sType = "H" Or tmClf.sType = "S" Then   'use hidden line (not pkg lines), and standard lines
                        gBuildFlights ilLoopOnRecd, llStdStartDates(), 1, 36, llTempProject(), 1, tgClf(), tgCff()
                    End If
                Next ilLoopOnRecd
                
                '10-14-16 calculate net if applicable
                If ilAgyCommPct > 0 Then
                    For ilLoopOnDates = 1 To 36
                        slAmount = gLongToStrDec(llTempProject(ilLoopOnDates), 2)
                        slAgyCommPct = gIntToStrDec(ilAgyCommPct, 4)
                        slStr = gMulStr(slAgyCommPct, slAmount)                       ' gross portion of possible split
                        slStr = gRoundStr(slStr, ".01", 2)
                        llTempProject(ilLoopOnDates) = llTempProject(ilLoopOnDates) - gStrDecToLong(slStr, 2)  'adjusted net
                                          
                        slAmount = gLongToStrDec(tlInstallInfo(ilLoopOnDates).lOrdered, 2)
                        slAgyCommPct = gIntToStrDec(ilAgyCommPct, 4)
                        slStr = gMulStr(slAgyCommPct, slAmount)                       ' gross portion of possible split
                        slStr = gRoundStr(slStr, ".01", 2)
                        tlInstallInfo(ilLoopOnDates).lOrdered = tlInstallInfo(ilLoopOnDates).lOrdered - gStrDecToLong(slStr, 2) 'adjusted net

                        slAmount = gLongToStrDec(tlInstallInfo(ilLoopOnDates).lInstallment, 2)
                        slAgyCommPct = gIntToStrDec(ilAgyCommPct, 4)
                        slStr = gMulStr(slAgyCommPct, slAmount)                       ' gross portion of possible split
                        slStr = gRoundStr(slStr, ".01", 2)
                        tlInstallInfo(ilLoopOnDates).lInstallment = tlInstallInfo(ilLoopOnDates).lInstallment - gStrDecToLong(slStr, 2) 'adjusted net

                        slAmount = gLongToStrDec(tlInstallDiscrep.lInstallment(ilLoopOnDates), 2)
                        slAgyCommPct = gIntToStrDec(ilAgyCommPct, 4)
                        slStr = gMulStr(slAgyCommPct, slAmount)                       ' gross portion of possible split
                        slStr = gRoundStr(slStr, ".01", 2)
                        tlInstallDiscrep.lInstallment(ilLoopOnDates) = tlInstallDiscrep.lInstallment(ilLoopOnDates) - gStrDecToLong(slStr, 2) 'adjusted net
                    Next ilLoopOnDates
                End If
                
                'take all the $ accumulated for the entire contract and put prepare it for output
                For ilLoopOnDates = 1 To 36
                    tlInstallInfo(ilLoopOnDates).lOrdered = tlInstallInfo(ilLoopOnDates).lOrdered + llTempProject(ilLoopOnDates)
                    tlInstallInfoAN(ilLoopOnDates).lOrdered = tlInstallInfoAN(ilLoopOnDates).lOrdered + llTempProject(ilLoopOnDates)
                    llTempProject(ilLoopOnDates) = 0
                Next ilLoopOnDates

                'done with the contract, write prepass record
                'For each InstallInfo entry, find the matching month in the InstallDiscre array.  Compare the billing $ against the installment $.
                'If matching, dont show for Discrep only version.  Month must be invoiced to show on Discrep only version.
                For ilLoopOnDates = 1 To 36
                    If tlInstallInfo(ilLoopOnDates).lOrdered + tlInstallInfo(ilLoopOnDates).lInstallment + tlInstallInfo(ilLoopOnDates).lBilling + tlInstallInfo(ilLoopOnDates).lRevenue <> 0 Then
                         For ilLoopForDiscrep = 1 To 36      'loop thru the Discrepancy table that has the installment and billing totals to see if they match.  If matching and discrep only,
                                                        'dont include the entry
                            If tlInstallInfo(ilLoopOnDates).lStartDate = tlInstallDiscrep.lStartDate(ilLoopForDiscrep) Then     'got the matching month
                                'compare the billed info against the predetermined installment amount; the trans type must be non blank (representing AN or IN) indicating it has been invoiced
                                ilInclude = False
                                If Not ilDiscrepOnly Then       'matching months, not discrep only so show everything
                                    ilInclude = True
                                Else
                                    'if discrepancy only and there isnt any billing, then it cant be a discrepancy to print so dont include it
                                    If (ilDiscrepOnly And Trim$(tlInstallDiscrep.sTranType(ilLoopForDiscrep)) = "") Or (tlInstallDiscrep.lBilling(ilLoopForDiscrep) = tlInstallDiscrep.lInstallment(ilLoopForDiscrep) And Trim$(tlInstallDiscrep.sTranType(ilLoopForDiscrep)) <> "" And ilDiscrepOnly = True) Then
                                        ilInclude = False                'not billed or they match, not a discrep
                                        Exit For
                                    Else
                                        ilInclude = True
                                    End If
                                End If
                                If ilInclude Then
                                    tmGrf.lChfCode = tgChf.lCode            'contract internal code
                                    'gPackDateLong tlInstallInfo(ilLoopOnDates).lStdMonth, tmGrf.iStartDate(0), tmGrf.iStartDate(1)        'start date of std month for sorting
                                    gPackDateLong llStdStartDates(ilLoopOnDates), tmGrf.iStartDate(0), tmGrf.iStartDate(1)        'start date of std month for sorting
                                    'tmGrf.lDollars(1) = tlInstallInfo(ilLoopOnDates).lOrdered
                                    'tmGrf.lDollars(2) = tlInstallInfo(ilLoopOnDates).lInstallment
                                    'tmGrf.lDollars(3) = tlInstallInfo(ilLoopOnDates).lBilling
                                    'tmGrf.lDollars(4) = tlInstallInfo(ilLoopOnDates).lRevenue
                                    tmGrf.lDollars(0) = tlInstallInfo(ilLoopOnDates).lOrdered
                                    tmGrf.lDollars(1) = tlInstallInfo(ilLoopOnDates).lInstallment
                                    tmGrf.lDollars(2) = tlInstallInfo(ilLoopOnDates).lBilling
                                    tmGrf.lDollars(3) = tlInstallInfo(ilLoopOnDates).lRevenue
                                    tmGrf.sGenDesc = tlInstallInfo(ilLoopOnDates).sTranType         'tran type (IN, HI, AN) from billing or revenue
                                    tmGrf.iYear = ilMonthIndex(ilLoopOnDates)          'month index (1-12) to show Jan-Dec
                                    tmGrf.lCode4 = 0            'no comments on IN transactions
                                    tmGrf.sBktType = ""        'cant have multiple comments, init flag
                                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                    Exit For
                                End If
                            End If
                        Next ilLoopForDiscrep
                    End If
                Next ilLoopOnDates

                For ilLoopOnDates = 1 To 36
                    If tlInstallInfoAN(ilLoopOnDates).lInstallment + tlInstallInfoAN(ilLoopOnDates).lBilling + tlInstallInfoAN(ilLoopOnDates).lRevenue <> 0 Then
                        For ilLoopForDiscrep = 1 To 36      'loop thru the table that has the installment and billing totals to see if they match.  If match and discrep only,
                                                            'dont include the entry
                            If tlInstallInfoAN(ilLoopOnDates).lStartDate = tlInstallDiscrep.lStartDate(ilLoopForDiscrep) Then     'got the matching month
                                'compare the billed info against the predetermined installment amount; the trans type must be non blank (representing AN or IN) indicating it has been invoiced
                                ilInclude = False

                                If Not ilDiscrepOnly Then       'matching months,not discrep only so show everything
                                    ilInclude = True
                                Else
                                    If (ilDiscrepOnly And Trim$(tlInstallDiscrep.sTranType(ilLoopForDiscrep)) = "") Or (tlInstallDiscrep.lBilling(ilLoopForDiscrep) = tlInstallDiscrep.lInstallment(ilLoopForDiscrep) And Trim$(tlInstallDiscrep.sTranType(ilLoopForDiscrep)) <> "" And ilDiscrepOnly = True) Then
                                        ilInclude = False                'not billed or they match, not a discrep
                                        Exit For
                                    Else
                                        ilInclude = True
                                    End If
                                End If

                                If ilInclude Then

                                    tmGrf.lChfCode = tgChf.lCode            'contract internal code
                                    'gPackDateLong tlInstallInfo(ilLoopOnDates).lStdMonth, tmGrf.iStartDate(0), tmGrf.iStartDate(1)        'start date of std month for sorting
                                    gPackDateLong llStdStartDates(ilLoopOnDates), tmGrf.iStartDate(0), tmGrf.iStartDate(1)        'start date of std month for sorting
                                    'tmGrf.lDollars(1) = tlInstallInfoAN(ilLoopOnDates).lOrdered
                                    'tmGrf.lDollars(2) = tlInstallInfoAN(ilLoopOnDates).lInstallment
                                    tmGrf.lDollars(0) = tlInstallInfoAN(ilLoopOnDates).lOrdered
                                    tmGrf.lDollars(1) = tlInstallInfoAN(ilLoopOnDates).lInstallment
                                    'if there already is an entry for the Month (IN or installment entry, not AN entry), then zero out the
                                    'ordered and installment $ so it wont be overstated
                                    For ilLoopOnIN = 1 To 36
                                        If ilLoopOnIN = ilLoopOnDates Then
                                            'tmGrf.lDollars(1) = 0
                                            'tmGrf.lDollars(2) = 0
                                            tmGrf.lDollars(0) = 0
                                            tmGrf.lDollars(1) = 0
                                            Exit For
                                        End If
                                    Next ilLoopOnIN
                                    'tmGrf.lDollars(3) = tlInstallInfoAN(ilLoopOnDates).lBilling
                                    'tmGrf.lDollars(4) = tlInstallInfoAN(ilLoopOnDates).lRevenue
                                    tmGrf.lDollars(2) = tlInstallInfoAN(ilLoopOnDates).lBilling
                                    tmGrf.lDollars(3) = tlInstallInfoAN(ilLoopOnDates).lRevenue
                                    tmGrf.sGenDesc = tlInstallInfoAN(ilLoopOnDates).sTranType         'tran type (IN, HI, AN) from billing or revenue
                                    tmGrf.iYear = ilMonthIndex(ilLoopOnDates)          'month index (1-12) to show Jan-Dec
                                    tmGrf.lCode4 = tlInstallInfoAN(ilLoopOnDates).lComment      'AN comment if one exists
                                    tmGrf.sBktType = tlInstallInfoAN(ilLoopOnDates).sCommentFlag     'flag to indicate mutliple comments if they exists
                                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                    Exit For
                                End If
                            End If
                        Next ilLoopForDiscrep
                    End If
                Next ilLoopOnDates
            End If                          'tgChf.sInstallDefined = "Y"
        Next ilCurrentRecd

        ilRet = btrClose(hmSbf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmGrf)
        btrDestroy hmSbf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        btrDestroy hmGrf

        Erase tlInstallInfo, tlInstallInfoAN, tmChfAdvtExt, tmSbfList, tmRvfList
    Exit Sub
End Sub
'
'       Create user options in prepass due to Encyypted fields
'
Public Sub gGenUserOptions()
   Dim ilRet As Integer
    hmUor = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUor, "", sgDBPath & "Uor.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmUor)
        btrDestroy hmUor
        Exit Sub
    End If
    imUorRecLen = Len(tmUor)
    
    hmUrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmUrf)
        btrDestroy hmUrf
        ilRet = btrClose(hmUor)
        btrDestroy hmUor
        Exit Sub
    End If
    imUrfRecLen = Len(tmUrf)
    
    tmUor.iGenDate(0) = igNowDate(0)
    tmUor.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmUor.lGenTime = lgNowTime

    ilRet = btrGetFirst(hmUrf, tmUrf, imUrfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)    'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE     '11-8-05 check for end of file to avoid loopin
        If tmUrf.iCode <> 1 And ((tmUrf.sDelete = "Y" And RptSel!ckcSelC7.Value = vbChecked) Or (tmUrf.sDelete = "N")) Then          'ignore CSI or any deleted users if not requested
            gUrfDecrypt tmUrf
            tmUor.tUor = tmUrf
            ilRet = btrInsert(hmUor, tmUor, imUorRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
        End If
        ilRet = btrGetNext(hmUrf, tmUrf, imUrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
    ilRet = btrClose(hmUrf)
    btrDestroy hmUrf
    ilRet = btrClose(hmUor)
    btrDestroy hmUor
    Exit Sub
End Sub
Public Sub gGenUserSummary()
    Dim ilRet As Integer
    Dim blInclude As Boolean
    Dim slDormant As String
    
    hmUor = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUor, "", sgDBPath & "Uor.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmUor)
        btrDestroy hmUor
        Exit Sub
    End If
    imUorRecLen = Len(tmUor)
    
    hmUrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmUrf)
        btrDestroy hmUrf
        ilRet = btrClose(hmUor)
        btrDestroy hmUor
        Exit Sub
    End If
    imUrfRecLen = Len(tmUrf)
    
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        btrDestroy hmSlf
        ilRet = btrClose(hmUrf)
        btrDestroy hmUrf
        ilRet = btrClose(hmUor)
        btrDestroy hmUor
        Exit Sub
    End If
    
    hmUst = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUst, "", sgDBPath & "ust.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmUst)
        btrDestroy hmUst
        ilRet = btrClose(hmSlf)
        btrDestroy hmSlf
        ilRet = btrClose(hmUrf)
        btrDestroy hmUrf
        ilRet = btrClose(hmUor)
        btrDestroy hmUor
        Exit Sub
    End If
    
    tmUor.iGenDate(0) = igNowDate(0)
    tmUor.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmUor.lGenTime = lgNowTime

    'Traffic
    ilRet = btrGetFirst(hmUrf, tmUrf, imUrfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)    'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE                  'check for end of file to avoid looping
        blInclude = False
        If tmUrf.iCode <> 1 Then    'ignore CSI
            If (tmUrf.sDelete = "N" And RptSel!rbcSelC6(0).Value = True) Then        'include Active only
                blInclude = True
            ElseIf (tmUrf.sDelete = "Y" And RptSel!rbcSelC6(1).Value = True) Then    'include Dormant only
                blInclude = True
            ElseIf RptSel!rbcSelC6(2).Value = True Then                              'include Both
                blInclude = True
            End If
        End If
        If blInclude Then
            gUrfDecrypt tmUrf
            tmUor.tUor = tmUrf
            tmUor.tUor.sWin(0) = "T"

            'grab Salesperson code info
            'ilRet = gBinarySearchSlf(tmUrf.iCode)
            'If ilRet > 0 Then
            '    tmUor.tUor.iSlfCode = tgMSlf(ilRet).iCode
            'End If
            
            ilRet = btrInsert(hmUor, tmUor, imUorRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
        End If
        ilRet = btrGetNext(hmUrf, tmUrf, imUrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
    'Affiliate
    tmUstUOR.iGenDate(0) = igNowDate(0)
    tmUstUOR.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmUstUOR.lGenTime = lgNowTime
    
    'clear last tmUrf record
    tmUrf.iCode = 0
    tmUrf.sRept = ""
    tmUrf.sCity = ""
    tmUrf.sName = ""
    tmUrf.sPhoneNo = ""
    tmUrf.iSlfCode = 0
    tmUrf.iSnfCode = 0
    tmUrf.sName = ""
    tmUrf.lEMailCefCode = 0
    tmUrf.sDelete = ""
    
    ilRet = btrGetFirst(hmUst, tmUst, imUrfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)    'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE                  'check for end of file to avoid looping
        blInclude = False
        'If tmUst.iCode <> 1 Then    'ignore CSI --> For the Affiliate User, it doesn't contain Counterpoint as record 1
        If (tmUst.iState = 0 And RptSel!rbcSelC6(0).Value = True) Then          'include Active only
            blInclude = True
        ElseIf (tmUst.iState = 1 And RptSel!rbcSelC6(1).Value = True) Then      'include Dormant only
            blInclude = True
        ElseIf RptSel!rbcSelC6(2).Value = True Then                             'include Both
            blInclude = True
        End If
        'set active/dormant flag
        slDormant = IIF(tmUst.iState = 0, "N", "Y")
        'End If
        If blInclude Then
            'gUrfDecrypt tmUrf
            tmUstUOR.tUor = tmUrf
            tmUstUOR.tUor.sWin(0) = "A"
            tmUstUOR.tUor.iCode = tmUst.iCode
            tmUstUOR.tUor.sRept = tmUst.sReportName             'Name on report
            tmUstUOR.tUor.sCity = tmUst.sCity
            tmUstUOR.tUor.sName = tmUst.sName                   'Sign on name
            tmUstUOR.tUor.sPassword = tmUst.sPassword
            tmUstUOR.tUor.sPhoneNo = tmUst.sPhoneNo
            tmUstUOR.tUor.iSnfCode = tmUst.iSnfCode
            tmUstUOR.tUor.lEMailCefCode = tmUst.lEMailCefCode   'Email address
            tmUstUOR.tUor.iSlfCode = 0                          'No sales people for Affiliate
            tmUstUOR.tUor.sDelete = slDormant
            
            ilRet = btrInsert(hmUor, tmUstUOR, imUorRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
        End If
        ilRet = btrGetNext(hmUst, tmUst, imUrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
    ilRet = btrClose(hmUst)
    btrDestroy hmUst
    ilRet = btrClose(hmSlf)
    btrDestroy hmSlf
    ilRet = btrClose(hmUrf)
    btrDestroy hmUrf
    ilRet = btrClose(hmUor)
    btrDestroy hmUor
    
    Exit Sub
End Sub

'
'           gCreateUserActivity - Gather Traffic & Affiliate Users Activity
'           by date span for all or selective active users
'           Activity is retrieved from UAF
'
Public Sub gGenUserActivity()
Dim ilError As Integer
Dim slStr As String
Dim llEarliestDate As Long
Dim llLatestDate As Long
Dim slStartDate As String
Dim slEndDate As String
Dim ilRet As Integer
Dim llLoopOnDate As Long
Dim llLoopOnUaf As Integer
Dim slLoopDate As String
Dim llUafStartTime As Long
Dim llUafEndTime As Long
Dim llInputStartTime As Long
Dim llInputEndTime As Long
Dim slSTime As String
Dim slETime As String
Dim slSDate As String
Dim slEDate As String
Dim ilTime(0 To 1) As Integer
ReDim ilUseCodes(0 To 0) As Integer
Dim ilInclCodes As Integer              'include or exclude the selected user codes
Dim ilUserCode As Integer
Dim ilUserOk As Integer
Dim ilTimeOk As Integer
Dim llUafStartDate As Long
Dim llUafEndDAte As Long
Dim llMid As Long
Dim llMinuteDiff As Long
Dim ilNoDays As Integer
Dim ilNoHours As Integer
Dim ilNoMinutes As Integer
Dim llSecDiff As Long
Dim ilABort As Integer
Dim llDescDate As Long
Dim llDescTime As Long

ReDim tlUAF(0 To 0) As UAF

        On Error GoTo gGenUserActivityErr
        hmUaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmUaf, "", sgDBPath & "Uaf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilError = ilRet
        End If
        imUafRecLen = Len(tmUaf)
        On Error GoTo gGenUserActivityErr
        gBtrvErrorMsg ilRet, "gGenUserActivity (btrOpen):" & "Uaf.Btr", RptSel
        On Error GoTo 0
        
        hmAfr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmAfr, "", sgDBPath & "Afr.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilError = ilRet
        End If
        imAfrRecLen = Len(tmAfr)
        On Error GoTo gGenUserActivityErr
        gBtrvErrorMsg ilRet, "gGenUserActivity (btrOpen):" & "Afr.mkd", RptSel
        On Error GoTo 0
    
        hmUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilError = ilRet
        End If
        imUrfRecLen = Len(tmUrf)
        On Error GoTo gGenUserActivityErr
        gBtrvErrorMsg ilRet, "gGenUserActivity (btrOpen):" & "Urf.btr", RptSel
        On Error GoTo 0
        
        hmUst = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmUst, "", sgDBPath & "Ust.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilError = ilRet
        End If
        imUstRecLen = Len(tmUst)
        On Error GoTo gGenUserActivityErr
        gBtrvErrorMsg ilRet, "gGenUserActivity (btrOpen):" & "Ust.mkd", RptSel
        On Error GoTo 0
        
        If ilError <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmAfr)
            btrDestroy hmAfr
            ilRet = btrClose(hmUaf)
            btrDestroy hmUaf
            ilRet = btrClose(hmUrf)
            btrDestroy hmUrf
            ilRet = btrClose(hmUst)
            btrDestroy hmUst
        End If
        
        slStr = RptSel!edcSelCFrom.Text                'Earliest date to retrieve from PRF or RVF
        llEarliestDate = gDateValue(slStr)
        slStr = RptSel!edcSelCFrom1.Text               'Latest date to retrieve from PRF or RVF
        llLatestDate = gDateValue(slStr)
        slStartDate = Format$(llEarliestDate, "m/d/yy")
        slEndDate = Format$(llLatestDate, "m/d/yy")
        
        slStr = RptSel!edcSelCTo.Text
        gPackTime slStr, ilTime(0), ilTime(1)
        gUnpackTimeLong ilTime(0), ilTime(1), False, llInputStartTime
        slStr = RptSel!edcSelCTo1.Text
        gPackTime slStr, ilTime(0), ilTime(1)
        gUnpackTimeLong ilTime(0), ilTime(1), True, llInputEndTime
        
        gObtainUrfUstCodes ilInclCodes, ilUseCodes(), 10, RptSel
        
        'filtering of the prepass temporary records
        tmAfr.lGenTime = lgNowTime
        tmAfr.iGenDate(0) = igNowDate(0)
        tmAfr.iGenDate(1) = igNowDate(1)
        llMid = 86400           '12M end of day
        
        'Loop on each date to gather all the activity for selected users
        'loop daily as too many records may be generated
        For llLoopOnDate = llEarliestDate To llLatestDate
            slLoopDate = Format$(llLoopOnDate, "m/d/yy")
            ReDim tlUAF(0 To 0) As UAF
            ilRet = gObtainUafByDate(RptSel, hmUaf, tlUAF(), slLoopDate)
            For llLoopOnUaf = LBound(tlUAF) To UBound(tlUAF) - 1
                tmUaf = tlUAF(llLoopOnUaf)

                gUnpackTimeLong tmUaf.iStartTime(0), tmUaf.iStartTime(1), False, llUafStartTime
                gUnpackTimeLong tmUaf.iEndTime(0), tmUaf.iEndTime(1), True, llUafEndTime
                slSTime = gFormatTimeLong(llUafStartTime, "A", "1")
                slETime = gFormatTimeLong(llUafEndTime, "A", "1")
                llDescTime = 99999 - llUafStartTime
                
                gUnpackDateLong tmUaf.iStartDate(0), tmUaf.iStartDate(1), llUafStartDate
                gUnpackDateLong tmUaf.iEndDate(0), tmUaf.iEndDate(1), llUafEndDAte
                slSDate = Format$(llUafStartDate, "m/d/yy")
                slEDate = Format$(llUafEndDAte, "m/d/yy")
                llDescDate = 99999 - llUafStartDate     'need to sort date descending (most current date first)
                
                ilTimeOk = True                 'all tasks that were incomplete will show
                ilABort = False                 'set in case task never completed
                
                If llUafEndDAte = 62093 Then           'task never ended, highest date allowed is stored in field, always show aborted tasks
                    ilABort = True
                    'no end date, test only start time against time input parameters
                    If llUafStartTime < llInputStartTime Or llUafStartTime > llInputEndTime Then
                        ilTimeOk = False
                    End If
                Else
                    If (llUafStartDate = llUafEndDAte) Then             'activity started & ended same date or never ended (test highest date allowed (12/31/2069))
                        If (llUafEndTime < llInputStartTime Or llUafStartTime > llInputEndTime) Then
                            ilTimeOk = False
                        End If
                    Else                        'task ended on a different day
                        'test date it started to see if activity falls within requested times
                        If llUafStartTime >= llInputStartTime And llUafStartTime <= llMid Then
                            ilTimeOk = True
                        Else
                            'first days time doesnt fall with the user requested times, test the end time
                            If llUafEndTime >= llInputStartTime And llUafEndTime <= llInputEndTime Then
                                ilTimeOk = True
                            Else
                                ilTimeOk = False
                            End If
                        End If
                        
                    End If
                End If
                
                tmAfr.lCrfCsfcode = 0       'set duration to 0 incase the task was never completed
                If (ilTimeOk) And (Not ilABort) Then            'determine the duration, as long as the task was completed

                    llSecDiff = DateDiff("s", slSDate & " " & slSTime, slEDate & " " & slETime)
'                    ilNoDays = llSecDiff \ 86400
'                    llMinuteDiff = llSecDiff - (CLng(ilNoDays) * 86400)
'                    ilNoHours = llMinuteDiff \ 3600
'                    llMinuteDiff = llSecDiff - ((CLng(ilNoDays) * 86400) + (ilNoHours * 3600))
'                    ilNoMinutes = llMinuteDiff \ 60
'                    llMinuteDiff = llSecDiff - ((CLng(ilNoDays) * 86400) + (ilNoHours * 3600) + (ilNoMinutes * 60))

                    tmAfr.lCrfCsfcode = llSecDiff
                End If
                ilUserCode = tmUaf.iUserCode     'assume traffic user code
                If tmUaf.sSystemType = "A" Then     'affiliate event, the code needs to be negated so it knows its a different table
                    ilUserCode = -tmUaf.iUserCode
                End If
                'the filter agyadvcodes rtn handles negated codes
                'determine if user selected
                ilUserOk = gFilterAgyAdvCodes(ilUserCode, ilInclCodes, ilUseCodes())
                If (ilUserOk) And (ilTimeOk) Then
                    'get the user record (traffic) to decrypt
                    'get the users Name On Report field if defined;otherwise use the signon name
                    tmAfr.sISCI = ""
                    If tmUaf.sSystemType = "A" Then
                        tmUstSrchKey.iCode = tmUaf.iUserCode
                        ilRet = btrGetEqual(hmUst, tmUst, imUstRecLen, tmUstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            If Trim$(tmUst.sReportName) = "" Then
                                tmAfr.sISCI = Trim$(tmUst.sName)
                            Else
                                tmAfr.sISCI = Trim$(tmUst.sReportName)
                            End If
                        End If
                    Else
                        tmUrfSrchKey.iCode = tmUaf.iUserCode
                        ilRet = btrGetEqual(hmUrf, tmUrf, imUrfRecLen, tmUrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            If Trim$(tmUrf.sRept) = "" Then
                                tmAfr.sISCI = Trim$(gDecryptField(tmUrf.sName))
                            Else
                                tmAfr.sISCI = Trim$(gDecryptField(tmUrf.sRept))
                            End If
                        End If
                    End If
                    
                    tmAfr.lAttCode = llDescDate
                    tmAfr.lAstCode = tmUaf.lCode
                    
                    'make time descending
                    slStr = Val(llDescTime)
                    Do While Len(slStr) < 6
                        slStr = "0" & slStr
                    Loop
                    tmAfr.sCart = Trim$(slStr)
                    'afrastcode = UAFcode
                    'afrattcode = start date (descending)
                    'afrCrfCsfCode - duration of event
                    'afrgendate - generation date (filter)
                    'afrgentime - generation time (filter)
                    'afrISCI - users name on report or user signon
                    'afrcart - start time (desc as a string)
                    tmAfr.sCallLetters = ""
                    ilRet = btrInsert(hmAfr, tmAfr, imAfrRecLen, INDEXKEY0)

                End If
            Next llLoopOnUaf
        Next llLoopOnDate
    
        Erase tlUAF
        
        ilRet = btrClose(hmUaf)
        btrDestroy hmUaf
        ilRet = btrClose(hmAfr)
        btrDestroy hmAfr
        ilRet = btrClose(hmUrf)
        btrDestroy hmUrf
        ilRet = btrClose(hmUst)
        btrDestroy hmUst
        Exit Sub

gGenUserActivityErr:
        ilRet = 1
        Resume Next
              
        Exit Sub
End Sub
'
'       check for valid Sales Source selection
'
'       <input> ilSlfCode - slf internal code
'               tlSofList - array of sales offices
'       <output>
'       Return true if valid sales source
Public Function mTestValidSalesSource(ilSlfCode As Integer, tlSofList() As SOFLIST, ilMnfSSCodes() As Integer) As Integer
Dim ilSlfInx As Integer
Dim ilMatchSSCode As Integer
Dim ilTemp As Integer
Dim ilOKSS As Integer

            ilOKSS = False
            ilSlfInx = gBinarySearchSlf(ilSlfCode)
            If ilSlfInx = -1 Then
                ilMatchSSCode = 0
            Else
                tmSlf = tgMSlf(ilSlfInx)
                ilMatchSSCode = mFindMatchSSCode(tmSlf.iSofCode, tlSofList())   'get the sales source for this sof trans.
            End If

            For ilTemp = 0 To UBound(ilMnfSSCodes) - 1
                If ilMnfSSCodes(ilTemp) = ilMatchSSCode Then
                    ilOKSS = True
                    Exit For
                End If
            Next ilTemp
            mTestValidSalesSource = ilOKSS
            
            Exit Function
End Function

Public Sub gGenUnpostedBarterStations()
Dim ilRet As Integer
Dim hlTxr As Integer
Dim tlTxr As TXR
Dim ilTxrReclen As Integer
Dim slFile As String
Dim slSDate As String
Dim slEDate As String
Dim llSDate As Long
Dim llEDate As Long
Dim slUnposted As String
Dim llChfCode As Long
Dim slCntrStatus As String
Dim slCntrType As String
Dim ilHOState As Integer
Dim ilCurrentRecd As Integer
Dim ilEnoughRoom As Integer
Dim ilRemChar As Integer
Dim slText As String
Dim slStr As String * 200
Dim ilSaveMonth As Integer
Dim ilUnPostedCount As Integer
Dim blTestHiddenLines As Boolean

        On Error GoTo gGenUnpostedBarterStationsErr:
        slFile = "Chf"
        hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmCHF)
            btrDestroy hmCHF
            Exit Sub
        End If
        imCHFRecLen = Len(tmChf)

        slFile = "Clf"
        hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            btrDestroy hmClf
            btrDestroy hmCHF
            Exit Sub
        End If
        imClfRecLen = Len(tmClf)

        slFile = "Cff"
        hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            btrDestroy hmCff
            btrDestroy hmClf
            btrDestroy hmCHF
            Exit Sub
        End If
        imCffRecLen = Len(tmCff)
        
        slFile = "Txr"
        hlTxr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hlTxr, "", sgDBPath & "Txr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hlTxr)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            btrDestroy hlTxr
            btrDestroy hmCff
            btrDestroy hmClf
            btrDestroy hmCHF
            Exit Sub
        End If
        ilTxrReclen = Len(tlTxr)
        
        slFile = "IIHF"
        hmIihf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmIihf, "", sgDBPath & "Iihf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmIihf)
            ilRet = btrClose(hlTxr)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            btrDestroy hmIihf
            btrDestroy hlTxr
            btrDestroy hmCff
            btrDestroy hmClf
            btrDestroy hmCHF
            Exit Sub
        End If
        imIihfRecLen = Len(tmIihf)
        
        'get the next standard bdcst billing dates
'        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slSDate
'        slSDate = gIncOneDay(slSDate)
'        slEDate = gObtainEndStd(slSDate)
'        llSDate = gDateValue(slSDate)
'        llEDate = gDateValue(slEDate)

        blTestHiddenLines = True
        If RptSel!ckcSelC5(0).Value = vbUnchecked Then
            blTestHiddenLines = False
        End If

        slStr = Trim$(RptSel!edcSelCFrom.Text)             'month in text form (jan..dec)
        gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
        If ilSaveMonth = 0 Then                           'input isn't text month name, try month #
            ilSaveMonth = Val(slStr)
        End If
        
        slStr = Trim$(str$(ilSaveMonth)) & "/15/" & Trim$(str$(RptSel!edcSelCTo.Text))      'format xx/xx/xxxx
        slSDate = gObtainStartStd(slStr)
        llSDate = gDateValue(slSDate)
        slStr = gObtainEndStd(slStr)
        llEDate = gDateValue(slStr)         '+ 1                      'increment for next month (why was this done to increment to next month?-remove+1)
        slEDate = Format$(llEDate, "m/d/yy")

        
        slCntrStatus = "HO"                 'statuses: sch hold + order
        slCntrType = "CVTRQ"         'all types: PI, DR, etc.  except PSA(p) and Promo(m)
        ilHOState = 1                       'get latest orders & revisions
        
        ilRet = gObtainCntrForDate(RptSel, slSDate, slEDate, slCntrStatus, slCntrType, ilHOState, tmChfAdvtExt())
        
        tlTxr.iGenDate(0) = igNowDate(0)        'gen date & time for crystal filter
        tlTxr.iGenDate(1) = igNowDate(1)
        tlTxr.lGenTime = lgNowTime
        
        For ilCurrentRecd = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
            llChfCode = tmChfAdvtExt(ilCurrentRecd).lCode

            slUnposted = mTestUnpostedBarterStations(llChfCode, llSDate, llEDate, ilUnPostedCount, blTestHiddenLines)
            If Trim$(slUnposted) <> "" Then     'create as many records as necessary to string out the vehicle names
                ilRemChar = Len(slUnposted)
                tlTxr.lSeqNo = 1
                tlTxr.lCsfCode = tmChfAdvtExt(ilCurrentRecd).lCntrNo
                tlTxr.iType = tmChfAdvtExt(ilCurrentRecd).iAdfCode
                tlTxr.iGeneric1 = ilUnPostedCount
                Do While ilRemChar > 0
                    
                    slText = Mid$(slUnposted, ((tlTxr.lSeqNo - 1) * 200 + 1), 200)
                    tlTxr.sText = Trim$(slText)
                    ilRet = btrInsert(hlTxr, tlTxr, Len(tlTxr), INDEXKEY0)
                    tlTxr.iGeneric1 = 0
                    ilRemChar = ilRemChar - 200
                    slText = ""
                    tlTxr.lSeqNo = tlTxr.lSeqNo + 1
                Loop
            End If
            
        Next ilCurrentRecd
        
        ilRet = btrClose(hlTxr)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        btrDestroy hlTxr
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        
        Erase tmChfAdvtExt
        Exit Sub
gGenUnpostedBarterStationsErr:
        On Error GoTo 0
        gBtrvErrorMsg ilRet, "gGenUnpostedBarterStations (btrOpen):" & Trim$(slFile), RptSel
        
        Exit Sub
 
End Sub
'
'       mTestUnpostedBarterStations - determine what stations within a period has not yet been posted
'       <input> llchfcode - internal contract code
'               llStartDate - start date of period to search
'               llEndDate - end date of period to search
'       <output> ilUnpostedCount - total stations unposted per contract
'
    Private Function mTestUnpostedBarterStations(llChfCode As Long, llStartDate As Long, llEndDate As Long, ilUnPostedCount As Integer, blTestHiddenLines As Boolean) As String
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim ilAdf As Integer
    Dim ilVff As Integer
    Dim ilVef As Integer
    Dim ilRet As Integer
    Dim llCffStartDate As Long
    Dim llCffEndDate As Long
    Dim slNotPosted As String
    Dim blOk As Boolean
    Dim llLineStartDate As Long
    Dim llLineEndDate As Long
    Dim llSpots As Long
    Dim tlCff As CFF
    Dim illoop As Integer
    Dim llDate As Long
    Dim llDate2 As Long
    Dim ilTemp As Integer
    
    slNotPosted = ""
    ilUnPostedCount = 0
    
    ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llChfCode, False, tgChfInv, tgClfInv(), tgCffInv())
    If ilRet Then
'        If tgChfInv.sBillCycle = "C" Then
'            llStartDate = lmStartCal
'            llEndDate = lmEndCal
'        ElseIf tgChfInv.sBillCycle = "W" Then
'            llStartDate = lmStartWk
'            llEndDate = lmEndWk
'        Else
'            llStartDate = lmStartStd
'            llEndDate = lmEndStd
'        End If

        For ilClf = LBound(tgClfInv) To UBound(tgClfInv) - 1 Step 1
            tmClf = tgClfInv(ilClf).ClfRec
            llSpots = 0
            gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llLineStartDate
            gUnpackDateLong tmClf.iEndDate(0), tmClf.iEndDate(1), llLineEndDate
            
            'If (tmClf.sType <> "O") And (tmClf.sType <> "A") And (tmClf.sType <> "E") Then
            'always exclude package lines
            If (tmClf.sType = "S") Or ((blTestHiddenLines) And (tmClf.sType = "H")) And (llLineEndDate >= llLineStartDate) Then        'always include standard lines; only test hidden lines if user requested
                ilVff = gBinarySearchVff(tmClf.iVefCode)
                If (ilVff <> -1) Then
                    If (tgVff(ilVff).sPostLogSource = "S") Then     'importing station invoices
               
                        'Dates and if Posted
                        ilCff = tgClfInv(ilClf).iFirstCff
                        Do While ilCff <> -1
                            tlCff = tgCffInv(ilCff).CffRec                       'use a temporary buffer for flight record
                            llCffStartDate = tgCffInv(ilCff).lStartDate
                            llCffEndDate = tgCffInv(ilCff).lEndDate
                            If (llCffEndDate >= llStartDate) And (llCffStartDate <= llEndDate) And (llCffStartDate <= llCffEndDate) Then   'flight must be with user entered dates and also not a cancel before start
                                'Test if IIHF exist
                                tmIihfSrchKey2.lChfCode = llChfCode
                                tmIihfSrchKey2.iVefCode = tmClf.iVefCode
                                gPackDateLong llStartDate, tmIihfSrchKey2.iInvStartDate(0), tmIihfSrchKey2.iInvStartDate(1)
                                ilRet = btrGetEqual(hmIihf, tmIihf, imIihfRecLen, tmIihfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet <> BTRV_ERR_NONE Then
                                                                  
                                    'backup start date to Monday
                                    illoop = gWeekDayLong(llCffStartDate)
                                    Do While illoop <> 0
                                        llCffStartDate = llCffStartDate - 1
                                        illoop = gWeekDayLong(llCffStartDate)
                                    Loop
                                    'only retrieve for projections, anything in the past has already
                                    'been invoiced and has been retrieved from history or receiv files
                                    'adjust the gather dates from flights: use flight start date or requested start date, whichever is later
                                    If llStartDate > llCffStartDate Then
                                        llCffStartDate = llStartDate
                                    End If
                                    'use flight end date or requsted end date, whichever is lesser
                                    If llEndDate < llCffEndDate Then
                                        llCffEndDate = llEndDate
                                    End If
                            
                                    For llDate = llCffStartDate To llCffEndDate Step 7
                                        'Loop on the number of weeks in this flight
                                        'calc week into of this flight to accum the spot count
                                        If tlCff.sDyWk = "W" Then            'weekly
                                            llSpots = tlCff.iSpotsWk + tlCff.iXSpotsWk
                                        Else                                        'daily
                                            If illoop + 6 < llCffEndDate Then           'we have a whole week
                                                llSpots = tlCff.iDay(0) + tlCff.iDay(1) + tlCff.iDay(2) + tlCff.iDay(3) + tlCff.iDay(4) + tlCff.iDay(5) + tlCff.iDay(6)
                                            Else
                                                llCffEndDate = llDate + 6
                                                If llDate > llCffEndDate Then
                                                    llCffEndDate = llCffEndDate       'this flight isn't 7 days
                                                End If
                                                For llDate2 = llDate To llCffEndDate Step 1
                                                    ilTemp = gWeekDayLong(llDate2)
                                                    llSpots = llSpots + tlCff.iDay(ilTemp)
                                                Next llDate2
                                            End If
                                        End If
                                    Next llDate
                                 
                                    If llSpots > 0 Then
                                        ilVef = gBinarySearchVef(tmClf.iVefCode)
                                        If ilVef <> -1 Then
                                            If slNotPosted = "" Then
                                                slNotPosted = Trim$(tgMVef(ilVef).sName)
                                                ilUnPostedCount = ilUnPostedCount + 1
                                            Else
                                                If InStr(slNotPosted, Trim$(tgMVef(ilVef).sName)) = 0 Then
                                                    slNotPosted = slNotPosted & ", " & Trim$(tgMVef(ilVef).sName)
                                                    ilUnPostedCount = ilUnPostedCount + 1
                                                End If
                                            End If
                                        Else
                                            If slNotPosted = "" Then
                                                slNotPosted = "vefCode = " & tgMVef(ilVef).iCode
                                            Else
                                                slNotPosted = slNotPosted & ", vefCode= " & tgMVef(ilVef).iCode
                                            End If
                                        End If
                                    End If
                                End If
                                Exit Do
                            End If
                            ilCff = tgCffInv(ilCff).iNextCff
                        Loop
                    End If
                End If
            End If
        Next ilClf
    End If
    mTestUnpostedBarterStations = slNotPosted
End Function
Public Sub gGenStatementTrans()
'****************************************************************
'       obtain all RVF Cash transactions for statements
'       Payments and JE use Latest cash date to include
'       Invoice and adjustments use latest bill date to include
'
'****************************************************************
'
'
    Dim ilRet As Integer    'Return status
    Dim llTranDate As Long
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim illoop As Integer
    Dim slDate As String
    Dim llLatestCashDate As Long
    Dim llLatestBillDate As Long
    Dim ilIncludeCodes As Integer
    Dim ilUseCodes() As Integer
    Dim ilAgfCode As Integer
    Dim tlCharTypeBuff As POPCHARTYPE
    Dim ilOk As Integer
    
    
        hmRvr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmRvr, "", sgDBPath & "Rvr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
        
        End If
        imRvrRecLen = Len(tmRvr)
        
        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            ilRet = btrClose(hmRvr)
            btrDestroy hmRvr
           
        End If
        imRvfRecLen = Len(tmRvf)
        
        tmRvr.iGenDate(0) = igNowDate(0)        'todays date used for retrieval/removal of records
        tmRvr.iGenDate(1) = igNowDate(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmRvr.lGenTime = lgNowTime

        'ttp 8820 2-28-18 dates were missing in header of the statements
'        slDate = RptSel!edcSelCFrom.Text   'Latest cash date
        slDate = RptSel!CSI_CalFrom.Text   'Latest cash date        8-29-19 use csi calendar control vs edit box

        llLatestCashDate = gDateValue(slDate)
'        slDate = RptSel!edcSelCTo.Text   'Latest bill date
        slDate = RptSel!CSI_CalTo.Text   'Latest bill date
        llLatestBillDate = gDateValue(slDate)
        
       
        'ReDim ilUseCodes(1 To 1) As Integer
        ReDim ilUseCodes(0 To 0) As Integer
        mObtainAgyAdvCodes ilIncludeCodes, ilUseCodes(), 2    'assume using lbcselection(2) for list box of direct & agencies
          
        imRvfRecLen = Len(tmRvf)
        btrExtClear hmRvf   'Clear any previous extend operation
        ilExtLen = Len(tmRvf)  'Extract operation record size
        
        '7-1-14 use key3 (rvftrandate) instead of key 0 (rvfagfcode)
        ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
            Call btrExtSetBounds(hmRvf, llNoRec, -1, "UC", "RVF", "") '"EG") 'Set extract limits (all records)
            tlCharTypeBuff.sType = "C"
            ilOffSet = gFieldOffset("Rvf", "RvfCashTrade")      'cash only
            ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Sub
            End If

            ilRet = btrExtAddField(hmRvf, 0, ilExtLen)  'Extract the whole record
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainRVF (btrExtAddField):" & "RVF.Btr", RptSel
            On Error GoTo 0
            ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo mRvfErr
                gBtrvErrorMsg ilRet, "gObtainRVF (btrExtGetNextExt):" & "RVF.Btr", RptSel
                On Error GoTo 0
                ilExtLen = Len(tmRvf)  'Extract operation record size
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    'first test for valid trans types (Invoices, adjustments, write-off & payments
                    gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llTranDate
                    'ttp 8820 - wrong tran type first char tested for payments (should be P, was testing I)
                    If ((tmRvf.sTranType = "IN" Or tmRvf.sTranType = "AN") And (llTranDate <= llLatestBillDate)) Or ((Left$(tmRvf.sTranType, 1) = "P" Or Left$(tmRvf.sTranType, 1) = "W") And (llTranDate <= llLatestCashDate)) Then
                        'valid payee selected?
                        If tmRvf.iAgfCode = 0 Then
                            ilAgfCode = -tmRvf.iAdfCode
                        Else
                            ilAgfCode = tmRvf.iAgfCode
                        End If
                        
                        ilOk = gFilterAgyAdvCodes(ilAgfCode, ilIncludeCodes, ilUseCodes())
                        'Create RVR record
                        If ilOk Then
                            On Error GoTo mRvfErr
                            LSet tmRvr = tmRvf
                            tmRvr.lCefCode = tgSaf(0).lStatementComment 'points to CXF
                           
                            ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)
                            If ilRet <> BTRV_ERR_NONE Then
                                On Error GoTo mRvfErr
                                gBtrvErrorMsg ilRet, "gGenStatementTrans (btrInsert):" & "rvf.Btr", RptSel
                            End If
                            On Error GoTo 0
                        Else
                            ilRet = ilRet
                        End If
                        
                    End If
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
        
   
    Exit Sub
mRvfErr:
    On Error GoTo 0
    Exit Sub
End Sub
Private Function mObtainPhfOrRvf(RptForm As Form, slEarliestDate As String, slLatestDate As String, tlTranType As TRANTYPES, tlRvf() As RVF, ilWhichFile As Integer) As Integer
'****************************************************************
'*      This routine is a copy of gObtainPhfOrRvf but has been
'*      revised for Sales Commissions on Collections
'
'*      Obtain all History OR Receivables transactions whose
'*      transaction date falls within the earliest and latest
'*      dates requested.
'*      Also test PIs whose entry dates fall within the requeted period,
'*      and whosde tran date fall into any month prior to the report month.
'       That test is to get the PO that have been applied.
'*      The entire file must be read to find those transactions.
'*      <input>  RptForm - Form calling this populate rtn
'*               slEarliestDate - get all trans starting with
'*                   this date
'*               slLatestDAte - get all trans equal or prior to this date
'*              ilWhichFile - 1 = PHF, 2 = RVF, 3= both
'*
'*      <I/O>    tlRvf() - array of matching Phf/Rvf recds
'*               funtion return - true if receivables populated
'*                       false if no receivables, error
'*
'*
'*          5-26-04 Exclude NTR "AN" transactions when NTR to be excluded
'****************************************************************
'
'    ilRet = mObtainPhfOrRvf (RptForm,  slEarliestDate, slLatestDate, tlTranType, tlRvf(),ilWhichFile)
'
    Dim ilRet As Integer    'Return status
    ReDim ilEarliestDate(0 To 1) As Integer
    ReDim ilLatestDate(0 To 1) As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim illoop As Integer
    Dim ilRVFUpper As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilStartFile As Integer
    Dim ilEndFile As Integer
    Dim ilLowLimit As Integer
    Dim blValidDates As Boolean
    Dim llTransDate As Long
    Dim llEntryDate As Long
    Dim llEarliestDate As Long
    Dim llLatestDate As Long
    
    If PeekArray(tlRvf).Ptr <> 0 Then
        ilLowLimit = LBound(tlRvf)
    Else
        ilLowLimit = 0
    End If
    
    ReDim tlRvf(ilLowLimit To ilLowLimit) As RVF

    If ilWhichFile = 1 Then
        ilStartFile = 1
        ilEndFile = 1
        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            mObtainPhfOrRvf = False
            Exit Function
        End If
    ElseIf ilWhichFile = 2 Then
        ilStartFile = 2
        ilEndFile = 2
        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            mObtainPhfOrRvf = False
            Exit Function
        End If
    Else
        ilStartFile = 1
        ilEndFile = 2
        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            mObtainPhfOrRvf = False
            Exit Function
        End If
    End If


    imRvfRecLen = Len(tlRvf(ilLowLimit))
    gPackDate slEarliestDate, ilEarliestDate(0), ilEarliestDate(1)
    llEarliestDate = gDateValue(slEarliestDate)
    gPackDate slLatestDate, ilLatestDate(0), ilLatestDate(1)
    llLatestDate = gDateValue(slLatestDate)
    btrExtClear hmRvf   'Clear any previous extend operation
    ilExtLen = Len(tlRvf(ilLowLimit))  'Extract operation record size
    'ilRVFUpper = UBound(tlRvf)             1-11-18 removed, unused
    For illoop = ilStartFile To ilEndFile         'pass 1- get PHF, pass 2 get RVF
        '7-1-14 use key3 (rvftrandate) instead of key 0 (rvfagfcode)
        ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
            Call btrExtSetBounds(hmRvf, llNoRec, -1, "UC", "RVF", "") '"EG") 'Set extract limits (all records)

            ilRet = btrExtAddField(hmRvf, 0, ilExtLen)  'Extract the whole record
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainRVF (btrExtAddField):" & "RVF.Btr", RptForm
            On Error GoTo 0
            ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo mRvfErr
                gBtrvErrorMsg ilRet, "gObtainRVF (btrExtGetNextExt):" & "RVF.Btr", RptForm
                On Error GoTo 0
                ilExtLen = Len(tmRvf)  'Extract operation record size
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    blValidDates = False
                    'first test for valid trans types (Invoices, adjustments, write-off & payments
'                    If ((Left$(tmRvf.sTranType, 1) = "I" Or tmRvf.sTranType = "HI") And tlTranType.iInv) Or (Left$(tmRvf.sTranType, 1) = "A" And tlTranType.iAdj) Or (Left$(tmRvf.sTranType, 1) = "W" And tlTranType.iWriteOff) Or (Left$(tmRvf.sTranType, 1) = "P" And tlTranType.iPymt) Then
                    If (Left$(tmRvf.sTranType, 1) = "W" And tlTranType.iWriteOff) Or (Left$(tmRvf.sTranType, 1) = "P" And tlTranType.iPymt) Then
                        'test for valid date.  Transaction dates must fall within the requested dates.
                        'PIs Entry dates must fall within the requested dates, and the transaction dates must be in a prior date of the requested dates
                        blValidDates = True
                        If Left$(tmRvf.sTranType, 1) = "W" Or tmRvf.sTranType = "PI" Then
                            blValidDates = False
                            gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llTransDate
                            gUnpackDateLong tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), llEntryDate
                            If Left$(tmRvf.sTranType, 1) = "P" Then '  PI :test Entry date within requested dates, and trans date must be in prior date to requested period
                                If llTransDate >= llEarliestDate And llTransDate <= llLatestDate Then
                                    blValidDates = True
                                Else            'journal entries tran date must be between requested period
                                    If (llEntryDate >= llEarliestDate And llEntryDate <= llLatestDate) And (llTransDate < llEarliestDate) Then
                                        blValidDates = True
                                    End If
                                End If
                            Else            ' Journal Entries (W) transaction dates must fall within requested dates
                                If llTransDate >= llEarliestDate And llTransDate <= llLatestDate Then
                                    blValidDates = True
                                End If
                            End If
                        Else            'include all POs
                            If tmRvf.sTranType = "PO" Then
                                blValidDates = True
                            End If
                        End If
                        If blValidDates Then
                            If (tlTranType.iNTR) Then       'NTR option, tested separately because it shouldnt be tested with Cash transactions
                                If tmRvf.iMnfItem > 0 Then
                                    tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                    ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                                Else            'its not an NTR
                                    'got valid trans type - test for Cash, Trade, Merchandising or Promotions
                                    '05-26-04 dont include NTR, exclude this if it is
                                    If tmRvf.iMnfItem = 0 And tmRvf.lSbfCode = 0 Then       'not NTR
                                        If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                            tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                            ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                                        End If
                                    Else
                                        'must be an installment record, so it should be included
                                        'If trans is an NTR, it must have the item type and pointer to SBF; otherwise its assumed to be an installment
                                        'if it has an SBF pointer only
                                        If tmRvf.iMnfItem = 0 And tmRvf.lSbfCode > 0 Then
                                            If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                                tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                                ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                                            End If
                                        End If
    
                                    End If
                                End If
                            Else
                                'got valid trans type - test for Cash, Trade, Merchandising or Promotions
                                If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                    tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                    ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                                End If
                            End If
                        End If
                    End If
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
        If ilWhichFile = 3 And illoop = 1 Then                           'if 1, then just finished history, go do Receivables
            btrExtClear hmRvf   'Clear any previous extend operation
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmRvf)
                btrDestroy hmRvf
                mObtainPhfOrRvf = False
                Exit Function
            End If
            imRvfRecLen = Len(tmRvf)
            'ilRVFUpper = UBound(tlRvf)     1-11-18 removed, unused .  causing overflow error

        End If
    Next illoop
    ilRet = btrClose(hmRvf)
    btrDestroy hmRvf
    mObtainPhfOrRvf = True
    Exit Function
mObtainPhfOrRvfErr:
    ilRet = 1
    Resume Next
mRvfErr:
    On Error GoTo 0
    mObtainPhfOrRvf = False
    Exit Function
End Function
'
'           Genenate Report of Station Posting Activity
'           Selections by Invoice date; User logged in date, vehicle
'
Public Sub gGenStationPostingRpt()
    Dim ilRet As Integer
    Dim ilError  As Integer
    Dim slInvStartDate As String
    Dim slInvEndDate As String
    Dim slLogInStartDate As String
    Dim slLogInEndDate As String
    Dim rst_StationPostInfo As ADODB.Recordset
    Dim ilIncludeCodes As Integer
    Dim ilUseCodes() As Integer
    Dim illoop As Integer
    Dim slVefCodes As String
    Dim slSQLQuery As String
    Dim slNowDate As String
    Dim llDate As Long
    
            tmGrf.iGenDate(1) = igNowDate(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            gUnpackDateLong igNowDate(0), igNowDate(1), llDate
            slNowDate = Format$(llDate, "ddddd")

'            slInvStartDate = RptSel!edcSelCFrom.Text        'inv start
            slInvStartDate = RptSel!CSI_CalFrom.Text        'inv start
            slInvStartDate = Format(gDateValue(slInvStartDate), "ddddd")    'insure mm/dd/yy
'            slInvEndDate = RptSel!edcSelCFrom1.Text        'inv end
            slInvEndDate = RptSel!CSI_CalTo.Text        'inv end
            slInvEndDate = Format(gDateValue(slInvEndDate), "ddddd")    'insure mm/dd/yy
'            slLogInStartDate = RptSel!edcSelCTo.Text 'log in start date
            slLogInStartDate = RptSel!CSI_CalFrom2.Text 'log in start date
            slLogInStartDate = Format(gDateValue(slLogInStartDate), "ddddd")    'insure mm/dd/yy
'            slLogInEndDate = RptSel!edcSelCTo1.Text     'log in end date
            slLogInEndDate = RptSel!CSI_CalTo2.Text     'log in end date
            slLogInEndDate = Format(gDateValue(slLogInEndDate), "ddddd")    'insure mm/dd/yy
    
            'set up query to include or exclude selcted vehicle codes
            slVefCodes = ""
            gObtainCodesForMultipleLists 0, tgVehicle(), ilIncludeCodes, ilUseCodes(), RptSel
            If RptSel!ckcAll.Value = vbUnchecked Then           'selective vehicles, not All
                For illoop = LBound(ilUseCodes) To UBound(ilUseCodes) - 1
                    If Trim$(slVefCodes) = "" Then
                        slVefCodes = str$(ilUseCodes(illoop))
                    Else
                        slVefCodes = slVefCodes & "," & str$(ilUseCodes(illoop))
                    End If
                Next illoop
          
                If ilIncludeCodes Then      'include codes (vs exclude)
                    slVefCodes = " and lrfpostvefcode in (" & slVefCodes & ")"
                Else                        'exclude codes
                    slVefCodes = " and lrfpostvefcode Not in (" & slVefCodes & ")"
                End If
            End If
            '  & " And " & Format$(slLogInStartDate, sgSQLDateForm) & "  <= ldfPostEndDate and " & Format$(slLogInEndDate, sgSQLDateForm) & " >= ldfPostStartDate" & slVefCodes
            slSQLQuery = "Select * from LRF_Log_Remote_Stn where lrfInvStartDate >= '" & Format$(slInvStartDate, sgSQLDateForm) & "' and lrfInvStartDate <= '" & Format$(slInvEndDate, sgSQLDateForm) & "'"
            slSQLQuery = slSQLQuery & " and '" & Format$(slLogInStartDate, sgSQLDateForm) & "' <= lrfPostEndDate and '" & Format$(slLogInEndDate, sgSQLDateForm) & "' >= lrfPostStartDate" & slVefCodes
            Set rst_StationPostInfo = gSQLSelectCall(slSQLQuery)
            While Not rst_StationPostInfo.EOF
                
                slSQLQuery = "INSERT INTO " & "GRF_Generic_Report"
                slSQLQuery = slSQLQuery & "(grfVefCode, grfGenDesc, "
                slSQLQuery = slSQLQuery & "grfStartDate, grfDate1, grfDate2, "
                slSQLQuery = slSQLQuery & "grfTime, grfMissedTime, "
                slSQLQuery = slSQLQuery & " grfGendate, grfGenTime) "
                
                slSQLQuery = slSQLQuery & " Values("
                slSQLQuery = slSQLQuery & rst_StationPostInfo!lrfPostVefCode & ",'" & Trim$(rst_StationPostInfo!lrfUserName) & "', "
                slSQLQuery = slSQLQuery & "'" & Format$(rst_StationPostInfo!lrfInvStartDate, sgSQLDateForm) & "', '" & Format$(rst_StationPostInfo!lrfPostStartDate, sgSQLDateForm) & "', '" & Format$(rst_StationPostInfo!lrfPostEndDate, sgSQLDateForm) & "', "
                slSQLQuery = slSQLQuery & "'" & Format$(rst_StationPostInfo!lrfPostStartTime, sgSQLTimeForm) & "', '" & Format$(rst_StationPostInfo!lrfPostEndTime, sgSQLTimeForm) & "', "
                slSQLQuery = slSQLQuery & "'" & Format$(slNowDate, sgSQLDateForm) & "', " & lgNowTime & ")"
                
                If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                    gHandleError "TrafficErrors.txt", "RptcrAlloc: mUpdateGrf"
                End If
   
                rst_StationPostInfo.MoveNext
            Wend

    
            rst_StationPostInfo.Close
        Exit Sub
gGenStationPostingErr:
    ilRet = -1
    Resume Next
    
End Sub

Public Function mExportCashDist(tlRvr As RVR) As String
    mExportCashDist = ""
    Dim slCSVString As String
    Dim slString As String
    Dim ilInt As Integer
    Dim illoop As Integer
    Dim llLong As Long
    Dim llDate As Long
    Dim tmPrfSrchKey As LONGKEY0
    Dim slStamp As String
    ReDim tlMnf(0 To 0) As MNF
    slCSVString = ""
    On Error GoTo ExportError
    '--------------
    'Participant
    If RptSel.rbcSelCSelect(2).Value = True Then
        '"Participant,Air Vehicle,Bill Vehicle,Source,Agency,Advertiser,Product,Contract,Inv#,Invoice Date,Paid Amount,Distribution %,Distribution Amount"
        
        gPDNToStr tmRvr.sNet, 2, slString
        If Val(slString) = 0 Or tlRvr.imnfOwner = 0 Then
            'Dont export these bad records
            Exit Function
        End If
        
        '--------------
        'Participant
        ilInt = gObtainMnfForType("H", slStamp, tlMnf())         'H=Vehicle Group
        For illoop = LBound(tlMnf) To UBound(tlMnf) - 1
            If tlMnf(illoop).iCode = tlRvr.imnfOwner Then
                slCSVString = slCSVString & """" & Trim(tlMnf(illoop).sName) & """"
                Exit For
            End If
        Next illoop
        '--------------
        'Air Vehicle
        ilInt = gBinarySearchVef(tlRvr.iAirVefCode)
        If ilInt = -1 Then
            slCSVString = slCSVString & ","
        Else
            slCSVString = slCSVString & ",""" & Trim(tgMVef(ilInt).sName) & """"
        End If
        '--------------
        'Bill Vehicle
        ilInt = gBinarySearchVef(tlRvr.iBillVefCode)
        If ilInt = -1 Then
            slCSVString = slCSVString & ","
        Else
            slCSVString = slCSVString & ",""" & Trim(tgMVef(ilInt).sName) & """"
        End If
        '--------------
        'Source
        ilInt = gObtainMnfForType("S", slStamp, tlMnf())       'Source Types
        For illoop = LBound(tlMnf) To UBound(tlMnf) - 1
            If tlMnf(illoop).iCode = tlRvr.iMnfSSCode Then
                slCSVString = slCSVString & ",""" & Trim(tlMnf(illoop).sName) & """"
                Exit For
            End If
        Next illoop
        '--------------
        'Agency
        ilInt = gBinarySearchAgf(tlRvr.iAgfCode)
        If ilInt = -1 Then
            slCSVString = slCSVString & ","
        Else
            slCSVString = slCSVString & "," & """" & Trim(tgCommAgf(ilInt).sName) & """"
        End If
        '--------------
        'Advertiser
        ilInt = gBinarySearchAdf(tlRvr.iAdfCode)
        If ilInt = -1 Then
            slCSVString = slCSVString & ","
        Else
            slCSVString = slCSVString & ",""" & Trim(tgCommAdf(ilInt).sName) & """"
        End If
        '--------------
        'Product
        tmPrfSrchKey.lCode = tlRvr.lPrfCode
        ilInt = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        If ilInt = BTRV_ERR_NONE Then
            slCSVString = slCSVString & ",""" & Trim$(tmPrf.sName) & """"
        Else
            slCSVString = slCSVString & ","
        End If
        '--------------
        'Contract
        slCSVString = slCSVString & "," & tlRvr.lCntrNo
        '--------------
        'Inv#
        slCSVString = slCSVString & "," & tlRvr.lInvNo
        '--------------
        'Invoice Date
        gUnpackDateLong tlRvr.iInvDate(0), tlRvr.iInvDate(1), llDate
        slCSVString = slCSVString & "," & Format(llDate, "ddddd")
        '--------------
        'Paid Amount
        slCSVString = slCSVString & "," & Format(-tlRvr.lDistAmt / 100, "#0.00")
        '--------------
        'Dist %
        slCSVString = slCSVString & "," & Format(tlRvr.iProdPct / 100, "#0.00") & "%"
        '--------------
        'Amount
        gPDNToStr tmRvr.sNet, 2, slString
        slCSVString = slCSVString & "," & Format(-(Val(slString)), "#0.00")
        
    End If
    
    '--------------
    'Write File
    Print #hmExport, slCSVString
    'Show some status
    lmExportCount = lmExportCount + 1
    If lmExportCount / 10 - Int(lmExportCount / 10) = 0 Then
        'show every 10
        RptSel.lacExport.Caption = "Exported " & lmExportCount & " records..."
        RptSel.lacExport.Refresh
    End If
    mExportCashDist = "Exported " & lmExportCount & " records..."
    Exit Function
    
ExportError:
    mExportCashDist = "Error:" & err & "-" & Error(err)
End Function

'TTP 10164 - Ageing summary by month - export option for Audacy A/R dump
Function mSaveAgeSummary(tlRvr As RVR) As String
    mSaveAgeSummary = ""
    Dim illoop As Integer
    'Sum rvrNet by Year/Month / Agency / Contract#
    Dim blFound As Boolean
    Dim slString As String
    Dim slGenDate As String
    Dim slAgeDate As String
    Dim ilDaysBehind As Integer
    slString = ""
    blFound = False
    slGenDate = ""
    slAgeDate = ""
    
    lmExportCount = lmExportCount + 1
    If lmExportCount / 10 - Int(lmExportCount / 10) = 0 Then
        'show every 10
        RptSel.lacExport.Caption = "Processing " & lmExportCount & " records..."
        RptSel.lacExport.Refresh
    End If
    mSaveAgeSummary = "Processing " & lmExportCount & " records..."
    
    'Check tmAgingSummary for existing Year/Month/Agency/Contract#, Add  - if Not append tmAgingSummary
    For illoop = 0 To UBound(tmAgingSummary)
        If tmAgingSummary(illoop).iYear = tlRvr.iAgingYear Then
            If tmAgingSummary(illoop).iMonth = tlRvr.iAgePeriod Then
                If tmAgingSummary(illoop).lContractNumber = tlRvr.lCntrNo Then
                    blFound = True
                    '(Add) Net
                    gPDNToStr tlRvr.sNet, 2, slString
                    tmAgingSummary(illoop).dBalance = tmAgingSummary(illoop).dBalance + Val(slString)
        
                    '(Max of) Days Behind = TempDaysBehind := ({RVR_Receivables_Rept.rvrGenDate} - {@AgeingPeriodToDate}=Date ({RVR_Receivables_Rept.rvrAgeYear}, {RVR_Receivables_Rept.rvrAgePeriod}, {RVR_Receivables_Rept.rvrDistAmt}) );
                    'gUnpackDate tlRvr.iGenDate(0), tlRvr.iGenDate(1), slGenDate
                    'slAgeDate = tlRvr.iAgePeriod & "/" & tlRvr.lDistAmt & "/" & tlRvr.iAgingYear
                    'ilDaysBehind = DateDiff("d", slAgeDate, slGenDate)
                    'ilDaysBehind = IIF(ilDaysBehind < 0, 0, ilDaysBehind)
                    'tmAgingSummary(UBound(tmAgingSummary)).iDaysBehind = IIF(ilDaysBehind < 0, 0, ilDaysBehind)
                    'If ilDaysBehind > tmAgingSummary(UBound(tmAgingSummary)).iDaysBehind Then tmAgingSummary(UBound(tmAgingSummary)).iDaysBehind = ilDaysBehind
        
                    Exit For
                End If
            End If
        End If
    Next illoop
    
    If blFound = False Then
        ReDim Preserve tmAgingSummary(0 To UBound(tmAgingSummary) + 1) As AGEINGSUMMARY
        tmAgingSummary(UBound(tmAgingSummary)).iYear = tlRvr.iAgingYear 'year
        tmAgingSummary(UBound(tmAgingSummary)).iMonth = tlRvr.iAgePeriod 'month
        tmAgingSummary(UBound(tmAgingSummary)).iAgencyCode = IIF(tlRvr.iAgfCode = 0, tlRvr.iAdfCode, tlRvr.iAgfCode) 'Agency = if {RVR_Receivables_Rept.rvragfCode} = 0 then   //direct    {ADF_Advertisers.adfName} Else    {AGF_Agencies.agfName}
        tmAgingSummary(UBound(tmAgingSummary)).iAdvertiserCode = tlRvr.iAdfCode 'advertiser
        tmAgingSummary(UBound(tmAgingSummary)).iSalesPersonCode = tlRvr.iSlfCode  'salesman
        tmAgingSummary(UBound(tmAgingSummary)).lProductCode = tlRvr.lPrfCode 'product
        tmAgingSummary(UBound(tmAgingSummary)).lContractNumber = tlRvr.lCntrNo 'Contract#
        tmAgingSummary(UBound(tmAgingSummary)).lInvoiceNumber = tlRvr.lInvNo   'Invoice#
        tmAgingSummary(UBound(tmAgingSummary)).iInvoiceDate(0) = tlRvr.iInvDate(0) 'Invoice Date
        tmAgingSummary(UBound(tmAgingSummary)).iInvoiceDate(1) = tlRvr.iInvDate(1)
        gPDNToStr tlRvr.sNet, 2, slString
        tmAgingSummary(UBound(tmAgingSummary)).dBalance = Val(slString) 'Net Amount
        'Days Behind = TempDaysBehind := ({RVR_Receivables_Rept.rvrGenDate} - {@AgeingPeriodToDate}=Date ({RVR_Receivables_Rept.rvrAgeYear}, {RVR_Receivables_Rept.rvrAgePeriod}, {RVR_Receivables_Rept.rvrDistAmt}) );
        'gUnpackDate tlRvr.iGenDate(0), tlRvr.iGenDate(1), slGenDate
        'slAgeDate = tlRvr.iAgePeriod & "/" & tlRvr.lDistAmt & "/" & tlRvr.iAgingYear
        'ilDaysBehind = DateDiff("d", slAgeDate, slGenDate)
        'ilDaysBehind = IIF(ilDaysBehind < 0, 0, ilDaysBehind)
        'tmAgingSummary(UBound(tmAgingSummary)).iDaysBehind = ilDaysBehind
    End If
End Function

'TTP 10164 - Ageing summary by month - export option for Audacy A/R dump
Function mExportAgeMonthSummary() As String
    mExportAgeMonthSummary = ""
    Dim slCSVString As String
    Dim slString As String
    Dim ilRet As Integer
    Dim ilInt As Integer
    Dim illoop As Integer
    Dim llLong As Long
    Dim llDate As Long
    Dim tmPrfSrchKey As LONGKEY0
    Dim slStamp As String
    ReDim tlMnf(0 To 0) As MNF
    On Error GoTo ExportError
    
    For illoop = 0 To UBound(tmAgingSummary)
        slCSVString = ""
        'Agency,Advertiser,AE,Product,Order Number,Invoice Number,Invoice Date,Balance
        If tmAgingSummary(illoop).dBalance <> 0 Then
            slCSVString = ""
            'Year/Month
            slCSVString = slCSVString & """" & MonthName(tmAgingSummary(illoop).iMonth, True) & " " & tmAgingSummary(illoop).iYear & """"
            '--------------
            'Agency name - if {RVR_Receivables_Rept.rvragfCode} = 0 then //direct  {ADF_Advertisers.adfName}  Else {AGF_Agencies.agfName}
            If tmAgingSummary(illoop).iAgencyCode = 0 Then
                'Direct, use Avertiser
                ilInt = gBinarySearchAdf(tmAgingSummary(illoop).iAdvertiserCode)
                If ilInt = -1 Then
                    slCSVString = slCSVString & ","
                Else
                    slCSVString = slCSVString & ",""" & Trim(tgCommAdf(ilInt).sName) & """"
                End If
            Else
                'Loop up Agency Name
                ilInt = gBinarySearchAgf(tmAgingSummary(illoop).iAgencyCode)
                If ilInt = -1 Then
                    slCSVString = slCSVString & ","
                Else
                    slCSVString = slCSVString & ",""" & Trim(tgCommAgf(ilInt).sName) & """"
                End If
            End If
            '--------------
            'Advertiser name
            ilInt = gBinarySearchAdf(tmAgingSummary(illoop).iAdvertiserCode)
            If ilInt = -1 Then
                slCSVString = slCSVString & ","
            Else
                slCSVString = slCSVString & "," & """" & Trim(tgCommAdf(ilInt).sName) & """"
            End If
                        
            '--------------
            'AE (salesperson name)
            If tmAgingSummary(illoop).iSalesPersonCode <> 0 Then
                'Lookup Salesperson
                tmSlfSrchKey.iCode = tmAgingSummary(illoop).iSalesPersonCode
                ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                slString = Trim$(tmSlf.sLastName) & ", " & Trim$(tmSlf.sFirstName)
                If igSlfFirstNameFirst Then
                    slString = Trim$(tmSlf.sFirstName) & " " & Trim$(tmSlf.sLastName)
                Else
                    slString = Trim$(tmSlf.sLastName) & ", " & Trim$(tmSlf.sFirstName)
                End If
                 slCSVString = slCSVString & ",""" & slString & """"
            Else
                slCSVString = slCSVString & ","
            End If
            
            '--------------
            'Product
            tmPrfSrchKey.lCode = tmAgingSummary(illoop).lProductCode
            ilInt = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            If ilInt = BTRV_ERR_NONE Then
                slCSVString = slCSVString & ",""" & Trim$(tmPrf.sName) & """"
            Else
                slCSVString = slCSVString & ","
            End If

            '--------------
            'Order Number
            slCSVString = slCSVString & "," & tmAgingSummary(illoop).lContractNumber
            
            '--------------
            'Invoice Number
            slCSVString = slCSVString & "," & tmAgingSummary(illoop).lInvoiceNumber
            
            '--------------
            'Invoice Date
            gUnpackDateLong tmAgingSummary(illoop).iInvoiceDate(0), tmAgingSummary(illoop).iInvoiceDate(1), llDate
            slCSVString = slCSVString & "," & Format(llDate, "mm/dd/yy")
            
            '--------------
            'Balance
            slCSVString = slCSVString & "," & Format(tmAgingSummary(illoop).dBalance, "#.00")
        
            '--------------
            'Write File
            Print #hmExport, slCSVString
            'Show some status
            lmExportCount = lmExportCount + 1
            If lmExportCount / 10 - Int(lmExportCount / 10) = 0 Then
                'show every 10
                RptSel.lacExport.Caption = "Exported " & lmExportCount & " records..."
                RptSel.lacExport.Refresh
            End If
            mExportAgeMonthSummary = "Exported " & lmExportCount & " records..."
            
        End If
    Next illoop
    
    Exit Function
    
ExportError:
    mExportAgeMonthSummary = "Error:" & err & "-" & Error(err)
End Function


'TTP 10118 -Billing Distribution Export to CSV
Function mExportInvDist(tlRvr As RVR) As String
    mExportInvDist = ""
    Dim slCSVString As String
    Dim slString As String
    Dim ilInt As Integer
    Dim illoop As Integer
    Dim llLong As Long
    Dim llDate As Long
    Dim tmPrfSrchKey As LONGKEY0
    Dim slStamp As String
    ReDim tlMnf(0 To 0) As MNF
    slCSVString = ""
    On Error GoTo ExportError
    
    'get position in the global array of vehicles: tgMVef
    Dim ilVefInt As Integer
    ilVefInt = gBinarySearchVef(tlRvr.iAirVefCode)
    'get position in the global array of Advertisers: tgCommAdf
    Dim ilAdfInt As Integer
    ilAdfInt = gBinarySearchAdf(tlRvr.iAdfCode)
    'get position in the global array of Agencies: tgCommAgf
    Dim ilAgfInt As Integer
    ilAgfInt = gBinarySearchAgf(tlRvr.iAgfCode)
    'get position in the global array of Salespersons: tgMSlf
    Dim sgSalespersonName As String
    Dim ilSlfInt As Integer
    ilSlfInt = gBinarySearchSlf(tmChf.iSlfCode(0))
    '--------------
    'Detail
    If RptSel!rbcSelCInclude(0).Value Then
        'TTP 10459 - Bill Dist. participant shows sales source:
        '"Participant,Sales Source,Airing Vehicle,Advertiser,Agency,Contract,Invoice Date,Inv #,Type,Gross Billed,Net Billed,Distribution %,Distribution Due,MissingSSFlag"
        
'        gPDNToStr tmRvr.sNet, 2, slString
'        If Val(slString) = 0 Then
'            'Dont export these bad records
'            Exit Function
'        End If
        '--------------
        'Account ID (TTP 10902)
        'Fix v81 TTP 10902 - per Jason email Thu 12/21/23 4:01 PM
        If ilAgfInt = -1 Then 'If no Agency then use Direct Advertiser
            slCSVString = slCSVString & ""
        Else
            slCSVString = slCSVString & """" & Trim((tgCommAgf(ilAgfInt).sCodeStn)) & """"
        End If
        
        '--------------
        'Participant
        If tlRvr.imnfOwner = 0 Then
            slString = "Unknown Participant"
        Else
            '{MNF_Multi_Names.mnfName}
            slString = mGetMnfName(tlRvr.imnfOwner)
        End If
        slCSVString = slCSVString & ",""" & Trim(slString) & """"
        
        '--------------
        'SalesSource
        If tlRvr.iMnfSSCode = 0 Then
            slString = "Unknown S/S"
        Else
            '{MNF_Multi_Names.mnfName}
            slString = mGetMnfName(tlRvr.iMnfSSCode)
        End If
        slCSVString = slCSVString & "," & """" & Trim(slString) & """"
        
        '--------------
        'Airing Vehicle
        'ilInt = gBinarySearchVef(tlRvr.iAirVefCode)
        If ilVefInt = -1 Then
            slCSVString = slCSVString & ","
        Else
            slCSVString = slCSVString & ",""" & Trim(tgMVef(ilVefInt).sName) & """"
        End If
        
        '--------------
        'Advertiser
        'ilInt = gBinarySearchAdf(tlRvr.iAdfCode)
        If ilAdfInt = -1 Then
            slCSVString = slCSVString & ","
        Else
            slCSVString = slCSVString & ",""" & Trim(tgCommAdf(ilAdfInt).sName) & """"
        End If
        
        '--------------
        'Advertiser Reference ID (TTP 10902)
        If ilAdfInt = -1 Then
            slCSVString = slCSVString & ","
        Else
            slCSVString = slCSVString & ",""" & Trim(mGetAdfxRefID(tgCommAdf(ilAdfInt).iCode)) & """"
        End If
        
        '--------------
        'Advertiser Political (TTP 10902)
        If ilAdfInt = -1 Then
            slCSVString = slCSVString & ","
        Else
            slCSVString = slCSVString & ",""" & Trim(tgCommAdf(ilAdfInt).sPolitical) & """"
        End If
        
        '--------------
        'Agency
        'ilInt = gBinarySearchAgf(tlRvr.iAgfCode)
        If ilAgfInt = -1 Then
            slCSVString = slCSVString & ","
        Else
            slCSVString = slCSVString & "," & """" & Trim(tgCommAgf(ilAgfInt).sName) & """"
        End If
        
        '--------------
        'Agency Reference ID (TTP 10902)
        If ilAgfInt = -1 Then
            slCSVString = slCSVString & ","
        Else
            slCSVString = slCSVString & ",""" & Trim(mGetAgfxRefID(tgCommAgf(ilAgfInt).iCode)) & """"
        End If
        
        '--------------
        'Contract
        slCSVString = slCSVString & "," & tlRvr.lCntrNo
        
        '--------------
        'Product (TTP 10902)
        slCSVString = slCSVString & ",""" & Trim(tmChf.sProduct) & """"
        
        '--------------
        'Sales Office (TTP 10902)
        'tmChf.iSlfCode (0) =>  tgMSlf(ilSlfInt).iSofCode
        slCSVString = slCSVString & ",""" & Trim(mGetSOFName(tgMSlf(ilSlfInt).iSofCode)) & """"
        
        '--------------
        'Salesperson ID (TTP 10902)
        slCSVString = slCSVString & "," & Trim(tmChf.iSlfCode(0)) & ""
        
        '--------------
        'SalesPerson (TTP 10902)
        'gObtainSalespersonName tmChf.iSlfCode(0), sgSalespersonName, True
        slCSVString = slCSVString & ",""" & Trim$(tgMSlf(ilSlfInt).sFirstName) & " " & Trim$(tgMSlf(ilSlfInt).sLastName) & """"
        
        '--------------
        'Invoice Date
        gUnpackDateLong tlRvr.iInvDate(0), tlRvr.iInvDate(1), llDate
        slCSVString = slCSVString & "," & Format(llDate, "ddddd")
        
        '--------------
        'Inv#
        slCSVString = slCSVString & "," & tlRvr.lInvNo
        
        '--------------
        'Trans Type
        'slCSVString = slCSVString & "," & tlRvr.sTranType & IIF(tlRvr.sCashTrade = "T", " T", "") '{@TradeFlag}
        slCSVString = slCSVString & ",""" & tlRvr.sTranType & """"
        
        '--------------
        'Gross Billed
        gPDNToStr tmRvr.sGross, 2, slString
        slCSVString = slCSVString & "," & slString
        
        '--------------
        'Net Billed
        gPDNToStr tmRvr.sNet, 2, slString
        slCSVString = slCSVString & "," & slString
        
        '--------------
        'Distribution %
        slCSVString = slCSVString & "," & Format(tlRvr.iProdPct / 100, "#0.00") & "%"
        
        '--------------
        'Distribution Due
        slCSVString = slCSVString & "," & Format(tlRvr.lDistAmt / 100, "#0.00")
        
        '--------------
        'MissingSSFlag
        slCSVString = slCSVString & ",""" & IIF(Trim(tlRvr.sSource) = "#", "*", "") & """"
        
        '--------------
        'Cash/Trade Flag (TTP 10902)
        slCSVString = slCSVString & ",""" & IIF(tlRvr.sCashTrade = "T", "T", "C") & """"
        
    Else
        'Summary
        '"Participant,Airing Vehicle,Gross,Net,Distribution Due"
        'this is done with mExportInvDist
    End If
    
    '--------------
    'Write File
    Print #hmExport, slCSVString
    'Show some status
    lmExportCount = lmExportCount + 1
    If lmExportCount / 10 - Int(lmExportCount / 10) = 0 Then
        'show every 10
        RptSel.lacExport.Caption = "Exported " & lmExportCount & " records..."
        RptSel.lacExport.Refresh
    End If
    mExportInvDist = "Exported " & lmExportCount & " records..."
    Exit Function
    
ExportError:
    mExportInvDist = "Error:" & err & "-" & Error(err)
End Function

Private Function mGetMnfName(ilMnfCode As Integer) As String
    Dim slSQLQuery As String
    Dim tmp_rst As ADODB.Recordset
    slSQLQuery = "Select mnfName from MNF_Multi_Names where mnfCode = " & ilMnfCode
    Set tmp_rst = gSQLSelectCall(slSQLQuery)
    If Not tmp_rst.EOF Then
        mGetMnfName = Trim$(tmp_rst!mnfName)
    Else
        mGetMnfName = ""
    End If
End Function

Function mSaveInvDistSummary(tlRvr As RVR) As String
    mSaveInvDistSummary = ""
    Dim illoop As Integer
    'Sum rvrNet by Year/Month / Agency / Contract#
    Dim blFound As Boolean
    Dim slString As String
    Dim slGenDate As String
    Dim slAgeDate As String
    Dim ilDaysBehind As Integer
    slString = ""
    blFound = False
    slGenDate = ""
    slAgeDate = ""
    
    lmExportCount = lmExportCount + 1
    If lmExportCount / 10 - Int(lmExportCount / 10) = 0 Then
        'show every 10
        RptSel.lacExport.Caption = "Processing " & lmExportCount & " records..."
        RptSel.lacExport.Refresh
    End If
    mSaveInvDistSummary = "Processing " & lmExportCount & " records..."
    
    'Check tmInvDistSummary for existing stuff
    For illoop = 1 To UBound(tmInvDistSummary)
        blFound = False
    
        If tmInvDistSummary(illoop).sCashTrade = tlRvr.sCashTrade Then
            If tmInvDistSummary(illoop).iParticipant = tlRvr.iMnfSSCode Then
                If tmInvDistSummary(illoop).iSalesSource = tlRvr.imnfOwner Then
                    If tmInvDistSummary(illoop).iAiringVehicle = tlRvr.iAirVefCode Then
                        blFound = True
                        '(Add) Net
                        gPDNToStr tlRvr.sNet, 2, slString
                        tmInvDistSummary(illoop).dNet = tmInvDistSummary(illoop).dNet + Val(slString)
                        '(Add) Gross
                        gPDNToStr tlRvr.sGross, 2, slString
                        tmInvDistSummary(illoop).dGross = tmInvDistSummary(illoop).dGross + Val(slString)
                        '(Add) Dist Due
                        tmInvDistSummary(illoop).lDistDue = tmInvDistSummary(illoop).lDistDue + tlRvr.lDistAmt
                        'Cash / Trade / Unknown
                        If tmInvDistSummary(UBound(tmInvDistSummary)).sCashTrade = "" Or tmInvDistSummary(UBound(tmInvDistSummary)).sCashTrade = "U" Then
                            If tlRvr.sCashTrade = "T" Then
                                tmInvDistSummary(UBound(tmInvDistSummary)).sCashTrade = "T"
                            ElseIf tlRvr.sCashTrade = "C" Then
                                tmInvDistSummary(UBound(tmInvDistSummary)).sCashTrade = "C"
                            Else
                                tmInvDistSummary(UBound(tmInvDistSummary)).sCashTrade = "U"
                            End If
                        End If
                        Exit For
                    End If
                End If
            End If
        End If
    Next illoop
    
    If blFound = False Then
        ReDim Preserve tmInvDistSummary(0 To UBound(tmInvDistSummary) + 1) As INVDISTSUMMARY
        'Cash / Trade / Unknown
        If tlRvr.sCashTrade = "T" Then
            tmInvDistSummary(UBound(tmInvDistSummary)).sCashTrade = "T"
        ElseIf tlRvr.sCashTrade = "C" Then
            tmInvDistSummary(UBound(tmInvDistSummary)).sCashTrade = "C"
        Else
            tmInvDistSummary(UBound(tmInvDistSummary)).sCashTrade = "U"
        End If
        tmInvDistSummary(UBound(tmInvDistSummary)).iAiringVehicle = tlRvr.iAirVefCode
        'TTP 10459 - Bill Dist. participant shows sales source:
        tmInvDistSummary(UBound(tmInvDistSummary)).iParticipant = tlRvr.imnfOwner
        tmInvDistSummary(UBound(tmInvDistSummary)).iSalesSource = tlRvr.iMnfSSCode
        'tmInvDistSummary(UBound(tmInvDistSummary)).iVehicle = tlRvr.iBillVefCode
        tmInvDistSummary(UBound(tmInvDistSummary)).lDistDue = tlRvr.lDistAmt  'Dist Due
        gPDNToStr tlRvr.sGross, 2, slString
        tmInvDistSummary(UBound(tmInvDistSummary)).dGross = Val(slString) 'Gross
        gPDNToStr tlRvr.sNet, 2, slString
        tmInvDistSummary(UBound(tmInvDistSummary)).dNet = Val(slString) 'Net
        tmInvDistSummary(UBound(tmInvDistSummary)).sCashTrade = Trim(tlRvr.sCashTrade)
        tmInvDistSummary(UBound(tmInvDistSummary)).sMissingSSFlag = IIF(Trim(tlRvr.sSource) = "#", "*", "")
    End If
End Function

'TTP 10118 -Billing Distribution Export to CSV
Function mExportInvDistSummary() As String
    mExportInvDistSummary = ""
    Dim slCSVString As String
    Dim slString As String
    Dim ilRet As Integer
    Dim ilInt As Integer
    Dim illoop As Integer
    Dim llLong As Long
    Dim llDate As Long
    Dim tmPrfSrchKey As LONGKEY0
    Dim slStamp As String
    ReDim tlMnf(0 To 0) As MNF
    On Error GoTo ExportError
    
    For illoop = 1 To UBound(tmInvDistSummary)
        slCSVString = ""
        slString = ""
        '"Participant,Sales Source,Airing Vehicle,Gross,Net,Distribution Due,MissingSSFlag"
        'TTP 10459 - Bill Dist. participant shows sales source
        '--------------
        'Participant
        If tmInvDistSummary(illoop).iParticipant = 0 Then
            slString = "Unknown Participant"
        Else
            '{MNF_Multi_Names.mnfName}
            slString = mGetMnfName(tmInvDistSummary(illoop).iParticipant)
        End If
        slCSVString = slCSVString & """" & Trim(slString) & """"
        
        '--------------
        'Cash/Trade (+ S/S)
        If tmInvDistSummary(illoop).sCashTrade = "T" Then
            slString = "Trade"
        ElseIf tmInvDistSummary(illoop).sCashTrade = "C" Then
            slString = "Cash"
        Else
            slString = "Unknown C/T"
        End If
        slString = slString & " Sold By "
        
        '--------------
        'Sales Source
        If tmInvDistSummary(illoop).iSalesSource = 0 Then
            slString = slString & "Unknown S/S"
        Else
            '{MNF_Multi_Names.mnfName}
            slString = slString & mGetMnfName(tmInvDistSummary(illoop).iSalesSource)
        End If
        slCSVString = slCSVString & ",""" & Trim(slString) & """"
        
        '--------------
        'Airing Vehicle
        ilInt = gBinarySearchVef(tmInvDistSummary(illoop).iAiringVehicle)
        If ilInt = -1 Then
            slCSVString = slCSVString & ","
        Else
            slCSVString = slCSVString & ",""" & Trim(tgMVef(ilInt).sName) & """"
        End If
        
        '--------------
        'Gross
        slCSVString = slCSVString & "," & Format(tmInvDistSummary(illoop).dGross, "#0.00")
        
        '--------------
        'Net
        slCSVString = slCSVString & "," & Format(tmInvDistSummary(illoop).dNet, "#0.00")
        
        '--------------
        'Distribution Due
        slCSVString = slCSVString & "," & Format(tmInvDistSummary(illoop).lDistDue / 100, "#0.00")
        
        '--------------
        'MissingSSFlag
        slCSVString = slCSVString & "," & tmInvDistSummary(illoop).sMissingSSFlag
        
        '--------------
        'Write File
        Print #hmExport, slCSVString
        'Show some status
        lmExportCount = lmExportCount + 1
        If lmExportCount / 10 - Int(lmExportCount / 10) = 0 Then
            'show every 10
            RptSel.lacExport.Caption = "Exported " & lmExportCount & " records..."
            RptSel.lacExport.Refresh
        End If
        mExportInvDistSummary = "Exported " & lmExportCount & " records..."
    Next illoop
    
    Exit Function
    
ExportError:
    mExportInvDistSummary = "Error:" & err & "-" & Error(err)
End Function


Function mGetAdfxRefID(ilAdfID As Integer) As String
    Dim slSql As String
    Dim myRsQuery As ADODB.Recordset
    mGetAdfxRefID = ""
    If ilAdfID = 0 Then Exit Function
    If ilAdfID = imLastAdfCode Then
        mGetAdfxRefID = smLastAdfName
        Exit Function
    End If
    
    mGetAdfxRefID = ""
    slSql = "select adfxRefId as Code  from ADFX_Advertisers where adfxCode = " & ilAdfID
    Set myRsQuery = gSQLSelectCall(slSql)
    If Not myRsQuery.EOF Then
        mGetAdfxRefID = myRsQuery!Code
        imLastAdfCode = ilAdfID
        smLastAdfName = myRsQuery!Code
    End If
End Function

Function mGetAgfxRefID(ilAgfID As Integer) As String
    Dim slSql As String
    Dim myRsQuery As ADODB.Recordset
    mGetAgfxRefID = ""
    If ilAgfID = 0 Then Exit Function
    If ilAgfID = imLastAgyCode Then
        mGetAgfxRefID = smLastAgyName
        Exit Function
    End If
    slSql = "select agfxRefId as Code from AGFX_Agencies where agfxCode = " & ilAgfID
    Set myRsQuery = gSQLSelectCall(slSql)
    If Not myRsQuery.EOF Then
        mGetAgfxRefID = myRsQuery!Code
        imLastAgyCode = ilAgfID
        smLastAgyName = myRsQuery!Code
    End If
End Function

Function mGetSOFName(ilSofCode As Integer) As String
    Dim slSql As String
    Dim myRsQuery As ADODB.Recordset
    mGetSOFName = ""
    If ilSofCode = 0 Then Exit Function
    
    If ilSofCode = imLastSofCode Then
        mGetSOFName = smLastSofName
        Exit Function
    End If
    
    slSql = "SELECT sofName"
    slSql = slSql & " From SOF_Sales_Offices"
    slSql = slSql & " Where sofCode = " & ilSofCode
    
    Set myRsQuery = gSQLSelectCall(slSql)
    If Not myRsQuery.EOF Then
        mGetSOFName = Trim(myRsQuery!sofName)
        smLastSofName = mGetSOFName
    End If
End Function
