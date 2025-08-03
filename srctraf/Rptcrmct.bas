Attribute VB_Name = "RPTCRMCT"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrmct.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Procedures (Marked)                                                              *
'*  mGetBudgetWks                                                                         *
'******************************************************************************************

Option Explicit
Option Compare Text

'Dim lmStartDates(1 To 15) As Long       'max 14 for 14 week quarter
Dim lmStartDates(0 To 15) As Long       'max 14 for 14 week quarter. Index zero ignored
Dim imIncludeCodes As Integer
Dim imUseCodes() As Integer
Dim imInclAdvtCodes As Integer      '6-21-18
Dim imUseAdvtCodes() As Integer
Dim tmVefDollars() As ADJUSTLIST
Dim imAdvt As Integer
Dim imBusCat As Integer                 'Date: 10/3/2019
Dim imProdProt As Integer               'Date: 10/3/2019
Dim imVehGroup As Integer               'Date: 10/3/2019
Dim imNTR As Integer
Dim imAirTime As Integer
Dim imHardCost As Integer
Dim imSlsp As Integer                   'true if slsp option
Dim imVehicle As Integer                'true if vehicle option
Dim imAgency As Integer                 'true if agency option
Dim imDead As Integer
Dim imWork As Integer
Dim imIncomplete As Integer
Dim imComplete As Integer
Dim imRev As Integer                    '11-7-16
Dim smGrossOrNet As String
Dim imDiscrepOnly As Integer
Dim imCreditCk As Integer
Dim imShowInternal As Integer       'show internal comment
Dim imShowOther As Integer          'show other comment
Dim imShowChgRsn As Integer         'show change reason comment
Dim imShowCancel As Integer         'show cancellation
Dim imInclAcq As Integer            'include acquisition costs

Dim imVehGroupMinor As Integer      'Veh Group minor sort
Dim imVehicleMinor As Integer       'Vehicle minor sort
Dim imProdProtMinor As Integer      'Prod Protection minor sort
Dim imBusCatMinor As Integer        'Bus Cat minor sort
Dim imAgencyMinor As Integer        'Agency minor sort
Dim imAdvtMinor As Integer          'Advertiser minor sort
Dim imSlspMinor As Integer          'Sales people minor sort

Dim hmAlf As Integer            'locked avails
Dim tmAlf As ALF
Dim tmAlfSrchkey1 As ALFKEY1
Dim imAlfRecLen As Integer

Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey1 As CHFKEY1    'CHF record image
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChf As CHF
Dim tlChfAdvtExt() As CHFADVTEXT
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
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
Dim hmRdf As Integer
Dim tmRdfSrchKey As INTKEY0
'Dayparts file handle
Dim imRdfRecLen As Integer      'RD record length
Dim tmRdf As RDF
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim tlMMnf() As MNF                    'array of MNF records for specific type
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVsf As Integer            'Virtual Vehicle file handle
Dim imVsfRecLen As Integer        'VSF record length
Dim hmSsf As Integer            'Spot Summary file handle
Dim tmSsf As SSF                'SSF record image
Dim imSsfRecLen As Integer
Dim hmSdf As Integer            'Spot detail file handle
Dim imSdfRecLen As Integer
Dim hmSmf As Integer            'Mg/outside file handle
Dim imSmfRecLen As Integer
Dim tmRcf As RCF
Dim hmRcf As Integer            'Rate Card file handle
Dim imRcfRecLen As Integer      'RCF record length
Dim tmRif As RIF
Dim hmRif As Integer            'Rate Card items file handle
Dim imRifRecLen As Integer      'RIF record length
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim tmZeroGrf As GRF              'initialized Generic recd
Dim tmAnr As ANR                    'prepass analysis file
Dim hmAnr As Integer
Dim imAnrRecLen As Integer        'ANR record length
Dim tmZeroAnr As ANR              'initialized Analysis recd
Dim tmBvf As BVF                  'Budgets by office & vehicle
Dim hmBvf As Integer
Dim imBvfRecLen As Integer        'BVF record length
Dim tmBvfVeh() As BVF               'Budget by vehicle
Dim tmPjf As PJF                  'Slsp Projections
Dim hmPjf As Integer
Dim tmPjfSrchKey As PJFKEY0       'Gen date and time
Dim imPjfRecLen As Integer        'PJF record length
'4-20-00
Dim tmScf As SCF                  'Slsp Commission file
Dim hmScf As Integer
Dim imScfRecLen As Integer        'SCF record length

'Quarterly Avails
Dim imStandard As Integer
Dim imRemnant As Integer    'True=Include Remnant
Dim imReserv As Integer  'true = include reservations
Dim imDR As Integer     'True =Include Direct Response
Dim imPI As Integer     'True=Include per Inquiry
Dim imPSA As Integer    'True=Include PSA
Dim imPromo As Integer  'True=Include Promo
Dim imHold As Integer   'true = include hold contracts
Dim imTrade As Integer  'true = include trade contracts
Dim imCash As Integer
Dim imOrder As Integer  'true = include Complete order contracts
'Log Calendar
Dim hmLcf As Integer            'Log Calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim imLcfRecLen As Integer        'LCF record length
'Copy inventory
'  Copy Product/Agency File
' Time Zone Copy FIle
'  Media code File
'  Rating Book File
Dim hmDrf As Integer        'Rating book file handle
Dim tmDrf As DRF            'DRF record image
Dim imDrfRecLen As Integer  'DRF record length
' Demo Plus File
Dim hmDpf As Integer        'Demo Plus file handle
Dim tmDpf As DPF            'DPF record image
Dim imDpfRecLen As Integer  'DPF record length
'  Research Estimates
Dim hmDef As Integer
Dim hmRaf As Integer
'  Receivables File
Dim hmRvf As Integer        'receivables file handle
Dim tmRvf As RVF            'RVF record image
Dim imRvfRecLen As Integer  'RVF record length

'  Special Billing  File
Dim hmSbf As Integer        'SBF file handle
Dim tmSbf As SBF            'SBF record image
Dim tmSbfSrchKey As LONGKEY0 'SBF key record image
Dim tmSbfSrchKey0 As SBFKEY0
Dim imSbfRecLen As Integer  'SBF record length

'  Receivables Report File
Dim tlSofList() As SOFLIST    'list of selling office codes and sales sourcecodes
Dim tlActList() As ACTLIST      'Sales Activity list containing advt, potential code, $
Dim tlSlsList() As SLSLIST      'Sales Analysis Summary
Type FLIGHT
    iLen As Integer                 '12-23-08
    iLine As Integer                '1-7-09
    lActPrice As Long
    'iSpots(1 To 14) As Integer      '2-4-00 year 2000 has 14 week qtr
    iSpots(0 To 14) As Integer      '2-4-00 year 2000 has 14 week qtr. Index zero ignored
    'lProj(1 To 14) As Long          '2-0-00 year 2000 has 14 week qtr
    lProj(0 To 14) As Long          '2-0-00 year 2000 has 14 week qtr. Index zero ignored
    sDescription As String * 40     'daypart description with/without overrides
End Type
Dim tmFlight() As FLIGHT

Dim tmRVFComm() As RVFCOMM
Dim tmAdvtComm() As RVFCOMM     'bonus version
Dim tmSalesGoalInfo() As SALESGOAL   'each slsp info regarding ytd for sales goal (sales comm bonus version only)

Type RVFCOMM
    sKey  As String * 50 'key for sorting: slsp code|cnt code|date     (no bonus version)
    '                                         6         8      5
    'key for sorting ( bonus version :    slsp code|advt name|ContrType\comm %
    '                                         6         30      1         5
    iSlfCode As Integer 'Salesperson code number (if direct: default to advt otherwisw default to agency)
    iSofCode As Integer
    lChfCode As Long     'Contract code
    lInvNo As Long      'Invoice number
    iAirVefCode As Integer 'Vehicle Code of Airing vehicle
    iTranDate(0 To 1) As Integer  'Transaction Date of Rate Card or zero if not superseded
    lGross As Long  'Gross amount (xx,xxx,xxx.xx)
    lNet As Long   'Net amount (xx,xxx,xxx.xx)
    lMerch As Long  'merch amt
    lPromo As Long  'promotions amt
    lAdjComm As Long    'calc from gross or net minus metch - promotions
    sCashTrade As String * 1    'C=Cash; T=Trade; M=Merchandising; P=Promotion
    iUseThisCommPct As Integer      'comm % to use in report for sorting
    iSlfCommPct As Integer          'non remnant slsp comm% under goal or sbf comm %
    iSlfOverCommPct As Integer      'non-remnant slsp comm% over goal
    lSlfSplit As Long               'revenue share %
    iRemUnderPct As Integer         'remnant slsp comm% under goal
    iRemOverPct As Integer          'remnant slsp comm% over goal
    iStartFiscalMonth As Integer    'fiscal month (1-12)
    iCurrentMonth As Integer        'fiscal year (2000, 2001, etc)
    iDiscrep As Integer             'error flag if discrepancy in vehicle/slsp commission (sub-companies)
    '                               0 if no error, 1 if discrepancy
    sNTR As String * 1              'N = ntr, else blank
    lGoal As Long               'sales goal from slf
    sType As String * 1         'contract header type (for remnant)
    lAcquisition As Long
    iAdfCode As Integer         'backlog may not have chf header, use recv/phf advertiser
    lContrNo As Long
    lRvfCode As Long
    lLastYear As Long           'bonus version
    lThisYearPrevMonth As Long  'bonus version
    lThisYearCurrent As Long    'bonus version
End Type

'Sales Commission Bonus Version; required because all advertisers need to be accounted for
'since only the advertisers that are billed in the current month are the only ones shown
'on the report
Type SALESGOAL
    'these 5 fields are for the salesperson heading for goal info
    dPrevYTDNN As Long        'total previous ytd net net by salesperson, non-remnant
    dPrevYTDGross As Long        'total previous gross by slsp, non-remnant
    dPrevYTDNet As Long          'total previous net by slsp, non-remnant
    dPrevYTDMerchPromo As Long   'total previous merch/promo by slsp, non-remnant
    dPrevYTDAcq As Long      'total previous acquisition costs by slsp, non-remnant
    dPrevYTDNNREmnant As Long  'total previous ytd net net remnant by slsp, non-remnant
    dYTDNNRemnant As Long       'total YTD (incl current month) by slsp , remnants
    dYTDNN As Long              'total ytd netnet, incl current month) by slsp, remnant
'    dPrevYTDNN As Double        'total previous ytd net net by salesperson, non-remnant
'    dPrevYTDGross As Double        'total previous gross by slsp, non-remnant
'    dPrevYTDNet As Double          'total previous net by slsp, non-remnant
'    dPrevYTDMerchPromo As Double   'total previous merch/promo by slsp, non-remnant
'    dPrevYTDAcq As Double      'total previous acquisition costs by slsp, non-remnant
'    dPrevYTDNNREmnant As Double  'total previous ytd net net remnant by slsp, non-remnant
'    dYTDNNRemnant As Double       'total YTD (incl current month) by slsp , remnants
'    dYTDNN As Double
End Type

''Billed and Booked Comparison Budget info
'Type BUDGETSBYVEF
'    iVefCode As Integer             'vehicle code for budget
'    iMnfCode As Integer             'vehicle group selected
'    lBudgetByMonth(1 To 13) As Long
'End Type

Dim lmChfCode As Long           '4-24-06 for Monthly Sales Activity to show info (i.e. product name ) from latest version
Dim lmSingleCntr As Long
Dim imInclNonPolit As Integer
Dim imInclPolit As Integer
'Dan M 7-30-08 added ntr/hard cost to report
Dim tmMnfNtr() As MNF
Const NOT_SELECTED = 0

'Sort selection for Insertion Order Activity report
'If any of these sort numbers change, the .rpt (InsertionActivity.rpt) must be changed to match
Const INSERT_ACT_ADV = 0
Const INSERT_ACT_AGY = 1
Const INSERT_ACT_AGYEST = 2
Const INSERT_ACT_CONTRACT = 3
Const INSERT_ACT_SENDER = 4
Const INSERT_ACT_SENTDATE = 5
Const INSERT_ACT_STATUS = 6

'Date: 8/30/2019 Create array of agy, advt, vehicle (etc) codes to include or exclude
'Dim imUseCodes() As Integer            'codes to include or exclude (for agy, advt, vehicle, etc)
Dim imUsevefcodes() As Integer          'array of vehicle codes to include/exclude
Dim imInclVefCodes As Integer           'flag to incl or exclude vehicle codes
Dim imUseAdfCodes() As Integer          'array of advt codes to include/exclude
Dim imInclAdfCodes As Integer           'flag to incl or exclude advt codes
Dim imUseAgfCodes() As Integer          'array of agy codes to include/exclude
Dim imInclAgfCodes As Integer           'flag to incl or exclude agy codes
Dim imUseCatCodes() As Integer          'array of bus category codes to include/exclude
Dim imInclCatCodes As Integer           'flag to incl or exclude bus category codes
Dim imUseProdCodes() As Integer         'array of prod prot codes to include/exclude
Dim imInclProdCodes As Integer          'flag to incl or exclude prod prot codes
Dim imUseSlfCodes() As Integer          'array of slsp codes to include/exclude
Dim imUseVefGrpCodes() As Integer       'array of vehicle group codes to include/exclude
Dim imInclVefGrpCodes As Integer        'flag to incl or exclude vehicle group codes

'TTP 10119
Dim lmExportCount As Long ' TTP 10252 - Ageing Summary by Month: overflow error when exporting
Dim hmExport As Integer
Dim smExportStatus As String
Dim smClientName As String
Dim tmMnfSrchKey As INTKEY0
Dim tmMnfList() As MNFLIST      'array of mnf codes for Missed reasons and billing rules


'               Test table for unique Sales Source and Vehicle in Billed and Booked Comparison
'               When budgets are included, get budgets only for those vehicles and sales sources.
'               In addition, if the vehicle spans multiple sales sources and using Sales source as major sort,
'               the budgets should be retrieved for each sales source.
'               If not using sales source as major sort, sales source do not apply
'                4-15-09
'
'           <input>
'                    ilSSMnfCode - Sales Source MNF code
'                    ilVefCode - vehicle code
'           Return - updated tgBBCompare array
Public Sub gTestSSForBBCompare(ilSSMnfCode As Integer, ilVefCode As Integer)

Dim ilUpper As Integer
Dim ilFound As Integer
Dim illoop As Integer
Dim ilListIndex As Integer
Dim ilTempSSMnfCode As Integer

        ilTempSSMnfCode = ilSSMnfCode
        If RptSelCt!ckcSelC13(2).Value = vbUnchecked Then         'do not use sales source as major sort
            ilTempSSMnfCode = 0                         'all sales sources are the same
        End If

        ilListIndex = RptSelCt!lbcRptType.ListIndex
        If ilListIndex <> CNT_BOBCOMPARE Then
            Exit Sub
        End If

        ilFound = False
        ilUpper = UBound(tgBBCompare)
        For illoop = 0 To ilUpper
            If tgBBCompare(illoop).iSSCode = ilTempSSMnfCode And tgBBCompare(illoop).iVefCode = ilVefCode Then
                ilFound = True
                Exit For
            End If
        Next illoop
        If Not ilFound Then
            tgBBCompare(ilUpper).iSSCode = ilTempSSMnfCode
            tgBBCompare(ilUpper).iVefCode = ilVefCode
            ilUpper = ilUpper + 1
            ReDim Preserve tgBBCompare(0 To ilUpper) As BBCOMPARE
        End If
        Exit Sub
End Sub
'
'                   Create Average Rate report prepass file.  This was formerly
'                   a Crystal only report.  The report would not process past 10%
'                   of the data.  Could not get it to run with all the groupings required.
'                   Generate GRF file by rate/line.  As each line is processed, an
'                   array is built in tmFlight containing the unique Spot rate, the
'                   # spots ordered in the requested 13 weeks, plus the $ ordered
'                   in each of the 13 weeks.  All records generated are for
'                   the affected year and quarter only.
'
'
'                   Created D.Hosaka 1/19/98
'
'           2-2-00 Increase number of flight buckets for report from 13 to 14 due to 14 week qtr in year 2000
'           3-28-05 show open/close bb notation on line
'           12-23-08 integrate and add prepass for Avg Spot Price report
'           6-15-11 AVg Rate , Avg Spot Price & Adv Units Ordered: add option to get air time (spots scheduled) vs rep contracts
Sub gCrAvgRateCt(Optional blExport As Boolean = False)
    Dim ilListIndex As Integer
    Dim illoop As Integer
    Dim ilTemp As Integer
    Dim ilTotalSpots As Integer
    Dim ilRet As Integer
    Dim ilYear As Integer
    'ReDim llStartDates(1 To 2) As Long            'temp array for last year vs this years range of dates
    ReDim llStartDates(0 To 2) As Long            'temp array for last year vs this years range of dates. Index zero ignored
    Dim slStartDate As String                       'llLYGetFrom or llTYGetFrom converted to string
    Dim slEndDate As String                         'llLYGetTo or LLTYGetTo converted to string
    Dim slNameCode As String
    Dim slCode As String
    Dim slCntrTypes As String                       'valid contract types to access
    Dim slCntrStatus As String                      'valid status (holds, orders, working, etc) to access
    Dim ilHOState As Integer                        'include unsch holds/orders, sch holds/orders
    Dim ilFound As Integer
    Dim llContrCode As Long                         'contr code from gObtainCntrforDate
    Dim ilCurrentRecd As Integer                    'index of contract being processed from tlChfAdvtExt
    Dim ilClf As Integer                            'index to line from tgClfCt
    Dim ilCorpStd As Integer                        '1 = corp, 2 = std
    Dim ilLoopFlt As Integer
    Dim ilFoundVeh As Integer
    Dim ilMajorSet As Integer
    Dim ilMinorSet As Integer
    Dim ilMnfMajorCode As Integer           'field used to sort the major sort with
    Dim ilmnfMinorCode As Integer
    Dim slDate As String
    Dim ilOk As Integer
    Dim ilSlfInx As Integer
    Dim ilSSMnfCode As Integer
    Dim ilUnits As Integer
    Dim ilGrossNetTNet As Integer   '9-23-09 0 = gross, 1 = net , 2 = tnet
    Dim ilPeriods As Integer
    Dim slStr As String
    Dim ilIncludeAgyCodes As Integer            '12-9-16
    Dim ilUseAgyCodes() As Integer
    Dim ilAgfCode As Integer
    Dim slRepeat As String
    Dim slFileName As String
    
    'TTP 10119 - Average 30 Rate Report - add option to export to CSV
    If blExport = True Then
        RptSelCt.lacExport.Caption = "Exporting..."
        RptSelCt.lacExport.Refresh
        DoEvents
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
            slFileName = ""
            slFileName = slFileName & "AvgRate-"
            'rbcSelCSelect (Month or Week)
            If RptSelCt!rbcSelCSelect(0).Value = True Then
                'by week
                slFileName = slFileName & "Weekly-"

                'rbcSelC7
                If RptSelCt!rbcSelC7(0).Value = True Then
                    slFileName = slFileName & "ByDaypart-"
                ElseIf RptSelCt!rbcSelC7(1).Value = True Then
                    slFileName = slFileName & "ByDaypartOV-"
                Else
                    slFileName = slFileName & "ByAgency-"
                End If
                
                slFileName = slFileName & RptSelCt!edcSelCTo.Text 'year
                slFileName = slFileName & "Q" & RptSelCt!edcSelCTo1.Text 'Quarter
                
                'rbcSelC9 (By Gross, Net, T-Net)
                If RptSelCt!rbcSelC9(0).Value = True Then
                    slFileName = slFileName & "-Gross"
                ElseIf RptSelCt!rbcSelC9(1).Value = True Then
                    slFileName = slFileName & "-Net"
                Else
                    slFileName = slFileName & "-TNet"
                End If
                                
                'edcSelCTo (Year)
                'edcSelCTo1 (qtr)
            Else
                'by month
                slFileName = slFileName & "Monthly-"
                
                'rbcSelC7
                If RptSelCt!rbcSelC7(0).Value = True Then
                    slFileName = slFileName & "ByDaypart-"
                ElseIf RptSelCt!rbcSelC7(1).Value = True Then
                    slFileName = slFileName & "ByDaypartOV-"
                Else
                    slFileName = slFileName & "ByAgency-"
                End If
        
                slStr = RptSelCt!edcSelCTo1.Text           'month in text form (jan..dec)
                gGetMonthNoFromString slStr, ilRet         'getmonth #
                If ilRet = 0 Then ilRet = Val(slStr)
                slFileName = slFileName & MonthName(ilRet, True) 'Month
                slFileName = slFileName & RptSelCt!edcSelCTo.Text 'Year
                slFileName = slFileName & "-" & RptSelCt!edcSelCFrom1.Text 'No months
                
                'rbcSelC9 (By Gross, Net, T-Net)
                If RptSelCt!rbcSelC9(0).Value = True Then
                    slFileName = slFileName & "-Gross"
                ElseIf RptSelCt!rbcSelC9(1).Value = True Then
                    slFileName = slFileName & "-Net"
                Else
                    slFileName = slFileName & "-TNet"
                End If
                
            End If
                        
            slFileName = slFileName & "-"
            slFileName = slFileName & Format(gNow, "mmddyy")
            slFileName = slFileName & slRepeat
            slFileName = slFileName & " " & gFileNameFilter2(Trim$(smClientName))
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
    End If

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)

    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmRdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imRdfRecLen = Len(tmRdf)

    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSof
        btrDestroy hmRdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imSofRecLen = Len(tmSof)
    
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAgf
        btrDestroy hmSof
        btrDestroy hmRdf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)


    Dim tlCntTypes As CNTTYPES
    'get all the dates needed to work with
    'slDate = RptSelCt!edcSelCFrom.Text               'effective date entred

    ilListIndex = RptSelCt!lbcRptType.ListIndex

    ilPeriods = 14                  'default to # periods, if avg rate could be changed to # of months to gather
    ilGrossNetTNet = 0              '9-23-09 default to gross
    '10-29-10 Advt units orderd, avg rates & avg spot prices all have gross, net & t-net options
    If RptSelCt!rbcSelC9(1).Value = True Then
        ilGrossNetTNet = 1
    ElseIf RptSelCt!rbcSelC9(2).Value = True Then
        ilGrossNetTNet = 2
    End If

    'ilRet = gBuildAcqCommInfo(RptSelCt)         'build acq rep commission table, if applicable

    If ilListIndex = CNT_AVGRATE Then
        ilYear = Val(RptSelCt!edcSelCTo.Text)           'year requested
        If RptSelCt!rbcSelCInclude(0).Value Then
            ilCorpStd = 1                               'corp flag for genl subtrn
        Else
            ilCorpStd = 2
        End If
        

        If RptSelCt!rbcSelCSelect(0).Value Then             '9-28-11 week option, generate 1 qtr output
            'This Years start/end quarter and year dates
            gGetStartEndQtr ilCorpStd, ilYear, igMonthOrQtr, slStartDate, slEndDate
            llStartDates(1) = gDateValue(slStartDate)
            llStartDates(2) = gDateValue(slEndDate)
            lmStartDates(1) = llStartDates(1)
            For ilTemp = 2 To 15
                lmStartDates(ilTemp) = lmStartDates(ilTemp - 1) + 7
            Next ilTemp
            ilTemp = (llStartDates(2) - llStartDates(1)) / 7
            
            'TTP 10119 - Average 30 Rate Report - add option to export to CSV
            If blExport = True Then
                'Generate Header
                If RptSelCt.rbcSelC7(0) Then 'Daypart
                    slStr = "Daypart,Vehicle,Advertiser,Product,Contract#,Ln,Len,30 Unit Rate,30 Unit Spots,Total,# Spots"
                End If
                If RptSelCt.rbcSelC7(1) Then 'Daypart w/Overrides
                    slStr = "Daypart w/Overrides,Vehicle,Advertiser,Product,Contract#,Ln,Len,30 Unit Rate,30 Unit Spots,Total,# Spots"
                End If
                If RptSelCt.rbcSelC7(2) Then 'Agency
                    slStr = "Agency,Vehicle,Advertiser,Product,Contract#,Ln,Len,30 Unit Rate,30 Unit Spots,Total,# Spots"
                End If
                ilPeriods = ilTemp          'previously defaulted to 14 weeks in qtr
                'Weeks
                For illoop = 1 To ilPeriods Step 1
                    slStr = slStr & "," & Format(lmStartDates(illoop), "mm/dd")
                Next illoop
                Print #hmExport, slStr
            Else
                If ilTemp = 14 Then       'send to crystal whether a 13 or 14 week quarter
                    If Not gSetFormula("WeeksInQtr", "14") Then
                        ilRet = btrClose(hmCff)
                        ilRet = btrClose(hmClf)
                        ilRet = btrClose(hmGrf)
                        ilRet = btrClose(hmCHF)
                        btrDestroy hmCff
                        btrDestroy hmClf
                        btrDestroy hmGrf
                        btrDestroy hmCHF
                        Exit Sub
                    End If
                Else
                    If RptSelCt!rbcOutput(4).Value = False Then
                        If Not gSetFormula("WeeksInQtr", "13") Then
                            ilRet = btrClose(hmCff)
                            ilRet = btrClose(hmClf)
                            ilRet = btrClose(hmGrf)
                            ilRet = btrClose(hmCHF)
                            btrDestroy hmCff
                            btrDestroy hmClf
                            btrDestroy hmGrf
                            btrDestroy hmCHF
                            Exit Sub
                        End If
                    End If
                    ilPeriods = 13          'previously defaulted to 14 weeks in qtr
                End If
            End If
        Else                                    '9-28-11 Month option, generate 12 months
            ilPeriods = Val(RptSelCt!edcSelCFrom1.Text)  '# months to show
            slDate = str$(igMonthOrQtr) & "/15/" & str$(igYear)
            slDate = gObtainStartStd(slDate)
            lmStartDates(1) = gDateValue(slDate)
            For illoop = 1 To ilPeriods + 1 Step 1
                slDate = Format(lmStartDates(illoop), "m/d/yy")
                slDate = gObtainEndStd(slDate)      'get the end of std bdcst month to calc each following months start date
                lmStartDates(illoop + 1) = gDateValue(slDate) + 1
            Next illoop
            
            slStartDate = Format$(lmStartDates(1), "m/d/yy")
            slEndDate = Format(lmStartDates(ilPeriods + 1) - 1, "m/d/yy")
            
            'TTP 10119 - Average 30 Rate Report - add option to export to CSV
            If blExport = True Then
                'Generate Header
                If RptSelCt.rbcSelC7(0) Then 'Daypart
                    slStr = "Daypart,Vehicle,Advertiser,Product,Contract#,Ln,Len,30 "" Unit Rate,30 "" Unit Spots,Total,# Spots"
                End If
                If RptSelCt.rbcSelC7(1) Then 'Daypart w/Overrides
                    slStr = "Daypart w/Overrides,Vehicle,Advertiser,Product,Contract#,Ln,Len,30 "" Unit Rate,30 "" Unit Spots,Total,# Spots"
                End If
                If RptSelCt.rbcSelC7(2) Then 'Agency
                    slStr = "Agency,Vehicle,Advertiser,Product,Contract#,Ln,Len,30 "" Unit Rate,30 "" Unit Spots,Total,# Spots"
                End If
                
                'months
                For illoop = igMonthOrQtr To ilPeriods + igMonthOrQtr - 1 Step 1
                    If illoop > 12 Then
                        slDate = MonthName(illoop - 12, True)
                    Else
                        slDate = MonthName(illoop, True)
                    End If
                    slStr = slStr & "," & slDate
                Next illoop
                Print #hmExport, slStr

            Else
                'slStr = Trim$(Str(ilPeriods))           '4-6-12 show only # periods requested
                If Not gSetFormula("WeeksInQtr", "12") Then
                    ilRet = btrClose(hmCff)
                    ilRet = btrClose(hmClf)
                    ilRet = btrClose(hmGrf)
                    ilRet = btrClose(hmCHF)
                    btrDestroy hmCff
                    btrDestroy hmClf
                    btrDestroy hmGrf
                    btrDestroy hmCHF
                    Exit Sub
                End If
            End If
        End If
    ElseIf ilListIndex = CNT_AVG_PRICES Then
        slDate = RptSelCt!CSI_CalFrom.Text      'Date: 11/9/2019 using CSI calendar for date enry --> edcSelCFrom.Text
        'always object 15 start periods (regardless if weekly or monthly), just to be consistent in subroutines
        If RptSelCt!rbcSelCSelect(0).Value Then  'set last sunday of first week
            lmStartDates(1) = gDateValue(slDate)
            For illoop = 2 To 15
                lmStartDates(illoop) = lmStartDates(illoop - 1) + 7
            Next illoop
            '4-10-10 obtain contracts based on the earliest/latest date span user requested
            slStartDate = Format$(lmStartDates(1), "m/d/yy")
            slEndDate = Format(lmStartDates(15) - 1, "m/d/yy")
        ElseIf RptSelCt!rbcSelCSelect(1).Value Then  'set last date of 12 standard periods
            slDate = gObtainStartStd(slDate)
            lmStartDates(1) = gDateValue(slDate)
            For illoop = 1 To 14 Step 1
                slDate = Format(lmStartDates(illoop), "m/d/yy")
                slDate = gObtainEndStd(slDate)      'get the end of std bdcst month to calc each following months start date
                lmStartDates(illoop + 1) = gDateValue(slDate) + 1
            Next illoop
            '4-10-10 obtain contracts based on the earliest/latest date span user requested
            slStartDate = Format$(lmStartDates(1), "m/d/yy")
            slEndDate = Format(lmStartDates(14) - 1, "m/d/yy")
        End If

        'build array of selling office codes and their sales sources.  This is the most major sort
        'in the Business Booked reports
        ilTemp = 0
        ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            ReDim Preserve tlSofList(0 To ilTemp) As SOFLIST
            tlSofList(ilTemp).iSofCode = tmSof.iCode
            tlSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
            ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            ilTemp = ilTemp + 1
        Loop
    
        'Date: 11/11/2019 Test for each option to set option flag for major/minor option
        imSlsp = False: imAdvt = False: imAgency = False: imBusCat = False: imProdProt = False: imVehicle = False: imVehGroup = False
        'Major sort
        If RptSelCt!cbcSet1.ListIndex = 4 Then              'sales people
            imSlsp = True
            'get the sales people codes selected
            gObtainCodesForMultipleLists 2, tgSalesperson(), imIncludeCodes, imUseCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet1.ListIndex = 0 Then          'advt
            imAdvt = True
            'get the Advt codes selected                         '6-21-18
            gObtainCodesForMultipleLists 5, tgAdvertiser(), imInclAdvtCodes, imUseAdvtCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet1.ListIndex = 1 Then          'agency
            imAgency = True
            'get the Agency codes selected
            gObtainCodesForMultipleLists 1, tgAgency(), imInclAgfCodes, imUseAgfCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet1.ListIndex = 2 Then          'bus cat
            imBusCat = True
            'get the bus cat codes selected
            gObtainCodesForMultipleLists 3, tgMNFCodeRpt(), imInclCatCodes, imUseCatCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet1.ListIndex = 3 Then          'prod prot
            imProdProt = True
            'get the prod prot codes selected
            gObtainCodesForMultipleLists 7, tgMnfCodeCT(), imInclProdCodes, imUseProdCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet1.ListIndex = 5 Then          'vehicle
            imVehicle = True
            'Date: 9/25/2019 build array of vehicles to include or exclude
            gObtainCodesForMultipleLists 6, tgVehicle(), imInclVefCodes, imUsevefcodes(), RptSelCt
            gAddDormantVefToExclList imInclVefCodes, tgMVef(), imUsevefcodes()          '8-4-17 if excluding vehicles, make sure dormant ones exluded since
        ElseIf RptSelCt!cbcSet1.ListIndex = 6 Then          'vehicle group
            imVehGroup = True
        End If
    
        'Minor sort
        imVehGroupMinor = False: imVehicleMinor = False: imProdProtMinor = False: imBusCatMinor = False: imAgencyMinor = False: imAdvtMinor = False: imSlspMinor = False
        If RptSelCt!cbcSet2.ListIndex = 5 Then              'sales people
            imSlspMinor = True
            'get the sales people codes selected
            gObtainCodesForMultipleLists 2, tgSalesperson(), imIncludeCodes, imUseCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet2.ListIndex = 1 Then          'advt
            imAdvtMinor = True
            'get the Advt codes selected                         '6-21-18
            gObtainCodesForMultipleLists 5, tgAdvertiser(), imInclAdvtCodes, imUseAdvtCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet2.ListIndex = 2 Then          'agency
            imAgencyMinor = True
            'get the Agency codes selected
            gObtainCodesForMultipleLists 1, tgAgency(), imInclAgfCodes, imUseAgfCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet2.ListIndex = 3 Then          'bus cat
            imBusCatMinor = True
            'get the bus cat codes selected
            gObtainCodesForMultipleLists 3, tgMNFCodeRpt(), imInclCatCodes, imUseCatCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet2.ListIndex = 4 Then          'prod prot
            imProdProtMinor = True
            'get the prod prot codes selected
            gObtainCodesForMultipleLists 7, tgMnfCodeCT(), imInclProdCodes, imUseProdCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet2.ListIndex = 6 Then          'vehicle
            imVehicleMinor = True
            'Date: 9/25/2019 build array of vehicles to include or exclude
            gObtainCodesForMultipleLists 6, tgVehicle(), imInclVefCodes, imUsevefcodes(), RptSelCt
            gAddDormantVefToExclList imInclVefCodes, tgMVef(), imUsevefcodes()          '8-4-17 if excluding vehicles, make sure dormant ones exluded since
        ElseIf RptSelCt!cbcSet2.ListIndex = 7 Then          'vehicle group
            imVehGroupMinor = True
        End If
    ElseIf ilListIndex = CNT_ADVT_UNITS Then            '1-7-09 add Advt units ordered as prepass
        slDate = RptSelCt!CSI_CalFrom.Text              'Date: 9/24/2019 using CSI calendar for date entry  --> edcSelCFrom.Text
        'always object 15 start periods (regardless if weekly or monthly), just to be consistent in subroutines
        lmStartDates(1) = gDateValue(slDate)
        For illoop = 2 To 15
            lmStartDates(illoop) = lmStartDates(illoop - 1) + 7
        Next illoop
        '9-23-09 gross, net, tnet option
'        If RptSelCt!rbcSelC9(1).Value = True Then
'            ilGrossNetTNet = 1
'        ElseIf RptSelCt!rbcSelC9(2).Value = True Then
'            ilGrossNetTNet = 2
'        End If
        
        'build array of selling office codes and their sales sources.  This is the most major sort
        'in the Business Booked reports
        ilTemp = 0
        ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            ReDim Preserve tlSofList(0 To ilTemp) As SOFLIST
            tlSofList(ilTemp).iSofCode = tmSof.iCode
            tlSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
            ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            ilTemp = ilTemp + 1
        Loop
        
        '4-19-10 obtain contracts based on the earliest/latest date span user requested
        slStartDate = Format(lmStartDates(1), "m/d/yy")
        ilPeriods = Val(RptSelCt!edcSelCFrom1.Text)
        slEndDate = Format$(lmStartDates(ilPeriods + 1) - 1, "m/d/yy")
        
        'Date: 10/2/2019 Test for each option to set option flag for major/minor option
        imSlsp = False: imAdvt = False: imAgency = False: imBusCat = False: imProdProt = False: imVehicle = False: imVehGroup = False
        'Major sort
        If RptSelCt!cbcSet1.ListIndex = 4 Then              'sales people
            imSlsp = True
            'get the sales people codes selected
            gObtainCodesForMultipleLists 2, tgSalesperson(), imIncludeCodes, imUseCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet1.ListIndex = 0 Then          'advt
            imAdvt = True
            'get the Advt codes selected                         '6-21-18
            gObtainCodesForMultipleLists 5, tgAdvertiser(), imInclAdvtCodes, imUseAdvtCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet1.ListIndex = 1 Then          'agency
            imAgency = True
            'get the Agency codes selected
            gObtainCodesForMultipleLists 1, tgAgency(), imInclAgfCodes, imUseAgfCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet1.ListIndex = 2 Then          'bus cat
            imBusCat = True
            'get the bus cat codes selected
            gObtainCodesForMultipleLists 3, tgMNFCodeRpt(), imInclCatCodes, imUseCatCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet1.ListIndex = 3 Then          'prod prot
            imProdProt = True
            'get the prod prot codes selected
            gObtainCodesForMultipleLists 7, tgMnfCodeCT(), imInclProdCodes, imUseProdCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet1.ListIndex = 5 Then          'vehicle
            imVehicle = True
            'Date: 9/25/2019 build array of vehicles to include or exclude
            gObtainCodesForMultipleLists 6, tgVehicle(), imInclVefCodes, imUsevefcodes(), RptSelCt
            gAddDormantVefToExclList imInclVefCodes, tgMVef(), imUsevefcodes()          '8-4-17 if excluding vehicles, make sure dormant ones exluded since
        ElseIf RptSelCt!cbcSet1.ListIndex = 6 Then          'vehicle group
            imVehGroup = True
        End If
    
        'Minor sort
        imVehGroupMinor = False: imVehicleMinor = False: imProdProtMinor = False: imBusCatMinor = False: imAgencyMinor = False: imAdvtMinor = False: imSlspMinor = False
        If RptSelCt!cbcSet2.ListIndex = 5 Then              'sales people
            imSlspMinor = True
            'get the sales people codes selected
            gObtainCodesForMultipleLists 2, tgSalesperson(), imIncludeCodes, imUseCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet2.ListIndex = 1 Then          'advt
            imAdvtMinor = True
            'get the Advt codes selected                         '6-21-18
            gObtainCodesForMultipleLists 5, tgAdvertiser(), imInclAdvtCodes, imUseAdvtCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet2.ListIndex = 2 Then          'agency
            imAgencyMinor = True
            'get the Agency codes selected
            gObtainCodesForMultipleLists 1, tgAgency(), imInclAgfCodes, imUseAgfCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet2.ListIndex = 3 Then          'bus cat
            imBusCatMinor = True
            'get the bus cat codes selected
            gObtainCodesForMultipleLists 3, tgMNFCodeRpt(), imInclCatCodes, imUseCatCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet2.ListIndex = 4 Then          'prod prot
            imProdProtMinor = True
            'get the prod prot codes selected
            gObtainCodesForMultipleLists 7, tgMnfCodeCT(), imInclProdCodes, imUseProdCodes(), RptSelCt
        ElseIf RptSelCt!cbcSet2.ListIndex = 6 Then          'vehicle
            imVehicleMinor = True
            'Date: 9/25/2019 build array of vehicles to include or exclude
            gObtainCodesForMultipleLists 6, tgVehicle(), imInclVefCodes, imUsevefcodes(), RptSelCt
            gAddDormantVefToExclList imInclVefCodes, tgMVef(), imUsevefcodes()          '8-4-17 if excluding vehicles, make sure dormant ones exluded since
        ElseIf RptSelCt!cbcSet2.ListIndex = 7 Then          'vehicle group
            imVehGroupMinor = True
        End If
    End If

    tlCntTypes.iHold = gSetCheck(RptSelCt!ckcSelC3(0).Value)
    tlCntTypes.iOrder = gSetCheck(RptSelCt!ckcSelC3(1).Value)
    tlCntTypes.iStandard = gSetCheck(RptSelCt!ckcSelC5(0).Value)
    tlCntTypes.iReserv = gSetCheck(RptSelCt!ckcSelC5(1).Value)
    tlCntTypes.iRemnant = gSetCheck(RptSelCt!ckcSelC5(2).Value)
    tlCntTypes.iDR = gSetCheck(RptSelCt!ckcSelC5(3).Value)
    tlCntTypes.iPI = gSetCheck(RptSelCt!ckcSelC5(4).Value)
    tlCntTypes.iPSA = gSetCheck(RptSelCt!ckcSelC5(5).Value)
    tlCntTypes.iPromo = gSetCheck(RptSelCt!ckcSelC5(6).Value)
    tlCntTypes.iTrade = gSetCheck(RptSelCt!ckcSelC6(0).Value)
    tlCntTypes.iNC = gSetCheck(RptSelCt!ckcSelC6(2).Value)
    tlCntTypes.iAirTime = gSetCheck(RptSelCt!ckcSelC10(0).Value)        '6-15-11 option to include sched spot lines vs rep lines
    tlCntTypes.iRep = gSetCheck(RptSelCt!ckcSelC10(1).Value)            '6-15-11 option to include rep lines vs sched spot lines

    imInclNonPolit = gSetCheck(RptSelCt!ckcSelC12(1).Value)
    imInclPolit = gSetCheck(RptSelCt!ckcSelC12(0).Value)

    slCntrTypes = gBuildCntTypes()      'Setup valid types of contracts to obtain based on user

    illoop = RptSelCt!cbcSet1.ListIndex     '6-13-02
'Date: 9/25/2019 modified to accomode major/minor sort for Advertiser Units report
'    ilMajorSet = gFindVehGroupInx(ilLoop, tgVehicleSets1())
    'Date: 11/9/2019 added Major/Minor sorts for AVG_PRICES
    If (ilListIndex = CNT_ADVT_UNITS) Or (ilListIndex = CNT_AVG_PRICES) Then
        If illoop = 6 Then
            illoop = RptSelCt!lbcSelection(12).ListIndex
            ilMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())
        Else
            ilMajorSet = illoop
        End If
    
        illoop = RptSelCt!cbcSet2.ListIndex     '6-13-02
        If RptSelCt!cbcSet2.ListIndex = 7 Then
            'ilLoop = RptSelCt!lbcSelection(4).ListIndex
            illoop = RptSelCt!lbcSelection(12).ListIndex            '3-18-16 chged to single selection box
            ilMinorSet = tgVehicleSets1(illoop).iCode
        Else
            ilMinorSet = illoop
        End If
    Else
        ilMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())
    End If

    slCntrStatus = ""
    If tlCntTypes.iHold Then
        slCntrStatus = "HG"             'sch holds & uns holds
    End If
    If tlCntTypes.iOrder Then
        slCntrStatus = slCntrStatus & "ON"  'sch orders & sch orders
    End If
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)
    sgCntrForDateStamp = ""             'init the time stamp to read in contracts upon re-entry

'Date: 9/25/2019 modified to accomode major/minor sort for Advertiser Units report
'    If RptSelCt!rbcSelCInclude(0).Value Then        'slsp
'        gObtainCodesForMultipleLists 2, tgSalesperson(), imIncludeCodes, imUseCodes(), RptSelCt
'    Else
'        gObtainCodesForMultipleLists 6, tgCSVNameCode(), imIncludeCodes, imUseCodes(), RptSelCt
'        gObtainAgyAdvCodes ilIncludeAgyCodes, ilUseAgyCodes(), 1, RptSelCt
'    End If

    'Date: 11/9/2019 added Major/Minor sorts for AVG_PRICES
    'If ilListIndex <> CNT_ADVT_UNITS Then
    If (ilListIndex <> CNT_ADVT_UNITS) And (ilListIndex <> CNT_AVG_PRICES) Then
        If RptSelCt!rbcSelCInclude(0).Value Then        'slsp
            gObtainCodesForMultipleLists 2, tgSalesperson(), imIncludeCodes, imUseCodes(), RptSelCt
        Else
            gObtainCodesForMultipleLists 6, tgCSVNameCode(), imIncludeCodes, imUseCodes(), RptSelCt
            gObtainAgyAdvCodes ilIncludeAgyCodes, ilUseAgyCodes(), 1, RptSelCt
        End If
    End If
    
    'Build array of possible contracts that fall into last year or this years quarter and build into array tlChfAdvtExt

    lmSingleCntr = 0                        '11-27099
    If RptSelCt!edcText.Text <> "" Then
        ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
        lmSingleCntr = Val(RptSelCt!edcText.Text)
        tmChfSrchKey1.lCntrNo = lmSingleCntr
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)

        If lmSingleCntr = tmChf.lCntrNo Then
            tlChfAdvtExt(0).lCode = tmChf.lCode
            ReDim Preserve tlChfAdvtExt(0 To 1) As CHFADVTEXT
        End If
    Else
        ilRet = gObtainCntrForDate(RptSelCt, slStartDate, slEndDate, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
    End If
    
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
        'project the $
        llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT())   'get the latest version of this contract
        ilFound = True
        
        'Date: 10/5/2020 - TTP 9984, Report filters ignored by when imVehicle or imVehGroup sorting.  Moved the following "11/9/2019" block to here, above the user report Filters
        'Date: 11/9/2019 added Major/Minor sorts to AVG_PRICES
        If (ilListIndex = CNT_ADVT_UNITS) Or (ilListIndex = CNT_AVG_PRICES) Then                '6-21-18
'            If (gTestIncludeExclude(tgChfCT.iAdfCode, imInclAdvtCodes, imUseAdvtCodes())) = False Then      'advertisers
'                ilFound = False
'            End If
            'major sort
            If imAdvt = True Then             'advertiser selectivity
                ilFound = gFilterLists(tlChfAdvtExt(ilCurrentRecd).iAdfCode, imInclAdvtCodes, imUseAdvtCodes())
            ElseIf imAgency = True Then      'agency selectivity
                ilFound = gFilterLists(tlChfAdvtExt(ilCurrentRecd).iAgfCode, imInclAgfCodes, imUseAgfCodes())
            ElseIf imBusCat = True Then      'business category selectivity
                ilFound = gFilterLists(tgChfCT.iMnfBus, imInclCatCodes, imUseCatCodes())
            ElseIf imProdProt = True Then    'product protection selectivity
                ilFound = gFilterLists(tgChfCT.iMnfComp(0), imInclProdCodes, imUseProdCodes())
            ElseIf imSlsp = True Then        'sales people selectivity
                ilFound = gFilterLists(tlChfAdvtExt(ilCurrentRecd).iSlfCode(0), imIncludeCodes, imUseCodes())
            ElseIf imVehicle = True Then     'vehicle selectivity
                ilFound = True
            ElseIf imVehGroup = True Then     'vehicle group selectivity
                ilFound = True
            End If
            
            'major sort filter was set, now set minor sort filter if minor set was selected
            If ilFound And ilMinorSet > 0 Then
                If imAdvtMinor = True Then             'advertiser selectivity
                    ilFound = gFilterLists(tlChfAdvtExt(ilCurrentRecd).iAdfCode, imInclAdvtCodes, imUseAdvtCodes())
                ElseIf imAgencyMinor = True Then      'agency selectivity
                    ilFound = gFilterLists(tlChfAdvtExt(ilCurrentRecd).iAgfCode, imInclAgfCodes, imUseAgfCodes())
                ElseIf imBusCatMinor = True Then      'business category selectivity
                    ilFound = gFilterLists(tgChfCT.iMnfBus, imInclCatCodes, imUseCatCodes())
                ElseIf imProdProtMinor = True Then    'product protection selectivity
                    ilFound = gFilterLists(tgChfCT.iMnfComp(0), imInclProdCodes, imUseProdCodes())
                ElseIf imSlspMinor = True Then        'sales people selectivity
                    ilFound = gFilterLists(tlChfAdvtExt(ilCurrentRecd).iSlfCode(0), imIncludeCodes, imUseCodes())
                ElseIf imVehicleMinor = True Then     'vehicle selectivity
                    ilFound = True
                ElseIf imVehGroupMinor = True Then     'vehicle group selectivity
                    ilFound = True
                End If
            End If
        End If
        
        'gObtainCntrForDate has only returned the contract types that user is allowed to see.
        'Now determine if user has excluded from that.
        If tgChfCT.sStatus = "H" Or tgChfCT.sStatus = "G" Then
            If Not tlCntTypes.iHold Then
                ilFound = False
            End If
        ElseIf tgChfCT.sStatus = "N" Or tgChfCT.sStatus = "O" Then
            If Not tlCntTypes.iOrder Then
                ilFound = False
            End If
        End If
        If tgChfCT.sType = "C" And Not tlCntTypes.iStandard Then
            ilFound = False
        ElseIf tgChfCT.sType = "V" And Not tlCntTypes.iReserv Then
            ilFound = False
        ElseIf tgChfCT.sType = "T" And Not tlCntTypes.iRemnant Then
            ilFound = False
        ElseIf tgChfCT.sType = "R" And Not tlCntTypes.iDR Then
            ilFound = False
        ElseIf tgChfCT.sType = "Q" And Not tlCntTypes.iPI Then
            ilFound = False
        End If
        If tgChfCT.iPctTrade = 100 And Not tlCntTypes.iTrade Then       'only exclude trade if 100%
            ilFound = False
        End If
                
        
        '12-9-16 setup selective agencies
        If (ilListIndex = CNT_AVGRATE) And (RptSelCt!rbcSelC7(2).Value) Then                'agency option for avg rate
            If tgChfCT.iAgfCode = 0 Then
                ilAgfCode = -tgChfCT.iAdfCode
            Else
                ilAgfCode = tgChfCT.iAgfCode
            End If
            If Not gFilterLists(ilAgfCode, ilIncludeAgyCodes, ilUseAgyCodes()) Then
                ilFound = False
            End If
        End If

        '12-26-08 Implement inclusion of politicals
        If mCheckAdvPolitical(tgChfCT.iAdfCode) Then          'its a political, include this contract?
             If Not imInclPolit Then
                ilFound = False
            End If
        Else                                                'not a political advt, include this contract?
             If Not imInclNonPolit Then
                ilFound = False
            End If
        End If

        'test sales source selectivity for Avg Spot Price report
'        ilSSMnfCode = 0          'if Average rate, it doesnt use the Sales Source
        
        'Date: 11/9/2019 Commented out; implemented drop down list for Major sort selection
'         If ilListIndex = CNT_AVG_PRICES Then
'            ilOk = False
'            ilSlfInx = gBinarySearchSlf(tgChfCT.iSlfCode(0)) 'return index to salesp record to get selling office
'
'            For ilTemp = LBound(tlSofList) To UBound(tlSofList)
'                If tlSofList(ilTemp).iSofCode = tgMSlf(ilSlfInx).iSofCode Then     'is the sales of this cnt equal to a sales office built in list?
'                    For ilLoop = 0 To RptSelCt!lbcSelection(3).ListCount - 1
'                        If RptSelCt!lbcSelection(3).Selected(ilLoop) Then
'                            slNameCode = tgMnfCodeCT(ilLoop).sKey         'sales source code
'                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                            If Val(slCode) = tlSofList(ilTemp).iMnfSSCode Then      'sales source of this contract match one selected?
'                                ilSSMnfCode = Val(slCode)
'                                ilOk = True
'                                Exit For
'                            End If
'                        End If
'                    Next ilLoop
'                End If
'            Next ilTemp
'            If Not ilOk Then
'                ilFound = False
'            End If
'        End If

        If ilFound Then
            'Loop thru all lines and project their $ from the flights
            For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                'ReDim tmFlight(1 To 1) As FLIGHT            'init the flight data
                ReDim tmFlight(0 To 0) As FLIGHT            'init the flight data
                tmClf = tgClfCT(ilClf).ClfRec
                If tmClf.sType = "S" Or tmClf.sType = "H" Then     'get standard and hidden lines only
                    'Determine spot counts for vehicle selected
                    ilFoundVeh = False
                    'Date: 11/18/2019 Commented out; implemented drop down list for Major sort selection
'                    If RptSelCt!rbcSelCInclude(0).Value And ilListIndex = CNT_AVG_PRICES Then         '12-9-16 slsp option (avg spot prices)
'                        If gFilterLists(tgChfCT.iSlfCode(0), imIncludeCodes, imUseCodes()) Then
'                            ilFoundVeh = True
'                        End If
'                    Else
                    'Date: 11/9/2019 added Major/Minor sorts for AVG_PRICES
                    If (ilListIndex = CNT_ADVT_UNITS) Or (ilListIndex = CNT_AVG_PRICES) Then
                        'Date: 10/6/2019    obtain major/minor codes
                        gGetVehGrpSets tmClf.iVefCode, ilMinorSet, ilMajorSet, ilmnfMinorCode, ilMnfMajorCode
                        'major sort
                        If imVehicle Then
                            'setup selective vehicles
                            ilFoundVeh = gFilterLists(tmClf.iVefCode, imInclVefCodes, imUsevefcodes())
                        ElseIf imAdvt Then
                            'setup selective advertisers
                            ilFoundVeh = gFilterLists(tlChfAdvtExt(ilCurrentRecd).iAdfCode, imInclAdvtCodes, imUseAdvtCodes())
                        ElseIf imAgency Then
                            'setup selective agency
                            ilFoundVeh = gFilterLists(tlChfAdvtExt(ilCurrentRecd).iAgfCode, imInclAgfCodes, imUseAgfCodes())
                        ElseIf imBusCat Then
                            'setup selective bus cat
                            ilFoundVeh = gFilterLists(tgChfCT.iMnfBus, imInclCatCodes, imUseCatCodes())
                        ElseIf imProdProt Then
                            'setup selective bus cat
                            ilFoundVeh = gFilterLists(tgChfCT.iMnfComp(0), imInclProdCodes, imUseProdCodes())
                        ElseIf imSlsp Then
                            'setup selective sales people
                            ilFoundVeh = gFilterLists(tgChfCT.iSlfCode(0), imIncludeCodes, imUseCodes())
                        ElseIf imVehGroup Then
                            'setup selective vehicle group
                            ilFoundVeh = mCheckForValidVGItem(ilmnfMinorCode, ilMnfMajorCode)
                        End If
                        'minor sort
                        If ilFoundVeh Then  'And ilMinorSet > 0 Then
                            If imVehicleMinor Then
                                'setup selective vehicles
                                ilFoundVeh = gFilterLists(tmClf.iVefCode, imInclVefCodes, imUsevefcodes())
                            ElseIf imAdvtMinor Then
                                'setup selective advertisers
                                ilFoundVeh = gFilterLists(tlChfAdvtExt(ilCurrentRecd).iAdfCode, imInclAdvtCodes, imUseAdvtCodes())
                            ElseIf imAgencyMinor Then
                                'setup selective agency
                                ilFoundVeh = gFilterLists(tlChfAdvtExt(ilCurrentRecd).iAgfCode, imInclAgfCodes, imUseAgfCodes())
                            ElseIf imBusCatMinor Then
                                'setup selective bus cat
                                ilFoundVeh = gFilterLists(tgChfCT.iMnfBus, imInclCatCodes, imUseCatCodes())
                            ElseIf imProdProtMinor Then
                                'setup selective bus cat
                                ilFoundVeh = gFilterLists(tgChfCT.iMnfComp(0), imInclProdCodes, imUseProdCodes())
                            ElseIf imSlspMinor Then
                                'setup selective sales people
                                ilFoundVeh = gFilterLists(tgChfCT.iSlfCode(0), imIncludeCodes, imUseCodes())
                            ElseIf imVehGroupMinor Then
                                'setup selective vehicle group
                                ilFoundVeh = mCheckForValidVGItem(ilmnfMinorCode, ilMnfMajorCode)
                            End If
                        End If
                    Else
                        'setup selective vehicles
                        ilFoundVeh = gFilterLists(tmClf.iVefCode, imIncludeCodes, imUseCodes())
                    End If

                    If ilFoundVeh Then
                        'Build rate & spot counts into tlFlight structure
                        'mFltRatesSpots ilClf, llStartDates(), 1, 2, 2, tlCntTypes.iNC
                        mFltRatesSpots ilClf, lmStartDates(), 1, ilPeriods + 1, 2, tlCntTypes.iNC, ilGrossNetTNet    '9-23-09

                        'Line has been analyzed for spot counts, build 1 recd per unique spot rate  & line id
                        For ilLoopFlt = LBound(tmFlight) To UBound(tmFlight) - 1 Step 1
                            'Create GRF records based on unique line/rate
                            ilTotalSpots = 0
                            For illoop = 1 To ilPeriods     '9-29-11 chnged to varaible # periods for Avg rate by month, defaulted to 14 for all other options
                                ilTotalSpots = ilTotalSpots + (tmFlight(ilLoopFlt).iSpots(illoop))
                                'these buckets for spots and $ are for Avg Spot Price report
                                'if Avg RAte report, overlay some of the fields
                                'tmGrf.iPerGenl(ilLoop) = tmFlight(ilLoopFlt).iSpots(ilLoop)
                                tmGrf.iPerGenl(illoop - 1) = tmFlight(ilLoopFlt).iSpots(illoop)
                                'tmGrf.lDollars(ilLoop) = tmFlight(ilLoopFlt).lProj(ilLoop)
                                tmGrf.lDollars(illoop - 1) = tmFlight(ilLoopFlt).lProj(illoop)
                            Next illoop
                            If ilTotalSpots > 0 Then

                                tmGrf.iYear = ilYear
                                tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
                                tmGrf.iGenDate(1) = igNowDate(1)
                                gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                tmGrf.lGenTime = lgNowTime
                                tmGrf.iRdfCode = tmClf.iRdfCode         'DP code
                                tmGrf.iAdfCode = tgChfCT.iAdfCode         'advertiser code
                                tmGrf.lChfCode = tgChfCT.lCode          'contract code
                                tmGrf.iCode2 = tmClf.iLen               'spot length
                                tmGrf.iSofCode = tmClf.iLine            'Line #
                                tmGrf.iSlfCode = tgChfCT.iSlfCode(0)    'primary slsp
                                tmGrf.sGenDesc = tmFlight(ilLoopFlt).sDescription   'daypart name or daypart with override
                                
                                If RptSelCt!rbcSelC7(2).Value Then      '12-9-16 agency selection
                                    If tgChfCT.iAgfCode = 0 Then
                                        ilRet = gBinarySearchAdf(tgChfCT.iAdfCode)
                                        tmGrf.sGenDesc = tgCommAdf(ilRet).sName
                                    Else
                                        ilRet = gBinarySearchAgf(tgChfCT.iAgfCode)
                                        tmGrf.sGenDesc = tgCommAgf(ilRet).sName
                                    End If
                                Else
                                    tmGrf.sGenDesc = tmFlight(ilLoopFlt).sDescription   'daypart name or daypart with override
                                End If
                                
                                tmGrf.iVefCode = tmClf.iVefCode
                                tmGrf.lLong = tmFlight(ilLoopFlt).lActPrice
                                ' determine open/close or bot
                                tmGrf.sBktType = ""        ' blank out open close flag
                                If tgSpf.sUsingBBs = "Y" Then
                                    If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then    'use closest avail, dont know if its open or close
                                        'any avail
                                        If tmClf.iBBOpenLen > 0 Then
                                            tmGrf.sBktType = "B"        'any avail, show as B since this way of using BB wont conflict with Open/close way
                                        End If
                                    Else
                                        If tmClf.iBBOpenLen > 0 And tmClf.iBBCloseLen > 0 Then     'both open & close
                                            tmGrf.sBktType = "B"
                                        ElseIf tmClf.iBBOpenLen > 0 Then
                                            tmGrf.sBktType = "O"
                                        ElseIf tmClf.iBBCloseLen > 0 Then
                                            tmGrf.sBktType = "C"
                                        Else
                                            tmGrf.sBktType = ""
                                        End If
                                    End If
                                End If
                                'Date: 11/9/2019 commented when implementing Major/Minor sorts for AVG_PRICES
'                                If ilListIndex = CNT_AVG_PRICES Then
'                                    tmGrf.lCode4 = ilSSMnfCode            'Average spot prices cannot do a total by Vehicle group unless special code is either
'                                                            'implemented in prepass or crystal since there is a subtotal for each spot length within the vehicle or slsp/office
'                                                            'The sort will only get the last spot length for the vehicle group
'                                                            'Instead, implement option to use Sales Source as major sort
'                                ElseIf ilListIndex = CNT_AVGRATE Then
                                If ilListIndex = CNT_AVGRATE Then
                                    gGetVehGrpSets tmClf.iVefCode, ilMinorSet, ilMajorSet, ilmnfMinorCode, ilMnfMajorCode   '6-13-02
                                    tmGrf.lCode4 = ilMnfMajorCode     'store the major sort for Crystal

                                    'tmGrf.lDollars(2) = tmFlight(ilLoopFlt).lActPrice \ 100   'actual spot price
                                    tmGrf.lDollars(1) = tmFlight(ilLoopFlt).lActPrice \ 100   'actual spot price
                                    'tmGrf.iPerGenl(15) = ilTotalSpots           'total spot count for quarter
                                    tmGrf.iPerGenl(14) = ilTotalSpots           'total spot count for quarter
                                    'tmGrf.lDollars(1) = (ilTotalSpots * tmFlight(ilLoopFlt).lActPrice) \ 100  'Total $ for quarter
                                    tmGrf.lDollars(0) = (ilTotalSpots * tmFlight(ilLoopFlt).lActPrice) \ 100  'Total $ for quarter
                                ElseIf ((ilListIndex = CNT_ADVT_UNITS) Or (ilListIndex = CNT_AVG_PRICES)) Then      'adjust units ordered based on option:  counts or unit counts
                                    'Date: 11/9/2019 commented out when implementing Major/Minor sorts for AVG_PRICES
'                                    If RptSelCt!rbcSelC4(1).Value = True Then   'do by 30" unit counts (vs spot counts which doesnt get adjusted)
'                                        'anything 30" or less is considered a unit
'                                        'anything greater will be increments of 30s (45" = 2 30s)
'                                        'each spot may represent more than 1 unit
'                                        ilUnits = tmFlight(ilLoopFlt).iLen \ 30
'                                        If tmFlight(ilLoopFlt).iLen Mod 30 <> 0 Then
'                                            ilUnits = ilUnits + 1
'                                        End If
'                                    Else
'                                        ilUnits = 1     'each spot represents 1 unit
'                                    End If
                                    If (ilListIndex = CNT_AVG_PRICES) Then
                                        ilUnits = 1     'each spot represents 1 unit
                                    Else
                                        If RptSelCt!rbcSelC4(1).Value = True Then   'do by 30" unit counts (vs spot counts which doesnt get adjusted)
                                            'anything 30" or less is considered a unit
                                            'anything greater will be increments of 30s (45" = 2 30s)
                                            'each spot may represent more than 1 unit
                                            ilUnits = tmFlight(ilLoopFlt).iLen \ 30
                                            If tmFlight(ilLoopFlt).iLen Mod 30 <> 0 Then
                                                ilUnits = ilUnits + 1
                                            End If
                                        Else
                                            ilUnits = 1     'each spot represents 1 unit
                                        End If
                                    End If
                                    For illoop = 1 To 14
                                        'tmGrf.iPerGenl(ilLoop) = tmFlight(ilLoopFlt).iSpots(ilLoop) * ilUnits
                                        tmGrf.iPerGenl(illoop - 1) = tmFlight(ilLoopFlt).iSpots(illoop) * ilUnits
                                    Next illoop
                                    
                                    'Date: 9/25/2019 update GRF with agent code, Cast/trade, major/minor sort codes
                                    'agency code
                                    tmGrf.iPerGenl(14) = tgChfCT.iAgfCode
                                    'Cash or Trade
                                    tmGrf.sDateType = IIF(tgChfCT.iPctTrade = 100, "T", "C")
                                    
                                    'major/minor sort codes
                                    If imVehicle Then
                                        tmGrf.lCode4 = tmClf.iVefCode
                                    ElseIf imAdvt Then
                                        tmGrf.lCode4 = tlChfAdvtExt(ilCurrentRecd).iAdfCode
                                    ElseIf imAgency Then
                                        tmGrf.lCode4 = tlChfAdvtExt(ilCurrentRecd).iAgfCode
                                    ElseIf imBusCat Then
                                        tmGrf.lCode4 = tgChfCT.iMnfBus
                                    ElseIf imProdProt Then
                                        tmGrf.lCode4 = tgChfCT.iMnfComp(0)
                                    ElseIf imSlsp Then
                                        tmGrf.lCode4 = tgChfCT.iSlfCode(0)
                                    ElseIf imVehGroup Then
                                        tmGrf.lCode4 = ilMnfMajorCode
                                    End If
                                    
                                    'minor sort
                                    If imVehicleMinor Then
                                        tmGrf.iPerGenl(15) = tmClf.iVefCode
                                    ElseIf imAdvtMinor Then
                                        tmGrf.iPerGenl(15) = tlChfAdvtExt(ilCurrentRecd).iAdfCode
                                    ElseIf imAgencyMinor Then
                                        tmGrf.iPerGenl(15) = tlChfAdvtExt(ilCurrentRecd).iAgfCode
                                    ElseIf imBusCatMinor Then
                                        tmGrf.iPerGenl(15) = tgChfCT.iMnfBus
                                    ElseIf imProdProtMinor Then
                                        tmGrf.iPerGenl(15) = tgChfCT.iMnfComp(0)
                                    ElseIf imSlspMinor Then
                                        tmGrf.iPerGenl(15) = tgChfCT.iSlfCode(0)
                                    ElseIf imVehGroupMinor Then
                                        tmGrf.iPerGenl(15) = ilmnfMinorCode
                                    End If
                                End If
                                
                                If blExport Then
                                    smExportStatus = mExportAvg30Rate(tmGrf, ilPeriods)
                                Else
                                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                End If
                                
                                For illoop = 1 To 14        'initalize buckets
                                    'tmGrf.iPerGenl(ilLoop) = 0
                                    tmGrf.iPerGenl(illoop - 1) = 0
                                    'tmGrf.lDollars(ilLoop) = 0
                                    tmGrf.lDollars(illoop - 1) = 0
                                Next illoop

                            End If
                        Next ilLoopFlt
                    End If
                End If                          'tmclf = H or tmclf = S
            Next ilClf                      'process nextline
        End If                              'if ilFound
    Next ilCurrentRecd
    Erase tlChfAdvtExt, tlSofList, tmFlight
    Erase imUseCodes, ilUseAgyCodes             '12-9-16
    Erase imUseVefGrpCodes                      'Date: 10/4/2019 vehicle group array
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmAgf)
    btrDestroy hmRdf
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmSof
    btrDestroy hmAgf
    
    'TTP 10119 - Average 30 Rate Report - add option to export to CSV
    If blExport = True Then
        If InStr(1, smExportStatus, "Error") > 0 Then
            RptSelCt.lacExport.Caption = "Export Failed:" & smExportStatus
        Else
            RptSelCt.lacExport.Caption = "Export Stored in- " & sgExportPath & slFileName
        End If
        Close #hmExport
    End If
End Sub
'
'           Test if vehicle groups used;  If it is, test to see if the
'           item selected matches the vehicle processing
'           return - true to process record
Private Function mCheckForValidVGItem(ByVal iMinorCode As Integer, ByVal iMajorCode As Integer) As Integer
Dim ilIncludeVehicleGroup As Integer
Dim ilLoopOnVG As Integer
Dim slStr As String
Dim ilRet As Integer
Dim ilVGListBox As Integer

    ilIncludeVehicleGroup = True
    'test for the vehicle group selectivity
    If igRptCallType = CONTRACTSJOB Then
        'Date: 11/22/2019 added check for AVG Spot Prices report
        If (RptSelCt!lbcRptType.ListIndex = CNT_ADVT_UNITS Or RptSelCt!lbcRptType.ListIndex = CNT_AVG_PRICES) Then
            ilVGListBox = 8
            If RptSelCt!lbcSelection(ilVGListBox).SelCount > 0 Then
                ilIncludeVehicleGroup = False
                For ilLoopOnVG = 0 To RptSelCt!lbcSelection(ilVGListBox).ListCount - 1 Step 1
                    If RptSelCt!lbcSelection(ilVGListBox).Selected(ilLoopOnVG) Then
                        slStr = tgMnfCodeCB(ilLoopOnVG).sKey            'sales comparison
                        ilRet = gParseItem(slStr, 2, "\", slStr)
                        'Determine which vehicle set to test and whether major or minor sort
                        If imVehGroup Then
                            'If tmGrf.lCode4 = Val(slStr) Then
                            If iMajorCode = Val(slStr) Then
                                ilIncludeVehicleGroup = True
                                Exit For
                            End If
                        Else
                            If imVehGroupMinor Then
                                'If tmGrf.iPerGenl(15) = Val(slStr) Then
                                If iMinorCode = Val(slStr) Then
                                    ilIncludeVehicleGroup = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next ilLoopOnVG
            End If
        End If
    End If
    mCheckForValidVGItem = ilIncludeVehicleGroup
End Function

'
'
'           gCrMakePlan - Prepass toCalculate Price Needed to
'                         Make Plan by Daypart for each vehicle
'                         Calculate budgets by daypart, gather
'                         inventory, spots sold by daypart .
'                         Vehicles will be compared against the
'                         selected budget (plan or forecast) for
'                         the same year's active rate card (which
'                         is also selected).  If the rate card year
'                         doesn't exist, no data is produced.  That
'                         is, if an old rate is on file, that is
'                         the effective one for contract input, but
'                         for this purpose, that year's rate card
'                         must exist on file.
'
'           4-24-01 calculate correct # weeks in a quarter, few as 12 as many as 14
'
'           Created: D Hosaka   7/11/97
Sub gCrMakePlan()
    Dim ilCalType As Integer            '1 = corp, 2 = std
    Dim ilStartMonth As Integer         'start month of first quarter (different for corp vs standard)
    Dim ilRet As Integer
    Dim ilIndex As Integer              'temp looping variable
    Dim slNameCode As String            'parsing temp string
    Dim slNameYear As String            'parsing temp string
    Dim slYear As String                'year to process budgets
    Dim slCode As String                'paring temp string
    ReDim ilBdYear(0 To 0) As Integer             'budget year for mBdGetBudgetDollars
    ReDim ilBdMnfCode(0 To 0) As Integer          'budget code for mBdGetBudgetDollars (always only 1 budget)
    Dim ilRCCode As Integer             'rc code
    Dim ilRCSelected As Integer
    Dim ilRcf As Integer                'rate card loop in list box
    Dim llRif As Long
    Dim llYearStart As Long             'years std start date
    Dim llYearEnd As Long               'years std end date
    Dim llInputStart As Long          'std start date requested
    Dim llInputEnd As Long           'std end date requested (not past end of year)
    Dim llQtrEnd As Long             '4-24-01 qtr end date
    ReDim ilStdInputStart(0 To 1) As Integer  'btrieve form of std start date requested (to put into ANR)
    Dim illoop As Integer               'temp
    Dim ilTemp As Integer               'temp
    Dim ilTempMinusOne As Integer
    Dim slStr As String
    Dim slStart As String               'start date of input
    Dim slEnd As String                 'end of std month of first month requested
    Dim llDate As Long                  'temp date
    Dim ilYear As Integer               'year entered by user
    Dim ilStartQtr As Integer
    Dim ilNoQtrs As Integer             '# qtrs to process
    'ReDim ilBdStartWks(1 To 15) As Integer  '4-24-01 (14 to 15) start week index for each period
    ReDim ilBdStartWks(0 To 15) As Integer  '4-24-01 (14 to 15) start week index for each period. Index zero ignored
                                        'i.e. for Weekly report the elements will be 1, 2, 3, etc.
                                        'for monthly, the elements will be the start week of the qtr, - 1,5,10,14,18,23...
    Dim ilLoopWks As Integer
    Dim ilVeh As Integer                '
    Dim llAvail As Long
    Dim ilProcessWk As Integer
    Dim ilStartOfPer As Integer
    Dim ilEndOfPer As Integer
    Dim ilSplitType As Integer      '0 = budget & r/c match, 1 = year based on budget with split r/c
    Dim ilAvailDefined As Integer   '0 = avail missing in period, 1 = avail defined in period
    Dim ilWk As Integer
    Dim ilMajorSet As Integer
    Dim ilMinorSet As Integer
    Dim ilmnfMinorCode As Integer
    Dim ilMnfMajorCode As Integer

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
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)
    hmBvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmBvf, "", sgDBPath & "Bvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmBvf)
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imBvfRecLen = Len(tmBvf)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imSdfRecLen = Len(hmSdf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imSmfRecLen = Len(hmSmf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imVefRecLen = Len(hmVef)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVsf)
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imVsfRecLen = Len(hmVsf)
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imSsfRecLen = Len(tmSsf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imMnfRecLen = Len(hmMnf)
    hmRcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRcf, "", sgDBPath & "Rcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRcf)
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imRcfRecLen = Len(hmRcf)
    hmRif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRif, "", sgDBPath & "Rif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRif)
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imRifRecLen = Len(hmRif)
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRdf)
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imRdfRecLen = Len(tmRdf)
    hmAnr = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnr, "", sgDBPath & "Anr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAnr)
        btrDestroy hmAnr
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imAnrRecLen = Len(tmAnr)
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmLcf)
        btrDestroy hmLcf
        btrDestroy hmAnr
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        Exit Sub
    End If
    imLcfRecLen = Len(tmLcf)
    slNameCode = tgRptSelBudgetCodeCT(igBSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 1, "\", slNameYear)
    ilRet = gParseItem(slNameYear, 1, "\", slYear)
    slYear = gSubStr("9999", slYear)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilBdMnfCode(0) = Val(slCode)            'always only 1 budget
    ilBdYear(0) = Val(slYear)               'always only 1 budget, but may have a rate card that spans fiscal year
    ilYear = Val(RptSelCt!edcSelCFrom.Text)          'retrieve year entered
    ilStartQtr = Val(RptSelCt!edcSelCFrom1.Text)           'starting qtr
    ilNoQtrs = Val(RptSelCt!edcSelCTo.Text)              '# qtrs to process
    
    illoop = RptSelCt!cbcSet1.ListIndex
    ilMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())
    
    If RptSelCt!rbcSelCSelect(0).Value Then          'corp, should have 2 rate cards selected
        ilIndex = 11
        ilCalType = 1
        ilSplitType = 1                             'corporate year, may have two rates cards
        ilRet = gGetCorpCalIndex(ilYear)
        ilStartMonth = ((ilStartQtr - 1) * 3 + tgMCof(ilRet).iStartMnthNo)
        If ilStartMonth > 12 Then
            ilStartMonth = ilStartMonth - 12
        End If
    Else
        ilIndex = 12
        ilCalType = 2
        ilSplitType = 0                             'only 1 rate card & budget is used
        ilStartMonth = (ilStartQtr - 1) * 3 + 1
    End If
    'ReDim tmMRif(1 To 1) As RIF
    ReDim tmMRif(0 To 0) As RIF
    For ilRCSelected = 0 To RptSelCt!lbcSelection(ilIndex).ListCount - 1 Step 1
        If RptSelCt!lbcSelection(ilIndex).Selected(ilRCSelected) Then            'see if rate card selected
            slNameCode = tgRateCardCode(ilRCSelected).sKey
            ilRet = gParseItem(slNameCode, 3, "\", slCode)
            ilRCCode = Val(slCode)
            For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
                If tgMRcf(ilRcf).iCode = ilRCCode Then
                    'see which rate cards are selected
                    'Build array (tmMRif) of all valid Rates for each Vehicle's daypart
                    For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                        If (ilRCCode = tgMRif(llRif).iRcfCode) And tgMRcf(ilRcf).iYear = tgMRif(llRif).iYear Then
                            'test for selective vehicle
                            For illoop = 0 To RptSelCt!lbcSelection(3).ListCount - 1 Step 1
                                If (RptSelCt!lbcSelection(3).Selected(illoop)) Then
                                    slNameCode = tgVehicle(illoop).sKey 'Traffic!lbcVehicle.List(ilVehicle)
                                    ilRet = gParseItem(slNameCode, 1, "\", slStr)
                                    ilRet = gParseItem(slStr, 3, "|", slStr)
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    If Val(slCode) = tgMRif(llRif).iVefCode Then
                                        tmMRif(UBound(tmMRif)) = tgMRif(llRif)
                                        'ReDim Preserve tmMRif(1 To UBound(tmMRif) + 1) As RIF
                                        ReDim Preserve tmMRif(0 To UBound(tmMRif) + 1) As RIF
                                        illoop = RptSelCt!lbcSelection(3).ListCount 'stop the loop
                                    End If
                                End If
                            Next illoop
                        End If
                    Next llRif
                End If
            Next ilRcf
        End If
    Next ilRCSelected
    gGetStartEndYear ilCalType, ilYear, slStart, slEnd
    llYearStart = gDateValue(slStart)
    llYearEnd = gDateValue(slEnd)

    gGetStartEndQtr ilCalType, ilYear, ilStartQtr, slStart, slEnd   'obtain the start & end dates of the year for the requested starting qtr

    llInputStart = gDateValue(slStart)
    llQtrEnd = gDateValue(slEnd)            '4-24-01 qtr end date for weekly run
    slStr = Trim$(str$((((ilStartQtr - 1) * 3) + 1) + (ilNoQtrs * 3) - 1)) + "/" + "15/" + Trim$(str$(igYear))
    'calculate the end date of the period to process
    If ilCalType = 1 Then
        llInputEnd = gDateValue(gObtainEndCorp(slStr, True))  'corp
    Else
        llInputEnd = gDateValue(gObtainEndStd(slStr))         'std
    End If
    'Obtain the latest date to gather data.  Also, setup week index array of designating the start week of each period (ignore for weekly option)
    slStr = Format$(llInputStart, "m/d/yy")
    ilBdStartWks(1) = (llInputStart - llYearStart) / 7 + 1     'first week to start gathering data
    For illoop = 1 To (ilNoQtrs * 3) '+ 1  4-24-01            'loop for # months for all qtrs requested
        'ilBdStartWks is array indicating first week of the period (for weekly request, each element will be incremented by one.
        'if quarter, each element is the start of a std brdcst month
        If ilCalType = 1 Then                   'corp
            slStr = gObtainStartCorp(slStr, True)
            slEnd = gObtainEndCorp(slStr, True)
            llDate = gDateValue(slEnd) + 1      'get to next month
            slStr = Format$(llDate, "m/d/yy")
        Else
            slStr = gObtainStartStd(slStr)
            slEnd = gObtainEndStd(slStr)
            llDate = gDateValue(slEnd) + 1      'get to next month
            slStr = Format$(llDate, "m/d/yy")
        End If
        ilBdStartWks(illoop + 1) = (llDate - llYearStart) / 7 + 1   'first week to start gathering data
    Next illoop
    'convert to btrieve for Crystal
    gPackDateLong llInputStart, ilStdInputStart(0), ilStdInputStart(1)
    If RptSelCt!rbcSelC4(0).Value Then          'weekly option (vs quarters)
        ilWk = 1
        ilTemp = (llInputStart - llYearStart) / 7 + 1
        For illoop = ilTemp To 14 + ilTemp  '4-24-01
            ilBdStartWks(ilWk) = illoop     'Only accum 1 week at a time for weekly option
            ilWk = ilWk + 1
        Next illoop
    End If
    mBdGetBudgetDollars ilSplitType, hmCHF, hmClf, hmCff, hmSdf, hmSmf, hmVef, hmVsf, hmSsf, hmBvf, hmLcf, ilBdMnfCode(), ilBdYear(), llInputStart, llInputEnd, tmMRif(), tgMRdf()
    'Create ANR pre-pass from weekly vehicle dayparts (tgDollarRec)
    For ilVeh = LBound(tgImpactRec) To UBound(tgImpactRec) - 1 Step 1  'create a record for each daypart
        tmAnr = tmZeroAnr
        tmAnr.iGenDate(0) = igNowDate(0)
        tmAnr.iGenDate(1) = igNowDate(1)
        'tmAnr.iGenTime(0) = igNowTime(0)
        'tmAnr.iGenTime(1) = igNowTime(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmAnr.lGenTime = lgNowTime
        tmAnr.iVefCode = tgImpactRec(ilVeh).iVefCode           'vehicle
        tmAnr.iRdfCode = tgImpactRec(ilVeh).iRdfCode           'daypart
        tmAnr.iMnfBudget = ilBdMnfCode(0)                  'budget code
        tmAnr.iYear = ilStartMonth                          'required to determine month column headings in Crystal
        tmAnr.iEffectiveDate(0) = ilStdInputStart(0)            'start date of requested period (used for weekly hdr dates)
        tmAnr.iEffectiveDate(1) = ilStdInputStart(1)
        gGetVehGrpSets tmAnr.iVefCode, ilMinorSet, ilMajorSet, ilmnfMinorCode, ilMnfMajorCode
        tmAnr.lExtra2 = ilMnfMajorCode          'vehicle group or none (0)
        ilIndex = ilNoQtrs * 3                                  'no of periods to process if monthly
        
        'if weekly, always print out 13 weeks (vs monthly, depends on how many quarters is being processed
        If RptSelCt!rbcSelC4(0).Value Then          'weekly option (vs quarters)
            'determine weeks in the quarter
           'ilIndex = 13                            '4-24-01 force 13 periods to create
            ilIndex = ((llQtrEnd + 1) - llInputStart) / 7 '4-24-01 calc true # weeks in qtr
            tmAnr.iExtra1 = ilIndex                 'update # weeks in qtr for Crystal
        End If
        For ilTemp = 1 To ilIndex Step 1                 'Loop thru array of weeks to accumulate - each entry is a start week for the period
            ilTempMinusOne = ilTemp - 1
            llAvail = 0
            ilProcessWk = False
            'Accum the # of weeks for each period
            illoop = (llInputStart - llYearStart) / 7 + 1
            ilStartOfPer = ilBdStartWks(ilTemp) - illoop + 1
            ilEndOfPer = ilBdStartWks(ilTemp + 1) - illoop
            'Gather # of weeks in each month to create montly values
            ilAvailDefined = 0
            For ilLoopWks = ilStartOfPer To ilEndOfPer Step 1
                ilProcessWk = True                      'found a week to accumulate
                tmAnr.lBudget(ilTempMinusOne) = tmAnr.lBudget(ilTempMinusOne) + tgDollarRec(ilLoopWks, tgImpactRec(ilVeh).iPtDollarRec).lBudget         'total budget for period
                tmAnr.lSold(ilTempMinusOne) = tmAnr.lSold(ilTempMinusOne) + tgDollarRec(ilLoopWks, tgImpactRec(ilVeh).iPtDollarRec).lDollarSold         'total $ sold this period
                '4-24-01 dont count avails if oversold
                If (tgDollarRec(ilLoopWks, tgImpactRec(ilVeh).iPtDollarRec).l30Inv - tgDollarRec(ilLoopWks, tgImpactRec(ilVeh).iPtDollarRec).l30Sold) > 0 Then
                    tmAnr.lInv(ilTempMinusOne) = tmAnr.lInv(ilTempMinusOne) + (tgDollarRec(ilLoopWks, tgImpactRec(ilVeh).iPtDollarRec).l30Inv - tgDollarRec(ilLoopWks, tgImpactRec(ilVeh).iPtDollarRec).l30Sold)    'totl availabilty this period
                End If
                'If at least one week in period has avails, set it to process (for crystal)
                If (tgDollarRec(ilLoopWks, tgImpactRec(ilVeh).iPtDollarRec).iAvailDefined <> 0) Then
                    ilAvailDefined = 1
                End If
            Next ilLoopWks
            If ilProcessWk Then                 'found a week to process, see if there's any available
                If tmAnr.lInv(ilTempMinusOne) < 1 Then             'no avails-- is it because they were not defined or is truly sold out?
                    If ilAvailDefined = 0 Then      'not defined
                        'Crystal report will show blank (not sold or a % value) .
                        'If we want report to show "Undefined" for weeks without programming, set this field to some value such as 2,
                        'then in all versions of Crystal need to test for 2 and put out Undef.
                    Else
                        tmAnr.iPctSellout(ilTempMinusOne) = 1   'defined, flag to denote sold out
                    End If
                Else                            'calc price needed
                    tmAnr.lPriceNeeded(ilTempMinusOne) = (tmAnr.lBudget(ilTempMinusOne) - tmAnr.lSold(ilTempMinusOne)) / tmAnr.lInv(ilTempMinusOne)
                End If
            End If
        Next ilTemp
        ilRet = btrInsert(hmAnr, tmAnr, imAnrRecLen, INDEXKEY0)
    Next ilVeh
    Erase ilBdStartWks, tgDollarRec, tgImpactRec
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmRif)
    ilRet = btrClose(hmRcf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmBvf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmAnr)
    ilRet = btrClose(hmLcf)
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCHF
        btrDestroy hmAnr
        btrDestroy hmLcf
End Sub
'
'
'           gCrPaperWork - Create paperwork Summary prepass in GRF
'
'           Read all the contract headers and find create an entry in GRF that meets all
'           the input selectivity criteria.  If selecting by vehicle, first get the contract
'           header that meets the filter; then read the entire contract lines to create GRF
'           entries for the lines
'
'           Created DH: 7-19-01
'           5-21-02 chg to get latest revision of orders/holds (i.e. if sch & unsch, show the unsched; multiple
'           versions of proposals will show
'           12-7-04 Add option to exclude CBS contracts/lines
'           2-3-06 check contract header flag (chfcbsorder) for C for cancel before start order
'
'   GRF fields:
'   grfGenDate - generation Date for filter
'   grfGenTime - generation time for filter
'   grfDollars(1) - Internal comment code
'   grfDollars(2) -  Other Comment code
'   grfDollars(3) - Change Reason comment code
'   grfDollars(4) - cancellation comment code
'   grfDollars(5) - Total gross
'   grfDollars(6) - SBF code for NTR
'   grfDollars(7) - Tax Rate 1 for line
'   grfDollars(8) - tax Rate 2 for line
'   grfDateGenl(1) = line start date
'   grfDateGenl(2) = line end date
'   grfVefCode = line vehicle code
'   grfChfCode = Contract code
'   grfCode2 - MNF (NTR) type
'   grfperGenl(1) = sequence # by contract (NTR follows air time)
'   grfPerGenl(2) = tax auto code for line
'
Sub gCrPaperWork()
Dim ilRet As Integer
Dim slStr As String
Dim ilTemp As Integer
Dim ilError As Integer
Dim slCntrStatus As String
Dim slCntrType As String
Dim ilHOState As Integer
Dim slEarliestActive As String
Dim slLatestActive As String
Dim ilCkcAll As Integer
ReDim ilCodes(0 To 0) As Integer        'table of selective advt, agy, slsp, or vehicle codes
Dim ilCurrentRecd As Integer
Dim ilFoundOne As Integer
Dim llContrCode As Long
Dim illoop As Integer
Dim llTempDate As Long
Dim ilClf As Integer
Dim llEarliestEntry As Long
Dim llLatestEntry As Long
Dim slEarliestEntry As String
Dim slLatestEntry As String
'ReDim llProject(1 To 2) As Long
ReDim llProject(0 To 2) As Long     'Index zero ignored
'ReDim llStartDates(1 To 2) As Long        'gather all dates of each contract.  Indicate earliest and latest dates here to
ReDim llStartDates(0 To 2) As Long        'gather all dates of each contract.  Indicate earliest and latest dates here to. Index zero ignored
Dim ilByLine As Integer             'true if by line, else false
Dim ilInvSort As Integer            'invoice sort code for selection
Dim slCode As String
Dim slNameCode As String
Dim llCntStartDate As Long
Dim llCntEndDate As Long
Dim ilInclCBS As Integer
Dim ilAtLeast1Vehicle As Integer
Dim llTax1Pct As Long
Dim llTax2Pct As Long
Dim slGrossNet As String
Dim ilTrfAgyAdvt As Integer
Dim ilVefInx As Integer

    ilInvSort = 0       'init invoice sort code incase nothing selected
    If (Not RptSelCt!rbcSelCSelect(3).Value Or RptSelCt!rbcSelC9(3).Value) And (Not RptSelCt!rbcSelCInclude(1).Value) Then      'not vehicle option and not detail
        'Determine if an invoice sort selected
        illoop = RptSelCt!cbcSet1.ListIndex
        slCode = "0"                    'force None answer for invoice sort codes
        slStr = ""
        If illoop > 0 Then
            slNameCode = tgMnfCodeCT(illoop - 1).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
            ilRet = gParseItem(slNameCode, 1, "\", slNameCode)    'Get application name
            slStr = "(Invoice sort for " & slNameCode & ")"
        End If
        ilInvSort = Val(slCode)
        'Send formula to crystal for description in heading
        If Not gSetFormula("InvSortDesc", "'" & slStr & "'") Then
            On Error GoTo gCrPaperWorkErr
            gBtrvErrorMsg ilRet, "gCrPaperWorkErr (Crystal):Adf", RptSelCt
            On Error GoTo 0
            Exit Sub
        End If
    End If

    ilError = mInitPaperWork(ilCodes())   'open files, setup filter parameters, and arrays of selective advt, agy, slsp or vehicles
    If ilError <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmSbf)
        btrDestroy hmGrf
        btrDestroy hmCHF
        btrDestroy hmSlf
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmCff
        btrDestroy hmAgf
        btrDestroy hmAdf
        btrDestroy hmSbf
        On Error GoTo gCrPaperWorkErr
        gBtrvErrorMsg ilRet, "gCrPaperWorkErr (btrOpen)", RptSelCt
        On Error GoTo 0
    End If

    slCntrStatus = ""                 'statuses: hold, order, unsch hold, uns order
    If imHold Then                  'exclude holds and uns holds
        slCntrStatus = "HG"             'include orders and uns orders
    End If
    If imOrder Then                  'exclude holds and uns holds
        slCntrStatus = slCntrStatus & "ON"             'include orders and uns orders
    End If
    If imDead Then
        slCntrStatus = slCntrStatus & "D"
    End If
    If imWork Or imRev Then                             '11-7-16
        slCntrStatus = slCntrStatus & "W"
    End If
    If imIncomplete Then
        slCntrStatus = slCntrStatus & "I"
    End If
    If imComplete Then
        slCntrStatus = slCntrStatus & "C"
    End If
    slCntrType = ""
    If imStandard Then
        slCntrType = "C"
    End If
    If imReserv And tgUrf(0).sResvType <> "H" Then
        slCntrType = slCntrType & "V"
    End If
    If imRemnant And tgUrf(0).sRemType <> "H" Then
        slCntrType = slCntrType & "T"
    End If
    If imDR And tgUrf(0).sDRType <> "H" Then
        slCntrType = slCntrType & "R"
    End If
    If imPI And tgUrf(0).sPIType <> "H" Then
        slCntrType = slCntrType & "Q"
    End If
    If imPSA Then
        slCntrType = slCntrType & "S"
    End If
    If imPromo Then
        slCntrType = slCntrType & "M"
    End If
    If slCntrType = "CVTRQSM" Then          'all types: PI, DR, etc.  except PSA(p) and Promo(m)
        slCntrType = ""                     'blank out string for "All"
    End If

    'ilHOState = 2                       '5-21-02 chged from 3 to 2 to get latest orders & revisions  (HOGN plus any revised orders WCI)
                                        'dup contract pending will no longer show, but dup versions of proposals will show
'    If imHold Or imOrder Then
'        ilHOState = 2
'    Else
'        ilHOState = 4                   '11-16-06  get only the C, W, or I
'    End If

'11-3-16 Allow to see Rev Working along with the order.
'If Hold or Order selected, with Working or Complete or Incomplete, then show all (show  order along with rev working).
'Previously, if order in rev working state, the Order overrides the rev working status; unless Hold/Order deselected and Working is selected.
'If Hold or Order selected with Working or complete or incomplete, then jut show the holds/orders
'If Working, Complete or Incomplete selected without Hold or Order, then show just Working, complete, Incomplete
If (imHold Or imOrder) And (imWork Or imComplete Or imIncomplete) Then
    ilHOState = 4
Else
    If (imHold Or imOrder) Then
        ilHOState = 2
    Else
        ilHOState = 4                   '11-16-06  get only the C, W, or I
    End If
End If
    If RptSelCt!ckcAll.Value = vbChecked Then     'all advertisrs selected?
        ilCkcAll = True
    Else
        ilCkcAll = False
    End If
    If RptSelCt!rbcSelCInclude(0).Value Then    'by Contract
        ilByLine = False
    Else
        ilByLine = True
    End If

    If RptSelCt!ckcSelC6(4).Value = vbChecked Then    'include CBS
        ilInclCBS = True
    Else
        ilInclCBS = False
    End If

    imNTR = False
    If RptSelCt!ckcSelC5(7).Value = vbChecked Then      'Include NTR info with slsp commissions and tax info
        imNTR = True
    End If

     ilRet = gObtainTrf()

    'Setup up earliest and latest active dates
    slEarliestActive = RptSelCt!CSI_CalFrom.Text        'Date: 12/19/2019 added CSI calendar control for date entries --> edcSelCFrom.Text   'Earliest
    If slEarliestActive = "" Then
        slEarliestActive = "1/5/1970" 'Monday
    Else
        llTempDate = gDateValue(slEarliestActive)    'insure year attached
        slEarliestActive = Format$(llTempDate, "m/d/yy")
    End If
    llStartDates(1) = gDateValue(slEarliestActive)      'setup date for flight testing
    slLatestActive = RptSelCt!CSI_CalTo.Text            'Date: 12/19/2019 added CSI calendar control for date entries -->  edcSelCFrom1.Text   'End date

    If (StrComp(slLatestActive, "TFN", 1) = 0) Or (Len(slLatestActive) = 0) Then
        slLatestActive = "12/29/2069"    'Sunday
    Else
        llTempDate = gDateValue(slLatestActive)   'insure that theres a year attached
        slLatestActive = Format$(llTempDate, "m/d/yy")
    End If
    llStartDates(2) = gDateValue(slLatestActive)        'setup date for flight testing
    'Setup up earliest and latest entered dates
    slEarliestEntry = RptSelCt!CSI_From1.Text           'Date: 12/19/2019 added CSI calendar control for date entries -->  edcSelCTo.Text   'Earliest
    If slEarliestEntry = "" Then
        slEarliestEntry = "1/5/1970" 'Monday
    Else
        llTempDate = gDateValue(slEarliestEntry)    'insure year attached
        slEarliestEntry = Format$(llTempDate, "m/d/yy")
    End If
    llEarliestEntry = gDateValue(slEarliestEntry)
    slLatestEntry = RptSelCt!CSI_To1.Text               'Date: 12/19/2019 added CSI calendar control for date entries -->  edcSelCTo1.Text   'End date

    If (StrComp(slLatestEntry, "TFN", 1) = 0) Or (Len(slLatestEntry) = 0) Then
        slLatestEntry = "12/29/2069"    'Sunday
    Else
        llTempDate = gDateValue(slLatestEntry)   'insure that theres a year attached
        slLatestEntry = Format$(llTempDate, "m/d/yy")
    End If
    llLatestEntry = gDateValue(slLatestEntry)
    ilRet = gObtainCntrForDate(RptSelCt, slEarliestActive, slLatestActive, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())

    'common GRF fields that wont change
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)
    'tmGrf.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
        If Not ilCkcAll Then
            ilFoundOne = False
            If imAdvt Then
                For illoop = 0 To UBound(ilCodes) - 1
                    If tlChfAdvtExt(ilCurrentRecd).iAdfCode = ilCodes(illoop) Then
                        ilFoundOne = True
                        Exit For
                    End If
                Next illoop
            ElseIf imAgency Then
                For illoop = 0 To UBound(ilCodes) - 1
                    If tlChfAdvtExt(ilCurrentRecd).iAgfCode = ilCodes(illoop) Then
                        ilFoundOne = True
                        Exit For
                    End If
                Next illoop
            ElseIf imSlsp Then
                For illoop = 0 To UBound(ilCodes) - 1
                    For ilTemp = 0 To 9
                        If tlChfAdvtExt(ilCurrentRecd).iSlfCode(ilTemp) = ilCodes(illoop) Then
                            ilFoundOne = True
                            Exit For
                        End If
                    Next ilTemp
                    If ilFoundOne Then      'go no futher if one already found to show
                        Exit For
                    End If
                Next illoop
            Else                            'vehicle- assume found one because all lines need to be read
                ilFoundOne = True
            End If
        Else
            ilFoundOne = True
        End If
        'The contract types & statuses have been already filtered  in gObtainCntrForDate
        If ilFoundOne Then

            'Retrieve the contract, schedule lines and flights
            llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT())
            If Not ilRet Then
                On Error GoTo gCrPaperWorkErr
                gBtrvErrorMsg ilRet, "gCrPaperWorkErr (gObtainCntr):", RptSelCt
                On Error GoTo 0
            End If

            ilFoundOne = True

            'check if entered date span selectivity
            gUnpackDateLong tgChfCT.iOHDDate(0), tgChfCT.iOHDDate(1), llTempDate        'contract date entered
            If llTempDate < llEarliestEntry Or llTempDate > llLatestEntry Then
                ilFoundOne = False
            End If

            gUnpackDateLong tgChfCT.iStartDate(0), tgChfCT.iStartDate(1), llCntStartDate
            gUnpackDateLong tgChfCT.iEndDate(0), tgChfCT.iEndDate(1), llCntEndDate
            '2-3-06 test header flag for CBS
            If (llCntStartDate > llCntEndDate Or tgChfCT.sCBSOrder = "C") And Not ilInclCBS Then       'Cancel before start order, and user doesnt want it on report
                ilFoundOne = False
            End If
            'check cash/trade filter
            If (imCash And Not imTrade) And tgChfCT.iPctTrade = 100 Then
                ilFoundOne = False
            End If
            If (imTrade And Not imCash) And tgChfCT.iPctTrade = 0 Then
                ilFoundOne = False
            End If
            
            '11-7-16 special testing for Working Proposals vs Rev working , both have "W" types
            If tgChfCT.sStatus = "W" Then
                If tgChfCT.iCntRevNo = 0 Then
                    If Not imWork Then
                        ilFoundOne = False
                    End If
                Else                'revision
                    If Not imRev Then
                        ilFoundOne = False
                    End If
                End If
            End If

            'check Discrepancy only filter
            If imDiscrepOnly And tgChfCT.sDiscrep <> "Y" Then
                ilFoundOne = False
            End If
            'check Credit checks only :  Need to read advt & agy

            'If imCreditCk Or ilInvSort > 0 Then
                If tgChfCT.iAgfCode > 0 Then
                    tmAgfSrchKey.iCode = tgChfCT.iAgfCode
                    ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching agy recd
                    If ilRet <> BTRV_ERR_NONE Then
                        On Error GoTo gCrPaperWorkErr
                        gBtrvErrorMsg ilRet, "gCrPaperWorkErr (getbtrEqual):Agf", RptSelCt
                        On Error GoTo 0
                    End If          'ilret = btrv_err_none
                Else
                    tmAgf.sCrdApp = "A"  'assume the agency is approved
                End If
                tmAdfSrchKey.iCode = tgChfCT.iAdfCode
                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching advt recd
                If ilRet <> BTRV_ERR_NONE Then
                    On Error GoTo gCrPaperWorkErr
                    gBtrvErrorMsg ilRet, "gCrPaperWorkErr (getbtrEqual):Adf", RptSelCt
                    On Error GoTo 0
                End If

                If imCreditCk Then          'credit checks only
                    If tmAgf.sCrdApp = "A" And tmAdf.sCrdApp = "A" Then        'both are approved, dont show
                        ilFoundOne = False
                    End If
                End If

                If ilInvSort > 0 Then       'selective invoice sort requested
                    If tmAgf.iMnfSort <> ilInvSort And tmAdf.iMnfSort <> ilInvSort Then
                        ilFoundOne = False
                    End If
                End If
            'End If

            If ilFoundOne Then
                If imShowInternal Then
                    'tmGrf.lDollars(1) = tgChfCT.lCxfInt
                    tmGrf.lDollars(0) = tgChfCT.lCxfInt
                End If
                If imShowOther Then
                    'tmGrf.lDollars(2) = tgChfCT.lCxfCode
                    tmGrf.lDollars(1) = tgChfCT.lCxfCode
                End If
                If imShowChgRsn Then
                    'tmGrf.lDollars(3) = tgChfCT.lCxfChgR
                    tmGrf.lDollars(2) = tgChfCT.lCxfChgR
                End If
                If imShowCancel Then
                    'tmGrf.lDollars(4) = tgChfCT.lCxfCanc
                    tmGrf.lDollars(3) = tgChfCT.lCxfCanc
                End If

                'tmGrf.iPerGenl(1) = 0                   'init the sequence # for multiple records within a contr
                tmGrf.iPerGenl(0) = 0                   'init the sequence # for multiple records within a contr
                If imVehicle Or ilByLine Then            'select:vehicle or show sch lines
                    ilAtLeast1Vehicle = False
                    For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1  'loop through all the lines in the contract
                        ilFoundOne = True
                        tmClf = tgClfCT(ilClf).ClfRec
                        gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llCntStartDate
                        gUnpackDateLong tmClf.iEndDate(0), tmClf.iEndDate(1), llCntEndDate
                        If llCntStartDate > llCntEndDate And Not ilInclCBS Then       'Cancel before start order and user doesnt want it shown on report
                            ilFoundOne = False
                        End If
                        
                        ilVefInx = gBinarySearchVef(tmClf.iVefCode)
                        'If ilVefInx <= 0 Then
                        If ilVefInx < 0 Then
                            ilFoundOne = False
                        Else
                            If RptSelCt!rbcSelC7(2).Value = True Then                'show acq only
                                'it must be selling, conventional, Rep vehicles (games, airing excluded)
                                'need to test the types because later, before the record is written, tests to see if valid vehicle selection and if
                                'it should be shown on Insertion Order.  If that question has never been answered in Vehicle options, it is defaulted to Yes.
                                'so the type must also be tested.  These are the only vehicles to have acq costs
                                If tgMVef(ilVefInx).sType <> "S" And tgMVef(ilVefInx).sType <> "C" And tgMVef(ilVefInx).sType <> "R" Then
                                    ilFoundOne = False
                                End If
                            End If
                        End If
                        
                        If ilFoundOne Then
                            llTax1Pct = 0
                            llTax2Pct = 0
                            ilAtLeast1Vehicle = True
                            tmGrf.iYear = tmClf.iLine
                            'tmGrf.iDateGenl(0, 1) = tmClf.iStartDate(0)     'start date of line
                            'tmGrf.iDateGenl(1, 1) = tmClf.iStartDate(1)
                            'tmGrf.iDateGenl(0, 2) = tmClf.iEndDate(0)       'end date of line
                            'tmGrf.iDateGenl(1, 2) = tmClf.iEndDate(1)
                            tmGrf.iDateGenl(0, 0) = tmClf.iStartDate(0)     'start date of line
                            tmGrf.iDateGenl(1, 0) = tmClf.iStartDate(1)
                            tmGrf.iDateGenl(0, 1) = tmClf.iEndDate(0)       'end date of line
                            tmGrf.iDateGenl(1, 1) = tmClf.iEndDate(1)
                            tmGrf.iVefCode = tmClf.iVefCode
                            tmGrf.sBktType = tmClf.sType
                            tmGrf.lChfCode = tgChfCT.lCode        'contract  code
                            If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Then
                                'determine the tax amt for future dates (after last billing date)
                                'a contract can be max 3 years
                                ilTrfAgyAdvt = gGetAirTimeTrfCode(tgChfCT.iAdfCode, tgChfCT.iAgfCode, tmClf.iVefCode)
                                If ilTrfAgyAdvt < 0 Then
                                    'tmGrf.iPerGenl(2) = 0
                                    tmGrf.iPerGenl(1) = 0
                                Else
                                    'tmGrf.iPerGenl(2) = ilTrfAgyAdvt            'auto code
                                    tmGrf.iPerGenl(1) = ilTrfAgyAdvt            'auto code
                                End If
                                gGetAirTimeTaxRates tgChfCT.iAdfCode, tgChfCT.iAgfCode, tmClf.iVefCode, llTax1Pct, llTax2Pct, slGrossNet
                            End If

                            If (ilCkcAll And imVehicle) Or (ilByLine And Not imVehicle) Then
                                '8-14-15 bypass creating record if showing acq and its not a barter vehicle
                                If ((RptSelCt!rbcSelC7(2).Value = True And gIsOnInsertions(tmClf.iVefCode) = True) Or (RptSelCt!rbcSelC7(2).Value = False)) Then
                                    If RptSelCt!rbcSelC7(2).Value Then              'show acq only
                                        'tmGrf.lDollars(5) = tmClf.lAcquisitionCost
                                        tmGrf.lDollars(4) = tmClf.lAcquisitionCost
                                    Else
                                        gBuildFlights ilClf, llStartDates(), 1, 2, llProject(), 1, tgClfCT(), tgCffCT()
                                        'write out record for contractline
                                        'tmGrf.lDollars(5) = llProject(1)                'total $ for line
                                        tmGrf.lDollars(4) = llProject(1)                'total $ for line
                                    End If
                                    tmGrf.iCode2 = 0                                'init the NTR Mnf type code
                                    'tmGrf.iPerGenl(1) = tmGrf.iPerGenl(1) + 1       ' sequence #within Contract
                                    tmGrf.iPerGenl(0) = tmGrf.iPerGenl(0) + 1       ' sequence #within Contract
                                    'tmGrf.lDollars(7) = llTax1Pct
                                    'tmGrf.lDollars(8) = llTax2Pct
                                    tmGrf.lDollars(6) = llTax1Pct
                                    tmGrf.lDollars(7) = llTax2Pct
                                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                End If
                            Else
                                For illoop = LBound(ilCodes) To UBound(ilCodes) - 1
                                    If (ilCodes(illoop) = tmClf.iVefCode) And ((RptSelCt!rbcSelC7(2).Value = True And gIsOnInsertions(tmClf.iVefCode) = True) Or (RptSelCt!rbcSelC7(2).Value = False)) Then
                                        If RptSelCt!rbcSelC7(2).Value Then          'show acq only
                                            'tmGrf.lDollars(5) = tmClf.lAcquisitionCost
                                            tmGrf.lDollars(4) = tmClf.lAcquisitionCost
                                        Else
                                            gBuildFlights ilClf, llStartDates(), 1, 2, llProject(), 1, tgClfCT(), tgCffCT()
                                            
                                            'write out record for contractline
                                            'tmGrf.lDollars(5) = llProject(1)                'total $ for line
                                            tmGrf.lDollars(4) = llProject(1)                'total $ for line
                                        End If

                                        'tmGrf.lDollars(6) = 0                   'init the SBF pointer
                                        tmGrf.lDollars(5) = 0                   'init the SBF pointer
                                        tmGrf.iCode2 = 0                        'init the NTR MNF type code
                                        'tmGrf.iPerGenl(1) = tmGrf.iPerGenl(1) + 1       ' sequence #within Contract
                                        tmGrf.iPerGenl(0) = tmGrf.iPerGenl(0) + 1       ' sequence #within Contract
                                        'tmGrf.lDollars(7) = llTax1Pct
                                        'tmGrf.lDollars(8) = llTax2Pct
                                        tmGrf.lDollars(6) = llTax1Pct
                                        tmGrf.lDollars(7) = llTax2Pct
                                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                        Exit For
                                    End If
                                Next illoop
                            End If
                        End If
                        llProject(1) = 0
                    Next ilClf                                      'loop thru schedule lines
                    'If Not ilAtLeast1Vehicle Then                   'at least 1 line exists
                    '    tmGrf.lChfCode = tgChfCT.lCode        'contract  code
                    '    tmGrf.iCode2 = 0
                    '    tmGrf.iPerGenl(1) = tmGrf.iPerGenl(1) + 1       ' sequence #within Contract
                        'ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    'End If
                Else                'only summary allowed
                    'write out record for contract header
                    tmGrf.lChfCode = tgChfCT.lCode        'contract  code
                    'tmGrf.iPerGenl(1) = tmGrf.iPerGenl(1) + 1       ' sequence #within Contract
                    tmGrf.iPerGenl(0) = tmGrf.iPerGenl(0) + 1       ' sequence #within Contract
                    tmGrf.lDollars(4) = tgChfCT.lInputGross
                    tmGrf.sBktType = ""                             'vehicle line type, n/a for summary
                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                End If

                'if including NTR, create a record for each
                If imNTR Then
                    tmSbfSrchKey0.lChfCode = tgChfCT.lCode
                    'tmSbfSrchKey0.iDate(0) = 0
                    'tmSbfSrchKey0.iDate(1) = 0
                    gPackDate slEarliestActive, tmSbfSrchKey0.iDate(0), tmSbfSrchKey0.iDate(1)

                    tmSbfSrchKey0.sTranType = " "
                    ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = tgChfCT.lCode)
                        If ilRet <> BTRV_ERR_NONE Then
                            On Error GoTo gCrPaperWorkErr
                            gBtrvErrorMsg ilRet, "gCrPaperWorkErr (GetNext):SBF", RptSelCt
                            On Error GoTo 0
                            Exit Sub
                        End If

                        gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llTempDate
                        'earliest parameter filtered with keyread, ignore all SBF records except the item billing
                        If llTempDate <= llStartDates(2) And tmSbf.sTranType = "I" Then
                            tmGrf.iVefCode = tmSbf.iBillVefCode
                            tmGrf.lChfCode = tgChfCT.lCode        'contract  code
                            'tmGrf.lDollars(6) = tmSbf.lCode             'SBF code
                            tmGrf.lDollars(5) = tmSbf.lCode             'SBF code
                            tmGrf.iCode2 = tmSbf.iMnfItem               'MNF item type
                            'tmGrf.iPerGenl(1) = tmGrf.iPerGenl(1) + 1       ' sequence #within Contract
                            tmGrf.iPerGenl(0) = tmGrf.iPerGenl(0) + 1       ' sequence #within Contract
                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                        End If
                        ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If

            End If              'lldate2 >= llweekstart and lldate <= llweekend
        End If                  'ilfoundone = true
        'init record
        'tmGrf.lDollars(1) = 0
        'tmGrf.lDollars(2) = 0
        'tmGrf.lDollars(3) = 0
        'tmGrf.lDollars(4) = 0
        'tmGrf.lDollars(6) = 0                       'sbf code
        
        tmGrf.lDollars(0) = 0
        tmGrf.lDollars(1) = 0
        tmGrf.lDollars(2) = 0
        tmGrf.lDollars(3) = 0
        tmGrf.lDollars(5) = 0                       'sbf code
        tmGrf.iCode2 = 0                            'mnf NTR type
        llTax1Pct = 0
        llTax2Pct = 0
    Next ilCurrentRecd                                      'loop for CHF records
    Erase tlChfAdvtExt, tgClfCT, tgCffCT
    sgCntrForDateStamp = ""
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmAdf)
    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmSlf
    btrDestroy hmMnf
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmAgf
    btrDestroy hmAdf
    btrDestroy hmCHF

    Exit Sub
gCrPaperWorkErr:
    Exit Sub
End Sub
'**********************************************************************
'
'
'       Pre-pass to Sales Activity Report. produce a report
'       of all new and modified activity, including contracts
'       whose status is Hold or Order, plus Salesperson
'       projection data.  Modifications are reflected by increases
'       and decreases from the previous week.  Any potential
'       business entered as projections are detailed and totaled
'       by their projection category.  The effective dte entered
'       filters the contracts and projections for the current week,
'       using the rollover date.  The date is always backed up
'       to a Monday date, which denotes the current week.
'       Increases/decreases are compared agains the previous rev#.
'       The starting qtr and year entred gathers dollars affectving
'       the requested quarter.
'
'       Created:  D.hosaka  12/96
'
'       9/25/97 Do not backup date to Monday.  Use whatever date entered
'               for 7 days.  the 7th day will be the rollover date to
'               gather projections
'       3-14-01 Determine week by the last sunday of the quarter's month, not 13
'               week increments (i.e. 1st qtr 2001 is only 12 weeks, not 13)
'       6-5-01 Add ability to request Sales Activity by Day not week
'              Allow user to enter a date range to select contracts to process
'              Change to obtain contracts by extended reads to speed up, which
'              changed the program flow.
'***************************************************************************
'
'
'
Sub gCrSalesActCt()
Dim ilListIndex As Integer          'report type:  Sales activity by qtr or contract
Dim ilRet As Integer
Dim llEarliestEntry As Long           'start date of all contracts modified or entered new
Dim llLatestEntry As Long             'end date of all contracts modified or entered new
Dim slEarliestEntry As String
Dim slLatestEntry As String
Dim slStr As String                     'temp
Dim llEnterDate As Long                 'contract header entred date
Dim llDate As Long                      'temp
Dim ilFound As Integer
Dim llContrCode As Long                 'contract code to retrieve contr with lines
Dim ilClf As Integer
Dim slAirOrder As String * 1             'O = bill as ordered, A = bill as aired
Dim slStartQtr As String                'active dates of contrct
Dim slEndQtr As String
'ReDim llProject(1 To 2) As Long
ReDim llProject(0 To 2) As Long         'Index zero ignored
'ReDim llStdStartDates(1 To 2) As Long       'only doing 1 qtr, need start date of 2nd (last) qtr
ReDim llStdStartDates(0 To 2) As Long       'only doing 1 qtr, need start date of 2nd (last) qtr. Index zero ignored
'ReDim llProjectCash(1 To 2) As Long    'Not used
'ReDim llProjectTrade(1 To 2) As Long   'Not used
Dim ilTemp As Integer
Dim ilfirstTime As Integer
Dim ilProcessCnt As Integer
Dim ilUpperAct As Integer               'total # of potential advt
Dim ilCalType As Integer                '0 = std, 1 = cal. month, 4 = corp
Dim llAmount As Long
Dim slTimeStamp As String
Dim ilYear As Integer
Dim ilSlfCode As Integer                'slsp processing this report
Dim slCntrStatus As String
Dim ilHOState As Integer
Dim slCntrTypes As String
Dim ilCurrentRecd As Integer
Dim ilCurrentPrev As Integer
ReDim ilTodayDate(0 To 1) As Integer
Dim slTodayDate As String
Dim slGrossOrNet As String * 1          '10-4-13
Dim ilAgyCommPct As Integer
Dim ilAgyInx As Integer
Dim llNoPenny As Long
Dim llTemp As Long
'2020-10-30 - TTP # 9955 - add AirTime, NTR, Hardcost Include options to report (Daily Sales Activity by Contract Report, Weekly Sales Activity by Qtr)
Dim tlSbf() As SBF
'TTP 10855 - prevent overflow due to too many NTR items
'Dim ilSbf As Integer
Dim llSbf As Long
Dim ilValidSbf As Boolean
Dim ilIsItHardCost As Boolean
ReDim tlMnf(0 To 0) As MNF
Dim tlSBFTypes As SBFTypes
Dim blSplitAtNtrHc As Boolean
Dim ltmpDollars0 As Long
Dim ltmpDollars1 As Long
Dim stmpBktType As String
Dim itmpPerGenl As Integer

hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    Exit Sub
End If
imCHFRecLen = Len(tmChf)
hmPjf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmPjf, "", sgDBPath & "Pjf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmCHF)
    btrDestroy hmPjf
    btrDestroy hmCHF
    Exit Sub
End If
imPjfRecLen = Len(tmPjf)
hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmCHF)
    btrDestroy hmGrf
    btrDestroy hmPjf
    btrDestroy hmCHF
    Exit Sub
End If
imGrfRecLen = Len(tmGrf)
hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmCHF)
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmPjf
    btrDestroy hmCHF
    Exit Sub
End If
imClfRecLen = Len(tmClf)
hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmCHF)
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmPjf
    btrDestroy hmCHF
    Exit Sub
End If
imCffRecLen = Len(tmCff)
hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmCHF)
    btrDestroy hmSof
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmPjf
    btrDestroy hmCHF
    Exit Sub
End If
imSofRecLen = Len(tmSof)
hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmCHF)
    btrDestroy hmSlf
    btrDestroy hmSof
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmPjf
    btrDestroy hmCHF
    Exit Sub
End If
imSlfRecLen = Len(tmSlf)
ilListIndex = RptSelCt!lbcRptType.ListIndex
'Determine contracts to process based on their entered and modified dates
If ilListIndex = CNT_DAILY_SALESACTIVITY Then         '6-5-01
    slStr = RptSelCt!CSI_CalFrom.Text   'Date: 11/22/2019 added CSI calendar control for date entries --> edcSelCTo.Text
    llEarliestEntry = gDateValue(slStr)
    slEarliestEntry = Format$(llEarliestEntry, "m/d/yy")
    slStr = RptSelCt!CSI_CalTo.Text     'Date: 11/22/2019 added CSI calendar control for date entries --> edcSelCTo1.Text
    llLatestEntry = gDateValue(slStr)
    slLatestEntry = Format$(llLatestEntry, "m/d/yy")
    'Setup start & end dates of contracts to gather
    llStdStartDates(1) = gDateValue("1/1/1970")
    llStdStartDates(2) = gDateValue("12/31/2069")
Else                        'CNT_SALES ACTIVITY ; Wkly Sales Act by Qtr
    slStr = RptSelCt!CSI_CalFrom.Text   'Date: 1/8/2020 added CSI calendar control for date entries --> edcSelCFrom.Text
    llEarliestEntry = gDateValue(slStr)
    slEarliestEntry = Format$(llEarliestEntry, "m/d/yy")
    llLatestEntry = llEarliestEntry + 6
    slLatestEntry = Format$(llLatestEntry, "m/d/yy")
    'Determine start and end dates of $ to gather

    '3-14-01 quarters should not be determined by 13 week increments
    'llStdStartDates(1) = lgOrigCntrNo                  'start date of qtr
    'llStdStartDates(2) = llStdStartDates(1) + 90       'get start of next qtr

    'Determine qtrs by by the last sunday of the month, not 13 week increments
    ilYear = Val(RptSelCt!edcSelCTo.Text)
    '2 = always std qtrs
    gGetStartEndQtr 2, ilYear, igMonthOrQtr, slStartQtr, slEndQtr
    llStdStartDates(1) = gDateValue(slStartQtr)
    llStdStartDates(2) = gDateValue(slEndQtr)
End If

slGrossOrNet = "G"          '10-4-13 implement net feature
If RptSelCt!rbcSelC7(1).Value = True Then       'net
    slGrossOrNet = "N"
End If

    '2020-10-30 - TTP # 9955 - add AirTime, NTR, Hardcost Include options to report (Daily Sales Activity by Contract Report, Weekly Sales Activity by Qtr)
    tlSBFTypes.iNTR = False          'include NTR billing
    tlSBFTypes.iInstallment = False      'exclude Installment billing
    tlSBFTypes.iImport = False           'exclude rep import billing
    imAirTime = True
    imNTR = False
    imHardCost = False
    'if Weekly Sales Activity by Month and either NTR or hard cost selected
    
    'test for NTR inclusion (or Hard Cost)
    If Not RptSelCt!ckcSelC13(0).Value = vbChecked Then      'include AirTime
        imAirTime = False
    End If
    If RptSelCt!ckcSelC13(1).Value = vbChecked Then      'include NTR
        imNTR = True
        tlSBFTypes.iNTR = True
    End If
    If RptSelCt!ckcSelC13(2).Value = vbChecked Then      'include Hard Cost
        tlSBFTypes.iNTR = True 'Needed for gObtainSBF
        imHardCost = True
    End If
    blSplitAtNtrHc = RptSelCt!ckcSelC10(0).Value = vbChecked
    
    If imNTR Or imHardCost Then
        hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmSbf)
            ilRet = btrClose(hmSlf)
            ilRet = btrClose(hmSof)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmPjf)
            ilRet = btrClose(hmCHF)
            btrDestroy hmSbf
            btrDestroy hmSlf
            btrDestroy hmSof
            btrDestroy hmCff
            btrDestroy hmClf
            btrDestroy hmGrf
            btrDestroy hmPjf
            btrDestroy hmCHF
            Exit Sub
        End If
        imSbfRecLen = Len(tmSbf)

        hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmMnf)
            ilRet = btrClose(hmSbf)
            ilRet = btrClose(hmSlf)
            ilRet = btrClose(hmSof)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmPjf)
            ilRet = btrClose(hmCHF)
            btrDestroy hmMnf
            btrDestroy hmSbf
            btrDestroy hmSlf
            btrDestroy hmSof
            btrDestroy hmCff
            btrDestroy hmClf
            btrDestroy hmGrf
            btrDestroy hmPjf
            btrDestroy hmCHF
            Exit Sub
        End If
        imMnfRecLen = Len(tmMnf)
        sgDemoMnfStamp = ""             'insure that the hard costs  are read
        ilRet = gObtainMnfForType("I", sgDemoMnfStamp, tlMnf())
    End If
    
'build array of selling office codes and their sales sources.  This is the most major sort
'in the Business Booked reports
ilTemp = 0
ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
Do While ilRet = BTRV_ERR_NONE
    ReDim Preserve tlSofList(0 To ilTemp) As SOFLIST
    tlSofList(ilTemp).iSofCode = tmSof.iCode
    tlSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
    ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    ilTemp = ilTemp + 1
Loop
'Populate the salespeople.  Only salesp running report can see his own stuff
ilRet = gObtainSalesperson()        'populated slsp list in tgMSlf
If tgUrf(0).iCode = 1 Or tgUrf(0).iCode = 2 Then    'guide or counterpoint password
    ilSlfCode = 0                   'allow guide & CSI to get all stuff
Else
    ilSlfCode = tgUrf(0).iSlfCode   'slsp gets to see only his own stuff
End If
ilfirstTime = True
slAirOrder = tgSpf.sInvAirOrder     'inv all contracts as aired or ordered
tmGrf = tmZeroGrf                'initialize new record
ilRet = btrGetFirst(hmCHF, tmChf, imCHFRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)  'get contracts by external contr # (rev #)
slCntrTypes = gBuildCntTypes()
slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order.
ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)
slTodayDate = Format$(gNow(), "m/d/yy")
gPackDate slTodayDate, ilTodayDate(0), ilTodayDate(1)
'obtain all contracts active from earliest requested date to today,
'but only process the contracts between requested period.  All these contracts
'are gathered because one that is modified after the requested period isnt found if
'the requested dates go backward
ilRet = gCntrForActiveOHD(RptSelCt, "", "", slEarliestEntry, slTodayDate, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
     ilProcessCnt = False

    'for each Contract process 2 entries in table of GRF records.
    'If only 1 record created in previous week, its a decrease.
    'if only 1 record created in current week, its an increase (New).
    'If both records created for previous and current week, show difference.
    'Write out 1 record into GRF file containing the final result of previous to current.
    'Do While ilRet = BTRV_ERR_NONE
    For ilCurrentPrev = 1 To 2
        If ilCurrentPrev = 1 Then               'current
            tmGrf = tmZeroGrf                'initialize new record
            llContrCode = gActivityCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llEarliestEntry, llLatestEntry, hmCHF, tmChf)
            If llContrCode = 0 Then                     'nothing in current week
                ilCurrentPrev = 2                       '

                Exit For                                'dont bother testing previous week
            End If
            gUnpackDate tmChf.iOHDDate(0), tmChf.iOHDDate(1), slStr
            llEnterDate = gDateValue(slStr)
            tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
            tmGrf.iGenDate(1) = igNowDate(1)
            'tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
            'tmGrf.iGenTime(1) = igNowTime(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            tmGrf.lGenTime = lgNowTime
            tmGrf.sBktType = "O"                  'orders flag (vs P = project), for sorting
            tmGrf.sDateType = "O"               'indicate this is Orders vs A,B,C projections for sorting
            tmGrf.lChfCode = tmChf.lCntrNo            'contract #
            tmGrf.iAdfCode = tmChf.iAdfCode         'advertiser code
            'tmGrf.iPerGenl(2) = tmChf.iCntRevNo
            'tmGrf.iPerGenl(4) = 0                   'assume Order (vs hold)
            tmGrf.iPerGenl(1) = tmChf.iCntRevNo
            tmGrf.iPerGenl(3) = 0                   'assume Order (vs hold)
            tmGrf.lCode4 = tmChf.lCxfChgR
            If tmChf.sStatus = "H" Or tmChf.sStatus = "G" Then      'if hold or unsch hold, set flag for Crystal
                'tmGrf.iPerGenl(4) = 1
                tmGrf.iPerGenl(3) = 1
            End If
            ilRet = BTRV_ERR_NONE
            If tmChf.iSlfCode(0) <> tmSlf.iCode Then        'only read slsp recd if not in mem already
                tmSlfSrchKey.iCode = tmChf.iSlfCode(0)         'find the slsp to obtain the sales source code
                ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            End If                                          'table of selling offices built into memory with its
            If ilRet = BTRV_ERR_NONE Then
                tmGrf.iSofCode = tmSlf.iSofCode
                'sales source
                For ilTemp = LBound(tlSofList) To UBound(tlSofList)
                    If tlSofList(ilTemp).iSofCode = tmSlf.iSofCode Then
                        tmGrf.iCode2 = tlSofList(ilTemp).iMnfSSCode          'Sales source
                        Exit For
                    End If
                Next ilTemp
            Else
                tmGrf.iSofCode = 0
                tmGrf.iCode2 = 0
            End If
            ilProcessCnt = True
        Else                                'previous week
            ilProcessCnt = False
            llContrCode = gPaceCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llEarliestEntry - 1, hmCHF, tmChf) 'find the prvious version of the cntr
            If llContrCode > 0 Then
                gUnpackDate tmChf.iOHDDate(0), tmChf.iOHDDate(1), slStr
                llEnterDate = gDateValue(slStr)
                If llEnterDate < llEarliestEntry Then      'must be in same period of activity, ignore
                    ilProcessCnt = True
                End If
            End If
        End If

        'If (llEnterDate <= llLatestEntry) And (tmChf.iSlfCode(0) = ilSlfCode Or ilSlfCode = 0) And (tmChf.sStatus = "H" Or tmChf.sStatus = "O" Or tmChf.sStatus = "G" Or tmChf.sStatus = "N") Then
        'do not test for the slsp that signed on.  Filtering of contr slsp is allowed to see has been done
        'in gCntrForActiveOHD
        If (llEnterDate <= llLatestEntry) And (tmChf.sStatus = "H" Or tmChf.sStatus = "O" Or tmChf.sStatus = "G" Or tmChf.sStatus = "N") Then
            'if llenterdate is less than llearliest entry, then it needs to be processed.  It could be
            'a contract that did not get carried over (in which case its a decrease)

            If (ilProcessCnt) Then
                llContrCode = tmChf.lCode
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT())
                ilAgyCommPct = 10000                    'assume gross requested
                ilAgyInx = gBinarySearchAgf(tgChfCT.iAgfCode)
                If ilAgyInx >= 0 Then
                    ilAgyCommPct = tgCommAgf(ilAgyInx).iCommPct
                    If slGrossOrNet = "N" Then
                        ilAgyCommPct = 10000 - ilAgyCommPct
                    Else
                        ilAgyCommPct = 10000                    'assume gross requested
                    End If
                End If

                If imAirTime Then
                    For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                        tmClf = tgClfCT(ilClf).ClfRec
                        If tmClf.sType = "S" Or tmClf.sType = "H" Then    'project standard or hidden lines (packages dont have vehicle groups)
                            gBuildFlights ilClf, llStdStartDates(), 1, 2, llProject(), 1, tgClfCT(), tgCffCT()
                        End If
                    Next ilClf
    
                    llNoPenny = llProject(1) / 100    'drop pennies
                    llTemp = llNoPenny * CDbl(ilAgyCommPct) / 100  'adjust for the agy comm carried in 2 places                               'calc agy comm
    
                    If ilCurrentPrev = 1 Then       'current version
                        ''tmGrf.lDollars(1) = llProject(1)    'save the current version so its not wiped out by the previous version $
                        'tmGrf.lDollars(1) = llTemp
                        tmGrf.lDollars(0) = llTemp
                        'tmGrf.iPerGenl(1) = 0               'flag as NEW
                        tmGrf.iPerGenl(0) = 0               'flag as NEW
                    Else
                        ''tmGrf.lDollars(2) = llProject(1)
                        'tmGrf.lDollars(2) = llTemp
                        tmGrf.lDollars(1) = llTemp
                        'tmGrf.iPerGenl(1) = 1                'there is a previous version
                        tmGrf.iPerGenl(0) = 1                'there is a previous version
                    End If
                    If blSplitAtNtrHc = False Then
                        If Not (imNTR And imHardCost) Then
                            If tmChf.sNTRDefined = "Y" Then
                                'TTP # 9955 - even when not including NTR or HardCode, show NT is this contract includes NTR
                                tmGrf.sGenDesc = "NT"
                            End If
                        End If
                    End If
                End If
                
                '2020-10-30 - TTP # 9955 - add AirTime, NTR, Hardcost Include options to report (Daily Sales Activity by Contract Report, Weekly Sales Activity by Qtr)
                If imNTR Or imHardCost Then
                    'Add NTR to the current or prev
                    If ilCurrentPrev = 1 Then       'current version
                        llTemp = 0
                        llTemp = mGetSalesActNTR(ilListIndex, Format$(llStdStartDates(1), "m/d/yy"), Format$(llStdStartDates(2), "m/d/yy"), tlSBFTypes, tlSbf(), tlMnf(), slGrossOrNet, ilAgyCommPct)
                        If blSplitAtNtrHc = True Then
                            'Split this NTR or HC into a separate GRF record, store the values and insert later
                            itmpPerGenl = 0 'flag as New (NTR/HC)
                            ltmpDollars0 = llTemp ' Get NTR $ for insert later
                        Else
                            'Dont split, just Add NTR or HC to the AirTime
                            tmGrf.iPerGenl(0) = 0  'flag as New (NTR/HC)
                            tmGrf.lDollars(0) = tmGrf.lDollars(0) + llTemp ' add NTR to Contract $
                            If llTemp <> 0 Then tmGrf.sGenDesc = "NT" 'Contract has NTR
                        End If
                    Else
                        llTemp = 0
                        llTemp = mGetSalesActNTR(ilListIndex, Format$(llStdStartDates(1), "m/d/yy"), Format$(llStdStartDates(2), "m/d/yy"), tlSBFTypes, tlSbf(), tlMnf(), slGrossOrNet, ilAgyCommPct)
                        If blSplitAtNtrHc = True Then
                            'Split this NTR or HC into a separate GRF record, store the values and insert later
                            ltmpDollars1 = llTemp ' Get NTR $ for insert later
                            itmpPerGenl = 1 'there is a previous version (NTR/HC)
                        Else
                            'Dont split, just Add NTR or HC to the AirTime
                            tmGrf.iPerGenl(0) = 1  'there is a previous version (NTR/HC)
                            tmGrf.lDollars(1) = tmGrf.lDollars(1) + llTemp ' add NTR to Contract $
                            If llTemp <> 0 Then tmGrf.sGenDesc = "NT" 'Contract has NTR
                        End If
                    End If
                End If
            Else
                'not ilProcessCnt
            End If
        Else
            'Not Date Range -or- not correct Order status (HOGN)
        End If
        llProject(1) = 0
    Next ilCurrentPrev
    'GenDate - generation date (key)
    'GenTime - generation time (key)
    'chfCode = Contract #
    'adfCode = Advetiser Code
    'Code2 - Sales Source
    'BktType - O = orders, P = projetions
    'DateType = A, B or C for projections
    'Code 4 = change reason
    'sofCode = slsp office code
    'PerGenl(1) - 0 = new, else modification
    'PerGenl(2) - Contract Rev #
    'PerGenl(3) - internal flags for detail processing, unused in crystal
    'PerGenl(4) - hold or unsch hold flag
    'lDolllars(1) - Amount
    'mWriteSlsAct            'format common fields in record
    'tmGrf.lDollars(1) = (tmGrf.lDollars(1) - tmGrf.lDollars(2)) / 100
    
    If blSplitAtNtrHc = True Then
        'Insert NTR GRF
        Dim tmpGrf As GRF
        tmpGrf = tmGrf 'Clone Existing GRF, but update with NTR values
        tmpGrf.lDollars(0) = ltmpDollars0
        tmpGrf.lDollars(1) = ltmpDollars1
        tmpGrf.lDollars(0) = (tmpGrf.lDollars(0) - tmpGrf.lDollars(1)) / 100
        tmpGrf.iPerGenl(0) = itmpPerGenl
        'tmpGrf.sBktType = "N"                  'NTR/HC
        tmpGrf.sGenDesc = "NT"
        If tmpGrf.lDollars(0) <> 0 Then
            ilRet = btrInsert(hmGrf, tmpGrf, imGrfRecLen, INDEXKEY0)
        End If
        ltmpDollars0 = 0
        ltmpDollars1 = 0
    End If
    
    tmGrf.lDollars(0) = (tmGrf.lDollars(0) - tmGrf.lDollars(1)) / 100
    'If tmGrf.lDollars(1) <> 0 Then      'only show contracts with +/- $
    If tmGrf.lDollars(0) <> 0 Then      'only show contracts with +/- $
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
    End If
Next ilCurrentRecd
If ilListIndex = CNT_SALESACTIVITY Then
    'Process the projections for the same weeks (effective date)
    'Build table of all valid projections by advt, potential code and slsp.
    'Need to determine if projection is new, increase or decrease.
    'Write out 1 record into GRF file containing the final result of previous to current.
    '10-7-13 add net option to Weekly Sales Activity by Cnt
    ilAgyCommPct = 10000                    'assume gross requested
    If slGrossOrNet = "N" Then              'net
        ilAgyCommPct = 8500                 'net assumed 85% for all
    End If
    
    ilCalType = 0                       'retrieve for std month
    ilUpperAct = 0
    ilRet = gObtainMnfForType("P", slTimeStamp, tlMMnf())   'populate Potential types (A,B,C)
    ReDim Preserve tlActList(0 To 0) As ACTLIST
    For ilClf = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
        tmPjfSrchKey.iSlfCode = tgMSlf(ilClf).iCode
        tmPjfSrchKey.iRolloverDate(0) = 0               'find all for this slsp
        tmPjfSrchKey.iRolloverDate(1) = 0
        ilRet = btrGetGreaterOrEqual(hmPjf, tmPjf, imPjfRecLen, tmPjfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'get all projection records
        Do While (ilRet = BTRV_ERR_NONE And tmPjf.iSlfCode = tgMSlf(ilClf).iCode) And (tmPjf.iSlfCode = ilSlfCode Or ilSlfCode = 0)
            gUnpackDate tmPjf.iRolloverDate(0), tmPjf.iRolloverDate(1), slStr                 'effec date of projected record
            llEnterDate = gDateValue(slStr)
            If llEnterDate >= llEarliestEntry - 7 And llEnterDate <= llLatestEntry And ilYear = tmPjf.iYear Then 'within previous and currnet weeks?
                ilFound = False
                llProject(1) = 0
                For llDate = llStdStartDates(1) To (llStdStartDates(2) - 1) Step 7
                    slStr = Format$(llDate, "m/d/yy")
                    llAmount = gGetWkDollars(ilCalType, slStr, tmPjf.lGross())
                    llProject(1) = llProject(1) + llAmount
                Next llDate
                For ilTemp = 0 To ilUpperAct - 1 Step 1
                    If tlActList(ilTemp).iAdfCode = tmPjf.iAdfCode And tlActList(ilTemp).iPotnCode = tmPjf.iMnfBus And tlActList(ilTemp).iSlfCode = tmPjf.iSlfCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTemp
                If Not (ilFound) Then                   'create new entry
                    ReDim Preserve tlActList(0 To ilUpperAct) As ACTLIST
                    tlActList(ilUpperAct).iAdfCode = tmPjf.iAdfCode
                    tlActList(ilUpperAct).iPotnCode = tmPjf.iMnfBus
                    tlActList(ilUpperAct).iSlfCode = tmPjf.iSlfCode
                    tlActList(ilUpperAct).lCxfChgR = tmPjf.lCxfChgR
                    If llEnterDate < llEarliestEntry Then               'previous entry
                        tlActList(ilUpperAct).iWeekFlag = 1
                        tlActList(ilUpperAct).lAmount = tlActList(ilUpperAct).lAmount - llProject(1)
                    Else
                        tlActList(ilUpperAct).iWeekFlag = 2             'current entry only
                        tlActList(ilUpperAct).lAmount = tlActList(ilUpperAct).lAmount + llProject(1)
                    End If

                    ilUpperAct = ilUpperAct + 1         'increment for next new record
                Else                                    'update existing entry, must be increase or decrease
                    tlActList(ilTemp).iWeekFlag = 3             'both prev & current exit
                    If llEnterDate < llEarliestEntry Then               'previous entry
                        tlActList(ilTemp).lAmount = tlActList(ilTemp).lAmount - llProject(1)
                    Else
                        tlActList(ilTemp).lAmount = tlActList(ilTemp).lAmount + llProject(1)
                    End If
                End If
            End If
            ilRet = btrGetNext(hmPjf, tmPjf, imPjfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Next ilClf                                          'next slsp
    tgChfCT.lCntrNo = 0                                   'contract #s do not apply for the projections

    'All projections for past and current weeks activity are built into memory.  Now  write them
    'out to disk.
    For ilTemp = 0 To ilUpperAct - 1 Step 1
        tmGrf.sBktType = "P"                  'orders flag (vs P = project), for sorting
        'tmGrf.iPerGenl(1) = 1                'assume modification to projection
        tmGrf.iPerGenl(0) = 1                'assume modification to projection
        If tlActList(ilTemp).iWeekFlag = 2 Then           'flag 2 denotes new only
            'tmGrf.iPerGenl(1) = 0
            tmGrf.iPerGenl(0) = 0
        End If
        'tmGrf.iPerGenl(2) = 0
        tmGrf.iPerGenl(1) = 0
        tmGrf.sDateType = " "
        tmGrf.lCode4 = tlActList(ilTemp).lCxfChgR
        For ilFound = LBound(tlMMnf) To UBound(tlMMnf) - 1 Step 1
            If tlMMnf(ilFound).iCode = tlActList(ilTemp).iPotnCode Then
                tmGrf.sDateType = tlMMnf(ilFound).sName     'indicate Potential code A,B,C
                Exit For
            End If
        Next ilFound
        tgChfCT.iAdfCode = tlActList(ilTemp).iAdfCode         'common routine assumes adv is in cntr buffer
        tgChfCT.iSlfCode(0) = tlActList(ilTemp).iSlfCode
        'tmGrf.lDollars(1) = tlActList(ilTemp).lAmount
        
        llNoPenny = tlActList(ilTemp).lAmount     'drop pennies
        llTemp = llNoPenny * CDbl(ilAgyCommPct) / 10000  'adjust for the agy comm carried in 2 places                               'calc agy comm

        'tmGrf.lDollars(1) = llTemp
        tmGrf.lDollars(0) = llTemp

        mWriteSlsAct          'format common fields in record
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
    Next ilTemp
End If
Erase tlActList, tlSofList, llProject, llStdStartDates
ilRet = btrClose(hmSlf)
ilRet = btrClose(hmSof)
ilRet = btrClose(hmCff)
ilRet = btrClose(hmClf)
ilRet = btrClose(hmGrf)
ilRet = btrClose(hmPjf)
ilRet = btrClose(hmCHF)
        btrDestroy hmSlf
        btrDestroy hmSof
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmPjf
        btrDestroy hmCHF
        
If imNTR Or imHardCost Then
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSbf)
    btrDestroy hmMnf
    btrDestroy hmSbf
End If
End Sub
'
'
'                   Create Sales Analysis Summary prepass file
'                   Generate GRF file by vehicle.  Each record  contains the vehicle,
'                   plan $, Business on Books for current years Qtr, OOB w/ holds for
'                   current years qtr, slsp projection $(pjf), Last years same week OOB (chf),
'                   and last years Actual $ (from contracts)
'
'                   4/8/98 Ignore 100% trade contracts
'       2-11-05 Chg array of PHF/RVF array from integer to long (prevent overflow)
Sub gCrSalesAna(llCurrStart As Long, llCurrEnd As Long, llPrevStart As Long, llPrevEnd As Long)
Dim slMnfStamp As String
Dim slAirOrder As String * 1                'from site pref - bill as air or ordered
'ReDim ilLikePct(1 To 3) As Integer             'most likely percentage from potential code A, B & C
ReDim ilLikePct(0 To 3) As Integer             'most likely percentage from potential code A, B & C. Index zero ignored
'ReDim ilLikeCode(1 To 3) As Integer           'mnf most likely auto increment code for A, B, C
ReDim ilLikeCode(0 To 3) As Integer           'mnf most likely auto increment code for A, B, C. Index zero ignored
Dim ilPotnInx As Integer                    'index to ilLikePct (which % to use)
Dim illoop As Integer
Dim ilTemp As Integer
Dim ilSlsLoop As Integer
Dim ilRet As Integer
ReDim ilRODate(0 To 1) As Integer           'Effective Date to match retrieval of Projection record
ReDim ilEnterDate(0 To 1) As Integer        'Btrieve format for date entered by user
Dim llClosestDate As Long                   'closest date to the rollover user entered date
Dim slDate As String
Dim slStr As String
Dim ilMonth As Integer
Dim ilYear As Integer
Dim llEnterFrom As Long                       'gather cnts whose entered date falls within llEnterFrom and llEnterTo
Dim llEnterTo As Long
'ReDim llProject(1 To 2) As Long               'projected $, only using 1 bucket, common rtn needs assumes array
ReDim llProject(0 To 2) As Long               'projected $, only using 1 bucket, common rtn needs assumes array. Index zero ignored
'ReDim llLYDates(1 To 2) As Long               'range of  qtr dates for contract retrieval (this year)
ReDim llLYDates(0 To 2) As Long               'range of  qtr dates for contract retrieval (this year). Index zero ignored
Dim slTYStartQtr As String
Dim slTYEndQtr As String
'ReDim llTYDates(1 To 2) As Long               'range of qtr dates for contract retrieval (last year)
ReDim llTYDates(0 To 2) As Long               'range of qtr dates for contract retrieval (last year). Index zero ignored
Dim slLYStartQtr As String
Dim slLYEndQtr As String
'ReDim llStartDates(1 To 2) As Long            'temp array for last year vs this years range of dates
ReDim llStartDates(0 To 2) As Long            'temp array for last year vs this years range of dates. Index zero ignored
Dim llLYGetFrom As Long                       'start date of last year
Dim llLYGetTo As Long                         'obtain last years qtr if cnt entered date equal/prior to this date (same time last year)
Dim slLYStartYr As String                     'start date of last year
Dim slLYEndYr As String
Dim llTYGetFrom As Long                       'start date of last year
Dim llTYGetTo As Long                         'obtain this years qtr if cnt entered date equal/prior to this date
Dim slTYStartYr As String                     'start date of this year
Dim slTYEndYr As String                       'end date of this year
Dim llDate As Long                            'temp date field
Dim ilLYRepeat As Integer
Dim ilRepeat As Integer
Dim ilBdMnfCode As Integer                      'budget name to get from mnf file
Dim ilBdYear As Integer                         'budget year to get from budget file
Dim slNameCode As String
Dim slYear As String
Dim slCode As String
Dim slCntrTypes As String                       'valid contract types to access
Dim slCntrStatus As String                      'valid status (holds, orders, working, etc) to access
Dim ilHOState As Integer                        'include unsch holds/orders, sch holds/orders
Dim ilFound As Integer
Dim ilStartWk As Integer                        'starting week index to gather budget data
Dim ilEndWk As Integer                          'ending week index to gather budgets
Dim ilFirstWk As Integer                        'true if week 0 needs to be added when start wk = 1
Dim ilLastWk As Integer                         'true if week 53 needs to be added when end wk = 52
Dim llContrCode As Long                         'contr code from gObtainCntrforDate
Dim ilCurrentRecd As Integer                    'index of contract being processed from tlChfAdvtExt
Dim ilPastFut As Integer                        'loop to process past contracts, then current contracts
Dim ilClf As Integer                            'index to line from tgClfCt
Dim llAdjust As Long                            'Adjusted gross using the potential codes most likely %
Dim ilCorpStd As Integer            '1 = corp, 2 = std
Dim ilBvfCalType As Integer            '0=std, 1 = reg, 2 & 3 = julian, 4 = corp for jan thru dec, 5 = corp for fiscal year
Dim ilPjfCalType As Integer            '0=std, 1 = reg, 2 & 3 = julian, 4 = corp for jan thru dec, 5 = corp for fiscal year
Dim ilFoundSls As Integer
Dim ilTY As Integer
Dim llLYActualNoPace As Long            'last years actual qtr $ (no pacing test)
Dim llProjYearEnd As Long               'projection recds end date for the standard year
Dim ilSaveStartWk As Integer
Dim ilSaveEndWk As Integer
Dim ilMinorSet As Integer
Dim ilMajorSet As Integer
Dim ilmnfMinorCode As Integer           'field used to sort the minor sort with
Dim ilMnfMajorCode As Integer           'field used to sort the major sort with
Dim llRvfLoop As Long                   '2-11-05
' 7-30-08 Dan M added Ntr/hard cost option
Dim blIncludeNTR As Boolean
Dim blIncludeHardCost As Boolean
Dim tlNTRInfo() As NTRPacing
Dim ilLowerboundNTR As Integer
Dim ilUpperboundNTR As Integer
Dim ilNTRCounter As Integer
Dim llSingleContract As Long
Dim llDateEntered As Long               'receivables entered date for pacing test
Dim blFailedMatchNtrOrHardCost As Boolean
Dim blFailedBecauseInstallment As Boolean
Dim slTemp As String
'single contract selectivity
If Val(RptSelCt!edcText.Text) > 0 And RptSelCt!edcText.Text <> " " Then
     llSingleContract = Val(RptSelCt!edcText.Text)
Else
    llSingleContract = NOT_SELECTED
End If
Dim tlTranType As TRANTYPES
'ReDim tlRvf(1 To 1) As RVF
ReDim tlRvf(0 To 0) As RVF
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)
    hmBvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmBvf, "", sgDBPath & "Bvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmBvf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imBvfRecLen = Len(tmBvf)

    hmPjf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPjf, "", sgDBPath & "Pjf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmBvf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmPjf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imPjfRecLen = Len(tmPjf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmBvf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmMnf
        btrDestroy hmPjf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    ' Dan M 7-30-08
    hmSbf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSbf)
        btrDestroy hmSbf
        btrDestroy hmMnf
        btrDestroy hmPjf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    ReDim tmMnfNtr(0 To 0) As MNF

    tlTranType.iAdj = True              'look only for adjustments in the History & Rec files
    tlTranType.iInv = False
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iCash = True
    tlTranType.iTrade = False
    tlTranType.iMerch = False
    tlTranType.iPromo = False
    tlTranType.iNTR = False         '9-17-02

    If RptSelCt!ckcSelC3(0).Value Or RptSelCt!ckcSelC3(1).Value Then    'don't waste time filling array if don't need.
        tlTranType.iNTR = True
    'set flags ntr or hard cost or both chosen
        If RptSelCt!ckcSelC3(0).Value = 1 Then
             blIncludeNTR = True
             tlTranType.iNTR = True
        End If
        If RptSelCt!ckcSelC3(1).Value = 1 Then
            blIncludeHardCost = True
            tlTranType.iNTR = True
        End If
        ilRet = gObtainMnfForType("I", "", tmMnfNtr())
        If ilRet <> True Then
            MsgBox "error retrieving MNF files", vbOKOnly + vbCritical
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If


    slAirOrder = tgSpf.sInvAirOrder     'inv all contracts as aired or ordered
    'ilLoop = RptSelCt!cbcSet1.ListIndex
    'ilMajorSet = tgVehicleSets1(ilLoop).icode
    illoop = RptSelCt!cbcSet1.ListIndex
    ilMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())

    'ilLoop = RptSelCt!cbcSet2.ListIndex
    'ilMinorSet = tgVehicleSets2(ilLoop).icode
    illoop = RptSelCt!cbcSet2.ListIndex
    ilMinorSet = gFindVehGroupInx(illoop, tgVehicleSets2())

    'ReDim tlMMnf(1 To 1) As MNF
    ReDim tlMMnf(0 To 0) As MNF
    'get all the Potential codes from MNF  and save their adjustment percentages
    ilRet = gObtainMnfForType("P", slMnfStamp, tlMMnf())
    'For ilLoop = 1 To UBound(tlMMnf) - 1 Step 1
    For illoop = LBound(tlMMnf) To UBound(tlMMnf) - 1 Step 1
        If Trim$(tlMMnf(illoop).sName) = "A" Then
            ilLikePct(1) = Val(tlMMnf(illoop).sUnitType)            'most likely percentage from potential code "A"
            ilLikeCode(1) = tlMMnf(illoop).iCode
        ElseIf Trim$(tlMMnf(illoop).sName) = "B" Then
                ilLikePct(2) = Val(tlMMnf(illoop).sUnitType)            'most likely percentage from potential code "B"
                ilLikeCode(2) = tlMMnf(illoop).iCode
        ElseIf Trim$(tlMMnf(illoop).sName) = "C" Then
                ilLikePct(3) = Val(tlMMnf(illoop).sUnitType)            'most likely percentage from potential code "C"
                ilLikeCode(3) = tlMMnf(illoop).iCode
        End If
    Next illoop
    'get all the dates needed to work with
    slDate = RptSelCt!CSI_CalFrom.Text                  'Date: 12/12/2019 added CSI calendar control for date entry --> edcSelCFrom.Text               'effective date entred
    'obtain the entered dates year based on the std month
    llTYGetTo = gDateValue(slDate)                     'gather contracts thru this date

    gPackDateLong llTYGetTo, ilEnterDate(0), ilEnterDate(1)    'get btrieve date format for entered to pass to record to show on hdr
    'setup Projection rollover date
    'gPackDate slDate, ilRODate(0), ilRODate(1)
    gGetRollOverDate RptSelCt, 2, slDate, llClosestDate   'send the lbcselection index to search, plust rollover date
    gPackDateLong llClosestDate, ilRODate(0), ilRODate(1)


    ilYear = Val(RptSelCt!edcSelCTo.Text)           'year requested
    If RptSelCt!rbcSelCInclude(0).Value Then
        ilCorpStd = 1                               'corp flag for genl subtrn
        ilPjfCalType = 4               'get week inx based on std for projections
        ilBvfCalType = 5               'get week inx based on fiscal dates
    Else
        ilCorpStd = 2
        ilPjfCalType = 0             'std month
        ilBvfCalType = 0             'both projections and budgets will be std
    End If

    'This Years start/end quarter and year dates
    gGetStartEndQtr ilCorpStd, ilYear, igMonthOrQtr, slTYStartQtr, slTYEndQtr
    llTYDates(1) = gDateValue(slTYStartQtr)
    llTYDates(2) = gDateValue(slTYEndQtr)
    gGetStartEndYear ilCorpStd, ilYear, slTYStartYr, slTYEndYr
    llTYGetFrom = gDateValue(slTYStartYr)
    'Last years start/end quarter and year dates
    gGetStartEndQtr ilCorpStd, ilYear - 1, igMonthOrQtr, slLYStartQtr, slLYEndQtr
    llLYDates(1) = gDateValue(slLYStartQtr)
    llLYDates(2) = gDateValue(slLYEndQtr)
    gGetStartEndYear ilCorpStd, ilYear - 1, slLYStartYr, slLYEndYr
    llLYGetFrom = gDateValue(slLYStartYr)
    'determine same time last year
    llLYGetTo = llLYGetFrom + (llTYGetTo - llTYGetFrom)
    'Determine the Budget name selected
    slNameCode = tgRptSelBudgetCodeCT(igBSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 1, "\", slStr)
    ilRet = gParseItem(slStr, 1, "\", slYear)
    slYear = gSubStr("9999", slYear)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilBdMnfCode = Val(slCode)
    ilBdYear = Val(slYear)

    'ReDim tlSlsList(1 To 1) As SLSLIST          'array of vehicles and their sales
    ReDim tlSlsList(0 To 0) As SLSLIST          'array of vehicles and their sales
    'gather all budget records by vehicle for the requested year, totaling by quarter
    If Not mReadBvfRec(hmBvf, ilBdMnfCode, ilBdYear, tmBvfVeh()) Then
        Exit Sub
    End If

    'get startwk & endwk to gather budgets
    gObtainWkNo ilBvfCalType, slTYStartQtr, ilStartWk, ilFirstWk          'determine the first week inx to accum (current year)
    gObtainWkNo ilBvfCalType, slTYEndQtr, ilEndWk, ilLastWk               'determine the last week inx to accum  (current year)

    'build budget information by vehicle in memory
    ilFound = False
    For illoop = LBound(tmBvfVeh) To UBound(tmBvfVeh) - 1 Step 1
        For ilSlsLoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1
            If tmBvfVeh(illoop).iVefCode = tlSlsList(ilSlsLoop).iVefCode Then
                ilFound = True
                Exit For
            End If
        Next ilSlsLoop
        If Not ilFound Then
            tlSlsList(UBound(tlSlsList)).iVefCode = tmBvfVeh(illoop).iVefCode
            ReDim Preserve tlSlsList(LBound(tlSlsList) To UBound(tlSlsList) + 1)
        End If
        'ilSlsLoop contains index to the correct vehicle
        For ilTemp = ilStartWk To ilEndWk Step 1
            tlSlsList(ilSlsLoop).lPlan = tlSlsList(ilSlsLoop).lPlan + tmBvfVeh(illoop).lGross(ilTemp)
        Next ilTemp
        If ilFirstWk Then       'adjust for the partial weeks at the beginning or end of the year
                                'due to corp or calendar months
            tlSlsList(ilSlsLoop).lPlan = tlSlsList(ilSlsLoop).lPlan + tmBvfVeh(illoop).lGross(0)
        End If
        If ilLastWk Then
            tlSlsList(ilSlsLoop).lPlan = tlSlsList(ilSlsLoop).lPlan + tmBvfVeh(illoop).lGross(53)
        End If
    Next illoop

    'use startwk & endwk to gather projections
    gObtainWkNo ilPjfCalType, slTYStartQtr, ilStartWk, ilFirstWk          'determine the first week inx to accum (current year)
    gObtainWkNo ilPjfCalType, slTYEndQtr, ilEndWk, ilLastWk               'determine the last week inx to accum  (current year)
    'gather all Slsp projection records for the matching rollover date (exclude current records)
    ReDim tmTPjf(0 To 0) As PJF
    ilRet = gObtainPjf(RptSelCt, hmPjf, ilRODate(), tmTPjf())                'Read all applicable Projection records into memory

    'Build slsp projection $ just gathered into vehicle buckets
    For illoop = LBound(tmTPjf) To UBound(tmTPjf) Step 1
        ilPotnInx = 0
        For ilFound = 1 To 3 Step 1
            If tmTPjf(illoop).iMnfBus = ilLikeCode(ilFound) Then
                ilPotnInx = ilFound
                Exit For
            End If
        Next ilFound
        If ilPotnInx > 0 Then           'potential code exists
            'Determine start & end dates of the standard year from the proj recd
            slStr = "12/15/" & Trim$(str$(tmTPjf(illoop).iYear))
            llProjYearEnd = gDateValue(gObtainEndStd(slStr))
            ilSaveStartWk = ilStartWk
            ilSaveEndWk = ilEndWk
            If ilStartWk > ilEndWk Then                             'end wk isnt greater than start week, Must be corp going into next proj year
                If llTYDates(2) > llProjYearEnd Then                 'must be corp, the last week extends into the next year
                    ilEndWk = 53
                    'ilEndWk is the value of week to start from, and go to the end of this recds year
                Else
                    'ilStartWk is value of week to use
                    ilStartWk = 1
                End If
            End If
            For ilSlsLoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1
                If tlSlsList(ilSlsLoop).iVefCode = tmTPjf(illoop).iVefCode Then
                    llAdjust = 0
                    For ilTemp = ilStartWk To ilEndWk Step 1
                        llAdjust = llAdjust + tmTPjf(illoop).lGross(ilTemp)
                    Next ilTemp
                    If ilFirstWk Then       'adjust for the partial weeks at the beginning or end of the year
                                            'due to corp or calendar months
                        'llAdjust = llAdjust + tlSlsList(ilSlsLoop).lProj + tmTPjf(ilLoop).lGross(0)
                        llAdjust = llAdjust + tmTPjf(illoop).lGross(0)
                    End If
                    If ilLastWk Then
                        llAdjust = llAdjust + tmTPjf(illoop).lGross(53)
                    End If
                    llAdjust = (llAdjust * ilLikePct(ilPotnInx)) \ 100  'adjust the gross based on the potential codes most likely %
                    tlSlsList(ilSlsLoop).lProj = tlSlsList(ilSlsLoop).lProj + llAdjust
                    Exit For
                End If
            Next ilSlsLoop
            ilStartWk = ilSaveStartWk
            ilEndWk = ilSaveEndWk
        End If
    Next illoop
    slCntrTypes = gBuildCntTypes()      'Setup valid types of contracts to obtain based on user
    slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order.
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)

    ilRet = gObtainPhfRvf(RptSelCt, slLYStartQtr, slTYEndQtr, tlTranType, tlRvf(), 0)
    For llRvfLoop = LBound(tlRvf) To UBound(tlRvf) - 1 Step 1
        tmRvf = tlRvf(llRvfLoop)
        'dan M 7-30-08 added single contract selectivity
        If llSingleContract = NOT_SELECTED Or llSingleContract = tmRvf.lCntrNo Then
            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slCode
            llDate = gDateValue(slCode)
            ilTY = False
            ilFound = False
            llLYActualNoPace = 0

            gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slTemp
            llDateEntered = gDateValue(slTemp)
            'Dan M 8-15-8 ntr/hard cost adjustments.  Is this record ntr/hard cost and do we want that?
            blFailedMatchNtrOrHardCost = False
            'Dan M 8-15-8 don't allow installment option "I"
            blFailedBecauseInstallment = False
            If tmRvf.sType = "I" Then
                blFailedBecauseInstallment = True
            End If
            If ((blIncludeNTR) Xor (blIncludeHardCost)) And tmRvf.iMnfItem > 0 Then      'one or the other is true, but not both (if both true, don't have to isolate anything)
                ilRet = gIsItHardCost(tmRvf.iMnfItem, tmMnfNtr())
            'if is hard cost but blincludentr  or isn't hard cost but blincludehardcost then it needs to be removed. set failedmatchntrorhardcost true
                If (ilRet And blIncludeNTR) Or ((Not ilRet) And blIncludeHardCost) Then
                    blFailedMatchNtrOrHardCost = True
                End If
            End If
            If Not (blFailedMatchNtrOrHardCost Or blFailedBecauseInstallment) Then  'if both false, continue

                If llDate >= llTYDates(1) And llDate <= llTYDates(2) Then   'dan m changed < to <= 8-18-08
                    ilTY = True
                    'If llDate <= llTYGetTo Then    'replaced with below dan m 8-15-08
                    If llDateEntered <= llTYGetTo Then
                        ilFound = True
                        gPDNToLong tmRvf.sGross, llProject(1)           'theres only 1qtr to gather
                    End If
                'if trans date not within current year, assume last year
                Else
                    If llDate >= llLYDates(1) And llDate <= llLYDates(2) Then       'dan m changed < to <= 8-18-08
                        If llDateEntered <= llLYGetTo Then          'from lldate to lldateentered dan m 8-18-08
                            ilFound = True
                            gPDNToLong tmRvf.sGross, llProject(1)       'theres only 1 qtr to gather
                            gPDNToLong tmRvf.sGross, llLYActualNoPace
                            llLYActualNoPace = llLYActualNoPace / 100    'drop pennies
                        End If
                    End If
                End If
            End If
            If ilFound Then
                'Read the contract
                tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tgChfCT, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
                'Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo <> tmRvf.lCntrNo Or (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M"))
                Do While (ilRet = BTRV_ERR_NONE) And (tgChfCT.lCntrNo = tmRvf.lCntrNo) And (tgChfCT.sSchStatus <> "F" And tgChfCT.sSchStatus <> "M")
                    ilRet = btrGetNext(hmCHF, tgChfCT, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If ((ilRet <> BTRV_ERR_NONE) Or (tgChfCT.lCntrNo <> tmRvf.lCntrNo)) Then  'phoney a header from the receivables record so it can be procesed
                    For illoop = 0 To 9
                        tgChfCT.iSlfCode(illoop) = 0
                        tgChfCT.lComm(illoop) = 0
                    Next illoop
                    tgChfCT.iPctTrade = 0
                    If tmRvf.sCashTrade = "T" Then
                        tgChfCT.iPctTrade = 100           'ignore trades   later
                    End If
                End If
                'Accumulate the $ projected into the vehicles buckets
                If llProject(1) <> 0 Then                            'ignore building any data whose lines didnt have $
                    llProject(1) = llProject(1) \ 100               'drop pennies
                    ilFoundSls = False
                    Do While Not ilFoundSls
                        For ilSlsLoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1
                            If tlSlsList(ilSlsLoop).iVefCode = tmRvf.iBillVefCode Then
                                If Not ilTY Then               'past year  (holds & orders are combined)
                                    'if beyond the effective date, it's still actuals for the qtr
                                    ilFoundSls = True
                                    tlSlsList(ilSlsLoop).lLYWeek = tlSlsList(ilSlsLoop).lLYWeek + llProject(1)
                                    'this record is last years qtr, accum the actuals
                                    tlSlsList(ilSlsLoop).lLYAct = tlSlsList(ilSlsLoop).lLYAct + llLYActualNoPace
                                    Exit For
                                Else                                'current year, holds and orders are added together
                                    ilFoundSls = True
                                    tlSlsList(ilSlsLoop).lTYAct = tlSlsList(ilSlsLoop).lTYAct + llProject(1)
                                    Exit For
                                End If
                            End If
                        Next ilSlsLoop
                        If Not ilFoundSls Then              'there wasnt a budget for this vehicle to begin with,
                                                            'no entry has been created
                            tlSlsList(UBound(tlSlsList)).iVefCode = tmRvf.iBillVefCode
                            If Not ilTY Then        'past year
                                ilFoundSls = True
                                tlSlsList(ilSlsLoop).lLYWeek = tlSlsList(ilSlsLoop).lLYWeek + llProject(1)
                                'this record is last years qtr, accum the actuals
                                tlSlsList(ilSlsLoop).lLYAct = tlSlsList(ilSlsLoop).lLYAct + llLYActualNoPace
                            Else
                                ilFoundSls = True
                                tlSlsList(ilSlsLoop).lTYAct = tlSlsList(ilSlsLoop).lTYAct + llProject(1)
                            End If
                            ReDim Preserve tlSlsList(LBound(tlSlsList) To UBound(tlSlsList) + 1)
                        End If
                    Loop                    'loop until a vehicle budget has been found
                End If
            End If                          'ilfound
        End If              'single contract selectivity
    Next llRvfLoop

    'Process last year, then this year.  Get all contracts for the active quarter dates and project their $ from the flights.
    'Build array of possible contracts that fall into last year or this years quarter and build into array tlChfAdvtExt
    'ilret = gObtainCntrForDate(RptSelCt, slLYStartYr, slTYEndQtr, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
    slTYEndYr = Format$(gDateValue(slTYEndYr) + 90, "m/d/yy")        'get an extra quarter to make sure all changes included
    ilRet = gObtainCntrForDate(RptSelCt, slLYStartYr, slTYEndYr, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
            '7-30-08 added single contract selectivity dan M
        If llSingleContract = NOT_SELECTED Or llSingleContract = tlChfAdvtExt(ilCurrentRecd).lCntrNo Then

            For ilPastFut = 1 To 2 Step 1
                If ilPastFut = 1 Then                       'past
                    'slStartDate = Format$(llLYDates(1), "m/d/yy")       'gather all cntrs whose start/end dates fall within requested qtr (last year)
                    'slEndDate = Format$(llLYDates(2), "m/d/yy")
                    llStartDates(1) = llLYDates(1)
                    llStartDates(2) = llLYDates(2)
                    llEnterFrom = llLYGetFrom                           'gather all cntrs whose entered date falls within these dates
                    llEnterTo = llLYGetTo
                    'If processing last year, need two sets of actuals:  1 is all the actuals for last year for the pacing period
                    '(that is any contracts whose entered date is equal/prior to user entered date for last year),
                    'The other actuals are for the entire quarter of last year.
                    ilLYRepeat = 2
               Else                                         'current
                    'slStartDate = Format$(llTYDates(1), "m/d/yy")        'gather all cntrs whose start/end dates fall within requested qtr (this year)
                    'slEndDate = Format$(llTYDates(2), "m/d/yy")
                    llStartDates(1) = llTYDates(1)
                    llStartDates(2) = llTYDates(2)
                    llEnterFrom = llTYGetFrom           'gather cnts whose entered date falls within these dates
                    llEnterTo = llTYGetTo
                    ilLYRepeat = 1
                End If
                'project the $
                llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
                'Got the correct header that is equal or prior to the effective date entered
                llContrCode = gPaceCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llEnterTo, hmCHF, tmChf)
                If llContrCode > 0 Then
                    'Retrieve the contract, schedule lines and flights
                    ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT())   'get the latest version of this contract
                    ilFound = False
                    gUnpackDateLong tgChfCT.iOHDDate(0), tgChfCT.iOHDDate(1), llAdjust      'date entered
                    If ilPastFut = 2 Then       'if current, need to test entered date against the requested effective
                        If llAdjust <= llEnterTo Then       'entered date must be entered thru effectve date
                            ilFound = True
                        End If
                    Else                        'Past
                        ilFound = True          'past get all cnts affecting the qtr to get actuals as well as same wee last year
                    End If
                Else
                    ilFound = False
                End If
                For ilRepeat = 1 To ilLYRepeat              'Process last years data twice:  once for the same week last year,
                    If ilPastFut = 1 And ilRepeat = 2 Then                'if processing last year, and its the
                        If tlChfAdvtExt(ilCurrentRecd).lCode <> tgChfCT.lCode Then    'dont reread header if the most current version for last year is the
                                                                                    'same as that just processed for same week lastyear
                            llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
                            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT())   'get the latest version of this contract
                            gUnpackDateLong tgChfCT.iOHDDate(0), tgChfCT.iOHDDate(1), llAdjust      'date entered
                            ilFound = True
                        End If
                    End If
                    If ilFound And tgChfCT.iPctTrade <> 100 Then      'ignore contracts 100% trade, process all others
                                                                'once for actuals qtr totals
                        'Loop thru all lines and project their $ from the flights
                        For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                            llProject(1) = 0                'init bkts to accum qtr $ for this line
                            tmClf = tgClfCT(ilClf).ClfRec
                            If tmClf.sType = "S" Or tmClf.sType = "H" Then
                                gBuildFlights ilClf, llStartDates(), 1, 2, llProject(), 1, tgClfCT(), tgCffCT()
                            End If
                            'If slAirOrder = "O" Then                'invoice all contracts as ordered
                            '    If tmClf.sType <> "H" Then          'ignore all hidden lines for ordered billing, should be Pkg or conventional lines
                            '        gBuildFlights ilClf, llStartDates(), 1, 2, llProject(), 1
                            '    End If
                            'Else                                    'inv all contracts as aired
                            '    If tmClf.sType = "H" Then             'but if from pkg and hidden line, ignore hidd
                            '        'if hidden, will project if assoc. package is set to invoice as aired (real)
                            '        For ilTemp = LBound(tgClfCt) To UBound(tgClfCt) - 1    'find the assoc. pkg line for these hidden
                            '            If tmClf.iPkLineNo = tgClfCt(ilTemp).ClfRec.iLine Then
                            '                If tgClfCt(ilTemp).ClfRec.sType = "A" Then        'does the pkg line reflect bill as aired?
                            '                    gBuildFlights ilClf, llStartDates(), 1, 2, llProject(), 1 'pkg bills as aired, project the hidden line
                            '                End If
                            '                Exit For
                            '            End If
                            '        Next ilTemp
                            '    Else                            'conventional, VV, or Pkg line
                            '        If tmClf.sType <> "A" Then  'if this package line to be invoiced aired (real times),
                            '                                    'it has already been projected above with the hidden line
                            '            gBuildFlights ilClf, llStartDates(), 1, 2, llProject(), 1
                            '        End If
                            '    End If
                            'End If
                            'Accumulate the $ projected into the vehicles buckets
                            If llProject(1) > 0 Then                            'ignore building any data whose lines didnt have $
                                llProject(1) = llProject(1) \ 100               'drop pennies
                                ilFoundSls = False
                                Do While Not ilFoundSls
                                    For ilSlsLoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1
                                        If tlSlsList(ilSlsLoop).iVefCode = tmClf.iVefCode Then
                                            If ilPastFut = 1 Then               'past year  (holds & orders are combined)
                                                'if beyond the effective date, it's still actuals for the qtr
                                                ilFoundSls = True
                                                If ilRepeat = 1 Then              'do last years week actuals
                                                'If llAdjust <= llEnterTo Then    'entered date is equal/prior to same time last year
                                                                                 'show it in the same time last year column
                                                    tlSlsList(ilSlsLoop).lLYWeek = tlSlsList(ilSlsLoop).lLYWeek + llProject(1)
                                                Else                              'do last years actuals totals
                                                    'this record is last years qtr, accum the actuals
                                                    tlSlsList(ilSlsLoop).lLYAct = tlSlsList(ilSlsLoop).lLYAct + llProject(1)
                                                End If
                                                Exit For
                                            Else                                'current year, holds and orders are added together
                                                ilFoundSls = True
                                                If tgChfCT.sStatus = "H" Or tgChfCT.sStatus = "G" Then    'hold or unsch hold
                                                    tlSlsList(ilSlsLoop).lTYActHold = tlSlsList(ilSlsLoop).lTYActHold + llProject(1)
                                                    Exit For
                                                Else                            'order or unsch order
                                                    tlSlsList(ilSlsLoop).lTYAct = tlSlsList(ilSlsLoop).lTYAct + llProject(1)
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next ilSlsLoop
                                    If Not ilFoundSls Then              'there wasnt a budget for this vehicle to begin with,
                                                                        'no entry has been created
                                        tlSlsList(UBound(tlSlsList)).iVefCode = tmClf.iVefCode
                                        ReDim Preserve tlSlsList(LBound(tlSlsList) To UBound(tlSlsList) + 1)
                                    End If
                                Loop                    'loop until a vehicle budget has been found
                            End If
                        Next ilClf                      'process nextline
                         ' Dan M 7-30-08 Add NTR/Hard Cost option
                        'step one. get ntr contracts and prepare loop   Does user want to see HardCost/NTR? ilfound(entry date check), and not a trade already done
                        If (blIncludeNTR Or blIncludeHardCost) And (tgChfCT.sNTRDefined = "Y") Then
                           'call routine to fill array with choice: ntr, hard cost, both.  changed startDates to what using above.
                           gNtrByContract llContrCode, llStartDates(1), llStartDates(2), tlNTRInfo(), tmMnfNtr(), hmSbf, blIncludeNTR, blIncludeHardCost, RptSelCt
                           ilLowerboundNTR = LBound(tlNTRInfo)
                           ilUpperboundNTR = UBound(tlNTRInfo)
                        'ntr or hard cost found?
                            If ilUpperboundNTR <> ilLowerboundNTR Then
                        'step two. process each element in loop like above.
                                For ilNTRCounter = ilLowerboundNTR To ilUpperboundNTR - 1 Step 1
                                        llProject(1) = 0
                                        llProject(1) = tlNTRInfo(ilNTRCounter).lSBFTotal
                                        'this ntr/hardcost has a value to write to our totals field.  Following is copied nearly verbatim from clf above.
                                        If llProject(1) > 0 Then                            'ignore building any data whose lines didnt have $
                                            llProject(1) = llProject(1) \ 100               'drop pennies
                                            ilFoundSls = False
                                            Do While Not ilFoundSls
                                                For ilSlsLoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1
                                                    If tlSlsList(ilSlsLoop).iVefCode = tlNTRInfo(ilNTRCounter).iVefCode Then
                                                        If ilPastFut = 1 Then               'past year  (holds & orders are combined)
                                                            'if beyond the effective date, it's still actuals for the qtr
                                                            ilFoundSls = True
                                                            If ilRepeat = 1 Then              'do last years week actuals
                                                            'If llAdjust <= llEnterTo Then    'entered date is equal/prior to same time last year
                                                                                             'show it in the same time last year column
                                                                tlSlsList(ilSlsLoop).lLYWeek = tlSlsList(ilSlsLoop).lLYWeek + llProject(1)
                                                            Else                              'do last years actuals totals
                                                                'this record is last years qtr, accum the actuals
                                                                tlSlsList(ilSlsLoop).lLYAct = tlSlsList(ilSlsLoop).lLYAct + llProject(1)
                                                            End If
                                                            Exit For
                                                        Else                                'current year, holds and orders are added together
                                                            ilFoundSls = True
                                                            If tgChfCT.sStatus = "H" Or tgChfCT.sStatus = "G" Then    'hold or unsch hold
                                                                tlSlsList(ilSlsLoop).lTYActHold = tlSlsList(ilSlsLoop).lTYActHold + llProject(1)
                                                                Exit For
                                                            Else                            'order or unsch order
                                                                tlSlsList(ilSlsLoop).lTYAct = tlSlsList(ilSlsLoop).lTYAct + llProject(1)
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                Next ilSlsLoop
                                                If Not ilFoundSls Then              'there wasnt a budget for this vehicle to begin with,
                                                                                    'no entry has been created
                                                    tlSlsList(UBound(tlSlsList)).iVefCode = tlNTRInfo(ilNTRCounter).iVefCode
                                                    ReDim Preserve tlSlsList(LBound(tlSlsList) To UBound(tlSlsList) + 1)
                                                End If
                                            Loop                    'loop until a vehicle budget has been found
                                    End If                      'end copied from above..llproject(1) has a value
                                Next ilNTRCounter
                            End If      'upperbound greater than lowerbound ntr--found an ntr
                        End If          'want to see ntr/hard cost and this contract has one?
                        'end ntr addition
                    End If                              'llAdjust falls within requested dates
                Next ilRepeat                   'Process last year twice
            Next ilPastFut
        End If          'single contract selectivity
    Next ilCurrentRecd
    Erase tlChfAdvtExt                      'Make sure last years contrcts are erased, go process this year
    'Build up fields in GRF that need to be set up only once.  These fields do not change across records.
    'Setup last year's qtr column heading
    'ilYear contains starting year
    ilMonth = RptSelCt!edcSelCTo1.Text              'month
    slDate = Trim$(str$(((ilMonth - 1) * 3 + 1))) & "/15/" & Trim$(str$(ilYear))
    slDate = gObtainStartStd(slDate)
    If ilMonth = 1 Then
        tmGrf.sGenDesc = "1st"
    ElseIf ilMonth = 2 Then
        tmGrf.sGenDesc = "2nd"
    ElseIf ilMonth = 3 Then
        tmGrf.sGenDesc = "3rd"
    Else
        tmGrf.sGenDesc = "4th"
    End If
    tmGrf.sGenDesc = Trim$(tmGrf.sGenDesc) & " Qtr" & str$(ilYear - 1)    'add Year

    slDate = Format$(llLYGetTo, "m/d/yy")
    gPackDate slDate, ilMonth, ilYear
    tmGrf.iDate(0) = ilMonth                'last year's week (for last years column heading)
    tmGrf.iDate(1) = ilYear
    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
    'tmGrf.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    tmGrf.iStartDate(0) = ilEnterDate(0)     'effective date entered
    tmGrf.iStartDate(1) = ilEnterDate(1)
    tmGrf.iCode2 = ilBdMnfCode                          'budget name
    'tmGrf.iPerGenl(1) = ilCorpStd             '1 = corp, 2 = std
    tmGrf.iPerGenl(0) = ilCorpStd             '1 = corp, 2 = std


    'Loop thru the vehicle table (tlSlsList) with all the gathered budget info, projected info, Last year info and
    'write out the prepaFileListBox to disk (one per vehicle)
    For illoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1         'write a record per vehicle
        'If tlSlsList(ilLoop).lPlan + tlSlsList(ilLoop).lTYAct + tlSlsList(ilLoop).lProj + tlSlsList(ilLoop).lLYWeek + tlSlsList(ilLoop).lLYAct <> 0 Then
        '7-11-01 test TYActHold and TYAct
        If tlSlsList(illoop).lPlan + tlSlsList(illoop).lTYActHold + tlSlsList(illoop).lTYAct + tlSlsList(illoop).lProj + tlSlsList(illoop).lLYWeek + tlSlsList(illoop).lLYAct <> 0 Then
            tmGrf.iVefCode = tlSlsList(illoop).iVefCode
            'tmGrf.lDollars(1) = tlSlsList(ilLoop).lPlan         'current year, plan $
            'tmGrf.lDollars(2) = tlSlsList(ilLoop).lTYAct        'current year, orders
            'tmGrf.lDollars(3) = tlSlsList(ilLoop).lTYActHold    'current year, holds
            'tmGrf.lDollars(4) = tlSlsList(ilLoop).lProj         'current rollover
            'tmGrf.lDollars(5) = tlSlsList(ilLoop).lLYWeek       'last years, same week
            'tmGrf.lDollars(6) = tlSlsList(ilLoop).lLYAct        'last years, actuals
            tmGrf.lDollars(0) = tlSlsList(illoop).lPlan         'current year, plan $
            tmGrf.lDollars(1) = tlSlsList(illoop).lTYAct        'current year, orders
            tmGrf.lDollars(2) = tlSlsList(illoop).lTYActHold    'current year, holds
            tmGrf.lDollars(3) = tlSlsList(illoop).lProj         'current rollover
            tmGrf.lDollars(4) = tlSlsList(illoop).lLYWeek       'last years, same week
            tmGrf.lDollars(5) = tlSlsList(illoop).lLYAct        'last years, actuals
            gGetVehGrpSets tmGrf.iVefCode, ilMinorSet, ilMajorSet, ilmnfMinorCode, ilMnfMajorCode
            'tmGrf.iPerGenl(2) = ilmnfMinorCode
            'tmGrf.iPerGenl(3) = ilMnfMajorCode
            tmGrf.iPerGenl(1) = ilmnfMinorCode
            tmGrf.iPerGenl(2) = ilMnfMajorCode
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        End If
    Next illoop
    sgCntrForDateStamp = ""         'clear out time stamp to reread contr if not exiting report selectivity
    Erase tlSlsList, tlMMnf
    Erase tmTPjf, tlChfAdvtExt, tmBvfVeh
    Erase tlRvf
    ilRet = btrClose(hmBvf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmMnf)
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        btrDestroy hmPjf
        btrDestroy hmMnf
End Sub
























'
'
'
'                       gCrVehCPPCPM - Generate CPPs or CPM for week specified
'                           based on the Rate card.  The book used is retrieve
'                           from the vehicle table (book last used).
'                           User may select all demos or selected demos for
'                           one or more vehicles.  the Rate Card is determined
'                           by whatever card is effective based on the entered date.
'
'                   5/13/98 - change to use rating book from the vehicle table (last one
'                           referenced).  Show list of rate cards to choose from rather
'                           than user having to type in name.  Replace the list of
'                           rating books with the rate card list.
'                   10-17-03 fix subscript out of range when the start date of the rate card (terms)
'                           is not in the same year as the rate card year (i.e. RC year 2003 and the
'                           start date stored in RC is 12/30/03 (should be 12/30/02)
'                   6-2-04 Implement research demo estimates
'                   6-30-04 Demo headers did not appear on report when the daypart didnt have any
'                           values
'                   9-11-04 Additional audience magnitudes to thousands & hundreds :  add tens and units
Sub gCrVehCPPCPM()
Dim llRif As Long                'loop variable for DP Rates
Dim ilTest As Integer
Dim ilFound As Integer
Dim ilFoundAgain As Integer
Dim ilDay As Integer
ReDim ilValidDays(0 To 6) As Integer
Dim llTPrice As Long
Dim ilNoWks As Integer
Dim llDate As Long
Dim slDate As String
Dim llWkPrice As Long
Dim ilRCCode As Integer
Dim ilRdf As Integer                'loop variable for DP
Dim ilDnfCode As Integer
Dim ilMnfDemo As Integer          'demo name code into mnf
Dim ilDemoLoop As Integer            'Index into Demo processing
Dim ilBookInx As Integer            'loop to create ANR records form BOOKGEN array
Dim ilRet As Integer
Dim slName As String
Dim ilEffYear As Integer             'year of rate card
Dim llEffDate As Long               'Effectve date entered
ReDim ilEffDate(0 To 1) As Integer    'effective date entered format for Crystal
Dim llRCStartDate As Long           'Rate Card Start Date
Dim llEffEndDate As Long            'effective end date (currently only 1 week span)
Dim slStr As String
Dim illoop As Integer
Dim ilRCWkNo As Integer             'week processing to gather rates from rif (currently only 1 week to process)
Dim slCode As String
Dim ilVeh As Integer              'loop variable for the vehicles to process
Dim ilSaveVeh As Integer            'vehicle code processing
'ReDim ilVehicles(1 To 1) As Integer 'array of valid vehicles to process
ReDim ilVehicles(0 To 0) As Integer 'array of valid vehicles to process
ReDim ilDemoList(0 To 0) As Integer  'array of demo categories to process
''ReDim tmMRif(0 To 0) As RIF
'ReDim tmMRif(1 To 1) As RIF
ReDim tmMRif(0 To 0) As RIF
ReDim tmBookGen(0 To 0) As BOOKGEN
ReDim tmTAnr(0 To 0) As ANR
'ReDim ilDPList(1 To 1) As Integer       'list of dayparts generated for vehicle processing
ReDim ilDPList(0 To 0) As Integer       'list of dayparts generated for vehicle processing
Dim slSaveBase As String
'the following variables are for the routines to retrieve the rating data
''ggetdemoAvgAud and gAvgAudToLnResearch

'ReDim ilWkSpotCount(1 To 1) As Integer
'ReDim llWkActPrice(1 To 1) As Long
'ReDim llWkAvgAud(1 To 1) As Long

Dim llPop As Long
Dim ilMnfSocEco As Integer
Dim llOvStartTime As Long
Dim llOvEndTime As Long
Dim llAvgAud As Long
Dim llPopEst As Long
Dim ilAdjustAudData As Integer  'multiplication factor for audience data magnitude (spfauddata)
Dim llRafCode As Long
Dim ilAudFromSource As Integer
Dim llAudFromCode As Long
Dim slGrossNet As String * 1        '1-30-19 implement gross net option
Dim llRate As Long
Dim slAmount As String
Dim slSharePct As String


'**** end of varaibles required for Research routines
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    hmRcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRcf, "", sgDBPath & "Rcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRcf)
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    imRcfRecLen = Len(tmRcf)
    hmRif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRif, "", sgDBPath & "Rif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRif)
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    imRifRecLen = Len(tmRif)
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRdf)
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    imRdfRecLen = Len(tmRdf)
    hmDrf = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        btrDestroy hmAnr
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    imDrfRecLen = Len(tmDrf)
    hmAnr = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnr, "", sgDBPath & "Anr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAnr)
        btrDestroy hmAnr
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    imAnrRecLen = Len(tmAnr)
    hmDpf = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDpf)
        btrDestroy hmDpf
        btrDestroy hmAnr
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    imDpfRecLen = Len(tmDpf)
    hmDef = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmDef
        btrDestroy hmDpf
        btrDestroy hmAnr
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    hmRaf = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmRaf
        btrDestroy hmDef
        btrDestroy hmDpf
        btrDestroy hmAnr
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    'setup global variable for Demo plus table (if any exists)
    lgDpfNoRecs = btrRecords(hmDpf)
    If lgDpfNoRecs = 0 Then
        lgDpfNoRecs = -1
    End If

    ilAdjustAudData = 1
    If RptSelCt!rbcSelC4(1).Value Then      'option by cpm
        If tgSpf.sSAudData = "H" Then             'magnitutde hundreds vs thousands
            ilAdjustAudData = 10
        ElseIf tgSpf.sSAudData = "N" Then   'tens
            ilAdjustAudData = 100
        ElseIf tgSpf.sSAudData = "U" Then    'units
            ilAdjustAudData = 1000
        End If
    End If


    slStr = RptSelCt!CSI_CalFrom.Text   'Date: 12/12/2019 added CSI calendar control for date entry --> edcSelCFrom.Text           'effective date
    'insure its a Monday
    llEffDate = gDateValue(slStr)
    ilDay = gWeekDayLong(llEffDate)
    Do While ilDay <> 0
        llEffDate = llEffDate - 1
        ilDay = gWeekDayLong(llEffDate)
    Loop
    slStr = Format$(llEffDate, "m/d/yy")
    gPackDate slStr, ilEffDate(0), ilEffDate(1)

    'get sunday date for the R/C year
    llEffEndDate = llEffDate + 6
    slStr = Format$(llEffEndDate, "m/d/yy")              'Sunday will always be the correct year since we're dealing with Standard month
    gPackDate slStr, ilDay, ilEffYear
    ilRet = gObtainRcfRifRdf()                'bring in 3 files in global arrays
    ilRCCode = 0                                'no r/c definition yet

    slGrossNet = "G"                                '1-30-19 implement gross or net
    If RptSelCt!rbcSelC11(1).Value Then           'net selected
        slGrossNet = "N"
    End If


    For ilTest = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
        tmRcf = tgMRcf(ilTest)
        For illoop = 0 To RptSelCt!lbcSelection(4).ListCount - 1 Step 1
            slStr = tgRateCardCode(illoop).sKey
            ilRet = gParseItem(slStr, 3, "\", slCode)
            If Val(slCode) = tgMRcf(ilTest).iCode Then
                If (RptSelCt!lbcSelection(4).Selected(illoop)) Then
                    ilRCCode = tgMRcf(ilTest).iCode
                    gUnpackDate tgMRcf(ilTest).iStartDate(0), tgMRcf(ilTest).iStartDate(1), slName
                    llRCStartDate = gDateValue(slName)
                    ilTest = UBound(tgMRcf)
                    slName = tgRateCardCode(illoop).sKey
                    ilRet = gParseItem(slName, 2, "\", slName)
                    If Not gSetFormula("RateCardUsed", "'" & slName & "'") Then
                        Exit Sub
                    End If
                    Exit For
                End If
            End If
        Next illoop
    Next ilTest

    'find matching R/c from list box to the one entered.  If found a match, retrieve the RC Code
    'For ilLoop = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
    '    slStr = RptSelCt!edcSelCFrom1.Text
    '    If Trim$(slStr) = Trim$(tgMRcf(ilLoop).sName) Then
    '        ilRCCode = tgMRcf(ilLoop).iCode
    '        gUnPackDate tgMRcf(ilLoop).iStartDate(0), tgMRcf(ilLoop).iStartDate(1), slName
    '        llRCStartDate = gDateValue(slName)
    '        Exit For
    '    End If
    'Next ilLoop
    'get Book pointer from the Book selected
    'For ilLoop = 0 To RptSelCt!lbcSelection(4).ListCount - 1 Step 1
    '    If (RptSelCt!lbcSelection(4).Selected(ilLoop)) Then
    '        slName = tgBookName(ilLoop).sKey 'Traffic!lbcVehicle.List(ilVehicle)
    '        ilRet = gParseItem(slName, 1, "\", slStr)
    '        ilRet = gParseItem(slStr, 3, "|", slStr)
    '        ilRet = gParseItem(slName, 2, "\", slCode)
    '        ilDnfCode = Val(slCode)             'Book name pointer
    '        Exit For
    '    End If
    'Next ilLoop
    ilRet = gObtainVef()                'retrieve all the vehicles for the latest Booked used
    If Not ilRet Then
        ilRet = btrClose(hmAnr)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmRif)
        ilRet = btrClose(hmRcf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmVef)
        btrDestroy hmRaf
        btrDestroy hmDef
        btrDestroy hmDpf
        btrDestroy hmAnr
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If

    'Build array (tmMRif) of all valid Rates for each Vehicle's daypart to cut down on amount of processing
    For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
        llTPrice = 0
        For ilTest = 0 To 53 Step 1
            llTPrice = llTPrice + tgMRif(llRif).lRate(ilTest)
        Next ilTest
        If (ilRCCode = tgMRif(llRif).iRcfCode) And (ilEffYear = tgMRif(llRif).iYear) And (llTPrice > 0) Then
            'test for selective vehicle
            For illoop = 0 To RptSelCt!lbcSelection(3).ListCount - 1 Step 1
                If (RptSelCt!lbcSelection(3).Selected(illoop)) Then
                    slName = tgVehicle(illoop).sKey 'Traffic!lbcVehicle.List(ilVehicle)
                    ilRet = gParseItem(slName, 1, "\", slStr)
                    ilRet = gParseItem(slStr, 3, "|", slStr)
                    ilRet = gParseItem(slName, 2, "\", slCode)
                    'Build vehicle table of ones to process
                    ilFound = False
                    For ilTest = LBound(ilVehicles) To UBound(ilVehicles) - 1 Step 1
                        If ilVehicles(ilTest) = Val(slCode) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilTest
                    If Not ilFound Then
                        ilVehicles(UBound(ilVehicles)) = Val(slCode)
                        ReDim Preserve ilVehicles(LBound(ilVehicles) To UBound(ilVehicles) + 1)
                    End If
                    'If vehicle code matches daypart rates vehicle code, save the rate image
                    If Val(slCode) = tgMRif(llRif).iVefCode Then
                        tmMRif(UBound(tmMRif)) = tgMRif(llRif)
                        'ReDim Preserve tmMRif(1 To UBound(tmMRif) + 1) As RIF
                        ReDim Preserve tmMRif(0 To UBound(tmMRif) + 1) As RIF
                        Exit For    'ilLoop = RptSelCt!lbcSelection(3).ListCount 'stop the loop
                    End If
                End If
            Next illoop
        End If
    Next llRif
    'Build table of all demos that will be obtained.  This list corresponds to the
    '13 set of buckets sent in ANR (ie. the 1st 13 demos will always be in the same
    'relative index of anr.ipctsellout; the next 13 will always be in the same
    'relative index of anr.ipctsellout; etc.)
    For ilDemoLoop = 0 To RptSelCt!lbcSelection(2).ListCount - 1 Step 1
        If (RptSelCt!lbcSelection(2).Selected(ilDemoLoop)) Then
            slName = tgRptSelDemoCodeCT(ilDemoLoop).sKey
            ilRet = gParseItem(slName, 2, "\", slCode)
            ilDemoList(UBound(ilDemoList)) = Val(slCode)                     'mnf code to Demo name
            ReDim Preserve ilDemoList(0 To UBound(ilDemoList) + 1)
        End If
    Next ilDemoLoop
    'Process 1 vehicle at a time - get cpp/cpm for all dayparts and all demos rquested.
    'Then build 1 record for each set of 13 demos by daypart and vehicle.
    For ilVeh = LBound(ilVehicles) To UBound(ilVehicles) - 1 Step 1
        ilSaveVeh = ilVehicles(ilVeh)
        For ilDemoLoop = 0 To UBound(ilDemoList) - 1 Step 1
        ilMnfDemo = ilDemoList(ilDemoLoop)
        ilDnfCode = 0
        'For ilBookInx = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If tgMVef(ilBookInx).iCode = ilSaveVeh Then
            ilBookInx = gBinarySearchVef(ilSaveVeh)
            If ilBookInx <> -1 Then
                ilDnfCode = tgMVef(ilBookInx).iDnfCode
        '        Exit For
            End If
        'Next ilBookInx

        'get population for each demo
        ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, ilMnfSocEco, ilMnfDemo, llPop)
            For llRif = LBound(tmMRif) To UBound(tmMRif) - 1 Step 1
                If (tmMRif(llRif).iVefCode = ilSaveVeh) And (tmMRif(llRif).iRcfCode = ilRCCode) Then
                    'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                    ilRdf = gBinarySearchRdf(tmMRif(llRif).iRdfCode)
                    If ilRdf <> -1 Then
                    If tmMRif(llRif).sBase <> "Y" And tmMRif(llRif).sBase <> "N" Then
                        slSaveBase = tgMRdf(ilRdf).sBase
                    Else
                        slSaveBase = tmMRif(llRif).sBase
                    End If
                    'Matching vehicle and Rate card found--
                    'this Rate items record must match the DP code found, and be part of a base DP that is not dormant
                    If tmMRif(llRif).iRdfCode = tgMRdf(ilRdf).iCode And Trim$(slSaveBase) <> "N" And tgMRdf(ilRdf).sState <> "D" Then
                        ilFound = False
                        For ilTest = LBound(tmBookGen) To UBound(tmBookGen) - 1 Step 1
                            If (tmBookGen(ilTest).iRdfCode = tmMRif(llRif).iRdfCode) And (tmBookGen(ilTest).iVefCode = ilSaveVeh) And (tmBookGen(ilTest).iMnfDemo = ilMnfDemo) Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilTest
                        If Not ilFound Then
                            'Build record into tmBookGen
                            tmBookGen(UBound(tmBookGen)).lPop = llPop
                            tmBookGen(UBound(tmBookGen)).iMnfDemo = ilMnfDemo
                            tmBookGen(UBound(tmBookGen)).iRdfCode = tmMRif(llRif).iRdfCode
                            tmBookGen(UBound(tmBookGen)).iVefCode = ilSaveVeh
                            'Get price
                            llTPrice = 0
                            ilNoWks = 0
                            For llDate = llEffDate To llEffEndDate Step 7
                                slDate = Format$(llDate, "m/d/yy")
                                ilRCWkNo = (llEffDate - llRCStartDate) \ 7 + 1
                                If ilRCWkNo = 1 Then
                                    llWkPrice = tmMRif(llRif).lRate(0) + tmMRif(llRif).lRate(1)
                                ElseIf ilRCWkNo = 52 Then
                                    llWkPrice = tmMRif(llRif).lRate(52) + tmMRif(llRif).lRate(53)
                                ElseIf ilRCWkNo > 1 And ilRCWkNo < 52 Then
                                    llWkPrice = tmMRif(llRif).lRate(ilRCWkNo)
                                End If

                                If slGrossNet = "N" Then
                                    slAmount = gLongToStrDec(llWkPrice, 2)
                                    slSharePct = gIntToStrDec(8500, 4)
                                    slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
                                    slStr = gRoundStr(slStr, ".01", 2)
                                    llWkPrice = gStrDecToLong(slStr, 2)                     'adjusted net
                                End If
                                
                                llTPrice = llTPrice + llWkPrice         'no pennies
                                ilNoWks = ilNoWks + 1
                            Next llDate
                            If ilNoWks > 0 Then
                                tmBookGen(UBound(tmBookGen)).lAvgPrice = llTPrice / ilNoWks
                            End If

                            'generate valid days this daypart is airing (pass to Audience routines)
                            For illoop = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                If (tgMRdf(ilRdf).iStartTime(0, illoop) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, illoop) <> 0) Then
                                    For ilDay = 0 To 6 Step 1
                                        'If (tgMRdf(ilRdf).sWkDays(ilLoop, ilDay + 1) = "Y") Then
                                        If (tgMRdf(ilRdf).sWkDays(illoop, ilDay) = "Y") Then
                                            ilValidDays(ilDay) = True
                                        Else
                                            ilValidDays(ilDay) = False
                                        End If
                                    Next ilDay
                                End If
                            Next illoop

                            If (ilDnfCode > 0) Then             'book exists
                                llRafCode = 0
                                ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, ilSaveVeh, ilMnfSocEco, ilMnfDemo, llEffDate, llEffEndDate, tmMRif(llRif).iRdfCode, llOvStartTime, llOvEndTime, ilValidDays(), "S", llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                tmBookGen(UBound(tmBookGen)).lAvgAud = llAvgAud
                                If tgSpf.sDemoEstAllowed = "Y" Then     'use estimates?
                                    tmBookGen(UBound(tmBookGen)).lPop = llPopEst
                                End If

                                ''Get Rating, avg audience , cpp, cpm
                                'ilWkSpotCount(1) = 1
                                'llWkActPrice(1) = tmBookGen(UBound(tmBookGen)).lAvgPrice
                                'llWkAvgAud(1) = llAvgAud

                                'If ((tmBookGen(UBound(tmBookGen)).lAvgAud > 0) And (tmBookGen(UBound(tmBookGen)).lAvgPrice > 0)) Then
                                    ilFoundAgain = False
                                    For illoop = LBound(ilDPList) To UBound(ilDPList) - 1 Step 1
                                        If ilDPList(illoop) = tmBookGen(UBound(tmBookGen)).iRdfCode Then
                                            ilFoundAgain = True
                                            Exit For
                                        End If
                                    Next illoop
                                    If Not ilFoundAgain Then
                                        ilDPList(UBound(ilDPList)) = tmBookGen(UBound(tmBookGen)).iRdfCode
                                        ReDim Preserve ilDPList(LBound(ilDPList) To UBound(ilDPList) + 1) As Integer
                                    End If
                                    ReDim Preserve tmBookGen(0 To UBound(tmBookGen) + 1) As BOOKGEN
                               ' End If
                            End If              'ilDnfCode > 0
                        End If                  'Not ilFound
                    End If                      'tmMRif(ilRif).iRdfCode = tgMRdf(ilRdf).icode And Trim$(slSaveBase) <> "N" And tgMRdf(ilRdf).sState <> "D"
                    'Next ilRdf                  'next DP
                    End If
                End If
            Next llRif                          'next DP rate
        Next ilDemoLoop
        'Vehicle is complete - Create ANR file from tmBookGen array
        'Outer loop - DAypart table - loop thru the unique dayparts for the vehicle and find  all the demos for that DP.
        'Create 1 record/daypart/vehicle for a max of 13 demos per record.
        'Create array of ANR records all built in memory for all demos for this 1 daypart before writing to disk.
        For ilRdf = LBound(ilDPList) To UBound(ilDPList) - 1 Step 1
            For ilBookInx = 0 To UBound(tmBookGen) - 1 Step 1           'loop thru all the demos for each daypart for this vehicle to create 1 record for a set of 13 demos
                If tmBookGen(ilBookInx).iRdfCode = ilDPList(ilRdf) Then
                    For illoop = 0 To UBound(ilDemoList) - 1 Step 1
                        If ilDemoList(illoop) = tmBookGen(ilBookInx).iMnfDemo Then
                            Exit For                    'obtain the index to this demo in demo array.  The demo from the current tmBookGen
                                                        'record must be the the same index across all dayparts
                        End If
                    Next illoop
                    If tmBookGen(ilBookInx).iMnfDemo = ilDemoList(illoop) Then
                        ilDay = illoop \ 13 + 1            'relative ANR record(+1), may need multiples images if more than 13 demos
                        If ilDay > UBound(tmTAnr) Then  'need to allocate the record
                            tmAnr = tmZeroAnr
                            tmTAnr(ilDay - 1).iGenDate(0) = igNowDate(0)
                            tmTAnr(ilDay - 1).iGenDate(1) = igNowDate(1)
                            'tmTAnr(ilDay - 1).iGenTime(0) = igNowTime(0)
                            'tmTAnr(ilDay - 1).iGenTime(1) = igNowTime(1)
                            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                            tmTAnr(ilDay - 1).lGenTime = lgNowTime
                            tmTAnr(ilDay - 1).iEffectiveDate(0) = ilEffDate(0)
                            tmTAnr(ilDay - 1).iEffectiveDate(1) = ilEffDate(1)
                            tmTAnr(ilDay - 1).iVefCode = tmBookGen(ilBookInx).iVefCode
                            tmTAnr(ilDay - 1).iRdfCode = tmBookGen(ilBookInx).iRdfCode
                            'tmTAnr(ilDay - 1).lRCPrice(1) = tmBookGen(ilBookInx).lAvgPrice * ilAdjustAudData
                            tmTAnr(ilDay - 1).lRCPrice(0) = tmBookGen(ilBookInx).lAvgPrice * ilAdjustAudData
                            tmTAnr(ilDay - 1).iMnfBudget = ilDnfCode
                            ReDim Preserve tmTAnr(0 To UBound(tmTAnr) + 1) As ANR
                        End If
                        ilFound = (illoop Mod 13) + 1    'get remainder to determine which index this demo (1-13) will be placed in
                        tmTAnr(ilDay - 1).lUpfPrice(ilFound - 1) = tmBookGen(ilBookInx).lAvgAud
                        tmTAnr(ilDay - 1).lMinPrice(ilFound - 1) = tmBookGen(ilBookInx).lCPP
                        tmTAnr(ilDay - 1).lMaxPrice(ilFound - 1) = tmBookGen(ilBookInx).lCPM
                        tmTAnr(ilDay - 1).lScatPrice(ilFound - 1) = tmBookGen(ilBookInx).lPop
                        'ilLoop = the mnfDemo to be used
                        ilTest = (illoop \ 13)       'find the set (of 13) within all demos selected
                        tmTAnr(ilDay - 1).iPctSellout(ilFound - 1) = ilDemoList((ilTest * 13) + ilFound - 1)
                        'tmTAnr(ilDay - 1).iPctSellout(1) = ilDemoList(ilTest * 13)      'always place the 1st group of 13 demos in the first demo name -
                        tmTAnr(ilDay - 1).iPctSellout(0) = ilDemoList(ilTest * 13)      'always place the 1st group of 13 demos in the first demo name -
                    End If
                End If
            Next ilBookInx
            For illoop = 0 To UBound(tmTAnr) - 1 Step 1
                'insure the demo column headers are set up.  If a daypart doesnt have data, a column header may be missing
                'ilMnfDemo = tmTAnr(ilLoop).iPctSellout(1)           '1st demo in group
                ilMnfDemo = tmTAnr(illoop).iPctSellout(0)           '1st demo in group
                For ilFoundAgain = LBound(ilDemoList) To UBound(ilDemoList) - 1
                    'look for the matching demo pointer in the list of selected demos to create headers
                    If ilMnfDemo = ilDemoList(ilFoundAgain) Then
                        For ilRCWkNo = 1 To 13
                            If ilFoundAgain + ilRCWkNo <= UBound(ilDemoList) Then
                                tmTAnr(illoop).iPctSellout(ilRCWkNo - 1) = ilDemoList(ilFoundAgain + ilRCWkNo - 1)
                            Else
                                tmTAnr(illoop).iPctSellout(ilRCWkNo - 1) = 0
                            End If
                        Next ilRCWkNo
                        Exit For
                    End If
                Next ilFoundAgain

                ilRet = btrInsert(hmAnr, tmTAnr(illoop), imAnrRecLen, INDEXKEY0)
            Next illoop
            ReDim Preserve tmTAnr(0 To 0) As ANR
        Next ilRdf                          'creat next 13 demos for DP
        ReDim tmBookGen(0 To 0) As BOOKGEN   'initialize for the next demo
        'ReDim Preserve ilDPList(1 To 1) As Integer      'initialize for next demo
        ReDim Preserve ilDPList(0 To 0) As Integer      'initialize for next demo
    Next ilVeh                            'next vehicle for next daypart
    Erase tmBookGen, tmMRif, ilVehicles, ilDPList
    ilRet = btrClose(hmAnr)
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmRif)
    ilRet = btrClose(hmRcf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmDrf)
    ilRet = btrClose(hmDpf)
    ilRet = btrClose(hmDef)
    btrDestroy hmAnr
    btrDestroy hmRdf
    btrDestroy hmRif
    btrDestroy hmRcf
    btrDestroy hmMnf
    btrDestroy hmVef
    btrDestroy hmDrf
    btrDestroy hmDpf
    btrDestroy hmDef
    btrDestroy hmRaf
End Sub
'
'                   mCumeInsert - build GRF record for
'                   cumulative Activity Report.  Table of vehicles
'                   built in memory containing vehicle code and
'                   12 monthly buckets.  this routine loops through
'                   the table and creates 1 record per vehicle per
'                   cash and trade to disk.
'                   Zero $ records are not created.
'                   <input>  ilListIndex = report option:  Daily Ssales Activity, Sales Placement, Weekly Sales Activity
'                            ilMonths - # of periods requested
'                             tmVefDollars - table in memory of the vehicles gathred
'                                   for the contract (all gross $)
'                            ilSaveSof - sales source from contracts slsp
'                            ilMajorSet - vehicle group
'                   <output> tmGrf - GRF record created to disk

'       tmgrf.igendate - generation date
'       tmgrf.igentime - generation time
'       tmgrf.iPerGenl(4) = vehicle group (Daily Sales Activity by Month)
'       tmgrf.lchfcode - contract code
'       tmgrf.ivefcode - vehicle code
'       tmgrf.isofcode = sales office code
'       tmgrf.islfcode = salesperson code
'       tmgrf.sDateType = cash (C) or trade (T)
'       tmgrf.lDollars(1-12) - Jan thru Dec projections
'
'                   6/9/97 d.Hosaka
'       12-7-06 Only Writeout # of months requested so that the totals across the months for an advertiser
'               are not overstated
Sub mCumeInsert(ilListIndex As Integer, ilPeriods As Integer, tmVefDollars() As ADJUSTLIST, ilSaveSof As Integer, ilMajorSet As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPct                         ilAgyComm                     llTemp                    *
'*  llNoPenny                                                                             *
'******************************************************************************************

Dim ilTemp As Integer               'Loop variable for number of vehicles to process
Dim llGross As Long                 'total of all 12 months each vehicle, if zero record not written
Dim ilTemp2 As Integer              'Cash/trade loop variable
Dim illoop As Integer               'loop variable for 12 months
Dim ilRet As Integer                'error return from btrieve
Dim ilZero As Integer            'write out the record even if the 12 months balance to zero,
                                    'but theres values in different months (i.e. Jan: 5500, Feb: -500)
Dim ilPerRequested As Integer       '#periods requested

    For ilTemp = 0 To UBound(tmVefDollars) - 1 Step 1
        tmGrf.lChfCode = lmChfCode      '4-24-06 need to show the info from the latest version of contract tgChfCT.lCode
        tmGrf.iVefCode = tmVefDollars(ilTemp).iVefCode
        tmGrf.iSofCode = ilSaveSof                  'sales source
        tmGrf.iSlfCode = tmVefDollars(ilTemp).iSlfCode  '4-6-11 split slsp feature
        tmGrf.sBktType = IIF(tmVefDollars(ilTemp).iNTRInd = 1, "N", "A")
        
        For ilTemp2 = 1 To 2                        'loop all vehicles for cash & trade (one order splits cash & trade)
            llGross = 0
            ilZero = True
            For illoop = 1 To 12
                tmGrf.lDollars(illoop - 1) = 0
            Next illoop
            ilPerRequested = 12                 'assume all 12 months
            If ilListIndex = CNT_SALESACTIVITY_SS Then
                ilPerRequested = ilPeriods
            End If

            For illoop = 1 To ilPerRequested        '12
                If ilTemp2 = 1 Then
                    tmGrf.lDollars(illoop - 1) = tmGrf.lDollars(illoop - 1) + tmVefDollars(ilTemp).lProject(illoop)
                    'dont add up all values because changes in one month to another could cause zero balance, but
                    'the report still needs to show changes from one month to another
                    If tmVefDollars(ilTemp).lProject(illoop) <> 0 Then
                        ilZero = False
                    End If
                Else
                    tmGrf.lDollars(illoop - 1) = tmGrf.lDollars(illoop - 1) + tmVefDollars(ilTemp).lProjectTrade(illoop)
                    'dont add up all values because changes in one month to another could cause zero balance, but
                    'the report still needs to show changes from one month to another
                    If tmVefDollars(ilTemp).lProjectTrade(illoop) <> 0 Then
                        ilZero = False
                    End If
                End If
                'llGross = llGross + tmGrf.lDollars(ilLoop)
            Next illoop                 '12 months totalled
            'If llGross <> 0 Then        'totalled all 12 months, dont write out zero $ records
            If Not ilZero Then
                If ilTemp2 = 1 Then         'cash trade flag
                    tmGrf.sDateType = "C"
                Else
                    tmGrf.sDateType = "T"
                End If

                'date and time genned need only be set the first time - remains the same
                tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
                tmGrf.iGenDate(1) = igNowDate(1)
                'tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
                'tmGrf.iGenTime(1) = igNowTime(1)
                gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                tmGrf.lGenTime = lgNowTime
                'vehicle group is for the daily sales activity by month
                'gGetVehGrpSets tmGrf.iVefCode, ilMajorSet, ilMajorSet, tmGrf.iPerGenl(4), tmGrf.iPerGenl(4)   '6-13-02
                gGetVehGrpSets tmGrf.iVefCode, ilMajorSet, ilMajorSet, tmGrf.iPerGenl(3), tmGrf.iPerGenl(3)   '6-13-02
                
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            End If
        Next ilTemp2                                'loop for cash vs trade
    Next ilTemp                             'ilTemp = 0 To UBound(tmVefDollars)
End Sub
'
'                   gFltRatesSpots - Loop through the flights of the schedule line
'                                   and build the projections dollars& spots into lmprojmonths array
'                                   This routine is a copy of gBuildFlight - some parameters
'                                   may not be necessary, but retained in case some generality
'                                   needed at a later time.
'                       Currently, this routine only processes for "Week" option.  The past
'                       is not an issue (that is, needing to know what has been billed - which
'                       uses ilFirstProjInx.  This will always be 1 for the time-being.)
'                   <input> ilclf = sched line index into tgClfCt
'                           llStdStartDates() - array of dates to build $ from flights
'                           ilFirstProjInx - index of 1st month/week to start projecting
'                           ilMaxInx - max # of buckets to loop thru (not relavant for weekly buckets)
'                           ilWkOrMonth - 1 = Month, 2 = Week
'                           ilNC - true if include No charge lines, else False
'                           ilGrossNetTNet :  0 = gross , 1 = net, 2 = tnet (9-23-09)
'                   <output> tmFlight - structure of type FLIGHT containing vehicle, # spots/week for
'                           each unique spot price from all flights for same schedule line.
'
'
'                   General routine to build flight $ into week, month, qtr buckets
'                   tgChfCT, tgClfCT, tgCffCT are the buffers holding the contract
Sub mFltRatesSpots(ilClf As Integer, llStdStartDates() As Long, ilFirstProjInx As Integer, ilMaxInx As Integer, ilWkOrMonth As Integer, ilNC, ilGrossNetTNet As Integer)
Dim ilCff As Integer
Dim slStr As String
Dim llFltStart As Long
Dim llFltEnd As Long
Dim illoop As Integer
Dim llDate As Long
Dim llDate2 As Long
Dim llSpots As Long
Dim ilTemp As Integer
Dim llStdStart As Long
Dim llStdEnd As Long
Dim ilMonthInx As Integer
Dim ilWkInx As Integer
Dim ilFlight As Integer
Dim ilFoundFlight As Integer
Dim ilLoop2 As Integer
Dim ilLoop3 As Integer
Dim ilShowOVDays As Integer
Dim ilShowOVTimes As Integer
Dim slOVStartTime As String
Dim slOVEndTime As String
Dim ilXMid As Integer
Dim slDescription As String
Dim llrunningStartTime As Long
Dim llrunningEndtime As Long
Dim slTempDays As String
Dim slDay As String
Dim slSpotCount As String
Dim ilRet As Integer
Dim tlCff As CFF
Dim ilListIndex As Integer
Dim ilCommPct As Integer
Dim slAmount As String
Dim slSharePct As String
Dim llTempActPrice As Long
Dim blAcqOK As Boolean
Dim ilAcqLoInx As Integer
Dim ilAcqHiInx As Integer
Dim ilAcqCommPct As Integer
Dim llAcqComm As Long
Dim llAcqNet As Long

    ilListIndex = RptSelCt!lbcRptType.ListIndex
    
    blAcqOK = gGetAcqCommInfoByVehicle(tgClfCT(ilClf).ClfRec.iVefCode, ilAcqLoInx, ilAcqHiInx) 'determine the starting and ending indices of acq percents for this lines vehicle

    'ReDim tmFlight(1 To 1) As FLIGHT
    ReDim tmFlight(0 To 0) As FLIGHT
    llStdStart = llStdStartDates(ilFirstProjInx)
    llStdEnd = llStdStartDates(ilMaxInx)
    ilCff = tgClfCT(ilClf).iFirstCff
    Do While ilCff <> -1
        tlCff = tgCffCT(ilCff).CffRec
        
         '9-23-09 gross, net or tnet option
        If ilGrossNetTNet = 0 Then
            llTempActPrice = tlCff.lActPrice
        Else            'net or tnet
            llTempActPrice = tlCff.lActPrice
            If tgChfCT.iAgfCode = 0 Then          'direct
                ilCommPct = 10000                'no commission
            Else
                ilCommPct = 8500         'default to commissionable if no agency found
                'see what the agency comm is defined as
                tmAgfSrchKey.iCode = tgChfCT.iAgfCode
                ilRet = btrGetEqual(hmAgf, tmAgf, Len(tmAgf), tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching agy recd
                If ilRet = BTRV_ERR_NONE Then
                    ilCommPct = (10000 - tmAgf.iComm)
                End If          'ilret = btrv_err_none
            End If
            slAmount = gLongToStrDec(llTempActPrice, 2)
            slSharePct = gIntToStrDec(ilCommPct, 4)
            slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
            slStr = gRoundStr(slStr, ".01", 2)
            llTempActPrice = gStrDecToLong(slStr, 2) 'adjusted net
'            If ilGrossNetTNet = 2 Then          'tnet, adjust for acquisition cost
'
'                llTempActPrice = llTempActPrice - tgClfCT(ilClf).ClfRec.lAcquisitionCost
'            End If
        End If
        'If (tlCff.lActPrice > 0) Or (tlCff.lActPrice = 0 And ilNC) Then
        If (llTempActPrice <> 0) Or (llTempActPrice = 0 And ilNC) Then      '9-15-10 adjust for acquisition cost

            gUnpackDate tlCff.iStartDate(0), tlCff.iStartDate(1), slStr
            llFltStart = gDateValue(slStr)
            'backup start date to Monday
            illoop = gWeekDayLong(llFltStart)
            Do While illoop <> 0
                llFltStart = llFltStart - 1
                illoop = gWeekDayLong(llFltStart)
            Loop
            gUnpackDate tlCff.iEndDate(0), tlCff.iEndDate(1), slStr
            llFltEnd = gDateValue(slStr)
            'the flight dates must be within the start and end of the projection periods,
            'not be a CAncel before start flight, and have a cost > 0
            'All data, past & future is retrieved from schedule lines
            If (llFltStart < llStdEnd And llFltEnd >= llStdStart) And (llFltEnd >= llFltStart) Then
                'adjust the gather dates from flights: use flight start date or requested start date, whichever is later
                If llStdStart > llFltStart Then
                    llFltStart = llStdStart
                End If
                'use flight end date or requsted end date, whichever is lesser
                If llStdEnd < llFltEnd Then
                    llFltEnd = llStdEnd
                End If

                If ilGrossNetTNet = 2 Then          'tnet, adjust for acquisition cost
                    ilAcqCommPct = gGetEffectiveAcqComm(llFltStart, ilAcqLoInx, ilAcqHiInx)     'if varying commissions for acq costs on Insertion order, get the % to be used to calc.
                    gCalcAcqComm ilAcqCommPct, tgClfCT(ilClf).ClfRec.lAcquisitionCost, llAcqNet, llAcqComm
                    llTempActPrice = llTempActPrice - llAcqNet
                End If
                For llDate = llFltStart To llFltEnd Step 7
                    'Loop on the number of weeks in this flight
                    'calc week into of this flight to accum the spot count
                    If tlCff.sDyWk = "W" Then            'weekly
                        llSpots = tlCff.iSpotsWk + tlCff.iXSpotsWk
                        slTempDays = gDayNames(tlCff.iDay(), tlCff.sXDay(), 2, slStr)            'slstr not needed when returned
                        slStr = ""
                        For illoop = 1 To Len(slTempDays) Step 1
                            slDay = Mid$(slTempDays, illoop, 1)
                            If slDay <> " " And slDay <> "," Then
                                slStr = Trim$(slStr) & Trim$(slDay)
                            End If
                        Next illoop
                    Else                                        'daily
                        If illoop + 6 < llFltEnd Then           'we have a whole week
                            llSpots = tlCff.iDay(0) + tlCff.iDay(1) + tlCff.iDay(2) + tlCff.iDay(3) + tlCff.iDay(4) + tlCff.iDay(5) + tlCff.iDay(6)
                        Else
                            llFltEnd = llDate + 6
                            If llDate > llFltEnd Then
                                llFltEnd = llFltEnd       'this flight isn't 7 days
                            End If
                            For llDate2 = llDate To llFltEnd Step 1
                                ilTemp = gWeekDayLong(llDate2)
                                llSpots = llSpots + tlCff.iDay(ilTemp)
                            Next llDate2
                        End If
                        slStr = ""
                        For illoop = 0 To 6
                            slSpotCount = Trim$(str$(tlCff.iDay(illoop)))
                            Do While Len(slSpotCount) < 4
                                slSpotCount = " " & slSpotCount
                            Loop
                            slStr = slStr & " " & slSpotCount
                        Next illoop
                    End If
                    
                   If ilListIndex = CNT_AVGRATE Then        'avg rate gathers data by daypart
                       'slStr contains days
                        tmRdfSrchKey.iCode = tmClf.iRdfCode
                        ilRet = btrGetEqual(hmRdf, tmRdf, Len(tmRdf), tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching r/c recd
                        If ilRet <> BTRV_ERR_NONE Then
                            slDescription = "Missing DP"
    
                        End If
                        
                        If RptSelCt!rbcSelC7(1).Value = True Then           'use daypart with overrides
                            For ilLoop2 = 0 To 6 Step 1
                                'If tmRdf.sWkDays(7, ilLoop2 + 1) = "Y" Then             'is DP is a valid day
                                If tmRdf.sWkDays(6, ilLoop2) = "Y" Then             'is DP is a valid day
                                    
                                    If tlCff.iDay(ilLoop2) >= 0 Then         '11-19-02 is flight a valid day? 0=invalid day, >1 = daily # spots
                                        ilShowOVDays = True
                                        Exit For
                                    Else
                                        ilShowOVDays = False
                                    End If
                                End If
                            Next ilLoop2
                            'Times
                            ilShowOVTimes = False
                            If ((tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0)) And ((tmClf.iEndTime(0) <> 1) Or (tmClf.iEndTime(1) <> 0)) Then
                                gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slOVStartTime       '7-8-05
                                gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slOVEndTime
                                ilShowOVTimes = True
                            Else
                                'Add times
                                ilXMid = False
                                'if there are multiple segments and it cross midnight, show the earliest start time and xmidnight end time;
                                'otherwise the first segments start and end times are shown
                                For ilLoop2 = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1 'Row
                                       If (tmRdf.iStartTime(0, ilLoop2) <> 1) Or (tmRdf.iStartTime(1, ilLoop2) <> 0) Then
                                        gUnpackTime tmRdf.iStartTime(0, ilLoop2), tmRdf.iStartTime(1, ilLoop2), "A", "1", slOVStartTime '7-8-05
                                        gUnpackTime tmRdf.iEndTime(0, ilLoop2), tmRdf.iEndTime(1, ilLoop2), "A", "1", slOVEndTime
                                        gUnpackTimeLong tmRdf.iEndTime(0, ilLoop2), tmRdf.iEndTime(1, ilLoop2), True, llrunningEndtime
                                        If llrunningEndtime = 86400 Then    'its 12M end of day
                                            ilXMid = True
                                        End If
                                        For ilLoop3 = ilLoop2 + 1 To UBound(tmRdf.iStartTime, 2)
                                            gUnpackTimeLong tmRdf.iStartTime(0, ilLoop3), tmRdf.iStartTime(1, ilLoop3), False, llrunningStartTime
                                             If llrunningStartTime = 0 And llrunningEndtime = 86400 Then
                                                If ilXMid Then
                                                    gUnpackTime tmRdf.iEndTime(0, ilLoop3), tmRdf.iEndTime(1, ilLoop3), "A", "1", slOVEndTime
                                                    Exit For
                                                End If
                                            Else
                                                gUnpackTimeLong tmRdf.iEndTime(0, ilLoop3), tmRdf.iEndTime(1, ilLoop3), True, llrunningEndtime
                                            End If
                                        Next ilLoop3
                                        Exit For
                                    End If
                                Next ilLoop2
                            End If
                            
                            If ilShowOVDays Or ilShowOVTimes Then
                                slDescription = RTrim$(slStr) & " " & Trim$(slOVStartTime) & "-" & Trim$(slOVEndTime)
                            Else
                                slDescription = Trim$(tmRdf.sName)
                            End If
                        Else            'use daypart name only; no overrides
                            slDescription = Trim$(tmRdf.sName)
                        End If
                        
                        
                        ilFoundFlight = False
                        For ilFlight = LBound(tmFlight) To UBound(tmFlight) - 1 Step 1
                            'If Trim$(tmFlight(ilFlight).sDescription) = slDescription And tmFlight(ilFlight).lActPrice = tlCff.lActPrice Then
                            If Trim$(tmFlight(ilFlight).sDescription) = slDescription And tmFlight(ilFlight).lActPrice = llTempActPrice Then        '9-23-09
                                
                                ilFoundFlight = True
                                Exit For
                            End If
                        Next ilFlight
                        If Not ilFoundFlight Then
                            tmFlight(UBound(tmFlight)).sDescription = Trim$(slDescription)
                            'tmFlight(UBound(tmFlight)).lActPrice = tlCff.lActPrice
                            tmFlight(UBound(tmFlight)).lActPrice = llTempActPrice       '9-23-09
                            
                            ilFlight = UBound(tmFlight)
                            ReDim Preserve tmFlight(LBound(tmFlight) To ilFlight + 1) As FLIGHT
                        End If
                    ElseIf ilListIndex = CNT_AVG_PRICES Then        'avg spot prices gathers by spot length
                        ilFoundFlight = False
                        For ilFlight = LBound(tmFlight) To UBound(tmFlight) - 1 Step 1
                            If Trim$(tmFlight(ilFlight).iLen) = tmClf.iLen Then
                                ilFoundFlight = True
                                Exit For
                            End If
                        Next ilFlight
                        If Not ilFoundFlight Then
                            tmFlight(UBound(tmFlight)).iLen = tmClf.iLen
                            'tmFlight(UBound(tmFlight)).lActPrice = tlCff.lActPrice
                            tmFlight(UBound(tmFlight)).lActPrice = llTempActPrice       '9-23-09
                            ilFlight = UBound(tmFlight)
                            ReDim Preserve tmFlight(LBound(tmFlight) To ilFlight + 1) As FLIGHT
                        End If
                    'Advt units ordered gathers by line & rate
                    ElseIf ilListIndex = CNT_ADVT_UNITS Then            '1-7-09
                        ilFoundFlight = False
                        For ilFlight = LBound(tmFlight) To UBound(tmFlight) - 1 Step 1
                            'If Trim$(tmFlight(ilFlight).iLine) = tmClf.iLine And tmFlight(ilFlight).lActPrice = tmCff.lActPrice Then
                            If Trim$(tmFlight(ilFlight).iLine) = tmClf.iLine And tmFlight(ilFlight).lActPrice = llTempActPrice Then            '9-23-09
                                ilFoundFlight = True
                                Exit For
                            End If
                        Next ilFlight
                        If Not ilFoundFlight Then
                            tmFlight(UBound(tmFlight)).iLine = tmClf.iLine
                            'tmFlight(UBound(tmFlight)).lActPrice = tlCff.lActPrice
                            tmFlight(UBound(tmFlight)).lActPrice = llTempActPrice       '9-23-09
                            tmFlight(UBound(tmFlight)).iLen = tmClf.iLen
                            ilFlight = UBound(tmFlight)
                            ReDim Preserve tmFlight(LBound(tmFlight) To ilFlight + 1) As FLIGHT
                        End If
                    End If
                    
                    If ilWkOrMonth = 1 Or ilWkOrMonth = 2 Then                     'monthly buckets
                        'determine month that this week belongs in, then accumulate the gross and net $
                        'currently, the projections are based on STandard bdcst
                        For ilMonthInx = ilFirstProjInx To ilMaxInx - 1 Step 1       'loop thru months to find the match
                            If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                                'tmFlight(ilFlight).lProj(ilMonthInx) = tmFlight(ilFlight).lProj(ilMonthInx) + (llSpots * tlCff.lActPrice)
                                tmFlight(ilFlight).lProj(ilMonthInx) = tmFlight(ilFlight).lProj(ilMonthInx) + (llSpots * llTempActPrice)        '9-23-09
                                tmFlight(ilFlight).iSpots(ilMonthInx) = tmFlight(ilFlight).iSpots(ilMonthInx) + (llSpots)
                                'llProject(ilMonthInx) = llProject(ilMonthInx) + (llSpots * tlCff.lActPrice)
                                Exit For
                            End If
                        Next ilMonthInx
                    Else                                    'weekly buckets
                        ilWkInx = (llDate - llStdStartDates(1)) \ 7 + 1
                        If ilWkInx > 0 Then
                            'llProject(ilWkInx) = llProject(ilWkInx) + (llSpots * tlCff.lActPrice)
                            'tmFlight(ilFlight).lProj(ilWkInx) = tmFlight(ilFlight).lProj(ilWkInx) + (llSpots * tlCff.lActPrice)
                            tmFlight(ilFlight).lProj(ilWkInx) = tmFlight(ilFlight).lProj(ilWkInx) + (llSpots * llTempActPrice)      '9-23-09
                            tmFlight(ilFlight).iSpots(ilWkInx) = tmFlight(ilFlight).iSpots(ilWkInx) + (llSpots)
                        End If
                    End If
                Next llDate                                     'for llDate = llFltStart To llFltEnd
            End If                                          '
        End If                                          'actprice > 0 or (actprice = 0 and ilNC)
        ilCff = tgCffCT(ilCff).iNextCff                   'get next flight record from mem
    Loop                                            'while ilcff <> -1
End Sub
'
'
'        mInitPaperWork - Open files required and setup report parameter filters
'
'       Return - true if some kind of error
Function mInitPaperWork(ilCodes() As Integer) As Integer
Dim ilRet As Integer
Dim ilError As Integer
Dim slNameCode As String
Dim slCode As String
Dim ilTemp As Integer
    mInitPaperWork = BTRV_ERR_NONE


    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitPaperWorkErr
    gBtrvErrorMsg ilRet, "mInitPaperWork, RptCrmct (OpenErr):" & "GRF.Btr", RptSelCt
    On Error GoTo 0
    imGrfRecLen = Len(tmGrf)

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitPaperWorkErr
    gBtrvErrorMsg ilRet, "mInitPaperWork, RptCrmct (OpenErr):" & "CHF.Btr", RptSelCt
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)

    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitPaperWorkErr
    gBtrvErrorMsg ilRet, "mInitPaperWork, RptCrmct (OpenErr):" & "SLF.Btr", RptSelCt
    On Error GoTo 0
    imSlfRecLen = Len(tmSlf)

    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitPaperWorkErr
    gBtrvErrorMsg ilRet, "mInitPaperWork, RptCrmct (OpenErr):" & "MNF.Btr", RptSelCt
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitPaperWorkErr
    gBtrvErrorMsg ilRet, "mInitPaperWork, RptCrmct (OpenErr):" & "CLF.Btr", RptSelCt
    On Error GoTo 0
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitPaperWorkErr
    gBtrvErrorMsg ilRet, "mInitPaperWork, RptCrmct (OpenErr):" & "CFF.Btr", RptSelCt
    On Error GoTo 0
    imCffRecLen = Len(tmCff)

    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitPaperWorkErr
    gBtrvErrorMsg ilRet, "mInitPaperWork, RptCrmct (OpenErr):" & "AGF.Btr", RptSelCt
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)

    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitPaperWorkErr
    gBtrvErrorMsg ilRet, "mInitPaperWork, RptCrmct (OpenErr):" & "ADF.Btr", RptSelCt
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)

    hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitPaperWorkErr
    gBtrvErrorMsg ilRet, "mInitPaperWork, RptCrmct (OpenErr):" & "SBF.Btr", RptSelCt
    On Error GoTo 0
    imSbfRecLen = Len(tmSbf)


    If RptSelCt!rbcSelC7(0).Value Then
        smGrossOrNet = "G"
    ElseIf RptSelCt!rbcSelC7(1).Value Then
        smGrossOrNet = "N"
    End If
    imAdvt = False
    imSlsp = False
    imVehicle = False
    imAgency = False
    If RptSelCt!rbcSelCSelect(0).Value Then
        imAdvt = True
    ElseIf RptSelCt!rbcSelCSelect(1).Value Then
        imAgency = True
    ElseIf RptSelCt!rbcSelCSelect(2).Value Then
        imSlsp = True
    ElseIf RptSelCt!rbcSelCSelect(3).Value Then
        imVehicle = True
    End If
    imHold = False
    imOrder = False
    imDead = False
    imWork = False
    imIncomplete = False
    imComplete = False
    imRev = False                                       '11-7-16 rev working
    'Determine Contract Statuses to include
    If RptSelCt!ckcSelC3(0).Value = vbChecked Then
        imHold = True
    End If
    If RptSelCt!ckcSelC3(1).Value = vbChecked Then
        imOrder = True
    End If
    If RptSelCt!ckcSelC3(2).Value = vbChecked Then
        imDead = True
    End If
    If RptSelCt!ckcSelC3(3).Value = vbChecked Then
        imWork = True
    End If
    If RptSelCt!ckcSelC3(4).Value = vbChecked Then
        imIncomplete = True
    End If
    If RptSelCt!ckcSelC3(5).Value = vbChecked Then
        imComplete = True
    End If
    If RptSelCt!ckcSelC3(6).Value = vbChecked Then              '11-7-16 Rev working
        imRev = True
    End If

    'Determine Contract Types to include
    imStandard = False
    imReserv = False
    imRemnant = False
    imDR = False
    imPI = False
    imPSA = False
    imPromo = False
    If RptSelCt!ckcSelC5(0).Value = vbChecked Then  'include std cntrs?
        imStandard = True
    End If
    If RptSelCt!ckcSelC5(1).Value = vbChecked Then   'include reserves?
        imReserv = True
    End If
    If RptSelCt!ckcSelC5(2).Value = vbChecked Then   'include remnants?
        imRemnant = True
    End If
    If RptSelCt!ckcSelC5(3).Value = vbChecked Then   'direct response?
        imDR = True
    End If
    If RptSelCt!ckcSelC5(4).Value = vbChecked Then   'per inquiry?
        imPI = True
    End If
    If RptSelCt!ckcSelC5(5).Value = vbChecked Then   'psa?
        imPSA = True
    End If
    If RptSelCt!ckcSelC5(6).Value = vbChecked Then  'promo?
        imPromo = True
    End If

    imTrade = False
    imCash = False
    If RptSelCt!rbcSelC4(2).Value Then
        imCash = True
        imTrade = True
    ElseIf RptSelCt!rbcSelC4(0).Value Then
        imCash = True
    Else
        imTrade = True
    End If
    imDiscrepOnly = False
    If RptSelCt!ckcSelC8(0).Value = vbChecked Then  '9-12-02
        imDiscrepOnly = True
    End If
    imCreditCk = False
    If RptSelCt!ckcSelC8(1).Value = vbChecked Then  '9-12-02
        imCreditCk = True
    End If
    imShowInternal = False
    imShowOther = False
    imShowChgRsn = False
    imShowCancel = False
    If RptSelCt!ckcSelC6(0).Value = vbChecked Then
        imShowInternal = True
    End If
    If RptSelCt!ckcSelC6(1).Value = vbChecked Then
        imShowOther = True
    End If
    If RptSelCt!ckcSelC6(2).Value = vbChecked Then
        imShowChgRsn = True
    End If
    If RptSelCt!ckcSelC6(3).Value = vbChecked Then
        imShowCancel = True
    End If

    'setup arrays of selected items
    If Not RptSelCt!ckcAll.Value = vbChecked Then   'build array of the selected advertisers
        If imAdvt Then
            For ilTemp = 0 To RptSelCt!lbcSelection(5).ListCount - 1 Step 1
                If RptSelCt!lbcSelection(5).Selected(ilTemp) Then              'selected advt
                    slNameCode = tgAdvertiser(ilTemp).sKey
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilCodes(UBound(ilCodes)) = Val(slCode)
                    ReDim Preserve ilCodes(0 To UBound(ilCodes) + 1)
                End If
            Next ilTemp
        ElseIf imAgency Then
            For ilTemp = 0 To RptSelCt!lbcSelection(1).ListCount - 1 Step 1
                If RptSelCt!lbcSelection(1).Selected(ilTemp) Then              'selected agency
                    slNameCode = tgAgency(ilTemp).sKey
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilCodes(UBound(ilCodes)) = Val(slCode)
                    ReDim Preserve ilCodes(0 To UBound(ilCodes) + 1)
                End If
            Next ilTemp
        ElseIf imSlsp Then
            For ilTemp = 0 To RptSelCt!lbcSelection(2).ListCount - 1 Step 1
                If RptSelCt!lbcSelection(2).Selected(ilTemp) Then              'selected slsp
                    slNameCode = tgSalesperson(ilTemp).sKey
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilCodes(UBound(ilCodes)) = Val(slCode)
                    ReDim Preserve ilCodes(0 To UBound(ilCodes) + 1)
                End If
            Next ilTemp
        Else
            For ilTemp = 0 To RptSelCt!lbcSelection(6).ListCount - 1 Step 1
                If RptSelCt!lbcSelection(6).Selected(ilTemp) Then              'selected vehicle
                    slNameCode = tgCSVNameCode(ilTemp).sKey
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilCodes(UBound(ilCodes)) = Val(slCode)
                    ReDim Preserve ilCodes(0 To UBound(ilCodes) + 1)
                End If
            Next ilTemp
        End If
    End If
    mInitPaperWork = ilError
    Exit Function

mInitPaperWorkErr:
    On Error GoTo 0
    mInitPaperWork = False

End Function
'
'
'         gCrCumeAct - Cumulative Sales Placement or Cumulative Sales Activity
'           Create prepass file for Cumulative
'           Activity Report.  Produce a report of all new and
'           modified activity, including contracts whose status
'           is Hold or Order.  Modifications are reflected by
'           increases and decreases from the previous week.  The
'           effetive date entered filters the contracts for the
'           current week (which is always a Monday date).  Increases/
'           decreases are compared against the previous rev #.
'           12 months of contract data is gathered.
'
'           d.hosaka 6 / 8 / 97
'           4/30/98 Change the way in which current and previous weeks
'           are gathered.
'           7/1/98 - Gather data by OHD date (gObtainCntrForOHD), not
'           Cntrs start/end date parameters
'           7/15/98 - gObtainCntrForOHD was filtering on ChfDelete flag (remove)
'           7-26-02 Implement Start/end date selectivity to use earliest date entered or
'           latest date entered.  More selectivity on sales source, sales office
'           and market (if no mkts, allow vehicle group selectivity)
'       3-31-05 Advt option and agency option didnt balance, because when
'           gathering for advt option, it was using the vehicle and rounding too often
'       4-19-05 above changed messed up option by advt where the vehicle names do not show
'           for each contract.  When they do show, they are wrong.
'       11-17-05 When order is changed from commissionable to direct (or vice versa), the
'           order did not show as a change since all calculations were done on gross.
'           Each line is calculated and cash and/or trade $ written to disk.  This could
'           cause more rounding than before.
'       3-8-06 When package line was encountered, only the calculation was bypased.
'           The remaining code was processed which erroneously created extra $
'       4-24-06 show info (i.e. product name) from latest version of order, not previous version
'       2-01-07 implement inclusion of ntr/hard cost into Daily & Weekly Sales Activity by month
Sub gCrCumeActCt(llStdStartDates() As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llNoPenny                     llTemp                                                  *
'******************************************************************************************

Dim ilRet As Integer                    'return flag from all I/O
Dim llEarliestEntry As Long             'start date of effective week
Dim llLatestEntry As Long               'end date of effective week
Dim slStartDate As String               'active dates to report data for (12 month span)
Dim slEndDate As String                 'active dates to reort date for (12 month span)
Dim slStr As String                     'temp string variable
Dim ilTemp As Integer                   'temp integer variable
Dim ilSlfCode As Integer                'slsp requesting report, slsp can only see his own stuff
                                        'CSI and guide always sees everything
Dim ilSearchSlf As Integer              'temp loop variable
Dim ilSaveSS As Integer                 'slsp sales source from cnt
Dim llEnterDate As Long                 'date entered from contract header
Dim ilProcessCnt As Integer             'process contract flag (true or false)
Dim ilSelect As Integer                 'index into lbcSelection array for selective adv, agy, demo, vehicle
Dim slNameCode As String                'Parsing temporary string
Dim slName As String
Dim slCode As String                    'Parsing temporary string
Dim llContrCode As Long                 'Current contract's internal code #
Dim ilClf As Integer                    'For loop variable to loop thru sch lines
Dim ilFound As Integer                  'flag if found a selective vehicle when All not checked
'ReDim llProject(1 To 13) As Long               '$ for each lines projection
ReDim llProject(0 To 13) As Long               '$ for each lines projection. Index zero ignored
'ReDim llProjectCash(1 To 13) As Long    'returned Cash portion of contract (100% if all cash order; otherwise its split)
ReDim llProjectCash(0 To 13) As Long    'returned Cash portion of contract (100% if all cash order; otherwise its split). Index zero ignored
'ReDim llProjectTrade(1 To 13) As Long   'returned Trade portion of contract (100% if all trade, otherwise its split)
ReDim llProjectTrade(0 To 13) As Long   'returned Trade portion of contract (100% if all trade, otherwise its split). Index zero ignored
'ReDim llAcquisition(1 To 13) As Long      'Acq unused, temp for subroutine to call
ReDim llAcquisition(0 To 13) As Long      'Acq unused, temp for subroutine to call. Index zero ignored
'Dim llCashShare(1 To 13) As Long      'slsp split for all cash months
ReDim llCashShare(0 To 13) As Long      'slsp split for all cash months. Index zero ignored
'Dim llTradeShare(1 To 13) As Long     'sls split for all trade months by cnt/vehicle
ReDim llTradeShare(0 To 13) As Long     'sls split for all trade months by cnt/vehicle. Index zero ignored
Dim ilFoundAgain As Integer             'found veh in memory table to store $
Dim ilVefIndex As Integer               'index to vehicle found in memory
Dim llSingleCntr As Long                'single contract # (entred by user)
Dim slGrossOrNet As String              'G = gross, N = net
Dim slCntrTypes As String               'valid types of contracts to obtain based on user
Dim slCntrStatus As String              'Holds, orders, unsch hold, unsch order.
Dim ilHOState As Integer                'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)
Dim ilCurrentRecd As Integer            'next recd to process from tlchfadvtext
Dim ilThisWk As Integer
Dim llDate As Long
Dim llDate2 As Long
Dim ilSortCode As Integer               'sort key for the option selected:  when processing line, this is the associated
                                        'advt code, agy code, demo code or vehicle whose data is built into a table before
                                        'writing to disk
Dim ilCurrentPrev As Integer            '2 passes thru contracts ( once for current week, once for previous version)
Dim ilListIndex As Integer              '7-26-02
Dim ilSplitNTR As Integer               'If We're splitting NTR then we will run two times (once with NTR and Once Without): 0=Normal, 1=loop twice
Dim ilCurrentNTRGroupSplit As Integer   'this is the NTR Split/Group loop
Dim slEarliestEntry As String
Dim slLatestEntry As String
Dim illoop As Integer
Dim ilSaveSof As Integer
Dim ilUseOriginal As Integer
Dim ilPass As Integer
Dim ilMajorSet As Integer
Dim slVGSelected As String
Dim ilAgyCommPct As Integer
Dim ilNewOnly As Integer                '11-09-06 Only print new contracts
Dim ilPeriods As Integer                '12-07-06 Detrmine # periods requested so that only those months are written to prepass; all other months intialized
Dim tlSBFTypes As SBFTypes
'ReDim tlMnf(1 To 1) As MNF
ReDim tlMnf(0 To 0) As MNF
Dim ilCalMonth As Integer             '3-5-10 true if cal month, else false for std or corp
Dim ilAdjustDays As Integer
ReDim llCalSpots(0 To 0) As Long        'init buckets for daily calendar values
ReDim llCalAmt(0 To 0) As Long
ReDim llCalAcqAmt(0 To 0) As Long
ReDim ilValidDays(0 To 6) As Integer
Dim tlPriceTypes As PRICETYPES
Dim ilOKtoSeeVeh As Integer
Dim ilMnfSubCo As Integer
Dim ilLoopOnSlsp As Integer
Dim ilMax As Integer
Dim llProcessPct As Long
Dim ilLoopOnPer As Integer
Dim llSlfSplit(0 To 9) As Long
Dim ilSlspCode(0 To 9) As Integer
Dim llSlfSplitRev(0 To 9) As Long
                   

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)

    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSof
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imSofRecLen = Len(tmSof)

    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSlf
        btrDestroy hmSof
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imSlfRecLen = Len(tmSlf)

    '11-17-05 agency read required to get the agy comm
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAgf
        btrDestroy hmSlf
        btrDestroy hmSof
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)

    ilListIndex = RptSelCt!lbcRptType.ListIndex


    tlSBFTypes.iNTR = False          'include NTR billing
    tlSBFTypes.iInstallment = False      'exclude Installment billing
    tlSBFTypes.iImport = False           'exclude rep import billing
    imAirTime = True
    imNTR = False
    imHardCost = False

    'if Weekly Sales Activity by Month and either NTR or hard cost selected
    If (ilListIndex = CNT_CUMEACTIVITY) And (RptSelCt!ckcSelC8(1).Value = vbChecked Or RptSelCt!ckcSelC8(2).Value = vbChecked) Then
        'test for NTR inclusion (or Hard Cost)
        If Not RptSelCt!ckcSelC8(0).Value = vbChecked Then      'include AirTime
            imAirTime = False
        End If
        If RptSelCt!ckcSelC8(1).Value = vbChecked Then      'include NTR
            imNTR = True
        End If
        If RptSelCt!ckcSelC8(2).Value = vbChecked Then      'include Hard Cost
            imHardCost = True
        End If
    End If

    'if Daily Sales Activity by Month and either NTR or hard cost selected
    If (ilListIndex = CNT_SALESACTIVITY_SS) And (RptSelCt!ckcSelC13(1).Value = vbChecked Or RptSelCt!ckcSelC13(2).Value = vbChecked) Then
        'test for NTR inclusion (or Hard Cost)
        If Not RptSelCt!ckcSelC13(0).Value = vbChecked Then      'include AirTime
            imAirTime = False
        End If
        If RptSelCt!ckcSelC13(1).Value = vbChecked Then      'include NTR
            imNTR = True
        End If
        If RptSelCt!ckcSelC13(2).Value = vbChecked Then      'include Hard Cost
            imHardCost = True
        End If
    End If

    If imNTR Or imHardCost Then
        tlSBFTypes.iNTR = True
        hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmSbf)
            ilRet = btrClose(hmAgf)
            ilRet = btrClose(hmSlf)
            ilRet = btrClose(hmSof)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmCHF)
            btrDestroy hmSbf
            btrDestroy hmAgf
            btrDestroy hmSlf
            btrDestroy hmSof
            btrDestroy hmCff
            btrDestroy hmClf
            btrDestroy hmGrf
            btrDestroy hmCHF
            Exit Sub
        End If
        imSbfRecLen = Len(tmSbf)

        hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmMnf)
            ilRet = btrClose(hmSbf)
            ilRet = btrClose(hmAgf)
            ilRet = btrClose(hmSlf)
            ilRet = btrClose(hmSof)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmCHF)
            btrDestroy hmMnf
            btrDestroy hmSbf
            btrDestroy hmAgf
            btrDestroy hmSlf
            btrDestroy hmSof
            btrDestroy hmCff
            btrDestroy hmClf
            btrDestroy hmGrf
            btrDestroy hmCHF
            Exit Sub
        End If
        imMnfRecLen = Len(tmMnf)

        sgDemoMnfStamp = ""             'insure that the hard costs  are read
        ilRet = gObtainMnfForType("I", sgDemoMnfStamp, tlMnf())
    End If

    If RptSelCt!rbcSelCInclude(0).Value Then                'adv
        ilSelect = 5
    ElseIf RptSelCt!rbcSelCInclude(1).Value Then            'agy
        ilSelect = 1
    ElseIf RptSelCt!rbcSelCInclude(2).Value Then            'demo
        ilSelect = 11
    Else                                                    'vehicles
        ilSelect = 6
    End If

    ilNewOnly = True
    If RptSelCt!ckcSelC12(0).Value = vbUnchecked Then       'get increases/decreases
        ilNewOnly = False
    End If

    slStr = RptSelCt!edcText.Text            '12-7-06 #periods, pass to insert routine to write only those periods requested
    ilPeriods = Val(slStr)

    If RptSelCt!rbcSelC7(0).Value Then
        slGrossOrNet = "G"
    Else
        slGrossOrNet = "N"
    End If
    tmGrf = tmZeroGrf                'initialize new record

    ilUseOriginal = False           'assume not to use original entry date so an Activity report can be produced vs
                    'New only contracts show
    'Determine contracts to process based on their entered and modified dates
    If ilListIndex = CNT_SALESACTIVITY_SS Then         '7-26-02 daily sales activity by month
        illoop = RptSelCt!cbcSet1.ListIndex
        ilMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())
        slVGSelected = ""                       'this for the vehicle group headers if one selected, a vehicle group
                                                'could still be selected even if its not a primary vehicle group sort
        'assume sort by vehicle group (rbcselc9(0).value = true)
        If ilMajorSet = 1 Then
            slStr = "P"
        ElseIf ilMajorSet = 2 Then
            slStr = "S"
        ElseIf ilMajorSet = 3 Then
            slStr = "M"
        ElseIf ilMajorSet = 4 Then
            slStr = "F"
        ElseIf ilMajorSet = 5 Then
            slStr = "R"
        Else
            slStr = "N"
        End If
        slVGSelected = Trim$(slStr)
        If RptSelCt!rbcSelC9(1).Value Then
            slStr = "O"         'sort by office
        ElseIf RptSelCt!rbcSelC9(2).Value Then
            slStr = "A"         'sort by advertiser
        End If
        If Not gSetFormula("Sortby", "'" & slStr & "'") Then
            Erase tlSofList, llProject, llStdStartDates, llAcquisition, llProjectTrade, llProjectCash, llCashShare, llTradeShare
            ilRet = btrClose(hmAgf)
            ilRet = btrClose(hmSlf)
            ilRet = btrClose(hmSof)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmCHF)
            Exit Sub
        End If
        If Not gSetFormula("VGSelected", "'" & slVGSelected & "'") Then
            Erase tlSofList, llProject, llStdStartDates, llAcquisition, llProjectTrade, llProjectCash, llCashShare, llTradeShare
            ilRet = btrClose(hmAgf)
            ilRet = btrClose(hmSlf)
            ilRet = btrClose(hmSof)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmCHF)
            Exit Sub
        End If


        slStr = RptSelCt!CSI_CalFrom.Text       'Date: 11/26/2017   added CSI calendar controls for date entries --> edcSelCFrom.Text       'order entry start date
        llEarliestEntry = gDateValue(slStr)
        slEarliestEntry = Format$(llEarliestEntry, "m/d/yy")
        slStr = RptSelCt!CSI_CalTo.Text         'Date: 11/26/2017   added CSI calendar controls for date entries --> edcSelCFrom1.Text      'order entry end date
        llLatestEntry = gDateValue(slStr)
        slLatestEntry = Format$(llLatestEntry, "m/d/yy")
        'Setup start & end dates of contracts to gather
        slStartDate = Format$(llStdStartDates(1), "m/d/yy")  'these are the dates that are to be projected
        slEndDate = Format$(llStdStartDates(13), "m/d/yy")

        If RptSelCt!rbcSelCSelect(0).Value Then       'use orig entry date
            ilUseOriginal = True
        End If
        slStr = RptSelCt!edcTopHowMany.Text              'single cntr #
        
        ilCalMonth = False
        If RptSelCt!rbcSelCInclude(2).Value = True Then     'calendar month slected
            ilCalMonth = True
        End If
    Else            'cumulative activity - weekly sales activity by month
        slStr = RptSelCt!CSI_CalFrom.Text       'Date: 1/8/2020 added CSI calendar control for date entries --> edcSelCFrom.Text
        llEarliestEntry = gDateValue(slStr)
        llLatestEntry = llEarliestEntry + 6
        'obtain the earliest and latest dates to show data for
        'slStartDate = Format$(llStdStartDates(1), "m/d/yy")
        'slEndDate = Format$(llStdStartDates(13), "m/d/yy")
        slStartDate = Format$(llEarliestEntry, "m/d/yy")
        slEndDate = Format$(llLatestEntry, "m/d/yy")
        slStr = RptSelCt!edcSelCFrom1.Text              'single cntr #
        ilCalMonth = False
        If RptSelCt!rbcSelCSelect(2).Value = True Then     'calendar month slected
            ilCalMonth = True
        End If

    End If
    'determine selective contract if entered
    'If slStr = "" Then
    '10-14-05 check for valid input
    If Not IsNumeric(slStr) Then
        llSingleCntr = 0
    Else
        llSingleCntr = CLng(slStr)
    End If

    'build array of selling office codes and their sales sources.  This is the most major sort
    'in the Business Booked reports
    ilTemp = 0
    ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tlSofList(0 To ilTemp) As SOFLIST
        tlSofList(ilTemp).iSofCode = tmSof.iCode
        tlSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
        ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilTemp = ilTemp + 1
     Loop

    'Populate the salespeople.  Only salesp running report can see his own stuff
    ilRet = gObtainSalesperson()        'populated slsp list in tgMSlf

    If tgUrf(0).iCode = 1 Or tgUrf(0).iCode = 2 Then    'guide or counterpoint password
        ilSlfCode = 0                   'allow guide & CSI to get all stuff
    Else
        ilSlfCode = tgUrf(0).iSlfCode   'slsp gets to see only his own stuff
    End If


    slCntrTypes = gBuildCntTypes()      'Setup valid types of contracts to obtain based on user
    slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order.
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)

    ilAdjustDays = (llStdStartDates(13) - llStdStartDates(1)) + 1
    ReDim llCalSpots(0 To ilAdjustDays) As Long        'init buckets for daily calendar values (spots unused in this report)
    ReDim llCalAmt(0 To ilAdjustDays) As Long
    ReDim llCalAcqAmt(0 To ilAdjustDays) As Long            'acq unused in this reprot
    For illoop = 0 To 6                         'days of the week
        ilValidDays(illoop) = True              'force alldays as valid
    Next illoop
    
    'dont bother trying to calculate rates that are 0, since this report doesnt need spot counts
    tlPriceTypes.iCharge = True     'Chargeable lines
    tlPriceTypes.iZero = False      '.00 lines
    tlPriceTypes.iADU = False     'adu lines
    tlPriceTypes.iBonus = False          'bonus lines
    tlPriceTypes.iNC = False          'N/C lines
    tlPriceTypes.iRecap = False        'recapturable
    tlPriceTypes.iSpinoff = False      'spinoff
    
    'Process contracts for last and this year
    ilPass = 2                          'assume this is an Activity report vs new contracts only
    'Process contracts for last and this year
    If ilListIndex = CNT_SALESACTIVITY_SS Then         '7-26-02
        If ilUseOriginal Then       'gather new contracts only
            ilRet = gCntrForOrigOHD(RptSelCt, slStartDate, slEndDate, slEarliestEntry, slLatestEntry, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
            ilPass = 1              'do only one pass to get new stuff
        Else                        'get increases/decreases
            ilRet = gObtainCntrForOHD(RptSelCt, slEarliestEntry, slLatestEntry, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
        End If
    Else
        'ilRet = gObtainCntrForDate(RptSelCt, slStartDate, slEndDate, slCntrStatus, slCntrTypes, ilHOState, tlchfAdvtExt())
        'Change 6/30/98 to gather cntrs based on OHD date, not start & end dates of order
        ilRet = gObtainCntrForOHD(RptSelCt, slStartDate, slEndDate, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
    End If
    
    '09/28/2020 - TTP # 9952 - IF include NTR, Add option to split NTR (or by default: leave NTR grouped together)
    ilSplitNTR = 0
    If ilListIndex = CNT_CUMEACTIVITY Or ilListIndex = CNT_SALESACTIVITY_SS Then
        If RptSelCt.rbcSelC14(1).Value = True Then ilSplitNTR = 1
    End If
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
        'get conts earliest and latest dates to see which year it spans
        gUnpackDate tlChfAdvtExt(ilCurrentRecd).iStartDate(0), tlChfAdvtExt(ilCurrentRecd).iStartDate(1), slStr
        llDate = gDateValue(slStr)
        gUnpackDate tlChfAdvtExt(ilCurrentRecd).iEndDate(0), tlChfAdvtExt(ilCurrentRecd).iEndDate(1), slStr
        llDate2 = gDateValue(slStr)
        llContrCode = 0
        ReDim tmVefDollars(0 To 0) As ADJUSTLIST            'init vehicle buckets
                
        For ilCurrentPrev = 1 To ilPass
            ilThisWk = False
            If ilCurrentPrev = 1 Then               'current
                'Test if active dates are within the span of the year to report
                'Remove this test 6/30/98
                'If llDate < llStdStartDates(13) - 1 And llDate2 >= llStdStartDates(1) Then     'does contr dates span this year

                llContrCode = gActivityCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llEarliestEntry, llLatestEntry, hmCHF, tmChf)
                If llContrCode = 0 Then                     'nothing in current week
                    ilCurrentPrev = 2                       '
                    Exit For                                'dont bother testing previous week
                Else
                    ilThisWk = True
                End If
                lmChfCode = tlChfAdvtExt(ilCurrentRecd).lCode   '4-24-06
                'Else
                '    Exit For                                    'nothing in period reporting
                'End If
            Else                                'previous week
                llContrCode = gPaceCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llEarliestEntry - 1, hmCHF, tmChf) 'find the prvious version of the cntr
                'pass 2 for previous revisions.  if New only requested, dump the $ accumulated to far and reject contract
                If ilNewOnly And llContrCode > 0 Then 'there is a previous version, ignore since only NEW requested
                    llContrCode = 0             'fake out no contract exists
                    ReDim tmVefDollars(0 To 0) As ADJUSTLIST            'init vehicle buckets
                End If
            End If

            If llContrCode > 0 Then
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT())
                'Determine Office and Sales Source for later filtering if option by SS/Market

                ilSaveSS = 0
                For ilSearchSlf = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
                    If tgMSlf(ilSearchSlf).iCode = tgChfCT.iSlfCode(0) Then
                    For ilTemp = 0 To UBound(tlSofList) Step 1
                        If tlSofList(ilTemp).iSofCode = tgMSlf(ilSearchSlf).iSofCode Then
                        ilSaveSS = tlSofList(ilTemp).iMnfSSCode
                        ilSaveSof = tlSofList(ilTemp).iSofCode
                        Exit For
                        End If
                    Next ilTemp
                    Exit For
                    End If
                Next ilSearchSlf

                ilProcessCnt = False
                If ilListIndex = CNT_CUMEACTIVITY Then              '7-26-02
                    If ilSelect <> 6 And Not RptSelCt!ckcAll.Value = vbChecked Then                     'for advt, agy or demo selectivity, filter out before going to lines
                        If ilSelect = 5 Then                              'advt option
                            For ilTemp = 0 To RptSelCt!lbcSelection(5).ListCount - 1 Step 1
                                If RptSelCt!lbcSelection(5).Selected(ilTemp) Then              'selected slsp
                                    slNameCode = tgAdvertiser(ilTemp).sKey 'Traffic!lbcAdvertiser.List(ilTemp)         'pick up slsp code
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    If Val(slCode) = tmChf.iAdfCode Then
                                        ilProcessCnt = True
                                        Exit For
                                    End If
                                End If
                            Next ilTemp
                        ElseIf ilSelect = 11 Then      'demo  option
                            For ilTemp = 0 To RptSelCt!lbcSelection(11).ListCount - 1 Step 1
                                If RptSelCt!lbcSelection(11).Selected(ilTemp) Then              'selected slsp
                                    slNameCode = tgRptSelDemoCodeCT(ilTemp).sKey    'RptSelCt!lbcCSVNameCode.List(ilTemp)         'pick up slsp code
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    If Val(slCode) = tmChf.iMnfDemo(0) Then
                                        ilProcessCnt = True
                                        Exit For
                                    End If
                                End If
                            Next ilTemp
                        Else                                'agy
                            For ilTemp = 0 To RptSelCt!lbcSelection(1).ListCount - 1 Step 1
                                If RptSelCt!lbcSelection(1).Selected(ilTemp) Then
                                    slNameCode = tgAgency(ilTemp).sKey
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    If Val(slCode) = tmChf.iAgfCode Then
                                        ilProcessCnt = True
                                        Exit For
                                    End If
                                End If
                            Next ilTemp
                        End If
                    Else
                        ilProcessCnt = True
                    End If
                Else        '7-26-02 CNT_SALESACTIVITY_SS Sales Activity by Month , selectivity by sales source, sales office & market or vehicle
                    'Filter Sales source
                    If Not RptSelCt!ckcAll.Value = vbChecked Then
                        For ilTemp = 0 To RptSelCt!lbcSelection(3).ListCount - 1
                            If RptSelCt!lbcSelection(3).Selected(ilTemp) Then
                                slNameCode = tgSalesperson(ilTemp).sKey         'sales source code
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                If Val(slCode) = ilSaveSS Then      'sales source of this contract match one selected?
                                    ilProcessCnt = True
                                    Exit For
                                End If
                            End If
                        Next ilTemp
                    Else
                        ilProcessCnt = True
                    End If
                    'filter Sales Office
                    If ilProcessCnt Then            'found a sales source, continue with sales office; otherwise this is a contract not to be processed
                        ilProcessCnt = False
                        If Not RptSelCt!ckcAllAAS.Value = vbChecked Then
                            For ilTemp = 0 To RptSelCt!lbcSelection(2).ListCount - 1
                                If RptSelCt!lbcSelection(2).Selected(ilTemp) Then
                                    slNameCode = tgSOCodeCT(ilTemp).sKey         'sales source code
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    If Val(slCode) = ilSaveSof Then      'sales office of this contract match one selected?
                                        ilProcessCnt = True
                                        Exit For
                                    End If
                                End If
                            Next ilTemp
                        Else
                            ilProcessCnt = True
                        End If
                    End If
                End If
                
                   ' If (ilProcessCnt) And (tgChfCT.iPctTrade <> 100) And (tmChf.iSlfCode(0) = ilSlfCode Or ilSlfCode = 0) And (llSingleCntr = 0 Or llSingleCntr = tmChf.lCntrNo) Then
                    'filtering of slsp has been done in gObtainCntrForOHD routine (user allowed to see contract)
                    If (ilProcessCnt) And (tgChfCT.iPctTrade <> 100) And (llSingleCntr = 0 Or llSingleCntr = tmChf.lCntrNo) Then
                        ilAgyCommPct = 10000                     'Assume gross requested
                        If tgChfCT.iAgfCode > 0 Then
                            tmAgfSrchKey.iCode = tgChfCT.iAgfCode
                            ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If slGrossOrNet = "N" Then
                                ilAgyCommPct = 10000 - tmAgf.iComm
                            End If
                        End If
    
                        gUnpackDate tgChfCT.iOHDDate(0), tgChfCT.iOHDDate(1), slStr
                        llEnterDate = gDateValue(slStr)
                
                        If imAirTime Then               'include air time
                            For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                                tmClf = tgClfCT(ilClf).ClfRec
                               
                                ilMax = mCumeHowManySlsp(ilListIndex, llSlfSplit(), ilSlspCode(), llSlfSplitRev(), tmClf.iVefCode)
    '                            ReDim llslfsplit(0 To 9) As Long           '4-20-00 slsp slsp share %
    '                            ReDim ilSlspCode(0 To 9) As Integer             '4-20-00
    '                            ReDim llSlfSplitRev(0 To 9) As Long
    '
    '                                If ilListIndex = CNT_SALESACTIVITY_SS And RptSelCt!ckcSelC12(2).Value = vbChecked Then
    '                                'do split slsp for Sales Activity
    '                                ilMnfSubco = gGetSubCmpy(tgChfCT, ilSlspCode(), llslfsplit(), tmClf.iVefCode, False, llSlfSplitRev())                                         '4-6-00
    '
    '                                ilMax = 10
    '                                'determine the maximum number of entries to process; possibly 0 % comm in the middle of the split slsp
    '                                'start from end to get last valid slsp
    '                                For ilLoopOnSlsp = 9 To 0 Step -1
    '                                    If llslfsplit(ilLoopOnSlsp) > 0 Then
    '
    '                                        Exit For
    '                                    Else
    '                                        ilMax = ilMax - 1
    '                                    End If
    '                                Next ilLoopOnSlsp
    '
    '                            Else        'no split slsp, use primary for 100%
    '                                llslfsplit(0) = 1000000
    '                                ilSlspCode(0) = tgChfCT.iSlfCode(0)
    '                                ilMax = 1
    '                            End If
    '
                        
                                ilFound = True
                                If ilListIndex = CNT_CUMEACTIVITY Then          '7-26-02
                                    If ilSelect = 6 And Not RptSelCt!ckcAll.Value = vbChecked Then
                                        ilFound = False
                                        For ilTemp = 0 To RptSelCt!lbcSelection(6).ListCount - 1 Step 1
                                            If RptSelCt!lbcSelection(6).Selected(ilTemp) Then              'selected slsp
                                                slNameCode = tgCSVNameCode(ilTemp).sKey    'RptSelCt!lbcCSVNameCode.List(ilTemp)         'pick up slsp code
                                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                                If Val(slCode) = tmClf.iVefCode Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            End If
                                        Next ilTemp
                                    End If
                                Else                '7-25 CNT_SALESACTIVITY_SS check for market (if it exists) or vehicle if no mkt exists
                                    'determine which vehicle group is selected
                                    'tmGrf.iPerGenl(4) = 0
                                    tmGrf.iPerGenl(3) = 0
                                    If ilMajorSet > 0 Then
                                        'filter the vehicle group
                                        'gGetVehGrpSets tmClf.iVefCode, ilMajorSet, ilMajorSet, tmGrf.iPerGenl(4), tmGrf.iPerGenl(4)   '6-13-02
                                        gGetVehGrpSets tmClf.iVefCode, ilMajorSet, ilMajorSet, tmGrf.iPerGenl(3), tmGrf.iPerGenl(3)   '6-13-02
                                    End If
                                    If Not RptSelCt!CkcAllveh.Value = vbChecked And RptSelCt!lbcSelection(6).Visible = True Then        'all items in vehicle group selected?
                                        ilFound = False
                                        For illoop = 0 To RptSelCt!lbcSelection(6).ListCount - 1 Step 1
                                            If RptSelCt!lbcSelection(6).Selected(illoop) Then
                                                slNameCode = tgMnfCodeCT(illoop).sKey
                                                ilRet = gParseItem(slNameCode, 1, "\", slName)
                                                ilRet = gParseItem(slName, 3, "|", slName)
                                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                                'Determine which vehicle set to test
                                                'If tmGrf.iPerGenl(4) = Val(slCode) Then
                                                If tmGrf.iPerGenl(3) = Val(slCode) Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            End If
                                        Next illoop
                                    End If
                                End If
                                ilOKtoSeeVeh = gUserAllowedVehicle(tmClf.iVefCode)  '03-17-10 check to see if users allowed to see this vehicle
                                If (ilFound) And (ilOKtoSeeVeh) Then                 'all vehicles of selective one found
                                    If ilListIndex = CNT_CUMEACTIVITY Then          '7-26-02
                                        'determine which sort is used, then create the key field to be tested against
                                        If ilSelect = 1 Then            'agy
                                            ilSortCode = tgChfCT.iAgfCode
                                        ElseIf ilSelect = 5 Then        'adv
                                            'ilSortCode = tgChfCT.iAdfCode
                                            ilSortCode = tmClf.iVefCode            '3-31-05 this was used as the sort for advertiser option;
                                                                                    'and using adfcode was commented out.  change again to use
                                                                                    'the adfcode so that rounding occurs less frequently, rather
                                                                                    'than on each vehicle
                                                                                    '4-19-05 change this back to use the vehicle code.  Advt option
                                                                                    'does not show the vehicle names properly
    
                                        ElseIf ilSelect = 6 Then        'vehicle
                                            ilSortCode = tmClf.iVefCode
                                        Else
                                            ilSortCode = tgChfCT.iMnfDemo(0)    'demo category
                                        End If
                                    Else                '7-25-02 CNT_SALESACTIVITY_SS
                                        ilSortCode = tmClf.iVefCode     'report is generated by
                                    End If
                                    For ilTemp = 1 To 12 Step 1 'init projection $ each time
                                        llProject(ilTemp) = 0
                                    Next ilTemp
                                     'init the cal buckets, if used
                                    For ilTemp = 0 To UBound(llCalSpots) - 1
                                        llCalSpots(ilTemp) = 0        'init buckets for daily calendar values
                                        llCalAmt(ilTemp) = 0
                                        llCalAcqAmt(ilTemp) = 0
                                    Next ilTemp
    
                                    If tmClf.sType = "S" Or tmClf.sType = "H" Then  'project standard or hidden lines (packages dont have vehicle groups)
                                        If ilCalMonth Then              'cal month option
                                            gCalendarFlights tgClfCT(ilClf), tgCffCT(), llStdStartDates(1), llStdStartDates(13), ilValidDays(), True, llCalAmt(), llCalSpots(), llCalAcqAmt(), tlPriceTypes
                                            gAccumCalFromDays llStdStartDates(), llCalAmt(), llCalAcqAmt(), False, llProject(), llAcquisition(), 13
                                        Else                            'corp or std month option
                                            gBuildFlights ilClf, llStdStartDates(), 1, 13, llProject(), 1, tgClfCT(), tgCffCT()
                                        End If
                                        mSplitCashTrade llProject(), llProjectCash(), llProjectTrade(), ilAgyCommPct, ilListIndex
                                    'End If
                                        
                                        For ilLoopOnSlsp = 0 To ilMax - 1
                                            llProcessPct = llSlfSplitRev(ilLoopOnSlsp)
        
                                            If llProcessPct > 0 Then
                                                'llproject(1-12) contains $ for this line
                                                'vehicles are build into memory with its 12 $ buckets
                                                
                                                '4-5-11 Split the salesperson share
                                                For ilTemp = 1 To 13
                                                    llCashShare(ilTemp) = 0
                                                    llTradeShare(ilTemp) = 0
                                                Next ilTemp
                                                
                                                mCumeSlspShare llProcessPct, llProjectCash(), llProjectTrade(), llCashShare(), llTradeShare()
                                               
                                                ilFoundAgain = False
                                                For ilTemp = 0 To UBound(tmVefDollars) - 1 Step 1
                                                    'If tmVefDollars(ilTemp).ivefcode = tmClf.ivefcode Then
                                                    'If tmVefDollars(ilTemp).iVefCode = ilSortCode Then
                                                    If tmVefDollars(ilTemp).iSortCode = ilSortCode And tmVefDollars(ilTemp).iSlfCode = ilSlspCode(ilLoopOnSlsp) Then      '4-19-05 use generalized sort field for comparisons
            
                                                        ilFoundAgain = True
                                                        ilVefIndex = ilTemp
                                                        Exit For                    '12-12-05
                                                    End If
                                                Next ilTemp
                                                If Not (ilFoundAgain) Then
                                                    'tmVefDollars(UBound(tmVefDollars)).ivefcode = tmClf.ivefcode
                                                    'tmVefDollars(UBound(tmVefDollars)).iVefCode = ilSortCode           '4-19-05
                                                    tmVefDollars(UBound(tmVefDollars)).iSortCode = ilSortCode           '4-19-05 use generalized sort field for comparisons
                                                    tmVefDollars(UBound(tmVefDollars)).iVefCode = tmClf.iVefCode        '4-19-05 save the vehicle for the output
                                                    tmVefDollars(UBound(tmVefDollars)).iSlfCode = ilSlspCode(ilLoopOnSlsp)  '4-6-11  salesperson split feature
                                                    'ilVefIndex = ilUpperVef
                                                    ilVefIndex = UBound(tmVefDollars)
                                                    ReDim Preserve tmVefDollars(0 To UBound(tmVefDollars) + 1) As ADJUSTLIST
                                                End If
                                                'now add or subtract depending if this contract is a new or mod
                                                If llEnterDate < llEarliestEntry Then      'previous version
                                                    For ilTemp = 1 To 12 Step 1
                                                        'tmVefDollars(ilVefIndex).lProject(ilTemp) = tmVefDollars(ilVefIndex).lProject(ilTemp) - llProjectCash(ilTemp)
                                                        'tmVefDollars(ilVefIndex).lProjectTrade(ilTemp) = tmVefDollars(ilVefIndex).lProjectTrade(ilTemp) - llProjectTrade(ilTemp)
                                                        tmVefDollars(ilVefIndex).lProject(ilTemp) = tmVefDollars(ilVefIndex).lProject(ilTemp) - llCashShare(ilTemp)
                                                        tmVefDollars(ilVefIndex).lProjectTrade(ilTemp) = tmVefDollars(ilVefIndex).lProjectTrade(ilTemp) - llTradeShare(ilTemp)
                                                    Next ilTemp
                                                Else                                        'current weeks data
                                                    For ilTemp = 1 To 12 Step 1
                                                        'tmVefDollars(ilVefIndex).lProject(ilTemp) = tmVefDollars(ilVefIndex).lProject(ilTemp) + llProjectCash(ilTemp)
                                                        'tmVefDollars(ilVefIndex).lProjectTrade(ilTemp) = tmVefDollars(ilVefIndex).lProjectTrade(ilTemp) + llProjectTrade(ilTemp)
                                                        tmVefDollars(ilVefIndex).lProject(ilTemp) = tmVefDollars(ilVefIndex).lProject(ilTemp) + llCashShare(ilTemp)
                                                        tmVefDollars(ilVefIndex).lProjectTrade(ilTemp) = tmVefDollars(ilVefIndex).lProjectTrade(ilTemp) + llTradeShare(ilTemp)
                                                    Next ilTemp
                                                End If
                                                tmVefDollars(ilVefIndex).iNTRInd = 0
                                            End If                      'llprocessPct > 0
                                        Next ilLoopOnSlsp
                                    End If                                  'ilfound
                                End If                                  '3-8-06 if tmclf.stype = "S" or tmclf.stype = "H"
                            Next ilClf                                  'for ilclf = lbound(tgClfCt) - ubound(tgClfCt)
                        'End If
                    End If
                    If imNTR Or imHardCost Then
                        mSalesActNTR ilListIndex, ilSelect, ilMajorSet, ilMajorSet, slGrossOrNet, llStdStartDates(), llEarliestEntry, tlSBFTypes, tlMnf()
                    End If
                    'all schedule lines complete, continue to see if there's another same contract # to process before writing to disk

                    End If              'ilprocesscnt
                End If                  'llContrcode > 0
            
            Next ilCurrentPrev
        
        mCumeInsert ilListIndex, ilPeriods, tmVefDollars(), ilSaveSS, ilMajorSet
    'Done with this contract write it out to disk
    Next ilCurrentRecd
        
    Erase tlSofList, llProject, tmVefDollars, llStdStartDates, llAcquisition, llProjectCash, llProjectTrade
    Erase llCalSpots, llCalAmt, llCalAcqAmt, ilValidDays
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmMnf)
    Exit Sub
End Sub
Sub mWriteSlsAct()
Dim ilRet As Integer
Dim ilTemp As Integer
        'GenDate - generation date (key)
        'GenTime - generation time (key)
        'chfCode = Contract #
        'adfCode = Advetiser Code
        'Code2 - Sales Source
        'BktType - O = orders, P = projetions
        'DateType = A, B or C for projections
        'Code 4 = change reason
        'sofCode = slsp office code
        'PerGenl(1) - 0 = new, else modification
        'PerGenl(2) - Contract Rev #
        'PerGenl(3) - internal flags for detail processing, unused in crystal
        'PerGenl(4) - hold or unsch hold flag
        'lDolllars(1) - Amount
        tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
        tmGrf.iGenDate(1) = igNowDate(1)
        'tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
        'tmGrf.iGenTime(1) = igNowTime(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrf.lGenTime = lgNowTime
        tmGrf.lChfCode = tgChfCT.lCntrNo            'contract #
        tmGrf.iAdfCode = tgChfCT.iAdfCode         'advertiser code
        'tmGrf.iPerGenl(4) = 0                   'assume Order (vs hold)
        tmGrf.iPerGenl(3) = 0                   'assume Order (vs hold)
        If tgChfCT.sStatus = "H" Or tgChfCT.sStatus = "G" Then      'if hold or unsch hold, set flag for Crystal
            'tmGrf.iPerGenl(4) = 1
            tmGrf.iPerGenl(3) = 1
        End If
        ilRet = BTRV_ERR_NONE
        If tgChfCT.iSlfCode(0) <> tmSlf.iCode Then        'only read slsp recd if not in mem already
            tmSlfSrchKey.iCode = tgChfCT.iSlfCode(0)         'find the slsp to obtain the sales source code
            ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        End If                                          'table of selling offices built into memory with its
        If ilRet = BTRV_ERR_NONE Then
            tmGrf.iSofCode = tmSlf.iSofCode
            'sales source
            For ilTemp = LBound(tlSofList) To UBound(tlSofList)
                If tlSofList(ilTemp).iSofCode = tmSlf.iSofCode Then
                    tmGrf.iCode2 = tlSofList(ilTemp).iMnfSSCode          'Sales source
                    Exit For
                End If
            Next ilTemp
        Else
            tmGrf.iSofCode = 0
            tmGrf.iCode2 = 0
        End If
End Sub

'
'
'              Monthly Sales Commission
'
'       5-11-04 Build prepass of year to date transactions (IN/AN) from RVF/PHF to
'               determine year to date sales commission along with monthly sales comm.
'
'           Build array of valid transations;
'           Loop thru array of transactions and split the sales share based on cnt header
'           and store in tmRvfComm array;
'           After all gathered and split, sort the array by slsp, cnt , tran date;
'           Process one slsp at a time to determine if the goal has been reached in the past,
'           or reached during the current month.  Purpose is to determine which percentage to
'           use before writing prepass record for Crystal.
'
'       11-8-04 Gather for std dates (start of year to end of requested month) regardless if
'           using corporate cal
'       11-30-04 Allow multiple months to be gathered
'       3-17-05 Obtain NTR Items to see if hard cost applicable
'       4-12-07 option to include/exclude Hard cost
'       10-07-07 implement tests for Installment bill types
'               Look at rvftype = "A" and ""

Public Function gBuildSlsCommCt()
    Dim ilRet As Integer
    Dim illoop  As Integer
    Dim slCode As String
    Dim ilFound As Integer
    Dim slStr As String
    Dim ilStartFiscalMonth As Integer
    Dim ilCurrentMonth As Integer
    Dim ilYear As Integer
    Dim ilLoopYear As Integer
    Dim llMonthEndDate As Long                'end date of current month
    Dim llCommCalcDate As Long              'start of current commission Month
    Dim llCalStartDate As Long              'start date of corp/std year
    Dim llLYCalStartDate As Long            'start date of last year corp/std
    Dim ilLYStartDate(0 To 1) As Integer 'start date of last year corp/std for prepass
    Dim llDate As Long
    Dim llTempStart As Long
    Dim llTempEnd As Long
    Dim slGrossOrNet As String              'Base commission on G = Gross or N = Net
    'ReDim tlSlf(1 To 1) As SLF
    ReDim tlSlf(0 To 0) As SLF
    Dim llSingleCntr As Long                        'single contract #
    Dim ilUseSlsComm As Integer                 'using sub-company commissions
    Dim llLoopOnRvf As Long
    Dim ilEarliestDate(0 To 1) As Integer
    Dim ilLatestDate(0 To 1) As Integer
    Dim ilExtLen As Integer
    Dim ilLoopOnFile As Integer
    Dim llNoRec As Long
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilOffSet As Integer
    Dim llRecPos As Long
    Dim llPrevNetNetRemnant As Long
    Dim dlPrevNetNet As Double      'Dim llPrevNetNet As Long   10-6-08

    Dim llStartInx As Long
    Dim llEndInx As Long
    Dim llSalesGoal As Long
    Dim ilCurrentSlf As Integer
    Dim llAdjustSlf As Long
    Dim ilfirstTime As Integer
    Dim tlTranType As TRANTYPES
    Dim slYear As String    '11-09-04
    Dim slMonth As String
    Dim slDay As String
    Dim ilPos1 As Integer
    Dim ilMonths As Integer '11-30-04
    Dim slTimeStamp As String   '3-17-05
    Dim ilIsItHardCost As Integer
    Dim ilBonusVersion As Integer   '8-27-08
    Dim ilIsItPolitical As Integer  '4-14-15      polit flag for advt test
    Dim ilOKtoSeeVeh As Integer     '11-22-17

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Exit Function
    End If
    imGrfRecLen = Len(tmGrf)
    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        btrDestroy hmGrf
        Exit Function
    End If
    imRvfRecLen = Len(tmRvf)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        btrDestroy hmRvf
        btrDestroy hmGrf
        Exit Function
    End If
    imCHFRecLen = Len(tmChf)

    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        btrDestroy hmCHF
        btrDestroy hmSlf
        btrDestroy hmRvf
        btrDestroy hmGrf
        Exit Function
    End If
    imSlfRecLen = Len(tmSlf)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    btrDestroy hmCHF
    btrDestroy hmSlf
    btrDestroy hmRvf
    btrDestroy hmGrf
    Exit Function
    End If
    imVefRecLen = Len(tmVef)

    hmScf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmScf, "", sgDBPath & "Scf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmScf)
    btrDestroy hmScf
    btrDestroy hmVef
    btrDestroy hmCHF
    btrDestroy hmSlf
    btrDestroy hmRvf
    btrDestroy hmGrf
    Exit Function
    End If
    imScfRecLen = Len(tmScf)

    hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmSbf)
    btrDestroy hmSbf
    btrDestroy hmScf
    btrDestroy hmVef
    btrDestroy hmCHF
    btrDestroy hmSlf
    btrDestroy hmRvf
    btrDestroy hmGrf
    Exit Function
    End If
    imSbfRecLen = Len(tmSbf)
    
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVsf)
        btrDestroy hmSbf
        btrDestroy hmScf
        btrDestroy hmVef
        btrDestroy hmCHF
        btrDestroy hmSlf
        btrDestroy hmRvf
        btrDestroy hmGrf
        Exit Function
    End If
    imVsfRecLen = Len(hmVsf)
    
    gBuildSlsCommCt = True


    '4-20-00
    'ReDim tlScf(1 To 1) As SCF
    ReDim tlScf(0 To 0) As SCF
    'ReDim tlSlfIndex(1 To 1) As SLFINDEX    '
    ReDim tlSlfIndex(0 To 0) As SLFINDEX    '
    ilRet = gObtainVef()                  'buildglobal vehicle table
    If ilRet = 0 Then
        btrDestroy hmScf
        btrDestroy hmVef
        btrDestroy hmCHF
        btrDestroy hmSlf
        btrDestroy hmRvf
        btrDestroy hmGrf
        btrDestroy hmVef
        Exit Function
    End If

    If RptSelCt!ckcSelC3(0).Value = vbChecked Then    'bonus version with new  & increased sales
        ilBonusVersion = True
    Else
        ilBonusVersion = False
    End If

    ilRet = gObtainMnfForType("I", slTimeStamp, tlMMnf())   'populate NTR types to determine hard cost applicable

    ilRet = gObtainSlf(RptSelCt, hmSlf, tlSlf())    'keep slsp in memory
    'ReDim (LBound(tlSlf) To UBound(tlSlf))      'retain same # of slsp entries in memory so that the Current years YTD $ can be accumulated
    ReDim tmSalesGoalInfo(LBound(tlSlf) To UBound(tlSlf))       'retain same # of slsp entries in memory so that the Current years YTD $ & sales goal info can be accumulated
    '4-20-00 see if selective contract #
    slStr = RptSelCt!edcSelCTo1.Text
    If slStr = "" Then
        llSingleCntr = 0
    Else
        llSingleCntr = CLng(slStr)
    End If
    slGrossOrNet = "N"                  'force comm on net for now
    'Determine calendar month requested, and retrieve all History and Receivables
    'records that fall within the beginning of the cal year and end of calendar month requested
    slStr = RptSelCt!edcSelCFrom.Text             'month in text form (jan..dec)
    ilYear = Val(RptSelCt!edcSelCFrom1.Text)
    gGetMonthNoFromString slStr, ilCurrentMonth      'getmonth #

    '11-30-04 determine how many months requested, previously was only 1 month
    slStr = RptSelCt!edcSelCTo.Text                '# periods
    ilMonths = Val(slStr)
    If ilMonths < 1 Or ilMonths > 12 Then
        ilMonths = 1
    End If
    slMonth = "JanFebMarAprMayJunJulAugSepOctNovDec"

    ilPos1 = ilCurrentMonth + ilMonths - 1    'add the number of months
    If ilPos1 > 12 Then
        ilPos1 = ilPos1 - 12
        ilYear = ilYear + 1             'wrap around to next year
    End If

    'use the last month requested
    slStr = Trim$(str$(ilPos1)) & "/1/" & str$(ilYear)
    'slStr = Trim$(Str$(ilCurrentMonth)) & "/1/" & Trim$(RptSelCt!edcSelCFrom1.Text)
    'slStr = gObtainEndCal(slStr)               'obtain cal month for end date to gather
    slStr = gObtainEndStd(slStr)               '11-9-04 obtain std month for end date to gather
    llMonthEndDate = gDateValue(slStr)


    slStr = RptSelCt!edcSelCFrom.Text             'month in text form (jan..dec)
    gGetMonthNoFromString slStr, ilRet        'getmonth #
    slStr = Trim$(str$(ilRet)) & "/1/" & Trim$(RptSelCt!edcSelCFrom1.Text)
    slStr = gObtainStartStd(slStr)   '4-20-00 chged to use start of bdcst date instead of end of bdcst month
                                             'any adjustments entered using the middle of the month were not included as current        If tgSpf.sRUseCorpCal = "Y" Then
    llCommCalcDate = gDateValue(slStr)      'std start date of month requested to be used to determine what $ is accumulated for previous months

    slStr = "1/1/" & Trim$(RptSelCt!edcSelCFrom1.Text)  'get beginning of year so start of std bdcst month can be determined
    slStr = gObtainStartStd(slStr)               '11-9-04 obtain std start date of year
    llCalStartDate = gDateValue(slStr)

    ilStartFiscalMonth = 1          'assume Jan-dec std year
    'If tgSpf.sRUseCorpCal = "Y" Then
    If RptSelCt!rbcSelC4(0).Value = True Then       '8-19-09 use corporate calendar

        ilRet = gObtainCorpCal()
        For ilLoopYear = 1 To 2         'find current/last year start
            ilFound = False
            ilYear = Val(RptSelCt!edcSelCFrom1.Text)
            If ilLoopYear = 2 Then              'last year
                ilYear = ilYear - 1
            End If
            slStr = Trim$(str$(ilCurrentMonth)) & "/15/" & Trim$(str(ilYear))
            llDate = gDateValue(slStr)      'find the month/year within the corporate calendar
            For illoop = LBound(tgMCof) To UBound(tgMCof)
                'gUnpackDate tgMCof(ilLoop).iStartDate(0, 1), tgMCof(ilLoop).iStartDate(1, 1), slStr
                gUnpackDate tgMCof(illoop).iStartDate(0, 0), tgMCof(illoop).iStartDate(1, 0), slStr
                llTempStart = gDateValue(slStr)
                'gUnpackDate tgMCof(ilLoop).iEndDate(0, 12), tgMCof(ilLoop).iEndDate(1, 12), slStr
                gUnpackDate tgMCof(illoop).iEndDate(0, 11), tgMCof(illoop).iEndDate(1, 11), slStr
                llTempEnd = gDateValue(slStr)
                If llDate >= llTempStart And llDate <= llTempEnd Then
                    ilFound = True
                    Exit For
                End If
             Next illoop
             If ilFound Then
                 'gUnpackDate tgMCof(ilLoop).iStartDate(0, 1), tgMCof(ilLoop).iStartDate(1, 1), slStr
                 gUnpackDate tgMCof(illoop).iStartDate(0, 0), tgMCof(illoop).iStartDate(1, 0), slStr
                 'get the month,day & year for this corporate start date so that it can
                 'be converted to standard month with the correct year
                 gObtainYearMonthDayStr slStr, False, slYear, slMonth, slDay
                 slStr = str$(tgMCof(illoop).iStartMnthNo) & "/15/" & Trim$(slYear)
                 slStr = gObtainStartStd(slStr)               '11-9-04 obtain std month for end date to gather
                 If ilLoopYear = 2 Then             'last year
                    llLYCalStartDate = gDateValue(slStr)
                 Else
                    llCalStartDate = gDateValue(slStr)
                 End If
                 ilStartFiscalMonth = tgMCof(illoop).iStartMnthNo   'start month of fiscal year
                 'llCalStartDate = llTempStart        'start of the corp year
             Else
                 MsgBox "Corporate calendar not found for requested date"
                 ilRet = btrClose(hmGrf)
                 ilRet = btrClose(hmRvf)
                 ilRet = btrClose(hmCHF)
                 ilRet = btrClose(hmSlf)
                 Exit Function
             End If
         Next ilLoopYear
     Else                       'not corp, get standard start date of last year
        ilYear = Val(RptSelCt!edcSelCFrom1.Text) - 1
        slStr = "1/15/" & Trim$(str(ilYear))
        slStr = gObtainStartStd(slStr)
        llLYCalStartDate = gDateValue(slStr)      'std start date of last year std
     End If

    gPackDateLong llLYCalStartDate, ilLYStartDate(0), ilLYStartDate(1)
    ' slstr = Str$(ilStartFiscalMonth) & "/1/" & Trim$(RptSelCt!edcSelCFrom1.Text)
     'llCalStartDate = gDateValue(slStr)
    ilUseSlsComm = False                    'not using sub-company commissions
    If tgSpf.sSubCompany = "Y" Then         'using commission by vehicle (sub-company)?
        '4-20-00 Build  slsp commission table
        ilRet = gObtainScf(RptSelCt, hmScf, tlScf(), llCalStartDate, llMonthEndDate, tlSlfIndex())
        ilUseSlsComm = True
    End If
    tlTranType.iAirTime = True
    tlTranType.iAdj = True
    tlTranType.iInv = True
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iCash = True
    tlTranType.iTrade = False
    tlTranType.iMerch = True
    tlTranType.iPromo = True
    tlTranType.iNTR = True

    If RptSelCt!rbcSelC9(0).Value = True Then           'Air time only
        tlTranType.iNTR = False
    ElseIf RptSelCt!rbcSelC9(1).Value = True Then
        tlTranType.iAirTime = False            'include ntr only
    End If
    '4-2-07 Hard cost option by user
    If RptSelCt!ckcSelC10(0).Value = vbChecked Then
        tlTranType.iHardCost = True
    Else
        tlTranType.iHardCost = False
    End If

    '7-16-08 Include Acquistion costs
    imInclAcq = True                        'assume to include acq costs
    If RptSelCt!ckcSelC12(0).Value = vbUnchecked Then     'exclude acq costs
        imInclAcq = False
    End If

    slStr = Format$(llCalStartDate, "m/d/yy")
    slCode = Format$(llMonthEndDate, "m/d/yy")

    gPackDate slStr, ilEarliestDate(0), ilEarliestDate(1)
    gPackDate slCode, ilLatestDate(0), ilLatestDate(1)

    If ilBonusVersion Then          'version to do new/excess over last year
        slStr = Format$(llCalStartDate, "m/d/yy")      'start date for this year
        ilRet = gSetFormula("TYStartDate", "'" & slStr & "'")  'need this year start date to accum this years net-net
        'adjust earliest date to retrieve info for last years info
        slStr = Format$(llLYCalStartDate, "m/d/yy")      'start date for last year
        ilRet = gSetFormula("LYStartDate", "'" & slStr & "'")  'need last year start date to accum last years net-net
        gPackDate slStr, ilEarliestDate(0), ilEarliestDate(1)
    End If


    ilRet = gSetFormula("GatherDates", "'" & slStr & " - " & slCode & "'")  'show the dates in the legend of the printed report

    gObtainCodesForMultipleLists 6, tgCSVNameCode(), imIncludeCodes, imUseCodes(), RptSelCt
    
    ReDim tmRVFComm(0 To 0) As RVFCOMM
    
    'ilRet = gBuildAcqCommInfo(RptSelCt)         'build acq rep commission table, if applicable


    For ilLoopOnFile = 1 To 2         'pass 1- get PHF, pass 2 get RVF
        btrExtClear hmRvf           'clear any previous extend operations
        ilExtLen = Len(tmRvf)       'extract record size
        ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
            Call btrExtSetBounds(hmRvf, llNoRec, -1, "UC", "RVF", "") '"EG") 'Set extract limits (all records)

            tlDateTypeBuff.iDate0 = ilEarliestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilEarliestDate(1)
            ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
            ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gBuildSlsCommCt (btrExtAddLogicConst):" & "Rvf.Btr", RptSelCt
            On Error GoTo 0


            tlDateTypeBuff.iDate0 = ilLatestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilLatestDate(1)
            ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
            ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gBuildSlsCommCt (btrExtAddLogicConst):" & "Rvf.Btr", RptSelCt
            On Error GoTo 0

            ilRet = btrExtAddField(hmRvf, 0, ilExtLen)  'Extract the whole record
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gBuildSlsCommCt (btrExtAddField):" & "RVF.Btr", RptSelCt
            On Error GoTo 0
            ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo mRvfErr
                gBtrvErrorMsg ilRet, "gBuildSlsCommCt (btrExtGetNextExt):" & "RVF.Btr", RptSelCt
                On Error GoTo 0
                ilExtLen = Len(tmRvf)  'Extract operation record size
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    'obtain the contract header to see if slsp allowed to see this contract
                    tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
                    'only look for HOGN statuses, look for the valid header belonging to the transaction
                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sStatus <> "H" And tmChf.sStatus <> "O" And tmChf.sStatus <> "G" And tmChf.sStatus <> "N")
                        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    gFakeChf tmRvf, tmChf
                    ilOKtoSeeVeh = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, CLng(tmRvf.iBillVefCode), tmChf.iSlfCode())
                    'first test for valid trans types (Invoices, adjustments, write-off & payments
                    If (llSingleCntr <> 0 And llSingleCntr = tmRvf.lCntrNo) Or llSingleCntr = 0 And ilOKtoSeeVeh = True Then
                        '10-7-07 test for inclusion of specific installment types or regular revenue
                        ilIsItPolitical = gIsItPolitical(tmRvf.iAdfCode)           '4-14-15 is tran political and ok to include polit, or not political and OK to include non-polit
                        If (ilIsItPolitical = True And RptSelCt!ckcSelC13(0).Value = vbChecked) Or (ilIsItPolitical = False And RptSelCt!ckcSelC13(1).Value = vbChecked) Then
                            If ((((Left$(tmRvf.sTranType, 1) = "I" Or tmRvf.sTranType = "HI") And tlTranType.iInv) Or (Left$(tmRvf.sTranType, 1) = "A" And tlTranType.iAdj) Or (Left$(tmRvf.sTranType, 1) = "W" And tlTranType.iWriteOff) Or (Left$(tmRvf.sTranType, 1) = "P" And tlTranType.iPymt)) And (Trim$(tmRvf.sType) = "" Or tmRvf.sType = "A")) Then
                                'If (tlTranType.iNTR) Then       'NTR option, tested separately because it shouldnt be tested with Cash transactions
                                If tmRvf.iMnfItem > 0 Then      'ntr billing
                                    If (tlTranType.iNTR) Then       'NTR option, tested separately because it shouldnt be tested with Cash transactions
                                        'test if hard cost, another option to include/exclude
                                        ilIsItHardCost = gIsItHardCost(tmRvf.iMnfItem, tlMMnf())
                                        If ilIsItHardCost Then           'its hard cost, should it be included?
                                            If tlTranType.iHardCost And tmRvf.sCashTrade <> "T" Then  'hard included as option and its not a trade.  all trades are ignored
                                                mSlspCommExt ilUseSlsComm, slGrossOrNet, tlScf(), tlSlf(), tlSlfIndex(), ilBonusVersion, ilStartFiscalMonth, ilCurrentMonth    'build record
                                            End If
                                        Else            'not hard cost
                                            If tmRvf.sCashTrade <> "T" Then     'all transaction except trade can be included
                                                mSlspCommExt ilUseSlsComm, slGrossOrNet, tlScf(), tlSlf(), tlSlfIndex(), ilBonusVersion, ilStartFiscalMonth, ilCurrentMonth     'build record
                                            End If
                                        End If
                                    'Else            'its not an NTR
                                    '    'got valid trans type - test for Cash, Trade, Merchandising or Promotions
                                   '    If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                    '        mSlspCommExt ilUseSlsComm, slGrossOrNet, tlScf(), tlSlf(), tlSlfIndex()     'build record
                                    '    End If
                                    End If
                                Else
                                    'got valid trans type - test for Cash, Trade, Merchandising or Promotions
                                    If tlTranType.iAirTime Then
                                        If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                            mSlspCommExt ilUseSlsComm, slGrossOrNet, tlScf(), tlSlf(), tlSlfIndex(), ilBonusVersion, ilStartFiscalMonth, ilCurrentMonth      'build record
                                        End If
                                    End If
                                End If
                            End If          'trantype
                        End If              'isitpolitical
                    End If
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
        If ilLoopOnFile = 1 Then                          'if 1, then just finished history, go do Receivables
            btrExtClear hmRvf   'Clear any previous extend operation
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmRvf)
                btrDestroy hmRvf
                gBuildSlsCommCt = False
                Exit Function
            End If
            imRvfRecLen = Len(tmRvf)
        End If
    Next ilLoopOnFile
    ilRet = btrClose(hmRvf)
    btrDestroy hmRvf


    'sort the array of all transactions for the year by slsp, advt & trans date.  Then loop thru
    'for one slsp at a time to determine what slsp commission % should be used .  If sales goal
    'made prior to the current reporting month, use the overage.  If goal made during the current
    'month, use under goal for detail; then use the overage % minus under goal (getting the difference)
    'for the overage calculated at the end of the detail.
    'The whole purpose of this is to setup the proper % to use and store in prepass, because of the
    'need to sort by the % at print time.  Crystal cannot dynamically change the % during print time
    'while it has to use it as calculations.
    If UBound(tmRVFComm) > 0 Then
        'Sort the array by slsp code, cntr, trans date
        ArraySortTyp fnAV(tmRVFComm(), 0), UBound(tmRVFComm), 0, LenB(tmRVFComm(0)), 0, LenB(tmRVFComm(0).sKey), 0

        'go thru all the transactions and setup the comm% to use based on if goal hit or not
        'if goal hit, use overage %
        ilfirstTime = True


        For llLoopOnRvf = LBound(tmRVFComm) To UBound(tmRVFComm) - 1
            If ilfirstTime Then
                ilfirstTime = False
                llStartInx = LBound(tmRVFComm)
                llEndInx = llStartInx
            End If
            dlPrevNetNet = 0
            llPrevNetNetRemnant = 0
            llSalesGoal = tmRVFComm(llLoopOnRvf).lGoal
            ilCurrentSlf = tmRVFComm(llLoopOnRvf).iSlfCode
            gUnpackDateLong tmRVFComm(llLoopOnRvf).iTranDate(0), tmRVFComm(llLoopOnRvf).iTranDate(1), llDate

            Do While ilCurrentSlf = tmRVFComm(llLoopOnRvf).iSlfCode

                gUnpackDateLong tmRVFComm(llLoopOnRvf).iTranDate(0), tmRVFComm(llLoopOnRvf).iTranDate(1), llDate
                'Previous NetNet:  accumulate remnant vs non-remnant year to date (minus current month)
                'to see if slsp has already reached sales goals
                If tmRVFComm(llLoopOnRvf).sType = "T" Then        'remnants in the past
                    If llDate >= llCalStartDate And llDate < llCommCalcDate Then
                        llPrevNetNetRemnant = llPrevNetNetRemnant + tmRVFComm(llLoopOnRvf).lNet - tmRVFComm(llLoopOnRvf).lMerch - tmRVFComm(llLoopOnRvf).lPromo '9-12-08 was subt merch twice instead of promo
                    End If
                Else
                    If llDate >= llCalStartDate And llDate < llCommCalcDate Then
                        dlPrevNetNet = dlPrevNetNet + tmRVFComm(llLoopOnRvf).lNet - tmRVFComm(llLoopOnRvf).lMerch - tmRVFComm(llLoopOnRvf).lPromo       '9-12-08 was subt merch twice instead of promo
                    End If
                End If
                llEndInx = llLoopOnRvf
                llLoopOnRvf = llLoopOnRvf + 1
            Loop

            'Loop thru the one slsp and adjust the slsp comm % to use in the report
            For llAdjustSlf = llStartInx To llEndInx
                gUnpackDateLong tmRVFComm(llAdjustSlf).iTranDate(0), tmRVFComm(llAdjustSlf).iTranDate(1), llDate
                If tmRVFComm(llAdjustSlf).sType = "T" Then        'remnants in the past
                    tmRVFComm(llAdjustSlf).iUseThisCommPct = tmRVFComm(llAdjustSlf).iRemUnderPct        'assume not reached goal yet
                    If llDate >= llCalStartDate And llDate < llCommCalcDate Then
                        'in the past, has slsp already met sales goals
                        If dlPrevNetNet >= llSalesGoal And llSalesGoal > 0 Then
                            'use the overage in the past
                            tmRVFComm(llAdjustSlf).iUseThisCommPct = tmRVFComm(llAdjustSlf).iRemOverPct
                        Else
                            tmRVFComm(llAdjustSlf).iUseThisCommPct = tmRVFComm(llAdjustSlf).iRemUnderPct
                        End If
                    End If
                Else                'non-remnant

                    tmRVFComm(llAdjustSlf).iUseThisCommPct = tmRVFComm(llAdjustSlf).iSlfCommPct     'assume not reached goal yet
                    If tmRVFComm(llAdjustSlf).sNTR <> "N" Then     'dont alter the comm for ntr
                        If llDate >= llCalStartDate And llDate < llCommCalcDate Then
                            'in the past, has slsp already met sales goals
                            If dlPrevNetNet >= llSalesGoal And llSalesGoal > 0 Then
                                'use the overage in the past
                                tmRVFComm(llAdjustSlf).iUseThisCommPct = tmRVFComm(llAdjustSlf).iSlfOverCommPct
                            Else
                                tmRVFComm(llAdjustSlf).iUseThisCommPct = tmRVFComm(llAdjustSlf).iSlfCommPct
                            End If
                        End If
                    End If
                End If
            Next llAdjustSlf
            llStartInx = llEndInx + 1
            llEndInx = llStartInx
            llLoopOnRvf = llLoopOnRvf - 1
        Next llLoopOnRvf

        'now write out all the transactions to disk

        If ilBonusVersion Then
            mBuildcommByAdvt llLYCalStartDate, llCommCalcDate, llMonthEndDate, llCalStartDate, tlSlf()
        Else
            For llLoopOnRvf = LBound(tmRVFComm) To UBound(tmRVFComm) - 1
                tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
                tmGrf.iGenDate(1) = igNowDate(1)
                gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                tmGrf.lGenTime = lgNowTime
                tmGrf.iDate(0) = ilLYStartDate(0)           'last year start date, also sent as formula
                tmGrf.iDate(1) = ilLYStartDate(1)
                tmGrf.lChfCode = tmRVFComm(llLoopOnRvf).lChfCode           'contr internal code
                tmGrf.iSlfCode = tmRVFComm(llLoopOnRvf).iSlfCode
                tmGrf.iSofCode = tmRVFComm(llLoopOnRvf).iSofCode                'office code
                'tmGrf.iDateGenl(0, 1) = tmRVFComm(llLoopOnRvf).iTranDate(0)    'date billed or paid
                'tmGrf.iDateGenl(1, 1) = tmRVFComm(llLoopOnRvf).iTranDate(1)
                tmGrf.iDateGenl(0, 0) = tmRVFComm(llLoopOnRvf).iTranDate(0)    'date billed or paid
                tmGrf.iDateGenl(1, 0) = tmRVFComm(llLoopOnRvf).iTranDate(1)
                tmGrf.iVefCode = tmRVFComm(llLoopOnRvf).iAirVefCode
                'tmGrf.lDollars(1) = tmRVFComm(llLoopOnRvf).lInvNo          'Invoice #
                'tmGrf.lDollars(2) = tmRVFComm(llLoopOnRvf).lGross          'Gross $
                'tmGrf.lDollars(3) = tmRVFComm(llLoopOnRvf).lNet             'Net$
                'tmGrf.lDollars(4) = tmRVFComm(llLoopOnRvf).lMerch          'Merchandising $
                'tmGrf.lDollars(5) = tmRVFComm(llLoopOnRvf).lPromo          'Promotions $
                'tmGrf.lDollars(6) = tmRVFComm(llLoopOnRvf).lAdjComm         'adjust $ (net - merch - promo)
                'tmGrf.lDollars(7) = tmRVFComm(llLoopOnRvf).iUseThisCommPct       'slsp % non-remnant under goal
                'tmGrf.lDollars(8) = tmRVFComm(llLoopOnRvf).lSlfSplit    'slsp revenue share
                'tmGrf.lDollars(9) = tmRVFComm(llLoopOnRvf).iRemUnderPct    '3-15-00 slsp remnant % under
                tmGrf.lDollars(0) = tmRVFComm(llLoopOnRvf).lInvNo          'Invoice #
                tmGrf.lDollars(1) = tmRVFComm(llLoopOnRvf).lGross          'Gross $
                tmGrf.lDollars(2) = tmRVFComm(llLoopOnRvf).lNet             'Net$
                tmGrf.lDollars(3) = tmRVFComm(llLoopOnRvf).lMerch          'Merchandising $
                tmGrf.lDollars(4) = tmRVFComm(llLoopOnRvf).lPromo          'Promotions $
                tmGrf.lDollars(5) = tmRVFComm(llLoopOnRvf).lAdjComm         'adjust $ (net - merch - promo)
                tmGrf.lDollars(6) = tmRVFComm(llLoopOnRvf).iUseThisCommPct       'slsp % non-remnant under goal
                tmGrf.lDollars(7) = tmRVFComm(llLoopOnRvf).lSlfSplit    'slsp revenue share
                tmGrf.lDollars(8) = tmRVFComm(llLoopOnRvf).iRemUnderPct    '3-15-00 slsp remnant % under

                'tmGrf.iPerGenl(1) = tmRVFComm(llLoopOnRvf).iStartFiscalMonth
                'tmGrf.iPerGenl(2) = tmRVFComm(llLoopOnRvf).iCurrentMonth
                'tmGrf.iPerGenl(6) = tmRVFComm(llLoopOnRvf).iDiscrep                       'error flag: discrepancy in vehicle/slsp commission (sub-companies)
                tmGrf.iPerGenl(0) = tmRVFComm(llLoopOnRvf).iStartFiscalMonth
                tmGrf.iPerGenl(1) = tmRVFComm(llLoopOnRvf).iCurrentMonth
                tmGrf.iPerGenl(5) = tmRVFComm(llLoopOnRvf).iDiscrep                       'error flag: discrepancy in vehicle/slsp commission (sub-companies)
                'if inconsistency with vehicle/sub-company definition, flag as error to show on report
                tmGrf.sDateType = tmRVFComm(llLoopOnRvf).sNTR
'                If imInclAcq Then                           'incl acq costs by user request
                    'tmGrf.lDollars(10) = tmRVFComm(llLoopOnRvf).lAcquisition        'acq has been zeroed in previous routine if user requested to exclude
                    tmGrf.lDollars(9) = tmRVFComm(llLoopOnRvf).lAcquisition        'acq has been zeroed in previous routine if user requested to exclude
'                Else
'                    tmGrf.lDollars(10) = 0
'                End If
                tmGrf.iAdfCode = tmRVFComm(llLoopOnRvf).iAdfCode        '5-6-07
                tmGrf.lCode4 = tmRVFComm(llLoopOnRvf).lContrNo          '5-6-07
                tmGrf.lLong = tmRVFComm(llLoopOnRvf).lRvfCode
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            Next llLoopOnRvf
            Erase tmRVFComm
        End If
    End If



    Erase tlSlf, tlScf, tlSlfIndex      '4-20-00
   
    ilRet = btrClose(hmRvf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmScf)
    ilRet = btrClose(hmSbf)
    btrDestroy hmRvf
    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmSlf
    btrDestroy hmVef
    btrDestroy hmScf
    btrDestroy hmSbf
    Exit Function

mRvfErr:
    On Error GoTo 0
    gBuildSlsCommCt = False
End Function
'
'
'           mSlspcommExt - History or Receivables record is ready to be processed
'           for Monthly Salesperson Commission Statement.
'           If any slsp revenue splits, the record must be split according to the
'           contract header splits, then the slsp % of commission to be determined.
'           Build into array so that it can be sorted and manipulated before
'           writing to GRF.
'           5-12-04
'
'           <input> iluseSlsComm = true if using Sub company commissions
'                   slGrossOrNet - G = gross commissions, N = net comm
'                   tlScf - array of Slsp comm table
'                   tlSlf - array of slsp records
'                   tlslfIndex - start/end indices of a slsp commission values
'                   ilBonusVersion -true to sort differently for bonus calculation
'                   Assume RVF/PHF record in common to module
'           3-17-05 determine if hard cost NTR item

Public Sub mSlspCommExt(ilUseSlsComm As Integer, slGrossOrNet As String, tlScf() As SCF, tlSlf() As SLF, tlSlfIndex() As SLFINDEX, ilBonusVersion As Integer, ilStartFiscalMonth As Integer, ilCurrentMonth As Integer)
Dim ilRet As Integer
Dim illoop  As Integer
Dim slNameCode As String
Dim slCode As String
Dim slStr As String
Dim slPct As String
Dim llAmt As Long
Dim llDate As Long
Dim ilTemp As Integer
Dim slAmount As String
Dim slDollar As String
Dim llDollar As Long
Dim ilFoundSlsp As Integer
Dim ilLoopSlsp As Integer
Dim ilMnfSubCo As Integer                   '4-20-00
Dim ilSlfRecd As Integer                    '4-20-00
Dim ilOKtoSeeVeh As Integer                 '11-13-03
Dim llRVFCOMMUpper As Long
Dim slTranDate As String
Dim ilHCLoop As Integer                     '3-17-05
Dim ilAdfInx As Integer
Dim slContractType As String * 1            'for bonus version only:  R = remnant, S = standard
Dim ilAcqCommPct As Integer
Dim ilAcqLoInx As Integer
Dim ilAcqHiInx As Integer
Dim llAcqNet As Long
Dim llAcqComm As Long
Dim blAcqOK As Boolean


    llRVFCOMMUpper = UBound(tmRVFComm)
    gPDNToLong tmRvf.sNet, llAmt
    gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slStr
    llDate = gDateValue(slStr)                  'convert trans date to test if within requested limits


    'valid record must be an "Invoice" or "Adjustment" type, non-zero amount, and transaction date within the start date of the
    'cal year and end date of the current cal month requested.  Filtered out thru gObtainPhfRvf
    'If llAmt <> 0 Then
        'get contract from history or rec file
        
        '11-22-17 this has been moved to the mainline section of code, avoid reading contract twice
'        tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
'        tmChfSrchKey1.iCntRevNo = 32000
'        tmChfSrchKey1.iPropVer = 32000
'        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
'        'Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo <> tmRvf.lCntrNo Or (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M"))
'        Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")
'             ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'        gFakeChf tmRvf, tmChf       '4-20-00
'        If ((ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus = "F" Or tmChf.sSchStatus = "M")) Then
         If ((tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus = "F" Or tmChf.sSchStatus = "M")) Then
            For illoop = 0 To 9 Step 1
                If tmChf.iSlfCode(illoop) > 0 Then
                    ilTemp = ilTemp + 1
                End If
            Next illoop
            If ilTemp = 1 Then                      'only 1 slsp, force to 100% (no splits)
                tmChf.lComm(0) = 1000000             'force 100.0000%
            End If
            '4-20-99
            ReDim llSlfSplit(0 To 9) As Long           '4-20-00 slsp slsp share %
            ReDim ilSlfCode(0 To 9) As Integer             '4-20-00
            ReDim ilslfcomm(0 To 9) As Integer             'slsp under comm %
            ReDim ilslfremnant(0 To 9) As Integer          'slsp under remnant %
            ReDim llSlfSplitRev(0 To 9) As Long         '2-2-04 unused in this report (need for subroutine)


            ilMnfSubCo = gGetSubCmpy(tmChf, ilSlfCode(), llSlfSplit(), tmRvf.iAirVefCode, ilUseSlsComm, llSlfSplitRev())     '7-14-00
            gGetCommByDates ilSlfCode(), ilslfcomm(), ilslfremnant(), tlSlfIndex(), tlScf(), tlSlf(), tmRvf.iAirVefCode, llDate, tmChf '7-14-00
            For illoop = 0 To 9 Step 1          'see if there are any split commissions
                tmRVFComm(llRVFCOMMUpper).lInvNo = tmRvf.lInvNo          'Invoice #
                tmRVFComm(llRVFCOMMUpper).lChfCode = tmChf.lCode           'contr internal code
                tmRVFComm(llRVFCOMMUpper).iTranDate(0) = tmRvf.iTranDate(0)    'date billed or paid
                tmRVFComm(llRVFCOMMUpper).iTranDate(1) = tmRvf.iTranDate(1)
                tmRVFComm(llRVFCOMMUpper).iAirVefCode = tmRvf.iAirVefCode     '4-20-00 airing
                tmRVFComm(llRVFCOMMUpper).iAdfCode = tmRvf.iAdfCode         '5-6-07, use the advt from receivables since a chf header may not exist for backlog trans
                tmRVFComm(llRVFCOMMUpper).lContrNo = tmRvf.lCntrNo          '5-6-07 use contr # from Rvf because backlog may not have a chf header
                tmRVFComm(llRVFCOMMUpper).lRvfCode = tmRvf.lCode
                If ilSlfCode(illoop) > 0 And llSlfSplit(illoop) > 0 Then

                    ilFoundSlsp = False
                    If RptSelCt!ckcAll.Value = vbChecked Then
                        ilFoundSlsp = True
                    Else
                        For ilLoopSlsp = 0 To RptSelCt!lbcSelection(2).ListCount - 1 Step 1
                            If RptSelCt!lbcSelection(2).Selected(ilLoopSlsp) Then              'selected slsp
                                slNameCode = tgSalesperson(ilLoopSlsp).sKey      'pick up slsp code
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                '4-20-00 If Val(slCode) = tmChf.islfCode(ilLoop) Then
                                If Val(slCode) = ilSlfCode(illoop) Then
                                    ilFoundSlsp = True
                                    Exit For
                                End If
                            End If
                        Next ilLoopSlsp
                    End If
                    ilOKtoSeeVeh = gUserAllowedVehicle(tmRvf.iBillVefCode)  '11-13-03 check to see if users allowed to see this billing vehicle
                    If ilOKtoSeeVeh Then            '3-15-10 test selective vehicle
                        ilOKtoSeeVeh = False
                        If gFilterLists(tmRvf.iAirVefCode, imIncludeCodes, imUseCodes()) Then
                            ilOKtoSeeVeh = True
                        End If
                    End If
                    If ilFoundSlsp And ilOKtoSeeVeh Then
                        For ilSlfRecd = LBound(tlSlf) To UBound(tlSlf)
                            If ilSlfCode(illoop) = tlSlf(ilSlfRecd).iCode Then
                                tmSlf = tlSlf(ilSlfRecd)
                                Exit For
                            End If
                        Next ilSlfRecd

                        tmRVFComm(llRVFCOMMUpper).iSlfCode = tmSlf.iCode    'slsp code
                        tmRVFComm(llRVFCOMMUpper).iSlfCommPct = ilslfcomm(illoop)       '4-20-00 slsp share (xxx.xxxx)

                        tmRVFComm(llRVFCOMMUpper).iSlfOverCommPct = tmSlf.iOverComm
                        tmRVFComm(llRVFCOMMUpper).iRemOverPct = tmSlf.iRemOverComm
                        tmRVFComm(llRVFCOMMUpper).lGoal = tmSlf.lSalesGoal

                        tmRVFComm(llRVFCOMMUpper).lGross = 0                         'init gross, net, merch & promo $ fields
                        tmRVFComm(llRVFCOMMUpper).lNet = 0
                        tmRVFComm(llRVFCOMMUpper).lMerch = 0
                        tmRVFComm(llRVFCOMMUpper).lPromo = 0
                        tmRVFComm(llRVFCOMMUpper).lAcquisition = 0
                        
                        If imInclAcq Then
                            If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then
                                ilAcqCommPct = 0
                                blAcqOK = gGetAcqCommInfoByVehicle(tmRvf.iAirVefCode, ilAcqLoInx, ilAcqHiInx)
                                ilAcqCommPct = gGetEffectiveAcqComm(llDate, ilAcqLoInx, ilAcqHiInx)     'lldate is trans date
                                gCalcAcqComm ilAcqCommPct, tmRvf.lAcquisitionCost, llAcqNet, llAcqComm
                          Else
                                llAcqNet = tmRvf.lAcquisitionCost
                            End If
                        Else
                            llAcqNet = 0
                        End If
                        If tmRvf.sCashTrade = "C" Then
                            '4-20-00 slPct = gLongToStrDec(tmChf.lComm(ilLoop), 4)
                            slPct = gLongToStrDec(llSlfSplit(illoop), 4)
                            gPDNToStr tmRvf.sGross, 2, slAmount
                            slDollar = gMulStr(slAmount, slPct)
                            tmRVFComm(llRVFCOMMUpper).lGross = Val(gRoundStr(slDollar, "01.", 0))

                            gPDNToStr tmRvf.sNet, 2, slAmount
                            slDollar = gMulStr(slAmount, slPct)
                            tmRVFComm(llRVFCOMMUpper).lNet = Val(gRoundStr(slDollar, "01.", 0))

'                            If imInclAcq Then
'                                slAmount = gLongToStrDec(tmRvf.lAcquisitionCost, 2)
'                                slDollar = gMulStr(slAmount, slPct)
'                                tmRVFComm(llRVFCOMMUpper).lAcquisition = Val(gRoundStr(slDollar, "01.", 0))
'                            End If
                          
                        ElseIf tmRvf.sCashTrade = "M" Then                  'merchandising
                            '4-20-00 slPct = gLongToStrDec(tmChf.lComm(ilLoop), 4)
                            slPct = gLongToStrDec(llSlfSplit(illoop), 4)
                            gPDNToStr tmRvf.sGross, 2, slAmount
                            slDollar = gMulStr(slAmount, slPct)
                            tmRVFComm(llRVFCOMMUpper).lMerch = Val(gRoundStr(slDollar, "01.", 0))
'                            If imInclAcq Then
'                                slAmount = gLongToStrDec(tmRvf.lAcquisitionCost, 2)
'                                slDollar = gMulStr(slAmount, slPct)
'                                tmRVFComm(llRVFCOMMUpper).lAcquisition = Val(gRoundStr(slDollar, "01.", 0))
'                            End If
                        Else                                                 'promotions
                            '4-20-00 slPct = gLongToStrDec(tmChf.lComm(ilLoop), 4)
                            slPct = gLongToStrDec(llSlfSplit(illoop), 4)
                            gPDNToStr tmRvf.sGross, 2, slAmount
                            slDollar = gMulStr(slAmount, slPct)
                            tmRVFComm(llRVFCOMMUpper).lPromo = Val(gRoundStr(slDollar, "01.", 0))
'                            If imInclAcq Then
'                                slAmount = gLongToStrDec(tmRvf.lAcquisitionCost, 2)
'                                slDollar = gMulStr(slAmount, slPct)
'                                tmRVFComm(llRVFCOMMUpper).lAcquisition = Val(gRoundStr(slDollar, "01.", 0))
'                            End If
                        End If

                        slAmount = gLongToStrDec(llAcqNet, 2)
                        slDollar = gMulStr(slAmount, slPct)
                        tmRVFComm(llRVFCOMMUpper).lAcquisition = Val(gRoundStr(slDollar, "01.", 0))
                        
                        slPct = gIntToStrDec(ilslfcomm(illoop), 2)      'convert slsp comm % to packed dec.

                        '12-20-02 If NTR, need to get the SBF record for the slsp commission Pct
                       If tmRvf.iMnfItem > 0 Then          'this indicates NTR
                            'retrieve the associated NTR record from SBF
                           tmSbfSrchKey.lCode = tmRvf.lSbfCode  '12-16-02
                           ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                           If ilRet = BTRV_ERR_NONE Then
                               slPct = gIntToStrDec(tmSbf.iCommPct, 2)
                               If tgSpf.sSubCompany <> "Y" Then         '4-8-04 using commission by vehicle (sub-company)?
                                   'if not, then retrive the commission from SBF; otherwise it has already been retrieved from the vehicle slsp comm table
                                   tmRVFComm(llRVFCOMMUpper).iSlfCommPct = tmSbf.iCommPct
                               End If
                           Else
                               slPct = ".00"
                               tmRVFComm(llRVFCOMMUpper).iSlfCommPct = 0
                           End If
                       End If
                       If slGrossOrNet = "G" Then
                           llDollar = tmRVFComm(llRVFCOMMUpper).lGross - tmRVFComm(llRVFCOMMUpper).lMerch - tmRVFComm(llRVFCOMMUpper).lPromo - tmRVFComm(llRVFCOMMUpper).lAcquisition
                       Else
                           llDollar = tmRVFComm(llRVFCOMMUpper).lNet - tmRVFComm(llRVFCOMMUpper).lMerch - tmRVFComm(llRVFCOMMUpper).lPromo - tmRVFComm(llRVFCOMMUpper).lAcquisition
                       End If
                       slAmount = gLongToStrDec(llDollar, 2)           'adjusted rate to calc slsp comm. from
                       slAmount = gMulStr(slAmount, slPct)

                       tmRVFComm(llRVFCOMMUpper).lAdjComm = Val(gRoundStr(slAmount, "01.", 0))

                       tmRVFComm(llRVFCOMMUpper).iSofCode = tmSlf.iSofCode                 'office code

                        tmRVFComm(llRVFCOMMUpper).lSlfSplit = llSlfSplit(illoop)      'slsp revenue share
                        tmRVFComm(llRVFCOMMUpper).iRemUnderPct = ilslfremnant(illoop)    '3-15-00 slsp remnant % under

                        tmRVFComm(llRVFCOMMUpper).iStartFiscalMonth = ilStartFiscalMonth
                        tmRVFComm(llRVFCOMMUpper).iCurrentMonth = ilCurrentMonth
                        tmRVFComm(llRVFCOMMUpper).iDiscrep = 0                       'error flag: discrepancy in vehicle/slsp commission (sub-companies)
                        'if inconsistency with vehicle/sub-company definition, flag as error to show on report
                        If ilMnfSubCo < 0 Then
                            tmRVFComm(llRVFCOMMUpper).iDiscrep = 1
                        End If

                        '2-13-03 flag ntr transactions
                        If tmRvf.iMnfItem > 0 Then
                            '3-17-05 determine if hard cost item
                            For ilHCLoop = LBound(tlMMnf) To UBound(tlMMnf) - 1
                                If tmRvf.iMnfItem = tlMMnf(ilHCLoop).iCode Then
                                    If Trim(tlMMnf(ilHCLoop).sCodeStn) = "Y" Then
                                        tmRVFComm(llRVFCOMMUpper).sNTR = "H"
                                    Else
                                        tmRVFComm(llRVFCOMMUpper).sNTR = "N"
                                    End If
                                    Exit For
                                End If
                            Next ilHCLoop
                        Else
                            tmRVFComm(llRVFCOMMUpper).sNTR = ""
                        End If

                        tmRVFComm(llRVFCOMMUpper).sType = tmChf.sType

                        gUnpackDateForSort tmRvf.iTranDate(0), tmRvf.iTranDate(1), slTranDate
                        slStr = Trim$(str$(tmChf.lCntrNo))
                        Do While Len(slStr) < 8
                            slStr = "0" & slStr
                        Loop
                        slCode = Trim$(str$(tmRVFComm(llRVFCOMMUpper).iSlfCode))
                        Do While Len(slCode) < 8
                            slCode = "0" & slCode
                        Loop

                        slPct = Trim$(str$(tmRVFComm(llRVFCOMMUpper).iSlfCommPct))
                        Do While Len(slPct) < 5
                            slPct = "0" & slPct
                        Loop

                        If ilBonusVersion Then
                            slContractType = "S"        'assume standard contract
                            'keep standard contracts vs remnant apart to do over goal on remnants
                            If tmRVFComm(llRVFCOMMUpper).sType = "T" Then       'remnant
                                slContractType = "R"
                            End If
                            ilAdfInx = gBinarySearchAdf(tmRvf.iAdfCode)
                            If ilAdfInx = -1 Then       'no matching advt reference
                                'tgCommAdf(ilAdfInx).sName = "Missing Advt"
                                tmRVFComm(llRVFCOMMUpper).sKey = slCode & "|" & "Missing Advt" & "|" & slContractType & "|" & slPct    'sort major to minor: slf code, AdvtName , comm %
                            Else
                                tmRVFComm(llRVFCOMMUpper).sKey = slCode & "|" & Trim$(tgCommAdf(ilAdfInx).sName) & "|" & slContractType & "|" & slPct    'sort major to minor: slf code, AdvtName , comm %
                            End If
                            'tmRVFComm(llRVFCOMMUpper).sKey = slCode & "|" & Trim$(tgCommAdf(ilAdfInx).sName) & "|" & slContractType & "|" & slPct    'sort major to minor: slf code, AdvtName , comm %
                        Else
                            tmRVFComm(llRVFCOMMUpper).sKey = slCode & "|" & slStr & "|" & slTranDate    'sort major to minor: slf code, contr # , tran date
                        End If

                        ReDim Preserve tmRVFComm(0 To llRVFCOMMUpper + 1) As RVFCOMM
                        llRVFCOMMUpper = llRVFCOMMUpper + 1
                        'ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    End If                          ' if ilFoundSlsp
                End If                              'lcomm(ilLoop) > 0
            Next illoop                             'loop for 10 slsp possible splits
        End If
    'End If                                          'contr # doesnt match or not a fully sched contr

    Erase llSlfSplit, ilSlfCode, ilslfcomm, ilslfremnant

End Sub
'
'           mSplitCashTrade - split the $ for cash & trade for the Sales Activity reports
'                           (Weekly Sales Act by Month & Daily Sales Act by Month)
'           <input> llProject() - array of gross $ to be split for 12 months
'                   slGrossorNet - G = gross, N = net (user requested option)
'                   ilListIndex - report Sales Activity option
'           <output> llProjectCash() - array of Cash $ for 12 months stored in gross or net $
'                    llProjectTrade() - array of Trade $ for 12 months based stored in gross or net $
'
'                   Gross and Net $ are calculated based on user input.  Also, if contract
'                   is part cash, part trade, each % is taken in account
'                   to see if it is commissionable when net requested.
Public Sub mSplitCashTrade(llProject() As Long, llProjectCash() As Long, llProjectTrade() As Long, ilAgyCommPct As Integer, ilListIndex As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llGross                       ilRet                         ilZero                    *
'*                                                                                        *
'******************************************************************************************

Dim ilTemp As Integer               'Loop variable for number of vehicles to process
Dim ilPct As Integer                '% of cash or trade portion
Dim ilAgyComm As Integer            '% due station (minus the 15% agy comm), either 100 or 85
Dim ilTemp2 As Integer              'Cash/trade loop variable
Dim illoop As Integer               'loop variable for 12 months
Dim llTemp As Long                  'temp long variable for math, rounding, etc.
Dim llNoPenny As Long               'project $ without pennies
                                    'but theres values in different months (i.e. Jan: 5500, Feb: -500)
'Dim llTempDollars(1 To 12) As Long
ReDim llTempDollars(0 To 12) As Long    'Index zero ignored
    For ilTemp2 = 1 To 2                        'loop all vehicles for cash & trade (one order splits cash & trade)
         For illoop = 1 To 12
            llTempDollars(illoop) = 0
        Next illoop
        For illoop = 1 To 12
            If ilTemp2 = 1 Then                     'loop to calc cash $, then trade $
                ilPct = 100 - tgChfCT.iPctTrade       'get cash portion or order
                ilAgyComm = ilAgyCommPct
                'ilAgyComm = 10000                     'Assume gross requested
                'If slGrossOrNet = "N" Then          '
                '    If tgChfCT.iAgfCode > 0 Then      'agency exists,  net- take out commission
                '        ilAgyComm = 10000 - tmAgf.iComm    '85, prev forced to 15%, but now obtained from agency where its carried in 2 places
                '    End If
                'End If
            Else                                'trade portion
                ilPct = tgChfCT.iPctTrade         'trade portion of order
                ilAgyComm = ilAgyCommPct
                'ilAgyComm = 10000                 'assume no commissionable on trade
                If tgChfCT.sAgyCTrade = "Y" Then  'trade portion is commissionable, is it Gross orNet requested
                    'If slGrossOrNet = "N" Then
                    '    ilAgyComm = 10000 - tmAgf.iComm '85 prev forced to 15%, but now obtained from agency where its carried in 2 places
                    'End If
                    ilAgyComm = ilAgyCommPct
                End If
            End If
            If ilListIndex = CNT_SALESACTIVITY_SS Then
            'dont drop the pennies, Crystal will do that
                llNoPenny = llProject(illoop)     'retain pennies
                llTemp = llNoPenny * CDbl(ilPct) / 100   'calc cash vs trade
                llTemp = llTemp * CDbl(ilAgyComm) / 10000    'calc agy comm; adj for the comm % carried in 2 places
            Else
                llNoPenny = llProject(illoop) / 100    'drop pennies
                llTemp = llNoPenny * CDbl(ilPct) / 100    'calc cash vs trade
                llTemp = llTemp * CDbl(ilAgyComm) / 10000   'adjust for the agy comm carried in 2 places                               'calc agy comm
            End If
            llTempDollars(illoop) = llTemp
        Next illoop
        'return the split cash & trade dollars by the user requested gross or net $
        For ilTemp = 1 To 12
            If ilTemp2 = 1 Then
                llProjectCash(ilTemp) = llTempDollars(ilTemp)
            Else
                llProjectTrade(ilTemp) = llTempDollars(ilTemp)
            End If
        Next ilTemp
    Next ilTemp2                                'loop for cash vs trade
End Sub
'
'
'           Create prepass to generate Locked Avails report.  Go thru
'           ALF by vehicle and date to show vehicle,dates locked for spot
'           or avails and times locked.
'           4-5-06
Public Sub gCrLockedAvails()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llLoopDate                    ilVefIndex                                              *
'******************************************************************************************


Dim ilVehicle As Integer
Dim ilVefCode As Integer
Dim ilRet As Integer
Dim slNameCode As String
Dim slName As String
Dim slCode As String
Dim ilFound As Integer
Dim llStartDate As Long
Dim llEndDate As Long
Dim slDate As String
Dim ilDate(0 To 1) As Integer
Dim llDate As Long

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)

    hmGrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmVef)
        btrDestroy hmGrf
        btrDestroy hmVef
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmAlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAlf, "", sgDBPath & "Alf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAlf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmVef)
        btrDestroy hmAlf
        btrDestroy hmGrf
        btrDestroy hmVef
        Exit Sub
    End If
    imAlfRecLen = Len(tmAlf)

    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime

    slDate = RptSelCt!CSI_CalFrom.Text  'Date: 12/5/2019 added CSI calendar control for date entries --> edcSelCFrom.Text
    llStartDate = gDateValue(slDate)
    gPackDate slDate, ilDate(0), ilDate(1)
    slDate = RptSelCt!CSI_CalTo.Text    'Date: 12/5/2019 added CSI calendar control for date entries --> edcSelCFrom1.Text
    llEndDate = gDateValue(slDate)
    ilFound = False
    For ilVehicle = 0 To RptSelCt!lbcSelection(3).ListCount - 1 Step 1
        If (RptSelCt!lbcSelection(3).Selected(ilVehicle)) Then
            slNameCode = tgVehicle(ilVehicle).sKey
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            'retrieve ALF by vehicle and date key
            tmAlfSrchkey1.iVefCode = ilVefCode
            tmAlfSrchkey1.iDate(0) = ilDate(0)      'search starting at user requested start date
            tmAlfSrchkey1.iDate(1) = ilDate(1)
            ilRet = btrGetGreaterOrEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchkey1, INDEXKEY1, BTRV_LOCK_NONE)
            gUnpackDateLong tmAlf.iDate(0), tmAlf.iDate(1), llDate
            Do While ilRet <> BTRV_ERR_END_OF_FILE And tmAlf.iVefCode = ilVefCode And llDate >= llStartDate And llDate <= llEndDate

                'grfgentime = generation time (set above)
                'grfgendate = generation date (set above )
                'determine week # of this date to sort the  week for vehicle together
                tmGrf.iYear = (llDate - llStartDate) \ 7 + 1
                tmGrf.iVefCode = ilVefCode
                tmGrf.iDate(0) = tmAlf.iDate(0)     'date of lock
                tmGrf.iDate(1) = tmAlf.iDate(1)
                tmGrf.iMissedTime(0) = tmAlf.iStartTime(0)
                tmGrf.iMissedTime(1) = tmAlf.iStartTime(1)
                tmGrf.iTime(0) = tmAlf.iEndTime(0)
                tmGrf.iTime(1) = tmAlf.iEndTime(1)
                gUnpackTimeLong tmAlf.iStartTime(0), tmAlf.iStartTime(1), False, tmGrf.lLong
                tmGrf.sBktType = tmAlf.sLockType
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                ilRet = btrGetNext(hmAlf, tmAlf, imAlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                gUnpackDateLong tmAlf.iDate(0), tmAlf.iDate(1), llDate
                Loop
        End If              'vehicle selected
    Next ilVehicle

    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAlf)

    btrDestroy hmGrf
    btrDestroy hmVef
    btrDestroy hmAlf
    Exit Sub
End Sub
'
'
'       mSalesActNTR - gather NTR for the Sales Activty report
'       <input> ilListIndex = report option
'               ilSelect - list box index based on user option selected:  6 = vehicle
'               ilMajorSet - major vehicle group selected (if applicable)
'               ilMinorSet - minor vehicle group selected (not applicable this report)
'               slGrossOrNet - G (gross), N = net
'               llStdStartDates() array of std start dates for 13 months
'
'       2-1-07 Implement NTR activity in Weekly Sales Activity by Month
Public Sub mSalesActNTR(ilListIndex As Integer, ilSelect As Integer, ilMajorSet As Integer, ilMinorSet As Integer, slGrossOrNet As String, llStdStartDates() As Long, llEarliestEntry As Long, tlSBFTypes As SBFTypes, tlMnf() As MNF)
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSbf As Integer
    Dim llSbf As Long
    ReDim tlSbf(0 To 0) As SBF
    Dim ilFound As Integer
    Dim ilTemp As Integer
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slCode As String
    Dim illoop As Integer
    'Dim llProject(1 To 12) As Long
    ReDim llProject(0 To 12) As Long    'Index zero ignored
    'Dim llProjectCash(1 To 12) As Long
    ReDim llProjectCash(0 To 12) As Long    'Index zero ignored
    'Dim llProjectTrade(1 To 12) As Long
    ReDim llProjectTrade(0 To 12) As Long   'Index zero ignored
    Dim llNTRBillDate As Long
    Dim ilVefIndex As Integer
    Dim slName As String
    Dim ilSortCode As Integer
    Dim ilFoundAgain As Integer
    Dim llEnterDate As Long
    Dim slSbfStart As String
    Dim slSbfEnd As String
    Dim ilValidSbf As Integer
    Dim ilIsItHardCost As Integer
    Dim ilAgyCommPct As Integer
    Dim ilOKtoSeeVeh As Integer
    Dim ilMax As Integer
    Dim llSlfSplit(0 To 9) As Long
    Dim ilSlspCode(0 To 9) As Integer
    Dim llSlfSplitRev(0 To 9) As Long
    Dim ilLoopOnSlsp As Integer
    Dim llProcessPct As Long
    'Dim llCashShare(1 To 12) As Long
    ReDim llCashShare(0 To 12) As Long  'Index zero ignored
    'Dim llTradeShare(1 To 12) As Long
    ReDim llTradeShare(0 To 12) As Long 'Index zero ignored
    Dim ilSplitNTR As Integer 'Split NTR? 0=No, 1=Yes

    ilSplitNTR = 0 '09/28/2020 - TTP 9952
    If ilListIndex = CNT_CUMEACTIVITY Or ilListIndex = CNT_SALESACTIVITY_SS Then
        If RptSelCt!rbcSelC14(1).Value = True Then ilSplitNTR = 1
    End If
    
    gUnpackDate tgChfCT.iOHDDate(0), tgChfCT.iOHDDate(1), slCode        'get the date entered
    llEnterDate = gDateValue(slCode)
    'get the earliest and latest report date to see if date entered falls within
    slSbfStart = Format$(llStdStartDates(1), "m/d/yy")
    slSbfEnd = Format$(llStdStartDates(13) - 1, "m/d/yy")
    ilRet = gObtainSBF(RptSelCt, hmSbf, tgChfCT.lCode, slSbfStart, slSbfEnd, tlSBFTypes, tlSbf(), 0) '11-28-06 add last parm to indicate which key to use

    For llSbf = LBound(tlSbf) To UBound(tlSbf) - 1 Step 1
        tmSbf = tlSbf(llSbf)
        ilValidSbf = False
        ilIsItHardCost = False
        ilIsItHardCost = gIsItHardCost(tmSbf.iMnfItem, tlMnf())
        If tmSbf.iMnfItem > 0 Then          'ntr of some kind
            If ilIsItHardCost Then          'hard cost, Include?
                If imHardCost Then
                    ilValidSbf = True
                End If
            Else                            'non-hard cost
                If imNTR Then               'include non-hard cost?
                    ilValidSbf = True
                End If
            End If
        End If

        ilFound = True
        If ilListIndex = CNT_CUMEACTIVITY Then          '7-26-02
            If ilSelect = 6 And Not RptSelCt!ckcAll.Value = vbChecked Then
                ilFound = False
                For ilTemp = 0 To RptSelCt!lbcSelection(6).ListCount - 1 Step 1
                    If RptSelCt!lbcSelection(6).Selected(ilTemp) Then              'selected slsp
                        slNameCode = tgCSVNameCode(ilTemp).sKey    'RptSelCt!lbcCSVNameCode.List(ilTemp)         'pick up slsp code
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = tmSbf.iAirVefCode Then
                            ilFound = True
                            Exit For
                        End If
                    End If
                Next ilTemp
            End If
        Else                '7-25 CNT_SALESACTIVITY_SS check for market (if it exists) or vehicle if no mkt exists
            'determine which vehicle group is selected
            'tmGrf.iPerGenl(4) = 0
            tmGrf.iPerGenl(3) = 0
            If ilMajorSet > 0 Then
                'filter the vehicle group
                'gGetVehGrpSets tmSbf.iAirVefCode, ilMajorSet, ilMajorSet, tmGrf.iPerGenl(4), tmGrf.iPerGenl(4)   '6-13-02
                gGetVehGrpSets tmSbf.iAirVefCode, ilMajorSet, ilMajorSet, tmGrf.iPerGenl(3), tmGrf.iPerGenl(3)   '6-13-02
            End If
            If Not RptSelCt!CkcAllveh.Value = vbChecked And RptSelCt!lbcSelection(6).Visible = True Then        'all items in vehicle group selected?
                ilFound = False
                For illoop = 0 To RptSelCt!lbcSelection(6).ListCount - 1 Step 1
                    If RptSelCt!lbcSelection(6).Selected(illoop) Then
                        slNameCode = tgMnfCodeCT(illoop).sKey
                        ilRet = gParseItem(slNameCode, 1, "\", slName)
                        ilRet = gParseItem(slName, 3, "|", slName)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        'Determine which vehicle set to test
                        'If tmGrf.iPerGenl(4) = Val(slCode) Then
                        If tmGrf.iPerGenl(3) = Val(slCode) Then
                            ilFound = True
                            Exit For
                        End If
                    End If
                Next illoop
            End If
        End If
        
        ilOKtoSeeVeh = gUserAllowedVehicle(tmSbf.iAirVefCode)  '03-17-10 check to see if users allowed to see this vehicle
        If (ilFound) And (ilValidSbf) And (ilOKtoSeeVeh) Then                 'all vehicles of selective one found
            
            ilMax = mCumeHowManySlsp(ilListIndex, llSlfSplit(), ilSlspCode(), llSlfSplitRev(), tmSbf.iAirVefCode)
            
            If ilListIndex = CNT_CUMEACTIVITY Then          '7-26-02
                'determine which sort is used, then create the key field to be tested against
                If ilSelect = 1 Then            'agy
                    ilSortCode = tgChfCT.iAgfCode
                ElseIf ilSelect = 5 Then        'adv
                    'ilSortCode = tgChfCT.iAdfCode
                    ilSortCode = tmSbf.iAirVefCode            '3-31-05 this was used as the sort for advertiser option;
                                                            'and using adfcode was commented out.  change again to use
                                                            'the adfcode so that rounding occurs less frequently, rather
                                                            'than on each vehicle
                                                            '4-19-05 change this back to use the vehicle code.  Advt option
                                                            'does not show the vehicle names properly

                ElseIf ilSelect = 6 Then        'vehicle
                    ilSortCode = tmSbf.iAirVefCode
                Else
                    ilSortCode = tgChfCT.iMnfDemo(0)    'demo category
                End If
            Else                '7-25-02 CNT_SALESACTIVITY_SS
                ilSortCode = tmSbf.iAirVefCode     'report is generated by
            End If
            For ilTemp = 1 To 12 Step 1 'init projection $ each time
                llProject(ilTemp) = 0
            Next ilTemp

            For ilTemp = 1 To 12 Step 1 'init projection $ each time
                gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llNTRBillDate
                If llNTRBillDate >= llStdStartDates(ilTemp) And llNTRBillDate < llStdStartDates(ilTemp + 1) Then
                    llProject(ilTemp) = tmSbf.iNoItems * tmSbf.lGross
                    Exit For
                End If
            Next ilTemp

            'determine if this transaction should have commission
            ilAgyCommPct = 10000                     'Assume gross requested
            If slGrossOrNet = "N" Then          '
                If tgChfCT.iAgfCode > 0 And tmSbf.sAgyComm = "Y" Then       'agency exists and NTR should be commissionable, net- take out commission
                    ilAgyCommPct = ilAgyCommPct - tmAgf.iComm '85, prev forced to 15%, but now obtained from agency where its carried in 2 places
                End If
            End If
            mSplitCashTrade llProject(), llProjectCash(), llProjectTrade(), ilAgyCommPct, ilListIndex
            
            For ilLoopOnSlsp = 0 To ilMax - 1
                llProcessPct = llSlfSplitRev(ilLoopOnSlsp)

                If llProcessPct > 0 Then
                    'llproject(1-12) contains $ for this line
                    'vehicles are build into memory with its 12 $ buckets
                    
                    '4-5-11 Split the salesperson share
                    For ilTemp = 1 To 12
                        llCashShare(ilTemp) = 0
                        llTradeShare(ilTemp) = 0
                    Next ilTemp
                    
                    mCumeSlspShare llProcessPct, llProjectCash(), llProjectTrade(), llCashShare(), llTradeShare()
                    'llproject(1-12) contains $ for this NTR item
                    'vehicles are build into memory with its 12 $ buckets
                    ilFoundAgain = False
                    If ilSplitNTR = 1 Then
                        'if Were Splitting NTR, then NTR needs to be its own Row
                        For ilTemp = 0 To UBound(tmVefDollars) - 1 Step 1
                            'If tmVefDollars(ilTemp).ivefcode = tmClf.ivefcode Then
                            'If tmVefDollars(ilTemp).iVefCode = ilSortCode Then
                            If tmVefDollars(ilTemp).iSortCode = ilSortCode And tmVefDollars(ilTemp).iSlfCode = ilSlspCode(ilLoopOnSlsp) And tmVefDollars(ilTemp).iNTRInd = 1 Then     '4-19-05 use generalized sort field for comparisons
                                ilFoundAgain = True
                                ilVefIndex = ilTemp
                                Exit For                    '12-12-05
                            End If
                        Next ilTemp
                    Else
                        For ilTemp = 0 To UBound(tmVefDollars) - 1 Step 1
                            'If tmVefDollars(ilTemp).ivefcode = tmClf.ivefcode Then
                            'If tmVefDollars(ilTemp).iVefCode = ilSortCode Then
                            If tmVefDollars(ilTemp).iSortCode = ilSortCode And tmVefDollars(ilTemp).iSlfCode = ilSlspCode(ilLoopOnSlsp) Then     '4-19-05 use generalized sort field for comparisons
                                ilFoundAgain = True
                                ilVefIndex = ilTemp
                                Exit For                    '12-12-05
                            End If
                        Next ilTemp
                    End If
                    If Not (ilFoundAgain) Then
                        'tmVefDollars(UBound(tmVefDollars)).ivefcode = tmClf.ivefcode
                        'tmVefDollars(UBound(tmVefDollars)).iVefCode = ilSortCode           '4-19-05
                        tmVefDollars(UBound(tmVefDollars)).iSortCode = ilSortCode           '4-19-05 use generalized sort field for comparisons
                        tmVefDollars(UBound(tmVefDollars)).iVefCode = tmSbf.iAirVefCode       '2-23-07 save the vehicle for the output
                        tmVefDollars(UBound(tmVefDollars)).iSlfCode = ilSlspCode(ilLoopOnSlsp)
                        
                        'ilVefIndex = ilUpperVef
                        ilVefIndex = UBound(tmVefDollars)
                        ReDim Preserve tmVefDollars(0 To UBound(tmVefDollars) + 1) As ADJUSTLIST
                        
                    End If
                    
                    'now add or subtract depending if this contract is a new or mod
                    If llEnterDate < llEarliestEntry Then      'previous version
                        For ilTemp = 1 To 12 Step 1
                            tmVefDollars(ilVefIndex).lProject(ilTemp) = tmVefDollars(ilVefIndex).lProject(ilTemp) - llCashShare(ilTemp)
                            tmVefDollars(ilVefIndex).lProjectTrade(ilTemp) = tmVefDollars(ilVefIndex).lProjectTrade(ilTemp) - llTradeShare(ilTemp)
                        Next ilTemp
                    Else                                        'current weeks data
                        For ilTemp = 1 To 12 Step 1
                            tmVefDollars(ilVefIndex).lProject(ilTemp) = tmVefDollars(ilVefIndex).lProject(ilTemp) + llCashShare(ilTemp)
                            tmVefDollars(ilVefIndex).lProjectTrade(ilTemp) = tmVefDollars(ilVefIndex).lProjectTrade(ilTemp) + llTradeShare(ilTemp)
                        Next ilTemp
                    End If
                    'Indicate this is a NTR Row. EitherWay: If this is a NTR only row we'll indicate this is a NTR record.  If this is a combined (Air Time + NTR Row), we'll indicate this record includes NTR
                    tmVefDollars(ilVefIndex).iNTRInd = 1
                End If
            Next ilLoopOnSlsp
        End If                                  '3-8-06 if tmclf.stype = "S" or tmclf.stype = "H"
    Next llSbf                                  'for ilSbf = lbound(tlSbf) - ubound(tlSbf)
End Sub
'       2020-10-30 - TTP # 9955 - add AirTime, NTR, Hardcost Include options to report (Daily Sales Activity by Contract Report, Weekly Sales Activity by Qtr)
'       mSalesActNTR - gather NTR for the Daily Sales Activty report by Contract and Weelky Sales Activity by Quarter Report
'       <input> ilListIndex = report option
Public Function mGetSalesActNTR(ilListIndex As Integer, sStartDate As String, sEndDate As String, tlSBFTypes As SBFTypes, tlSbf() As SBF, tlMnf() As MNF, slGrossOrNet As String, ilAgyCommPct As Integer) As Long
    Dim ilRet As Integer
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim ilValidSbf As Integer
    Dim ilIsItHardCost As Integer
    Dim llNTR As Long
    Dim llNoPenny As Long
    Dim llTemp As Long
    ilRet = gObtainSBF(RptSelCt, hmSbf, tgChfCT.lCode, sStartDate, sEndDate, tlSBFTypes, tlSbf(), 0)
    For llSbf = LBound(tlSbf) To UBound(tlSbf) - 1 Step 1
        tmSbf = tlSbf(llSbf)
        ilValidSbf = False
        ilIsItHardCost = False
        ilIsItHardCost = gIsItHardCost(tmSbf.iMnfItem, tlMnf())
        If tmSbf.iMnfItem > 0 Then          'ntr of some kind
            If ilIsItHardCost Then          'hard cost, Include?
                If imHardCost Then
                    ilValidSbf = True
                End If
            Else                            'non-hard cost
                If imNTR Then               'include non-hard cost?
                    ilValidSbf = True
                End If
            End If
        End If
        
        If ilValidSbf = True Then
            'determine if this transaction should have commission
            If tmSbf.sAgyComm = "Y" And slGrossOrNet = "N" Then
                llNoPenny = (tmSbf.iNoItems * tmSbf.lGross) / 100    'drop pennies
                llNTR = llNTR + llNoPenny * CDbl(ilAgyCommPct) / 100   'adjust for the agy comm carried in 2 places
            Else
                llNTR = llNTR + (tmSbf.iNoItems * tmSbf.lGross)
            End If
             
        End If
    Next llSbf                                  'for ilSbf = lbound(tlSbf) - ubound(tlSbf)
    mGetSalesActNTR = llNTR
End Function
'
'
'           Create a Paperwork summary for tax information only.
'           Select contracts between a given date span, and show the
'           applicable taxes by vehicle.  Include NTR items which
'           could have different taxes.  Show Vehicle, agy, advt,
'           tax rates 1/2, tax descriptions, NTR type and bill dates
'           Bill dates will be NTR billing date or schedule line start/end dates.
'
Public Sub gCrPaperWkTax()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilTemp                        ilFound                       tlCntTypes                *
'*                                                                                        *
'******************************************************************************************

Dim illoop As Integer
Dim ilRet As Integer
Dim llStartDate As Long
Dim llEndDate As Long
Dim slStartDate As String                       'llLYGetFrom or llTYGetFrom converted to string
Dim slEndDate As String                         'llLYGetTo or LLTYGetTo converted to string
Dim slNameCode As String
Dim slCode As String
Dim slCntrTypes As String                       'valid contract types to access
Dim slCntrStatus As String                      'valid status (holds, orders, working, etc) to access
Dim ilHOState As Integer                        'include unsch holds/orders, sch holds/orders
Dim llContrCode As Long                         'contr code from gObtainCntrforDate
Dim ilCurrentRecd As Integer                    'index of contract being processed from tlChfAdvtExt
Dim ilClf As Integer                            'index to line from tgClfCt
Dim ilFoundVeh As Integer
Dim ilTrfAgyAdvt As Integer
Dim llTax1Pct As Long
Dim llTax2Pct As Long
Dim slGrossNet As String
Dim tlSbf() As SBF
Dim tlSBFType As SBFTypes
Dim llSbf As Long
Dim llLineStart As Long
Dim llLineEnd As Long
Dim llSingleCntr As Long
Dim llTemp As Long


    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
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
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSbf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imSbfRecLen = Len(tmSbf)


    ilRet = gObtainTrf()

    tlSBFType.iNTR = True
    tlSBFType.iInstallment = False
    tlSBFType.iImport = False

    slNameCode = RptSelCt!edcTopHowMany.Text
    llSingleCntr = Val(slNameCode)      'selective contract #

    slStartDate = RptSelCt!CSI_CalFrom.Text         'Date: 12/20/2019 added CSI calendar control for date entry --> edcSelCFrom.Text
    llStartDate = gDateValue(slStartDate)
    slStartDate = Format(llStartDate, "m/d/yy")   'make sure string start date has a year appended in case not entered with input

    slEndDate = RptSelCt!CSI_CalTo.Text             'Date: 12/20/2019 added CSI calendar control for date entry --> edcSelCFrom1.Text
    llEndDate = gDateValue(slEndDate)
    slEndDate = Format(llEndDate, "m/d/yy")    'make sure string end date has a year appended in case not entered with input

    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime

    slCntrTypes = gBuildCntTypes()      'Setup valid types of contracts to obtain based on user

    slCntrStatus = ""
    slCntrStatus = "HOGN"             'sch/unsch holds & uns holds
    ilHOState = 2                      'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)
    sgCntrForDateStamp = ""            'init the time stamp to read in contracts upon re-entry
    'Build array of possible contracts that fall into last year or this years quarter and build into array tlChfAdvtExt
    If llSingleCntr > 0 Then
        ReDim tlChfAdvtExt(0 To 1) As CHFADVTEXT
        tmChfSrchKey1.lCntrNo = llSingleCntr
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If ilRet = BTRV_ERR_NONE Then
            tlChfAdvtExt(0).lCode = tmChf.lCode
        End If
    Else
        ilRet = gObtainCntrForDate(RptSelCt, slStartDate, slEndDate, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
    End If

    'grf fields
    'grfGenDate - generation date
    'grfGenTime - generation time
    'grfchfCode - contract auto code
    'grfVefCode = vehicle code from line or NTR
    'grfSofCode = TRF auto code
    'grfCode2 - NTR Item type
    'grfStartDate - start date of line or NTR item
    'grfDate - end date of line or NTR item (same as startdate)
    'grfGenDesc - Vehicle code (5 char) or SBF Date (5 char)
    '
    'Major to minor sorting in crystal report:
    'Vehicle sort code, vehicle name, payee (direct followed by agency name),
    '   Advertiser name, contract, air time then NTR, NTR in date order
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
        llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT())   'get the latest version of this contract
            'Loop thru all lines and project their $ from the flights

            For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                tmClf = tgClfCT(ilClf).ClfRec
                If tmClf.sType = "S" Or tmClf.sType = "H" Then     'get standard and hidden lines only
                    'span of line end dates must be within the requested period
                    'first decide if its Cancel Before Start
                    gUnpackDate tmClf.iStartDate(0), tmClf.iStartDate(1), slNameCode
                    llLineStart = gDateValue(slNameCode)
                    gUnpackDate tmClf.iEndDate(0), tmClf.iEndDate(1), slNameCode
                    llLineEnd = gDateValue(slNameCode)

                    'gUnpackDate tmcff.iStartDate(0), tmcff.iStartDate(1), slStr
                    'backup start date to Monday
                    illoop = gWeekDayLong(llLineStart)
                    Do While illoop <> 0
                        llLineStart = llLineStart - 1
                        illoop = gWeekDayLong(llLineStart)
                    Loop
                    'the flight dates must be within the start and end of the requested period,
                    'not be a CAncel before start flight
                    If (llLineStart < llEndDate And llLineEnd >= llStartDate) And (llLineEnd >= llLineStart) Then

                        ilFoundVeh = False
                        'the vehicles in this list are only those with tax defined
                        For illoop = 0 To RptSelCt!lbcSelection(11).ListCount - 1 Step 1
                            If RptSelCt!lbcSelection(11).Selected(illoop) Then
                                slNameCode = tgSellNameCode(illoop).sKey    'RptSelCt!lbcCSVNameCode.List(ilLoop)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                                If tmClf.iVefCode = Val(slCode) Then
                                    ilFoundVeh = True
                                    Exit For
                                End If
                            End If
                        Next illoop

                        If ilFoundVeh Then
                            'need the tax auto code to pull off description
                            ilTrfAgyAdvt = gGetAirTimeTrfCode(tgChfCT.iAdfCode, tgChfCT.iAgfCode, tmClf.iVefCode)
                            If ilTrfAgyAdvt <= 0 Then
                                tmGrf.iSofCode = 0
                            Else
                                tmGrf.iSofCode = ilTrfAgyAdvt            'auto code
                            End If
                            gGetAirTimeTaxRates tgChfCT.iAdfCode, tgChfCT.iAgfCode, tmClf.iVefCode, llTax1Pct, llTax2Pct, slGrossNet

                            tmGrf.sGenDesc = Trim$(str$(tmClf.iVefCode))
                            Do While Len(Trim$(tmGrf.sGenDesc)) < 5
                                tmGrf.sGenDesc = "0" & tmGrf.sGenDesc
                            Loop
                            tmGrf.lChfCode = tgChfCT.lCode          'contr code
                            tmGrf.iVefCode = tmClf.iVefCode
                            tmGrf.iCode2 = 0                        'init NTR type for air time
                            tmGrf.iStartDate(0) = tmClf.iStartDate(0)     'sch line start date
                            tmGrf.iStartDate(1) = tmClf.iStartDate(1)
                            tmGrf.iDate(0) = tmClf.iEndDate(0)                        'schline end date
                            tmGrf.iDate(1) = tmClf.iEndDate(1)

                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                        End If
                    End If                  'lines pass request
                End If                      'tmclf = H or tmclf = S
            Next ilClf                      'process nextline
             'obtain all the NTR entries
            ReDim tlSbf(0 To 0) As SBF
            ilRet = gObtainSBF(RptSelCt, hmSbf, tgChfCT.lCode, slStartDate, slEndDate, tlSBFType, tlSbf(), 0)
            For llSbf = LBound(tlSbf) To UBound(tlSbf)
                ilFoundVeh = False
                'the vehicles in this list are only those with tax defined
                For illoop = 0 To RptSelCt!lbcSelection(11).ListCount - 1 Step 1
                    If RptSelCt!lbcSelection(11).Selected(illoop) Then
                        slNameCode = tgSellNameCode(illoop).sKey    'RptSelCt!lbcCSVNameCode.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                        If tlSbf(llSbf).iBillVefCode = Val(slCode) Then
                            ilFoundVeh = True
                            Exit For
                        End If
                    End If
                Next illoop
                If ilFoundVeh Then
                    'need the tax auto code to pull off description
                    ilTrfAgyAdvt = gGetAirTimeTrfCode(tgChfCT.iAdfCode, tgChfCT.iAgfCode, tlSbf(llSbf).iBillVefCode)
                    If ilTrfAgyAdvt < 0 Then
                        tmGrf.iSofCode = 0
                    Else
                        tmGrf.iSofCode = ilTrfAgyAdvt            'auto code
                    End If
                    gGetAirTimeTaxRates tgChfCT.iAdfCode, tgChfCT.iAgfCode, tlSbf(llSbf).iBillVefCode, llTax1Pct, llTax2Pct, slGrossNet
                    tmGrf.lChfCode = tgChfCT.lCode          'contr code
                    tmGrf.iVefCode = tlSbf(llSbf).iBillVefCode      'billing vehicle
                    tmGrf.iCode2 = tlSbf(llSbf).iMnfItem            'NTR type description
                    tmGrf.iStartDate(0) = tlSbf(llSbf).iDate(0)     'bill date of NTR
                    tmGrf.iStartDate(1) = tlSbf(llSbf).iDate(1)
                    tmGrf.iDate(0) = tlSbf(llSbf).iDate(0)          'make end date same as start
                    tmGrf.iDate(1) = tlSbf(llSbf).iDate(1)
                    gUnpackDateLong tlSbf(llSbf).iDate(0), tlSbf(llSbf).iDate(1), llTemp
                    tmGrf.sGenDesc = Trim$(str$(llTemp))
                    Do While Len(Trim(tmGrf.sGenDesc)) < 5
                        tmGrf.sGenDesc = "0" & tmGrf.sGenDesc
                    Loop
                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                End If
            Next llSbf
    Next ilCurrentRecd
    Erase tlChfAdvtExt
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    btrDestroy hmSbf
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmCHF
End Sub
'
'           Billed and Booked Comparisons - gather budgets by vehicle
'           <input>
'                    llstdStartDates()- array of monthly start/end dates to gather (all year sent)
'                    ilStartMonth - starting month to process
Public Sub gBudgetsForBOBCompare(llStdStartDates() As Long, ilStartMonth As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStartYr                     slEndYr                       slTemp                    *
'*  ilMnfCode                     ilStartWk                     ilEndWk                   *
'*  ilFirstWk                     ilLastWk                      ilBudgetLoop              *
'*  ilSSCode                      tlBudgetsByVef                                          *
'******************************************************************************************

Dim ilYear As Integer
Dim ilRet As Integer
Dim ilCorpStd As Integer
Dim ilBvfCalType As Integer
Dim llStartDate As Long
Dim llEndDate As Long
Dim illoop As Integer
Dim slStr As String
Dim ilTemp As Integer
Dim ilFound As Integer
Dim ilMajorSet As Integer
Dim ilmnfMinorCode As Integer           'vehicle group assigned to vehicle that user selected (ie. mkt name, participant name, format name, etc)
Dim ilMnfMajorCode As Integer
Dim ilVehicle As Integer
Dim slNameCode As String
Dim slCode As String
Dim ilNumberMonths As Integer
Dim ilPeriod As Integer
ReDim ilVefList(0 To 0) As Integer
'Dim ilStartWks(1 To 13) As Integer      'array of start weeks for the std or corp reporting months
ReDim ilStartWks(0 To 13) As Integer      'array of start weeks for the std or corp reporting months. Index zero ignored
'Dim ilTempStartWks(1 To 13) As Integer
ReDim ilTempStartWks(0 To 13) As Integer    'Index zero ignored
'Dim ilEndWks(1 To 13) As Integer        'array of end weeks for the std or corp reporting months
ReDim ilEndWks(0 To 13) As Integer        'array of end weeks for the std or corp reporting months. Index zero ignored
'Dim ilTempEndWks(1 To 13) As Integer
ReDim ilTempEndWks(0 To 13) As Integer  'Index zero ignored
Dim ilWeeks As Integer
Dim lTotalBudget As Long
Dim hlGrf As Integer
Dim hlSof As Integer
Dim ilSofTemp As Integer            '4-14-09
Dim ilSSTemp As Integer
ReDim tlSSList(0 To 0) As Integer
Dim ilSSandVehicle As Integer
Dim ilFoundSS As Integer

    If igBSelectedIndex <= 0 Then                              'budgets not selected
        Exit Sub
    End If


    hmBvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmBvf, "", sgDBPath & "Bvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmBvf)
        btrDestroy hmBvf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imBvfRecLen = Len(tmBvf)

    hlGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlGrf)
        ilRet = btrClose(hmBvf)
        btrDestroy hlGrf
        btrDestroy hmBvf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
'
'    hlSof = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hlSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hlSof)
'        ilRet = btrClose(hlGrf)
'        ilRet = btrClose(hmBvf)
'        btrDestroy hlSof
'        btrDestroy hlGrf
'        btrDestroy hmBvf
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'
'    ilSofTemp = 0
'    ilSSTemp = 0
'
'    If RptSelCt!ckcSelC13(2).Value = vbChecked Then         'use sales source as major sort
'        ReDim tlSSList(0 To 0) As Integer
'
'        ilRet = btrGetFirst(hlSof, tmSof, Len(tmSof), INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'        Do While ilRet = BTRV_ERR_NONE
'            ReDim Preserve tlSofList(0 To ilSofTemp) As SOFLIST
'            tlSofList(ilSofTemp).iSofCode = tmSof.iCode
'            tlSofList(ilSofTemp).iMnfSSCode = tmSof.iMnfSSCode
'            ilFound = False
'            For ilTemp = 0 To ilSSTemp
'                If tlSSList(ilTemp) = tmSof.iMnfSSCode Then
'                    ilFound = True
'                    Exit For
'                End If
'            Next ilTemp
'            If Not ilFound Then
'                tlSSList(ilSSTemp) = tmSof.iMnfSSCode
'                ilSSTemp = ilSSTemp + 1
'                ReDim Preserve tlSSList(0 To ilSSTemp) As Integer
'
'            End If
'
'            ilRet = btrGetNext(hlSof, tmSof, Len(tmSof), BTRV_LOCK_NONE, SETFORREADONLY)
'            ilSofTemp = ilSofTemp + 1
'        Loop
'    Else                        'do not use sales source in sort
'        ReDim tlSSList(0 To 1) As Integer           'fake out entry with sales source as zero
'        'zero out the sales source for this version of the report; budgets wont work since no matching linked fields
'        'with the sales source field (grfsofcode)
'        tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for retrieval/removal of records
'        tmGrf.iGenDate(1) = igNowDate(1)
'        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
'        tmGrf.lGenTime = lgNowTime
'
'        ilRet = btrGetFirst(hlGrf, tmGrf, Len(tmGrf), INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'        Do While ilRet = BTRV_ERR_NONE
'            If tmGrf.lGenTime = lgNowTime And tmGrf.iGenDate(0) = igNowDate(0) And tmGrf.iGenDate(1) = igNowDate(1) Then
'                tmGrf.iSofCode = 0
'                ilRet = btrUpdate(hlGrf, tmGrf, Len(tmGrf))
'            End If
'            ilRet = btrGetNext(hlGrf, tmGrf, Len(tmGrf), BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'    End If
'


    illoop = RptSelCt!cbcSet1.ListIndex
    ilMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())

    slStr = RptSelCt!edcSelCTo.Text
    ilNumberMonths = Val(slStr)

    ilYear = Val(RptSelCt!edcSelCFrom.Text)           'year requested
    If RptSelCt!rbcSelC9(0).Value Then         'corp
        ilRet = gGetCorpCalIndex(ilYear)
        ilCorpStd = 2
        ilBvfCalType = 5               'get week inx based on fiscal dates
    Else                                'std
        ilCorpStd = 1
        ilBvfCalType = 0             'both projections and budgets will be std
    End If

    'the array of dates are start dates
    llStartDate = llStdStartDates(1)
    llEndDate = llStdStartDates(ilNumberMonths + 1) - 1

    ilPeriod = 2                        'assume months (vs qtrs)
    'get the entire years start/end weeks for 12 months
    gGetStartEndWeeksCt ilCorpStd, ilPeriod, igBYear, ilTempStartWks(), ilTempEndWks()       '11-30-16
    'setup start/end weeks for only the months requested
    ilTemp = 1
    For illoop = ilStartMonth To 12
        ilStartWks(ilTemp) = ilTempStartWks(illoop)
        ilEndWks(ilTemp) = ilTempEndWks(illoop)
        ilTemp = ilTemp + 1
    Next illoop

     For ilVehicle = 0 To RptSelCt!lbcSelection(6).ListCount - 1 Step 1
        slNameCode = tgCSVNameCode(ilVehicle).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        'include if selection is by vehicle and it is selected, or if not by vehicle include them all
        If (RptSelCt!lbcSelection(6).Selected(ilVehicle)) Then       'only build those vehicles selected
            ilVefList(UBound(ilVefList)) = Val(slCode)
            ReDim Preserve ilVefList(0 To UBound(ilVefList) + 1) As Integer
        End If
    Next ilVehicle

    'gather all budget records by vehicle for the requested year
    If Not mReadBvfRec(hmBvf, igBSelectedIndex, igBYear, tmBvfVeh()) Then        '11-30-16 get the budget year
        Exit Sub
    End If

    'use startwk & endwk to gather budgets
    'gObtainWkNo ilBvfCalType, slStartYr, ilStartWk, ilFirstWk          'determine the first week inx to accum (current year)
    'gObtainWkNo ilBvfCalType, slEndYr, ilEndWk, ilLastWk               'determine the last week inx to accum  (current year)
    ilFound = False

    For illoop = LBound(tmBvfVeh) To UBound(tmBvfVeh) - 1 Step 1

        If tmBvfVeh(illoop).sSplit = "D" Then                   'only direct budgets, no split

            For ilVehicle = 0 To UBound(ilVefList) - 1 Step 1
                If ilVefList(ilVehicle) = tmBvfVeh(illoop).iVefCode Then    'only build those vehicles selected
                    'Create budgets for the same vehicle for each sales source
'                    tmGrf.iSofCode = 0                  'init sales source field
'                    For ilSSTemp = 0 To UBound(tlSSList) - 1
'                        'for each sales source that is defined, see if there was data with the vehicle
'                        ilFoundSS = False
'                        For ilSSandVehicle = 0 To UBound(tgBBCompare) - 1
'                            If ilVefList(ilVehicle) = tgBBCompare(ilSSandVehicle).iVefCode And tlSSList(ilSSTemp) = tgBBCompare(ilSSandVehicle).iSSCode Then
'                                ilFoundSS = True
'                                Exit For
'                            End If
'                        Next ilSSandVehicle
'
'                        If ilFoundSS Then
'                            tmGrf.iSofCode = tlSSList(ilSSTemp)
        '                    For ilTemp = LBound(tlSofList) To UBound(tlSofList)
        '                        If tlSofList(ilTemp).iSofCode = tmBvfVeh(ilLoop).iSofCode Then
        '                            'tmGrf.iSofCode = tlSofList(ilTemp).iMnfSSCode
        '                            tmGrf.iSofCode = tlSofList(ilTemp).iSofCode
        '                            Exit For
        '                        End If
        '                    Next ilTemp


                            tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for retrieval/removal of records
                            tmGrf.iGenDate(1) = igNowDate(1)
                            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                            tmGrf.lGenTime = lgNowTime
                            tmGrf.sDateType = "C"                'all budgets are cash (for sort)
                            'tmGrf.iPerGenl(13) = 1                  'flag to indicate data from budgets (vs contracts/recv.)
                            'tmGrf.iPerGenl(14) = igBSelectedIndex          'selected budget
                            tmGrf.iPerGenl(12) = 1                  'flag to indicate data from budgets (vs contracts/recv.)
                            tmGrf.iPerGenl(13) = igBSelectedIndex          'selected budget
                            tmGrf.iVefCode = tmBvfVeh(illoop).iVefCode
                            'tmGrf.iPerGenl(8) = igPeriods           'no periods to print
                            tmGrf.iPerGenl(7) = igPeriods           'no periods to print
                            gGetVehGrpSets tmBvfVeh(illoop).iVefCode, 0, ilMajorSet, ilmnfMinorCode, ilMnfMajorCode
                            'tmGrf.iPerGenl(4) = ilMnfMajorCode      'vehicle group sort
                            tmGrf.iPerGenl(3) = ilMnfMajorCode      'vehicle group sort
                            lTotalBudget = 0
                            For ilTemp = 1 To 17            'initalize buckets to accumulate weeks within months requested
                                tmGrf.lDollars(ilTemp - 1) = 0
                            Next ilTemp

                            For ilTemp = 1 To ilNumberMonths
                                For ilWeeks = ilStartWks(ilTemp) To ilEndWks(ilTemp)
                                    tmGrf.lDollars(ilTemp - 1) = tmGrf.lDollars(ilTemp - 1) + tmBvfVeh(illoop).lGross(ilWeeks)
                                    lTotalBudget = lTotalBudget + tmBvfVeh(illoop).lGross(ilWeeks)
                                Next ilWeeks
                            Next ilTemp
                            If lTotalBudget <> 0 Then
                                ilRet = btrInsert(hlGrf, tmGrf, Len(tmGrf), INDEXKEY0)
                            End If
                        'Exit For
'                        End If
'                    Next ilSSTemp           'create budget for same vehicle, next sales source
                End If
                'Exit For
            Next ilVehicle
        End If
    Next illoop


    Erase ilVefList, tmBvfVeh, tgBBCompare
    Erase ilStartWks, ilEndWks

    ilRet = btrClose(hmBvf)
    ilRet = btrClose(hlGrf)
'    ilRet = btrClose(hlSof)
    btrDestroy hmBvf
    btrDestroy hlGrf
'    btrDestroy hlSof


End Sub

Public Sub mGetBudgetWks(ilCalType As Integer, ilPeriod As Integer, ilYear As Integer, ilStartWk() As Integer, ilEndWk() As Integer, llStdStartDates() As Long, ilNumberMonths As Integer) 'VBC NR
Dim slDate As String 'VBC NR
Dim slStart As String 'VBC NR
Dim slEnd As String 'VBC NR
Dim slPrevStart As String 'VBC NR
Dim illoop As Integer 'VBC NR
Dim ilWkNo As Integer 'VBC NR
Dim ilRet As Integer 'VBC NR
Dim ilAdjust As Integer 'VBC NR
Dim ilCorpWeeks As Integer 'VBC NR
    If ilCalType = 1 Then                    'std 'VBC NR
        'determine # weeks in each period (standard  month)
        slDate = "1/15/" & Trim$(str$(ilYear)) 'VBC NR
        slStart = gObtainStartStd(slDate) 'VBC NR
        slPrevStart = slStart 'VBC NR
        If ilPeriod = 2 Then          'std month 'VBC NR
            For illoop = 1 To 13 Step 1 'VBC NR
                slEnd = gObtainEndStd(slStart) 'VBC NR
                If illoop = 1 Then 'VBC NR
                    ilStartWk(1) = 1 'VBC NR
                End If 'VBC NR
                ilEndWk(illoop) = (gDateValue(slEnd) - gDateValue(slPrevStart) + 1) \ 7 'VBC NR
                slStart = gIncOneDay(slEnd) 'VBC NR
                If illoop < 13 Then 'VBC NR
                    ilStartWk(illoop + 1) = ilEndWk(illoop) + 1 'VBC NR
                End If 'VBC NR
            Next illoop 'VBC NR
            slStart = gIncOneDay(slEnd) 'VBC NR
        ElseIf ilPeriod = 3 Then          'std quarter 'VBC NR
                For illoop = 1 To 4 Step 1 'VBC NR
                    For ilWkNo = 1 To 3 Step 1 'VBC NR
                        slEnd = gObtainEndStd(slStart) 'VBC NR
                        slStart = gIncOneDay(slEnd) 'VBC NR
                    Next ilWkNo 'VBC NR
                    If illoop = 1 Then 'VBC NR
                        ilStartWk(1) = 1 'VBC NR
                    End If 'VBC NR
                    ilEndWk(illoop) = (gDateValue(slEnd) - gDateValue(slPrevStart) + 1) \ 7 'VBC NR
                    slStart = gIncOneDay(slEnd) 'VBC NR
                    ilStartWk(illoop + 1) = ilEndWk(illoop) + 1 'VBC NR
                Next illoop 'VBC NR
        Else                                               'std week 'VBC NR
            slDate = RptSelCt!edcSelCTo.Text                 'user input 'VBC NR
            ilRet = ((gDateValue(slDate) - gDateValue(slStart)) \ 7 + 1)    'get week index 'VBC NR
            For illoop = 1 To 13 Step 1 'VBC NR
                If illoop = 1 Then 'VBC NR
                    ilStartWk(1) = 1 'VBC NR
                    If ilRet <> 1 Then 'VBC NR
                        ilStartWk(1) = ilRet 'VBC NR
                    End If 'VBC NR
                End If 'VBC NR
                slEnd = gObtainNextSunday(slDate)        'obtain end of week 'VBC NR
                ilEndWk(illoop) = (((gDateValue(slEnd) - gDateValue(slStart) + 1)) \ 7) 'VBC NR
                slDate = gIncOneDay(slEnd) 'VBC NR
                If illoop < 13 Then 'VBC NR
                    ilStartWk(illoop + 1) = ilEndWk(illoop) + 1 'VBC NR
                End If 'VBC NR
            Next illoop 'VBC NR
            slStart = gIncOneDay(slEnd) 'VBC NR
        End If 'VBC NR
    End If                                      'endif std 'VBC NR
    If ilCalType = 2 Then            'corporate calendar 'VBC NR
        'determine # weeks in each period corp period
        ilRet = gGetCorpCalIndex(ilYear) 'VBC NR
        'gUnpackDate tgMCof(ilRet).iStartDate(0, 1), tgMCof(ilRet).iStartDate(1, 1), slStart         'convert last bdcst billing date to string 'VBC NR
        'gUnpackDate tgMCof(ilRet).iEndDate(0, 12), tgMCof(ilRet).iEndDate(1, 12), slEnd 'VBC NR
        gUnpackDate tgMCof(ilRet).iStartDate(0, 0), tgMCof(ilRet).iStartDate(1, 0), slStart         'convert last bdcst billing date to string 'VBC NR
        gUnpackDate tgMCof(ilRet).iEndDate(0, 11), tgMCof(ilRet).iEndDate(1, 11), slEnd 'VBC NR

        'slDate = "1/15/" & Trim$(Str$(ilYear))
        'slStart = gObtainStartCorp(slDate, True)

        slEnd = gObtainEndCorp(slStart, True) 'VBC NR
        ilStartWk(1) = 1 'VBC NR
        If ilPeriod = 2 Then      'corp month  (vs qtr) 'VBC NR
            For illoop = 1 To 12 Step 1 'VBC NR
                If illoop = 1 Then 'VBC NR
                    ilEndWk(1) = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7 'VBC NR
                    ilAdjust = gWeekDayStr(slStart) 'VBC NR
                    If ilAdjust <> 0 Then 'VBC NR
                        ilEndWk(1) = ilEndWk(1) + 1   'adjust for week of 1/1 thru sunday, plus the 'VBC NR
                                                                'remainder from divide
                    End If 'VBC NR
                Else 'VBC NR
                    slStart = gIncOneDay(slEnd) 'VBC NR
                    slEnd = gObtainEndCorp(slStart, True) 'VBC NR
                    ilStartWk(illoop) = ilEndWk(illoop - 1) + 1 'VBC NR
                    ilEndWk(illoop) = (ilStartWk(illoop) + ((gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7)) - 1 'VBC NR
                End If 'VBC NR
            Next illoop 'VBC NR
            slStart = gIncOneDay(slEnd) 'VBC NR
        ElseIf ilPeriod = 3 Then              'corp qtr 'VBC NR
            For illoop = 1 To 4 Step 1 'VBC NR
                For ilWkNo = 1 To 3 Step 1 'VBC NR
                    If illoop = 1 And ilWkNo = 1 Then 'VBC NR
                        ilCorpWeeks = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7 'VBC NR
                        ilAdjust = gWeekDayStr(slStart) 'VBC NR
                        If ilAdjust <> 0 Then 'VBC NR
                            ilCorpWeeks = ilCorpWeeks + 1   'adjust for week of 1/1 thru sunday, plus the 'VBC NR
                                                                'remainder from divide
                        End If 'VBC NR
                    Else 'VBC NR
                        slStart = gIncOneDay(slEnd) 'VBC NR
                        slEnd = gObtainEndCorp(slStart, True) 'VBC NR
                        ilCorpWeeks = ilCorpWeeks + ((gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7) 'VBC NR
                    End If 'VBC NR
                Next ilWkNo 'VBC NR
                ilEndWk(illoop) = ilCorpWeeks 'VBC NR
                ilStartWk(illoop + 1) = ilEndWk(illoop) + 1 'VBC NR
                slStart = gIncOneDay(slEnd) 'VBC NR
            Next illoop 'VBC NR
        Else                                'corp week 'VBC NR
            slDate = RptSelCt!edcSelCTo.Text          'user input date 'VBC NR
            ilRet = ((gDateValue(slDate) - gDateValue(slStart)) \ 7 + 1)    'get week index 'VBC NR
            For illoop = 1 To 13 Step 1 'VBC NR
                    If illoop = 1 And ilRet <> 1 Then       'preset to "1" earlier 'VBC NR
                        ilStartWk(1) = ilRet 'VBC NR
                    End If 'VBC NR
                    ilEndWk(illoop) = ilStartWk(illoop) 'VBC NR
                    If illoop < 13 Then 'VBC NR
                        ilStartWk(illoop + 1) = ilEndWk(illoop) + 1 'VBC NR
                    End If 'VBC NR
                    slEnd = gObtainNextSunday(slDate) 'VBC NR
                    slDate = gIncOneDay(slEnd) 'VBC NR
            Next illoop 'VBC NR
            slDate = gIncOneDay(slEnd) 'VBC NR
        End If 'VBC NR
    End If 'VBC NR
End Sub 'VBC NR
'
'           Sales Commmission - bonus version that rewards commissions by
'           new and renewals that exceed last year
'           mBuildCommByAdvt
'           <input> llLYStartDate- last year start date
'                   llCurrMonthStart - this month start date
'                   llCurrMonthEnd  - this month end date
'                   llTYStartDate - this year start date (corp or std)
'                   tlSlf() - array of slsp records
Private Sub mBuildcommByAdvt(llLYStartDate As Long, llCurrMonthstart As Long, llCurrMonthEnd As Long, llTYStartDate As Long, tlSlf() As SLF)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llEndInx                      ilCurrMonthStart                                        *
'******************************************************************************************

Dim llLoopOnRvf As Long
Dim ilPrevSlf As Integer
Dim slPrevAdvtName As String
Dim ilPrevCommPct As Integer
Dim ilAdfInx As Integer
Dim llUpperAdvt As Long
Dim llTranDate As Long
Dim slAdvtName As String
Dim slRvfAdvtName As String
Dim illoop As Integer
Dim llStartInx As Long
Dim llLoopOnAdvt As Long
Dim ilRet As Integer
Dim slPrevType As String * 1
Dim slCurrentType As String * 1



        ReDim tmAdvtComm(0 To 0) As RVFCOMM
        llUpperAdvt = UBound(tmAdvtComm)
        tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
        tmGrf.iGenDate(1) = igNowDate(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrf.lGenTime = lgNowTime

        'gPackDateLong llLYStartDate, tmGrf.iDate(0), tmGrf.iDate(1)   'last year start date for crystal

        gPackDateLong llCurrMonthstart, tmGrf.iDate(0), tmGrf.iDate(1)   'last year start date for crystal

        ilPrevSlf = -1
        slPrevAdvtName = ""
        ilPrevCommPct = -1
        slPrevType = ""
        'Pass 1:  look only for current months billing.  Discard anyting in the past.
        'create entry for each unique key by slsp code/advertiser name (not code), and comm pct and remnant/vs non-remnant
        For llLoopOnRvf = LBound(tmRVFComm) To UBound(tmRVFComm) - 1
            gUnpackDateLong tmRVFComm(llLoopOnRvf).iTranDate(0), tmRVFComm(llLoopOnRvf).iTranDate(1), llTranDate
            If llTranDate >= llCurrMonthstart Then

                ilAdfInx = gBinarySearchAdf(tmRVFComm(llLoopOnRvf).iAdfCode)
                If ilAdfInx = -1 Then
                    slAdvtName = "Missing Advt"
                Else
                    slAdvtName = tgCommAdf(ilAdfInx).sName
                End If
                If ilPrevSlf = -1 Then          'first time
                    ilPrevSlf = tmRVFComm(llLoopOnRvf).iSlfCode
                    ilPrevCommPct = tmRVFComm(llLoopOnRvf).iUseThisCommPct
                    slPrevAdvtName = Trim$(slAdvtName)
                    slPrevType = "S"                                'assume standard cnt vs remnant, need to keep apart for overage on sales goals
                    If tmRVFComm(llLoopOnRvf).sType = "T" Then      'remnant
                        slPrevType = "R"
                    End If

                    tmAdvtComm(llUpperAdvt).iSlfCode = tmRVFComm(llLoopOnRvf).iSlfCode
                    tmAdvtComm(llUpperAdvt).iAdfCode = tmRVFComm(llLoopOnRvf).iAdfCode
                    tmAdvtComm(llUpperAdvt).iUseThisCommPct = tmRVFComm(llLoopOnRvf).iUseThisCommPct
                    tmAdvtComm(llUpperAdvt).iSofCode = tmRVFComm(llLoopOnRvf).iSofCode
                    tmAdvtComm(llUpperAdvt).sType = tmRVFComm(llLoopOnRvf).sType
                End If

                'accumulate slsp This month year to date sales)
                For illoop = LBound(tlSlf) To UBound(tlSlf) - 1
                    If tlSlf(illoop).iCode = tmRVFComm(llLoopOnRvf).iSlfCode Then
                        If tmRVFComm(llLoopOnRvf).sType = "T" Then      'remnant
                            tmSalesGoalInfo(illoop).dYTDNNRemnant = tmSalesGoalInfo(illoop).dYTDNNRemnant + (tmRVFComm(llLoopOnRvf).lNet - (tmRVFComm(llLoopOnRvf).lMerch + tmRVFComm(llLoopOnRvf).lPromo + tmRVFComm(llLoopOnRvf).lAcquisition))
                        Else                'non remnant
                            tmSalesGoalInfo(illoop).dYTDNN = tmSalesGoalInfo(illoop).dYTDNN + (tmRVFComm(llLoopOnRvf).lNet - (tmRVFComm(llLoopOnRvf).lMerch + tmRVFComm(llLoopOnRvf).lPromo + tmRVFComm(llLoopOnRvf).lAcquisition))
                        End If

                        Exit For
                    End If
                Next illoop
                slCurrentType = "S"         'assume standard contract
                If tmRVFComm(llLoopOnRvf).sType = "T" Then  'remnant, need to keep separate entries
                    slCurrentType = "R"
                End If

                'see if matching key, then accum buckets are create new entry
                'match on slsp code, comm%, advertiser name, and contract type (only whether standard or remnant)
                If ilPrevSlf <> tmRVFComm(llLoopOnRvf).iSlfCode Or ilPrevCommPct <> tmRVFComm(llLoopOnRvf).iUseThisCommPct Or slPrevAdvtName <> Trim$(slAdvtName) Or slPrevType <> slCurrentType Then

                    llUpperAdvt = llUpperAdvt + 1
                    ReDim Preserve tmAdvtComm(0 To llUpperAdvt) As RVFCOMM
                    ilPrevSlf = tmRVFComm(llLoopOnRvf).iSlfCode
                    ilPrevCommPct = tmRVFComm(llLoopOnRvf).iUseThisCommPct
                    slPrevAdvtName = Trim$(slAdvtName)  'Trim$(tmAdf.sName)
                    slPrevType = "S"                'assume standard cnt type
                    If tmRVFComm(llLoopOnRvf).sType = "T" Then          'contract type (remnant, etc)
                        slPrevType = "R"                    'remnant
                    End If
                    tmAdvtComm(llUpperAdvt).iSlfCode = tmRVFComm(llLoopOnRvf).iSlfCode
                    tmAdvtComm(llUpperAdvt).iAdfCode = tmRVFComm(llLoopOnRvf).iAdfCode
                    tmAdvtComm(llUpperAdvt).iUseThisCommPct = tmRVFComm(llLoopOnRvf).iUseThisCommPct
                    tmAdvtComm(llUpperAdvt).iSofCode = tmRVFComm(llLoopOnRvf).iSofCode
                    tmAdvtComm(llUpperAdvt).sType = tmRVFComm(llLoopOnRvf).sType

                End If
                'accumulate all this months info as everything else has been discarded
                tmAdvtComm(llUpperAdvt).lGross = tmAdvtComm(llUpperAdvt).lGross + tmRVFComm(llLoopOnRvf).lGross
                tmAdvtComm(llUpperAdvt).lNet = tmAdvtComm(llUpperAdvt).lNet + tmRVFComm(llLoopOnRvf).lNet
                tmAdvtComm(llUpperAdvt).lMerch = tmAdvtComm(llUpperAdvt).lMerch + tmRVFComm(llLoopOnRvf).lMerch
                tmAdvtComm(llUpperAdvt).lPromo = tmAdvtComm(llUpperAdvt).lPromo + tmRVFComm(llLoopOnRvf).lPromo
                tmAdvtComm(llUpperAdvt).lAcquisition = tmAdvtComm(llUpperAdvt).lAcquisition + tmRVFComm(llLoopOnRvf).lAcquisition

                If tmAdvtComm(llUpperAdvt).lSlfSplit = -1 Then
                    'already set, different share % for the same advt on this slp
                Else            'are the slsp splits the same and its not zero?
                    If tmAdvtComm(llUpperAdvt).lSlfSplit <> tmRVFComm(llLoopOnRvf).lSlfSplit And tmAdvtComm(llUpperAdvt).lSlfSplit <> 0 Then
                        'different slsp splits from one contract to another, set as a mixed share %
                        tmAdvtComm(llUpperAdvt).lSlfSplit = -1
                    Else
                        tmAdvtComm(llUpperAdvt).lSlfSplit = tmRVFComm(llLoopOnRvf).lSlfSplit    'slsp revenue share
                    End If
                End If
                tmAdvtComm(llUpperAdvt).iRemUnderPct = tmRVFComm(llLoopOnRvf).iRemUnderPct    '3-15-00 slsp remnant % under

                tmAdvtComm(llUpperAdvt).iDiscrep = tmRVFComm(llLoopOnRvf).iDiscrep               'error flag: discrepancy in vehicle/slsp commission (sub-companies)
                tmAdvtComm(llUpperAdvt).iCurrentMonth = tmRVFComm(llLoopOnRvf).iCurrentMonth             'for sales goal info, show starting-ending month of previous $
                tmAdvtComm(llUpperAdvt).iStartFiscalMonth = tmRVFComm(llLoopOnRvf).iStartFiscalMonth

             End If              'llTranDate < llCurrStartDate
        Next llLoopOnRvf
        ReDim Preserve tmAdvtComm(LBound(tmAdvtComm) To UBound(tmAdvtComm) + 1)

        'Pass 2:  Look for only the past; discard the current month
        'Accumulate the last year advt totals and this year YTD advt totals for those advt billed in the current month.
        'if advt not found in tmAdvtComm array if transaction is for last year, discard last year
        'Accum current year to date total to show % of sales goal reached (all advertisers).
        'AccumPrevious YTD gross, net merch/promo and acquisition costs for all advertisers to show for sales goal info

        llStartInx = LBound(tmAdvtComm)
         For llLoopOnRvf = LBound(tmRVFComm) To UBound(tmRVFComm) - 1
            gUnpackDateLong tmRVFComm(llLoopOnRvf).iTranDate(0), tmRVFComm(llLoopOnRvf).iTranDate(1), llTranDate
            If llTranDate < llCurrMonthstart Then           'in the past, accum last year and this year to date
                'current year to date
                If llTranDate > llTYStartDate Then              'this entry is this year, prior month(s)
                    'accumulate slsp previous year to date sales (current year)
                    For illoop = LBound(tlSlf) To UBound(tlSlf) - 1
                        If tlSlf(illoop).iCode = tmRVFComm(llLoopOnRvf).iSlfCode Then
                            'tmSlfPreviousYTD(ilLoop) = tmSlfPreviousYTD(ilLoop) + (tmRVFComm(llLoopOnRvf).lNet - (tmRVFComm(llLoopOnRvf).lMerch + tmRVFComm(llLoopOnRvf).lPromo + tmRVFComm(llLoopOnRvf).lAcquisition))
                            If tmRVFComm(llLoopOnRvf).sType = "T" Then      'remnant
                                tmSalesGoalInfo(illoop).dPrevYTDNNREmnant = tmSalesGoalInfo(illoop).dPrevYTDNNREmnant + (tmRVFComm(llLoopOnRvf).lNet - (tmRVFComm(llLoopOnRvf).lMerch + tmRVFComm(llLoopOnRvf).lPromo + tmRVFComm(llLoopOnRvf).lAcquisition))
                                tmSalesGoalInfo(illoop).dYTDNNRemnant = tmSalesGoalInfo(illoop).dYTDNNRemnant + (tmRVFComm(llLoopOnRvf).lNet - (tmRVFComm(llLoopOnRvf).lMerch + tmRVFComm(llLoopOnRvf).lPromo + tmRVFComm(llLoopOnRvf).lAcquisition))
                            Else                'non remnant
                                tmSalesGoalInfo(illoop).dPrevYTDNN = tmSalesGoalInfo(illoop).dPrevYTDNN + (tmRVFComm(llLoopOnRvf).lNet - (tmRVFComm(llLoopOnRvf).lMerch + tmRVFComm(llLoopOnRvf).lPromo + tmRVFComm(llLoopOnRvf).lAcquisition))
                                tmSalesGoalInfo(illoop).dPrevYTDGross = tmSalesGoalInfo(illoop).dPrevYTDGross + tmRVFComm(llLoopOnRvf).lGross
                                tmSalesGoalInfo(illoop).dPrevYTDNet = tmSalesGoalInfo(illoop).dPrevYTDNet + tmRVFComm(llLoopOnRvf).lNet
                                tmSalesGoalInfo(illoop).dPrevYTDMerchPromo = tmSalesGoalInfo(illoop).dPrevYTDMerchPromo + (tmRVFComm(llLoopOnRvf).lMerch + tmRVFComm(llLoopOnRvf).lPromo)
                                tmSalesGoalInfo(illoop).dPrevYTDAcq = tmSalesGoalInfo(illoop).dPrevYTDAcq + tmRVFComm(llLoopOnRvf).lAcquisition
                                tmSalesGoalInfo(illoop).dYTDNN = tmSalesGoalInfo(illoop).dYTDNN + (tmRVFComm(llLoopOnRvf).lNet - (tmRVFComm(llLoopOnRvf).lMerch + tmRVFComm(llLoopOnRvf).lPromo + tmRVFComm(llLoopOnRvf).lAcquisition))

                            End If

                            Exit For
                        End If
                    Next illoop
                Else                    'last year
                End If

                'find a matching advt and slsp entry in the newly built advt comm array so that the
                'last year and year to date advt totals can be built
                'Need to do all the advertiser testing by name due to allowing duplicate advt names
                'but with different addresses
                ilAdfInx = gBinarySearchAdf(tmRVFComm(llLoopOnRvf).iAdfCode)
                If ilAdfInx = -1 Then
                    slRvfAdvtName = "Missing Advt"
                Else
                    slRvfAdvtName = Trim$(tgCommAdf(ilAdfInx).sName)
                End If

                'see if matching slsp and advertisers to accum last year data
                For llLoopOnAdvt = llStartInx To UBound(tmAdvtComm) - 1
                    'the advt array is in sls code order, if exceeded the code in matching, the advt doesn
                    'exist for the current month; discard the past entry whether last year or prev ytd for this year
'                    If tmRVFComm(llLoopOnRvf).iSlfCode > tmAdvtComm(llLoopOnAdvt).iSlfCode Then
'                        llStartInx = llLoopOnAdvt        'start next search from where it left off to help speed up process
'                        Exit For
'                    Else
                        If tmRVFComm(llLoopOnRvf).iSlfCode = tmAdvtComm(llLoopOnAdvt).iSlfCode Then
                            ilAdfInx = gBinarySearchAdf(tmAdvtComm(llLoopOnAdvt).iAdfCode)
                            If ilAdfInx = -1 Then
                                slAdvtName = "Missing Advt"
                            Else
                                slAdvtName = tgCommAdf(ilAdfInx).sName
                            End If
                            If Trim$(slRvfAdvtName) = Trim$(slAdvtName) Then
                                'is this transaction for last year or current year to date?

                                If llTranDate < llTYStartDate Then      'last year
                                    'accumulate all this months info as everything else has been discarded
                                    tmAdvtComm(llLoopOnAdvt).lLastYear = tmAdvtComm(llLoopOnAdvt).lLastYear + (tmRVFComm(llLoopOnRvf).lNet - (tmRVFComm(llLoopOnRvf).lMerch + tmRVFComm(llLoopOnRvf).lPromo + tmRVFComm(llLoopOnRvf).lAcquisition))
                                Else                                      'current year to date (not including current month)
                                    If llTranDate < llCurrMonthstart Then
                                        tmAdvtComm(llLoopOnAdvt).lThisYearPrevMonth = tmAdvtComm(llLoopOnAdvt).lThisYearPrevMonth + (tmRVFComm(llLoopOnRvf).lNet - (tmRVFComm(llLoopOnRvf).lMerch + tmRVFComm(llLoopOnRvf).lPromo + tmRVFComm(llLoopOnRvf).lAcquisition))
                                    Else
                                        ilRet = ilRet
                                    End If
                                End If
                                Exit For
                            End If
'                        End If
                    End If
                Next llLoopOnAdvt
            End If              'llTranDate < llCurrStartDate
        Next llLoopOnRvf

        'pass 3:  create prepass record in GRF.  Each record is a unique slsp/advt name/comm % rate
        'The previous ytd (all advertisers by slsp) will be stored in every record for slsp
        For llLoopOnAdvt = LBound(tmAdvtComm) To UBound(tmAdvtComm) - 1
            'tmGrf.lDollars(2) = tmAdvtComm(llLoopOnAdvt).lGross          'Gross $
            'tmGrf.lDollars(3) = tmAdvtComm(llLoopOnAdvt).lNet             'Net$
            'tmGrf.lDollars(4) = tmAdvtComm(llLoopOnAdvt).lMerch          'Merchandising $
            'tmGrf.lDollars(5) = tmAdvtComm(llLoopOnAdvt).lPromo          'Promotions $
            'tmGrf.lDollars(10) = tmAdvtComm(llLoopOnAdvt).lAcquisition
            'tmGrf.lDollars(7) = tmAdvtComm(llLoopOnAdvt).iUseThisCommPct       'slsp % non-remnant under goal
            'tmGrf.lDollars(8) = tmAdvtComm(llLoopOnAdvt).lSlfSplit    'slsp revenue share
            ''tmGrf.lDollars(9) = tmAdvtComm(llLoopOnAdvt).iRemUnderPct    '3-15-00 slsp remnant % under
            'tmGrf.lDollars(15) = tmAdvtComm(llLoopOnAdvt).lNet - tmAdvtComm(llLoopOnAdvt).lMerch - tmAdvtComm(llLoopOnAdvt).lPromo - tmAdvtComm(llLoopOnAdvt).lAcquisition   'current month net -net
            'tmGrf.lDollars(16) = tmAdvtComm(llLoopOnAdvt).lLastYear             'for advt in current month
            'tmGrf.lDollars(17) = tmAdvtComm(llLoopOnAdvt).lThisYearPrevMonth    'for advt in current month

            tmGrf.lDollars(1) = tmAdvtComm(llLoopOnAdvt).lGross          'Gross $
            tmGrf.lDollars(2) = tmAdvtComm(llLoopOnAdvt).lNet             'Net$
            tmGrf.lDollars(3) = tmAdvtComm(llLoopOnAdvt).lMerch          'Merchandising $
            tmGrf.lDollars(4) = tmAdvtComm(llLoopOnAdvt).lPromo          'Promotions $
            tmGrf.lDollars(9) = tmAdvtComm(llLoopOnAdvt).lAcquisition
            tmGrf.lDollars(6) = tmAdvtComm(llLoopOnAdvt).iUseThisCommPct       'slsp % non-remnant under goal
            tmGrf.lDollars(7) = tmAdvtComm(llLoopOnAdvt).lSlfSplit    'slsp revenue share
            ''tmGrf.lDollars(9) = tmAdvtComm(llLoopOnAdvt).iRemUnderPct    '3-15-00 slsp remnant % under
            tmGrf.lDollars(14) = tmAdvtComm(llLoopOnAdvt).lNet - tmAdvtComm(llLoopOnAdvt).lMerch - tmAdvtComm(llLoopOnAdvt).lPromo - tmAdvtComm(llLoopOnAdvt).lAcquisition   'current month net -net
            tmGrf.lDollars(15) = tmAdvtComm(llLoopOnAdvt).lLastYear             'for advt in current month
            tmGrf.lDollars(16) = tmAdvtComm(llLoopOnAdvt).lThisYearPrevMonth    'for advt in current month

            'place the previous YTD , and other info for sales goal info for the slsp in every prepass record; required to show
            'for each slsp
            tmGrf.iSlfCode = tmAdvtComm(llLoopOnAdvt).iSlfCode
            For llStartInx = LBound(tlSlf) To UBound(tlSlf) - 1
                If tlSlf(llStartInx).iCode = tmGrf.iSlfCode Then
                    ''tmGrf.lDollars(18) = tmSlfPreviousYTD(llStartInx)        'total $ all advertisers for previous YTD by slsp
                    'tmGrf.lDollars(18) = (tmSalesGoalInfo(llStartInx).dPrevYTDNN + 50) \ 100 'total previous ytd net net (current year, all advt) for slsp
                    'tmGrf.lDollars(11) = (tmSalesGoalInfo(llStartInx).dPrevYTDGross + 50) \ 100 'total previous ytd Gross (current year, all advt) for slsp
                    'tmGrf.lDollars(12) = (tmSalesGoalInfo(llStartInx).dPrevYTDNet + 50) \ 100 'total previous ytd  net (current year, all advt) for slsp
                    'tmGrf.lDollars(13) = (tmSalesGoalInfo(llStartInx).dPrevYTDMerchPromo + 50) \ 100 'total previous ytd Merch & promo (current year, all advt) for slsp
                    'tmGrf.lDollars(14) = (tmSalesGoalInfo(llStartInx).dPrevYTDAcq + 50) \ 100 'total previous ytd Acq (current year, all advt) for slsp
                    ''tmgrf.ldollars(6) is used differently between bonus and non-bonus report versions
                    'tmGrf.lDollars(6) = (tmSalesGoalInfo(llStartInx).dPrevYTDNNREmnant + 50) \ 100 'total previous ytd Acq (current year, all advt) for slsp
                    'tmGrf.lDollars(1) = (tmSalesGoalInfo(llStartInx).dYTDNNRemnant + 50) \ 100 'total this month YTD remnants net net
                    'tmGrf.lDollars(9) = (tmSalesGoalInfo(llStartInx).dYTDNN + 50) \ 100   'total this month YTD non-remnants net net
                    
                    ''tmGrf.lDollars(18) = tmSlfPreviousYTD(llStartInx)        'total $ all advertisers for previous YTD by slsp
                    tmGrf.lDollars(17) = (tmSalesGoalInfo(llStartInx).dPrevYTDNN + 50) \ 100 'total previous ytd net net (current year, all advt) for slsp
                    tmGrf.lDollars(10) = (tmSalesGoalInfo(llStartInx).dPrevYTDGross + 50) \ 100 'total previous ytd Gross (current year, all advt) for slsp
                    tmGrf.lDollars(11) = (tmSalesGoalInfo(llStartInx).dPrevYTDNet + 50) \ 100 'total previous ytd  net (current year, all advt) for slsp
                    tmGrf.lDollars(12) = (tmSalesGoalInfo(llStartInx).dPrevYTDMerchPromo + 50) \ 100 'total previous ytd Merch & promo (current year, all advt) for slsp
                    tmGrf.lDollars(13) = (tmSalesGoalInfo(llStartInx).dPrevYTDAcq + 50) \ 100 'total previous ytd Acq (current year, all advt) for slsp
                    ''tmgrf.ldollars(6) is used differently between bonus and non-bonus report versions
                    tmGrf.lDollars(5) = (tmSalesGoalInfo(llStartInx).dPrevYTDNNREmnant + 50) \ 100 'total previous ytd Acq (current year, all advt) for slsp
                    tmGrf.lDollars(0) = (tmSalesGoalInfo(llStartInx).dYTDNNRemnant + 50) \ 100 'total this month YTD remnants net net
                    tmGrf.lDollars(8) = (tmSalesGoalInfo(llStartInx).dYTDNN + 50) \ 100   'total this month YTD non-remnants net net
                    Exit For
                End If
            Next llStartInx

            'tmGrf.iPerGenl(1) = tmAdvtComm(llLoopOnAdvt).iStartFiscalMonth
            'tmGrf.iPerGenl(2) = tmAdvtComm(llLoopOnAdvt).iCurrentMonth
            'tmGrf.iPerGenl(6) = tmAdvtComm(llLoopOnAdvt).iDiscrep                       'error flag: discrepancy in vehicle/slsp commission (sub-companies)
            tmGrf.iPerGenl(0) = tmAdvtComm(llLoopOnAdvt).iStartFiscalMonth
            tmGrf.iPerGenl(1) = tmAdvtComm(llLoopOnAdvt).iCurrentMonth
            tmGrf.iPerGenl(5) = tmAdvtComm(llLoopOnAdvt).iDiscrep                       'error flag: discrepancy in vehicle/slsp commission (sub-companies)
            'if inconsistency with vehicle/sub-company definition, flag as error to show on report

            tmGrf.iAdfCode = tmAdvtComm(llLoopOnAdvt).iAdfCode
            tmGrf.sBktType = tmAdvtComm(llLoopOnAdvt).sType     'contract type (remnant vs non-remnant)

            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        Next llLoopOnAdvt
        Erase tmRVFComm, tmAdvtComm
        Erase tmSalesGoalInfo
        Exit Sub

End Sub
'
'           mGetSlspShare - calculate the slsp share for the 12 periods
'           <input>  llProcessPct - slsp share %
'                   llProjectCash() - array of Projected cash $
'                   llProjectTrade()- array of projected trade $
'           <output>  None
'           Return in llProjectedCash() - remainder of projected cash after slsp has taken share
'           return in llProjectedTrade() - remainder of projected trade after slsp has taken share
'
Public Sub mCumeSlspShare(llProcessPct As Long, llProjectCash() As Long, llProjectTrade() As Long, llCashShare() As Long, llTradeShare() As Long)
Dim ilLoopOnPer As Integer
Dim slAmount As String
Dim slSharePct As String
Dim slStr As String

            For ilLoopOnPer = 1 To 12
                                               
                slAmount = gLongToStrDec(llProjectCash(ilLoopOnPer), 2)
                slSharePct = gLongToStrDec(llProcessPct, 4)                 'slsp or owner split share in %, else 100.0000%
                'slStr = gDivStr(gMulStr(slSharePct, slAmount), "100")                      ' gross portion of possible split
                slStr = gMulStr(slSharePct, slAmount)                      ' gross portion of possible split
                llCashShare(ilLoopOnPer) = Val(gRoundStr(slStr, "1", 0))
                
                slAmount = gLongToStrDec(llProjectTrade(ilLoopOnPer), 2)
                slSharePct = gLongToStrDec(llProcessPct, 4)                 'slsp or owner split share in %, else 100.0000%
                'slStr = gDivStr(gMulStr(slSharePct, slAmount), "100")        ' gross portion of possible split
                slStr = gMulStr(slSharePct, slAmount)        ' gross portion of possible split
                llTradeShare(ilLoopOnPer) = Val(gRoundStr(slStr, "1", 0))

                
            Next ilLoopOnPer
        Exit Sub
End Sub
'
'           Determine how many slsp splits for the Sales Activity
'           <input>  ilListIndex - report index
'                    llSlfsplit() - array of slsp commission splits from cnt header
'                    ilSlspCode() - array of slsp codes
'                    llSlfSplitRev()-array of slsp revenue % from contract header
'                    ilVefCode - vehicle to setup for # of slsp splits
'           <output> updated llslfsplit array
'                    updated ilSlspCode array
'                    updated llSlfSplitRev array
'           return - # of slsp splits
Function mCumeHowManySlsp(ilListIndex As Integer, llSlfSplit() As Long, ilSlspCode() As Integer, llSlfSplitRev() As Long, ilVefCode As Integer) As Integer
Dim ilMnfSubCo As Integer
Dim ilSlfSplitRev(0 To 9) As Long
Dim ilMax As Integer
Dim ilLoopOnSlsp As Integer
Dim ilTemp As Integer


'        ReDim llslfsplit(0 To 9) As Long           '4-20-00 slsp slsp share %
'        ReDim ilSlspCode(0 To 9) As Integer             '4-20-00
'        ReDim llSlfSplitRev(0 To 9) As Long
        For ilTemp = 0 To 9
            llSlfSplit(ilTemp) = 0
            ilSlspCode(ilTemp) = 0
            llSlfSplitRev(ilTemp) = 0
        Next ilTemp
         
        ilMax = 1
        If ilListIndex = CNT_SALESACTIVITY_SS And RptSelCt!ckcSelC12(2).Value = vbChecked Then
             'do split slsp for Sales Activity
             ilMnfSubCo = gGetSubCmpy(tgChfCT, ilSlspCode(), llSlfSplit(), ilVefCode, False, llSlfSplitRev())                                         '4-6-00
                                                    
            ilMax = 10
            'determine the maximum number of entries to process; possibly 0 % comm in the middle of the split slsp
            'start from end to get last valid slsp
            For ilLoopOnSlsp = 9 To 0 Step -1
                If ilSlspCode(ilLoopOnSlsp) > 0 Then
            
                    Exit For
                Else
                    ilMax = ilMax - 1
                End If
            Next ilLoopOnSlsp
            'handle case where theres only 1 slsp but no percentage entered, must be 100% (shouldnt happen)
            If (ilMax <= 0) And llSlfSplitRev(0) = 0 Then      'for some reason there is no revenue % stored in header
                llSlfSplitRev(0) = 1000000
                ilMax = 1
            End If
        Else        'no split slsp, use primary for 100%
            llSlfSplitRev(0) = 1000000
            ilSlspCode(0) = tgChfCT.iSlfCode(0)
            ilMax = 1
        End If
        mCumeHowManySlsp = ilMax
        Exit Function
End Function
'
'
'       Create prepass for Contract Verification.  Gather all contracts that are active
'       between the requested user entered date span.
'       Select user requested contract Verification states:  Not Verified, Send to Agy,
'       or Verified
'       Output results from GRF showing printed by advertiesr name within verification state.
'       In addition, verification date, contract #, agency, salesperson, total gross and contract start/end dates are shown.
Public Sub gCreateContractVerify()
Dim ilRet As Integer
Dim llStartDate As Long
Dim llEndDate As Long
Dim slStartDate As String                       'llLYGetFrom or llTYGetFrom converted to string
Dim slEndDate As String                         'llLYGetTo or LLTYGetTo converted to string
Dim slCntrTypes As String                       'valid contract types to access
Dim slCntrStatus As String                      'valid status (holds, orders, working, etc) to access
Dim ilHOState As Integer                        'include unsch holds/orders, sch holds/orders
Dim llContrCode As Long                         'contr code from gObtainCntrforDate
Dim ilCurrentRecd As Integer                    'index of contract being processed from tlChfAdvtExt
Dim blIncludeNotVerified As Boolean                'true to include contracts not verified
Dim blIncludeVerified As Boolean                   'true to include contracts verified
Dim blIncludeSenttoAgy As Boolean                  'true to include contracts sent to agency
Dim blIncludeState As Boolean

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    blIncludeNotVerified = gSetIncludeExcludeCkc(RptSelCt!ckcSelC3(0))       'include Not verified
    blIncludeVerified = gSetIncludeExcludeCkc(RptSelCt!ckcSelC3(2))       'include  verified
    blIncludeSenttoAgy = gSetIncludeExcludeCkc(RptSelCt!ckcSelC3(1))       'include send to agy
    
    slStartDate = RptSelCt!edcSelCFrom.Text
    llStartDate = gDateValue(slStartDate)
    slStartDate = Format(llStartDate, "m/d/yy")   'make sure string start date has a year appended in case not entered with input

    slEndDate = RptSelCt!edcSelCFrom1.Text
    llEndDate = gDateValue(slEndDate)
    slEndDate = Format(llEndDate, "m/d/yy")    'make sure string end date has a year appended in case not entered with input

    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
      
    slCntrTypes = gBuildCntTypes()      'Setup valid types of contracts to obtain based on user

    slCntrStatus = ""
    slCntrStatus = "HOGN"             'sch/unsch holds & uns holds
    ilHOState = 2                      'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)
    sgCntrForDateStamp = ""            'init the time stamp to read in contracts upon re-entry
    
    ilRet = gObtainCntrForDate(RptSelCt, slStartDate, slEndDate, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())

    'grf fields
    'grfGenDate - generation date
    'grfGenTime - generation time
    'grfchfCode - contract auto code
    
'10-6-15 comment out mainline; contract verification fields renamed for other usage
'    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
'        llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
'        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT())   'get the latest version of this contract
'
'        blIncludeState = False
'        If blIncludeNotVerified And tgChfCT.sVerifyFlag = "N" Then      'not verified sttate
'            blIncludeState = True
'        End If
'        If blIncludeSenttoAgy And tgChfCT.sVerifyFlag = "S" Then      'Sent to Agy sttate
'            blIncludeState = True
'        End If
'        If blIncludeVerified And tgChfCT.sVerifyFlag = "V" Then      ' verified sttate
'            blIncludeState = True
'        End If
'
'        If blIncludeState Then
'            tmGrf.lChfCode = tgChfCT.lCode          'contr code
'
'            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
'        End If
'    Next ilCurrentRecd
    
    Erase tlChfAdvtExt
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    btrDestroy hmGrf
    btrDestroy hmCHF
End Sub
'
'           create activity of what insertion orders were sent based
'           on matching the current external version of an order, and comparing
'           it the external rev # sent in the lines
'
Public Sub gCreateInsertionActivity()
 Dim ilRet As Integer
Dim blGetAllVersions As Boolean
Dim slCntrType As String            'HOGN (sch & unsched holds & orders)
Dim ilHOState As Integer            'get latest version
Dim slCntrStatus As String          'get all contract types (std, pi, dr, etc)
Dim slDate As String                'temp
Dim llDate As Long                  'temp
Dim llEarliestDate As Long          'earliest date user requested (sent or active )
Dim llLatestDate As Long            'latest date user requested (sent or active)
Dim slEarliestDate As String
Dim slLatestDate As String
Dim ilCurrentRecd As Integer        'loop for contracts gathered
Dim llCurrentCntrNo As Long         'contract # processing
Dim llContrCode As Long             'contract code processing
Dim blAtLeastOneNotSent As Boolean        'flag to indicate at least 1 vehicle was not sent
Dim llDateSent As Long
Dim llChfStartDate As Long
Dim llChfEndDate As Long
Dim llUserStartDateSent As Long         'user entered start date sent
Dim llUserEndDateSent As Long           'user entered end date sent
Dim llUserActiveStartDate As Long       'user entered active start date
Dim llUserActiveEndDate As Long         'user entered active end date
Dim ilWhichRate As Integer              ' = 1     0 = use spot rate, 1 = use acq rate, 2 = if acq non-zero, use it.  otherwise if 0, default to line rate
Dim ilWeekOrMonth As Integer            '= 1       'place in month bkts, else using weekly calculates index into an weekly bucket array
'Dim llProjectDates(1 To 2) As Long      'start & end dates to test to gather $
ReDim llProjectDates(0 To 2) As Long      'start & end dates to test to gather $. Index zero ignored
'Dim llProjectGross(1 To 2) As Long
ReDim llProjectGross(0 To 2) As Long    'Index zero ignored
'Dim llProjectSpots(1 To 2) As Long
ReDim llProjectSpots(0 To 2) As Long    'Index zero ignored
Dim ilClf As Integer                    'loop on lines
Dim ilVefInx As Integer                 'index into tgMVef array
Dim slVehiclesSent As String
Dim slVehiclesNotSent As String
Dim hlTxr As Integer
Dim tlTxr As TXR
Dim hlUrf As Integer
Dim tlUrf As URF
Dim tlSrchKey0 As INTKEY0
Dim blNothingSent As Boolean
Dim blProcessThis As Boolean
Dim ilLen As Integer
Dim ilStartPos As Integer
Dim ilMaxFieldLen As Integer

            hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmCHF)
                btrDestroy hmCHF
                Exit Sub
            End If
            imCHFRecLen = Len(tmChf)
        
            hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmGrf)
                ilRet = btrClose(hmCHF)
                btrDestroy hmGrf
                btrDestroy hmCHF
                Exit Sub
            End If
            imGrfRecLen = Len(tmGrf)
            
            hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmClf)
                ilRet = btrClose(hmGrf)
                ilRet = btrClose(hmCHF)
                btrDestroy hmClf
                btrDestroy hmGrf
                btrDestroy hmCHF
                Exit Sub
            End If
            imClfRecLen = Len(tmClf)
            
            hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmCff)
                ilRet = btrClose(hmClf)
                ilRet = btrClose(hmGrf)
                ilRet = btrClose(hmCHF)
                btrDestroy hmCff
                btrDestroy hmClf
                btrDestroy hmGrf
                btrDestroy hmCHF
                Exit Sub
            End If
            imCffRecLen = Len(tmCff)
            
            hlTxr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hlTxr, "", sgDBPath & "txr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hlTxr)
                ilRet = btrClose(hmCff)
                ilRet = btrClose(hmClf)
                ilRet = btrClose(hmGrf)
                ilRet = btrClose(hmCHF)
                btrDestroy hlTxr
                btrDestroy hmCff
                btrDestroy hmClf
                btrDestroy hmGrf
                btrDestroy hmCHF
                Exit Sub
            End If

            hlUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hlUrf, "", sgDBPath & "urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hlUrf)
                ilRet = btrClose(hlTxr)
                ilRet = btrClose(hmCff)
                ilRet = btrClose(hmClf)
                ilRet = btrClose(hmGrf)
                ilRet = btrClose(hmCHF)
                btrDestroy hlUrf
                btrDestroy hlTxr
                btrDestroy hmCff
                btrDestroy hmClf
                btrDestroy hmGrf
                btrDestroy hmCHF
                Exit Sub
            End If
            
            'ReDim tlChfAdvtExt(1 To 1) As CHFADVTEXT
            ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
            lmSingleCntr = Val(RptSelCt!edcText.Text)
            blGetAllVersions = False
            If RptSelCt!rbcSelC11(1).Value = True Then          'get all versions?
                blGetAllVersions = True
            End If
            
            'get the earliest and latest dates from user requests
            slDate = RptSelCt!CSI_CalFrom.Text              'Date: 12/18/2019 added CSI calendar control for date entry --> edcSelCFrom.Text              'user sent start date
            llEarliestDate = gDateValue(slDate)
            llUserStartDateSent = llEarliestDate            'user entered start date sent
            slDate = RptSelCt!CSI_CalTo.Text                'Date: 12/18/2019 added CSI calendar control for date entry --> edcSelCFrom1.Text             'user sent end date
            llLatestDate = gDateValue(slDate)
            llDate = gDateValue(slDate)
            llUserEndDateSent = llDate                  'user entered end date sent
                        
            slDate = RptSelCt!CSI_From1.Text                'Date: 12/18/2019 added CSI calendar control for date entry --> edcSelCTo.Text            'user active start date
            llUserActiveStartDate = gDateValue(slDate)         'user entered active start date
            slDate = RptSelCt!CSI_To1.Text                  'Date: 12/18/2019 added CSI calendar control for date entry --> edcSelCTo1.Text           'user active end date
            llDate = gDateValue(slDate)
            llUserActiveEndDate = llDate                    'user entered active end date
            
            If llEarliestDate > llUserActiveStartDate Then
                llEarliestDate = llUserActiveStartDate          'set earlier start date to gather
            End If
            If llLatestDate < llUserActiveEndDate Then          'set later date to gather
                llLatestDate = llUserActiveEndDate
            End If
            
            slEarliestDate = Format$(llEarliestDate, "m/d/yy")
            slLatestDate = Format$(llLatestDate, "m/d/yy")
            If lmSingleCntr > 0 Then                   'single contract entered?
                'ReDim tlChfAdvtExt(1 To 2) As CHFADVTEXT
                ReDim tlChfAdvtExt(0 To 1) As CHFADVTEXT
                llCurrentCntrNo = gSingleContract(hmCHF, tmChf, lmSingleCntr)

                'tlChfAdvtExt(1).lCode = tmChf.lCode
                'tlChfAdvtExt(1).lVefCode = tmChf.lVefCode
                'tlChfAdvtExt(1).lCntrNo = lmSingleCntr
                tlChfAdvtExt(0).lCode = tmChf.lCode
                tlChfAdvtExt(0).lVefCode = tmChf.lVefCode
                tlChfAdvtExt(0).lCntrNo = lmSingleCntr
            Else
                slCntrType = ""                     'blank out string for "All"
                ilHOState = 2                       'get latest orders & revisions
                slCntrStatus = "HOGN"               'include sch and uns holds & orders
                sgCntrForDateStamp = ""            'init the time stamp to read in contracts upon re-entry

                ilRet = gObtainCntrForDate(RptSelCt, slEarliestDate, slLatestDate, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())
            End If
            
            ilWhichRate = 1     '0 = use spot rate, 1 = use acq rate, 2 = if acq non-zero, use it.  otherwise if 0, default to line rate
            ilWeekOrMonth = 1       'place in month bkts, else using weekly calculates index into an weekly bucket array
            
            tmGrf.lGenTime = lgNowTime
            tmGrf.iGenDate(0) = igNowDate(0)
            tmGrf.iGenDate(1) = igNowDate(1)
            tlTxr.lGenTime = lgNowTime
            tlTxr.iGenDate(0) = igNowDate(0)
            tlTxr.iGenDate(1) = igNowDate(1)
            
            'Contracts gathered encompassing the largest date span entered that are active.
            'Date must be active in the user entered active date span to continue processing.
            'Contract to be shown on report only if  a schedule line has a BArter defined schedule line
            'chfEDSSentExtRevNo flag of -2 indicates to ignore contract.  client that is already on v7 with feature off, and then turning on
            'chfEdsSentExtRevNo flag of -1 indicates its New, nothing sent
            'All other chfEdsSentExtRevNo flags need to be processed to see if anything has been sent
            'Compare chfEdsSentExtRevNo with chfExtRevNo:   Compare lines ext # with header ext #.  If matching, line has been sent; otherwise line has not been sent.
            '   Keep track if all lines of order sent, or a partial set.  If nothing sent, show Not Yet Sent on report.  If Partial set sent, show the list of vehicles sent.
            '   If all Sent, show All sent on report
            'If chfEdsSentExtRevNo does not match chfExtRevNo: Ignore, no change
            For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
            
                llCurrentCntrNo = tlChfAdvtExt(ilCurrentRecd).lCntrNo
                llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT(), False)
                Do While llCurrentCntrNo = tgChfCT.lCntrNo
                    blProcessThis = False
                    gUnpackDateLong tgChfCT.iEDSSentDate(0), tgChfCT.iEDSSentDate(1), llDateSent
                    gUnpackDateLong tgChfCT.iStartDate(0), tgChfCT.iStartDate(1), llChfStartDate
                    gUnpackDateLong tgChfCT.iEndDate(0), tgChfCT.iEndDate(1), llChfEndDate
                    If ((llUserActiveStartDate <= llChfEndDate And llUserActiveEndDate >= llChfStartDate) Or (llUserActiveEndDate >= llChfStartDate And llUserActiveEndDate <= llChfEndDate)) And (tgChfCT.iEDSSentExtRevNo <> -2) Then    'test if requested dates span the active date of agreement. -2 indicates totally ignore
                        'test if contract active dates within requested span
                        'loop on schedule lines to determine if it is a barter vehicle and was sent
                        blAtLeastOneNotSent = False
                        'these dates will be used as the limits to gather gross acq $ from the
                        llProjectDates(1) = llChfStartDate
                        llProjectDates(2) = llChfEndDate
                        slVehiclesSent = ""
                        slVehiclesNotSent = ""
                        blNothingSent = True
                        llProjectGross(1) = 0
                        llProjectSpots(1) = 0
                        blProcessThis = True
                        tlTxr.lSeqNo = 0
                        If tgChfCT.iExtRevNo >= 0 Then          '-1 indicates new, nothing sent
                            If llDateSent >= llUserStartDateSent And llDateSent <= llUserEndDateSent Then         'is contract date sent within requested span.
                                For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                                    tmClf = tgClfCT(ilClf).ClfRec
                                    If (tmClf.sType = "S" Or tmClf.sType = "H") And (gIsOnInsertions(tmClf.iVefCode) = True) Then      'project standard or hidden lines if its a barter
                                        ilVefInx = gBinarySearchVef(tmClf.iVefCode)
                                        If tgChfCT.iEDSSentExtRevNo = tgChfCT.iExtRevNo Then           'sent (contracts extrev# is same as the eds ext rev #, see if all barter lines sent
                                            If tgChfCT.iEDSSentExtRevNo = tmClf.iEDSSentExtRevNo Then   'now check the lines to see if everything went out
                                                blNothingSent = False
                                                If Trim$(slVehiclesSent) = "" Then
                                                    slVehiclesSent = Trim$(tgMVef(ilVefInx).sName)
                                                Else
                                                    If InStr(slVehiclesSent, tgMVef(ilVefInx).sName) = 0 Then       '0 = vehicle not found, add it to string; otherwise its a duplicate vehiclename
                                                        slVehiclesSent = slVehiclesSent & ", " & Trim$(tgMVef(ilVefInx).sName)
                                                    End If
                                                End If
                                            Else                'vehicle not sent (unused now, maybe for later)
                                                blAtLeastOneNotSent = True
                                                If Trim$(slVehiclesNotSent) = "" Then
                                                    slVehiclesNotSent = Trim$(tgMVef(ilVefInx).sName)
                                                Else
                                                    If InStr(slVehiclesNotSent, tgMVef(ilVefInx).sName) = 0 Then       '0 = vehicle not found, add it to string; otherwise its a duplicate vehiclename
                                                        slVehiclesNotSent = slVehiclesNotSent & ", " & Trim$(tgMVef(ilVefInx).sName)
                                                    End If
                                                End If
                                            End If
                                            
                                        Else        'contract not sent at all
                                            blNothingSent = True
                                        End If
                                        'accum entire orders acquisition gross
                                        gBuildFlightSpotsAndRevenue ilClf, llProjectDates(), 1, 2, llProjectGross(), llProjectSpots(), ilWeekOrMonth, ilWhichRate, tgClfCT(), tgCffCT(), "G"
                                    End If      'clftype = S or clftype = H
                                Next ilClf
                            Else                    'not sent - gather the acq gross
                                For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                                    tmClf = tgClfCT(ilClf).ClfRec
                                    If (tmClf.sType = "S" Or tmClf.sType = "H") And (gIsOnInsertions(tmClf.iVefCode) = True) Then      'project standard or hidden lines if its a barter
                                        'accum entire orders acquisition gross
                                        gBuildFlightSpotsAndRevenue ilClf, llProjectDates(), 1, 2, llProjectGross(), llProjectSpots(), ilWeekOrMonth, ilWhichRate, tgClfCT(), tgCffCT(), "G"
                                    Else
                                        ilClf = ilClf
                                    End If      'clftype = S or clftype = H
                                Next ilClf
                            End If          'lldate sent >= lluserstartdatesend and lldate send <= lluserenddatesent
                        Else
                            'nothing sent
                            'blProcessThis has been set to true
                            'blNothingSent has been set to true
                        End If
                    End If          'contract Active date test
                    If blProcessThis Then
                        tmGrf.lChfCode = tgChfCT.lCode
                        'tmGrf.lDollars(1) = llProjectGross(1)           'total gross $
                        tmGrf.lDollars(0) = llProjectGross(1)           'total gross $
                        tmGrf.sBktType = "A"                             'status, assume all sent
                        If blNothingSent Then                           'if nothing sent, show Not Yet Sent
                            tmGrf.sBktType = "N"                        'nothing sent
                        Else                                            'something sent,  all or partial?
                            If blAtLeastOneNotSent Then                  'at least one not sent, its partial
                                tmGrf.sBktType = "P"                        'partial
                                'need to create a txr record to link to to show the list of vehicles sent
                                
                                If Trim$(slVehiclesSent) <> "" Then
                                    ilStartPos = 1
                                    ilLen = Len(Trim$(slVehiclesSent))
                                    Do While ilLen > 0
                                        ilMaxFieldLen = 200
                                        If ilLen < 200 Then
                                            ilMaxFieldLen = ilLen
                                        End If
                                        
                                        tlTxr.sText = Mid$(slVehiclesSent, ilStartPos, ilMaxFieldLen)
                                        tlTxr.lCsfCode = tgChfCT.lCode
                                        tlTxr.lSeqNo = tlTxr.lSeqNo + 1
                                        ilRet = btrInsert(hlTxr, tlTxr, Len(tlTxr), INDEXKEY0)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            Exit Do
                                        Else
                                            ilLen = ilLen - 200             '200 is max field length that has been written
                                            ilStartPos = ilStartPos + ilMaxFieldLen
                                        End If
                                    Loop
                                End If          'trim$(slvehiclessent) <> ""
                            End If              'blatleastonenotsent
                        End If                  'blnothingsent
                        
                        'convert date & time to numbers for sorting
                        'gUnpackDateLong tgChfCT.iEDSSentDate(0), tgChfCT.iEDSSentDate(1), tmGrf.lDollars(2)     'sent date
                        gUnpackDateLong tgChfCT.iEDSSentDate(0), tgChfCT.iEDSSentDate(1), tmGrf.lDollars(1)     'sent date
                        'gUnpackTimeLong tgChfCT.iEDSSentTime(0), tgChfCT.iEDSSentTime(1), False, tmGrf.lDollars(3)    'sent time
                        gUnpackTimeLong tgChfCT.iEDSSentTime(0), tgChfCT.iEDSSentTime(1), False, tmGrf.lDollars(2)    'sent time

                        tmGrf.sGenDesc = ""         'EDS sent user field, it is encrypted
                        tlSrchKey0.iCode = tgChfCT.iEDSSentUrfCode
                        ilRet = btrGetEqual(hlUrf, tlUrf, Len(tlUrf), tlSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                        '   grfgentime - generation time for filter to crystal
                        '   grfgendate - generation date for filter to crystal
                        '   grfchfcode - internal contract code
                        '   grfdollars(1) - aquisition total gross $
                        '   grfdollars(2) - EDS sent date for sorting
                        '   grfdollars(3) - eds sent time for sorting
                        '   grfbkttype - status of eds (N = not sent, A = all sent, P = partially sent
                            tmGrf.sGenDesc = Trim$(gDecryptField(tlUrf.sRept))
                        End If
                        
                        'create the primary record to print
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                        
                    End If                      'blprocessthis
                    
                    If blGetAllVersions Then        'all versions option
                        If tgChfCT.iCntRevNo > 0 Then       'this is R0, no other previous versions exist
                            tmChfSrchKey1.iCntRevNo = tgChfCT.iCntRevNo
                            tmChfSrchKey1.iPropVer = tgChfCT.iPropVer
                            tmChfSrchKey1.lCntrNo = tgChfCT.lCntrNo
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'set the key pointer to get previous version properly
                            If ilRet = BTRV_ERR_NONE Then
                                tmChfSrchKey1.iCntRevNo = tgChfCT.iCntRevNo - 1
                                tmChfSrchKey1.iPropVer = 32000
                                tmChfSrchKey1.lCntrNo = tgChfCT.lCntrNo
                                ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)       'get the previous version header
                                llContrCode = tmChf.lCode
                                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT(), False)      'get the entire contract
                            Else            'error on the read of previous version, exit this contract
                                Exit Do
                            End If
                        Else                'already version 0, no more to process; exit this contract
                            Exit Do
                        End If
                    Else                    'only process latest version, exit this contract
                        Exit Do
                    End If
                Loop            'while tgcntct.lcntrno = llCurrentCntrNo
            Next ilCurrentRecd
            
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hlTxr)
            ilRet = btrClose(hlUrf)
            btrDestroy hmCff
            btrDestroy hmClf
            btrDestroy hmGrf
            btrDestroy hmCHF
            btrDestroy hlTxr
            btrDestroy hlUrf
            Exit Sub
End Sub
'
'           create activity of what ProposalXML orders were sent based
'           on matching the current external version of an order, and comparing
'           it the external rev # sent in the lines
'
Public Sub gCreateXMLActivity()
Dim ilRet As Integer
Dim blGetAllVersions As Boolean
Dim slCntrType As String            'HOGN (sch & unsched holds & orders)
Dim ilHOState As Integer            'get latest version
Dim slCntrStatus As String          'get all contract types (std, pi, dr, etc)
Dim slDate As String                'temp
Dim llDate As Long                  'temp
Dim llEarliestDate As Long          'earliest date user requested (sent or active )
Dim llLatestDate As Long            'latest date user requested (sent or active)
Dim slEarliestDate As String
Dim slLatestDate As String
Dim ilCurrentRecd As Integer        'loop for contracts gathered
Dim llCurrentCntrNo As Long         'contract # processing
Dim llContrCode As Long             'contract code processing
Dim blAtLeastOneNotSent As Boolean        'flag to indicate at least 1 vehicle was not sent
Dim llDateSent As Long
Dim llChfStartDate As Long
Dim llChfEndDate As Long
Dim llUserStartDateSent As Long         'user entered start date sent
Dim llUserEndDateSent As Long           'user entered end date sent
Dim llUserActiveStartDate As Long       'user entered active start date
Dim llUserActiveEndDate As Long         'user entered active end date
Dim ilWhichRate As Integer              ' = 1     0 = use spot rate, 1 = use acq rate, 2 = if acq non-zero, use it.  otherwise if 0, default to line rate
Dim ilWeekOrMonth As Integer            '= 1       'place in month bkts, else using weekly calculates index into an weekly bucket array
'Dim llProjectDates(1 To 2) As Long      'start & end dates to test to gather $
ReDim llProjectDates(0 To 2) As Long      'start & end dates to test to gather $. Index zero ignored
'Dim llProjectGross(1 To 2) As Long
ReDim llProjectGross(0 To 2) As Long    'Index zero ignored
'Dim llProjectSpots(1 To 2) As Long
ReDim llProjectSpots(0 To 2) As Long    'Index zero ignored
Dim ilClf As Integer                    'loop on lines
Dim ilVefInx As Integer                 'index into tgMVef array
Dim ilVffInx As Integer
Dim slVehiclesSent As String
Dim slVehiclesNotSent As String
Dim hlTxr As Integer
Dim tlTxr As TXR
Dim hlUrf As Integer
Dim tlUrf As URF
Dim tlSrchKey0 As INTKEY0
Dim blNothingSent As Boolean
Dim blProcessThis As Boolean
Dim ilLen As Integer
Dim ilStartPos As Integer
Dim ilMaxFieldLen As Integer

            hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmCHF)
                btrDestroy hmCHF
                Exit Sub
            End If
            imCHFRecLen = Len(tmChf)
        
            hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmGrf)
                ilRet = btrClose(hmCHF)
                btrDestroy hmGrf
                btrDestroy hmCHF
                Exit Sub
            End If
            imGrfRecLen = Len(tmGrf)
            
            hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmClf)
                ilRet = btrClose(hmGrf)
                ilRet = btrClose(hmCHF)
                btrDestroy hmClf
                btrDestroy hmGrf
                btrDestroy hmCHF
                Exit Sub
            End If
            imClfRecLen = Len(tmClf)
            
            hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmCff)
                ilRet = btrClose(hmClf)
                ilRet = btrClose(hmGrf)
                ilRet = btrClose(hmCHF)
                btrDestroy hmCff
                btrDestroy hmClf
                btrDestroy hmGrf
                btrDestroy hmCHF
                Exit Sub
            End If
            imCffRecLen = Len(tmCff)
            
            hlTxr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hlTxr, "", sgDBPath & "txr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hlTxr)
                ilRet = btrClose(hmCff)
                ilRet = btrClose(hmClf)
                ilRet = btrClose(hmGrf)
                ilRet = btrClose(hmCHF)
                btrDestroy hlTxr
                btrDestroy hmCff
                btrDestroy hmClf
                btrDestroy hmGrf
                btrDestroy hmCHF
                Exit Sub
            End If

            hlUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hlUrf, "", sgDBPath & "urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hlUrf)
                ilRet = btrClose(hlTxr)
                ilRet = btrClose(hmCff)
                ilRet = btrClose(hmClf)
                ilRet = btrClose(hmGrf)
                ilRet = btrClose(hmCHF)
                btrDestroy hlUrf
                btrDestroy hlTxr
                btrDestroy hmCff
                btrDestroy hmClf
                btrDestroy hmGrf
                btrDestroy hmCHF
                Exit Sub
            End If
            
            'ReDim tlChfAdvtExt(1 To 1) As CHFADVTEXT
            ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
            lmSingleCntr = Val(RptSelCt!edcText.Text)
            blGetAllVersions = False
            If RptSelCt!rbcSelC11(1).Value = True Then          'get all versions?
                blGetAllVersions = True
            End If
            
            'get the earliest and latest dates from user requests
            slDate = RptSelCt!CSI_CalFrom.Text              'Date: 1/2/2020 added CSI calendar control for date entries --> edcSelCFrom.Text              'user sent start date
            llEarliestDate = gDateValue(slDate)
            llUserStartDateSent = llEarliestDate            'user entered start date sent
            slDate = RptSelCt!CSI_CalTo.Text                'Date: 1/2/2020 added CSI calendar control for date entries --> edcSelCFrom1.Text             'user sent end date
            llLatestDate = gDateValue(slDate)
            llDate = gDateValue(slDate)
            llUserEndDateSent = llDate                  'user entered end date sent
                        
            slDate = RptSelCt!CSI_From1.Text                'Date: 1/2/2020 added CSI calendar control for date entries --> edcSelCTo.Text            'user active start date
            llUserActiveStartDate = gDateValue(slDate)         'user entered active start date
            slDate = RptSelCt!CSI_To1.Text                  'Date: 1/2/2020 added CSI calendar control for date entries --> edcSelCTo1.Text           'user active end date
            llDate = gDateValue(slDate)
            llUserActiveEndDate = llDate                    'user entered active end date
            
            If llEarliestDate > llUserActiveStartDate Then
                llEarliestDate = llUserActiveStartDate          'set earlier start date to gather
            End If
            If llLatestDate < llUserActiveEndDate Then          'set later date to gather
                llLatestDate = llUserActiveEndDate
            End If
            
            slEarliestDate = Format$(llEarliestDate, "m/d/yy")
            slLatestDate = Format$(llLatestDate, "m/d/yy")
            If lmSingleCntr > 0 Then                   'single contract entered?
                'ReDim tlChfAdvtExt(1 To 2) As CHFADVTEXT
                ReDim tlChfAdvtExt(0 To 1) As CHFADVTEXT
                llCurrentCntrNo = gSingleContract(hmCHF, tmChf, lmSingleCntr)

                'tlChfAdvtExt(1).lCode = tmChf.lCode
                'tlChfAdvtExt(1).lVefCode = tmChf.lVefCode
                'tlChfAdvtExt(1).lCntrNo = lmSingleCntr
                tlChfAdvtExt(0).lCode = tmChf.lCode
                tlChfAdvtExt(0).lVefCode = tmChf.lVefCode
                tlChfAdvtExt(0).lCntrNo = lmSingleCntr
            Else
                slCntrType = ""                     'blank out string for "All"
                ilHOState = 2                       'get latest orders & revisions
                slCntrStatus = "HOGN"               'include sch and uns holds & orders
                sgCntrForDateStamp = ""            'init the time stamp to read in contracts upon re-entry

                ilRet = gObtainCntrForDate(RptSelCt, slEarliestDate, slLatestDate, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())
            End If
            
            ilWhichRate = 1     '0 = use spot rate, 1 = use acq rate, 2 = if acq non-zero, use it.  otherwise if 0, default to line rate
            ilWeekOrMonth = 1       'place in month bkts, else using weekly calculates index into an weekly bucket array
            
            tmGrf.lGenTime = lgNowTime
            tmGrf.iGenDate(0) = igNowDate(0)
            tmGrf.iGenDate(1) = igNowDate(1)
            tlTxr.lGenTime = lgNowTime
            tlTxr.iGenDate(0) = igNowDate(0)
            tlTxr.iGenDate(1) = igNowDate(1)
            
            'Contracts gathered encompassing the largest date span entered that are active.
            'Date must be active in the user entered active date span to continue processing.
            'Contract to be shown on report only if  a schedule line has a BArter defined schedule line
            'chfEDSSentExtRevNo flag of -2 indicates to ignore contract.  client that is already on v7 with feature off, and then turning on
            'All other chfEdsSentExtRevNo flags need to be processed to see if anything has been sent
            'Compare chfEdsSentExtRevNo with chfExtRevNo:   Compare lines ext # with header ext #.  If matching, line has been sent; otherwise line has not been sent.
            '   Keep track if all lines of order sent, or a partial set.  If nothing sent, show Not Yet Sent on report.  If Partial set sent, show the list of vehicles sent.
            '   If all Sent, show All sent on report
            'If chfEdsSentExtRevNo does not match chfExtRevNo: Ignore, no change
            For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
            
                llCurrentCntrNo = tlChfAdvtExt(ilCurrentRecd).lCntrNo
                llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT(), False)
                Do While llCurrentCntrNo = tgChfCT.lCntrNo
                    blProcessThis = False
                    gUnpackDateLong tgChfCT.iXMLSentDate(0), tgChfCT.iXMLSentDate(1), llDateSent
                    gUnpackDateLong tgChfCT.iStartDate(0), tgChfCT.iStartDate(1), llChfStartDate
                    gUnpackDateLong tgChfCT.iEndDate(0), tgChfCT.iEndDate(1), llChfEndDate
                    If ((llUserActiveStartDate <= llChfEndDate And llUserActiveEndDate >= llChfStartDate) Or (llUserActiveEndDate >= llChfStartDate And llUserActiveEndDate <= llChfEndDate)) And (llChfStartDate < llChfEndDate) And (tgChfCT.iXMLSentExtRevNo <> -2) Then     'test if requested dates span the active date of agreement. -2 indicates totally ignore
                        'test if contract active dates within requested span
                        'loop on schedule lines to determine if it is a barter vehicle and was sent
                        blAtLeastOneNotSent = False
                        'these dates will be used as the limits to gather gross acq $ from the
                        llProjectDates(1) = llChfStartDate
                        llProjectDates(2) = llChfEndDate
                        slVehiclesSent = ""
                        slVehiclesNotSent = ""
                        blNothingSent = True
                        llProjectGross(1) = 0
                        llProjectSpots(1) = 0
                        blProcessThis = True
                        tlTxr.lSeqNo = 0
                        If tgChfCT.iExtRevNo >= 0 Then          '
                            If llDateSent >= llUserStartDateSent And llDateSent <= llUserEndDateSent Then         'is contract date sent within requested span.
                                For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                                    tmClf = tgClfCT(ilClf).ClfRec
                                    ilVefInx = gBinarySearchVef(tmClf.iVefCode)
                                 
                                    If (tmClf.sType = "S" Or tmClf.sType = "H") And tgMVef(ilVefInx).sType = "R" And gIsInsertionExport(tmClf.iVefCode) = True Then           'project standard or hidden lines if its a barter
                                        
                                        If (tgChfCT.iXMLSentExtRevNo = tgChfCT.iExtRevNo) Then            'sent (contracts extrev# is same as the eds ext rev #, see if all barter lines sent
                                            If (tgChfCT.iXMLSentExtRevNo = tmClf.iXMLSentExtRevNo) And tgChfCT.iXMLSentDate(1) > 0 Then    'now check the lines to see if everything went out
                                                blNothingSent = False
                                                If Trim$(slVehiclesSent) = "" Then
                                                    slVehiclesSent = Trim$(tgMVef(ilVefInx).sName)
                                                Else
                                                    If InStr(slVehiclesSent, tgMVef(ilVefInx).sName) = 0 Then       '0 = vehicle not found, add it to string; otherwise its a duplicate vehiclename
                                                        slVehiclesSent = slVehiclesSent & ", " & Trim$(tgMVef(ilVefInx).sName)
                                                    End If
                                                End If
                                            Else                'vehicle not sent (unused now, maybe for later)
                                                blAtLeastOneNotSent = True
                                                If Trim$(slVehiclesNotSent) = "" Then
                                                    slVehiclesNotSent = Trim$(tgMVef(ilVefInx).sName)
                                                Else
                                                    If InStr(slVehiclesNotSent, tgMVef(ilVefInx).sName) = 0 Then       '0 = vehicle not found, add it to string; otherwise its a duplicate vehiclename
                                                        slVehiclesNotSent = slVehiclesNotSent & ", " & Trim$(tgMVef(ilVefInx).sName)
                                                    End If
                                                End If
                                            End If
                                            
                                        Else        'contract not sent at all
                                            'ext sent rev # not equal to the external rev #, hasnt been sent
                                            blNothingSent = True
                                        End If
                                    End If      'clftype = S or clftype = H and veftype = "R"
                                Next ilClf
                                If blNothingSent = False Then           'something was found to be sent, see if everything sent
                                    If blAtLeastOneNotSent Then
                                        blNothingSent = True            'show All sent or nothing sent
                                    End If
                                End If
                            Else                    'not sent - gather the acq gross
                                'no $ required to print
                                ilClf = ilClf
                                blProcessThis = False
                                For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                                    tmClf = tgClfCT(ilClf).ClfRec
                                    ilVefInx = gBinarySearchVef(tmClf.iVefCode)
                                    If (tmClf.sType = "S" Or tmClf.sType = "H") And tgMVef(ilVefInx).sType = "R" And gIsInsertionExport(tmClf.iVefCode) = True Then      'have to have at least 1 rep vehicle to process as Not Sent
                                        'accum entire orders acquisition gross
                                        'gBuildFlightSpotsAndRevenue ilClf, llProjectDates(), 1, 2, llProjectGross(), llProjectSpots(), ilWeekOrMonth, ilWhichRate, tgClfCT(), tgCffCT(), "G"
                                        blProcessThis = True
                                        
                                    End If      'clftype = S or clftype = H
                                Next ilClf
                                
                                    
                            End If          'lldate sent >= lluserstartdatesend and lldate send <= lluserenddatesent
                        End If      'tgChfCT.iExtRevNo >= 0
                    End If          'contract Active date test
                    If blProcessThis Then
                        tmGrf.lChfCode = tgChfCT.lCode
                        'tmGrf.lDollars(1) = llProjectGross(1)           'total gross $
                        tmGrf.lDollars(0) = llProjectGross(1)           'total gross $
                        tmGrf.sBktType = "A"                             'status, assume all sent
                        If blNothingSent Then                           'if nothing sent, show Not Yet Sent
                            tmGrf.sBktType = "N"                        'nothing sent
'                        Else                                            'something sent,  all or partial?
'                            tmGrf.sBktType = "N"                        'all or nothing, dont care about partials
'                            If blAtLeastOneNotSent Then                  'at least one not sent, its partial
'                                tmGrf.sBktType = "P"                        'partial
'                                'need to create a txr record to link to to show the list of vehicles sent
'
'                                If Trim$(slVehiclesSent) <> "" Then
'                                    ilStartPos = 1
'                                    ilLen = Len(Trim$(slVehiclesSent))
'                                    Do While ilLen > 0
'                                        ilMaxFieldLen = 200
'                                        If ilLen < 200 Then
'                                            ilMaxFieldLen = ilLen
'                                        End If
'
'                                        tlTxr.sText = Mid$(slVehiclesSent, ilStartPos, ilMaxFieldLen)
'                                        tlTxr.lCsfCode = tgChfCT.lCode
'                                        tlTxr.lSeqNo = tlTxr.lSeqNo + 1
'                                        ilRet = btrInsert(hlTxr, tlTxr, Len(tlTxr), INDEXKEY0)
'                                        If ilRet <> BTRV_ERR_NONE Then
'                                            Exit Do
'                                        Else
'                                            ilLen = ilLen - 200             '200 is max field length that has been written
'                                            ilStartPos = ilStartPos + ilMaxFieldLen
'                                        End If
'                                    Loop
'                                End If          'trim$(slvehiclessent) <> ""
'                            End If              'blatleastonenotsent
                        End If                  'blnothingsent
                        
                        If (RptSelCt!ckcSelC12(0).Value = vbChecked And tmGrf.sBktType = "A") Or (RptSelCt!ckcSelC12(1).Value = vbChecked And tmGrf.sBktType = "N") Then
                            'convert date & time to numbers for sorting
                            'gUnpackDateLong tgChfCT.iXMLSentDate(0), tgChfCT.iXMLSentDate(1), tmGrf.lDollars(2)     'sent date
                            gUnpackDateLong tgChfCT.iXMLSentDate(0), tgChfCT.iXMLSentDate(1), tmGrf.lDollars(1)     'sent date
                            'gUnpackTimeLong tgChfCT.iXMLSentTime(0), tgChfCT.iXMLSentTime(1), False, tmGrf.lDollars(3)    'sent time
                            gUnpackTimeLong tgChfCT.iXMLSentTime(0), tgChfCT.iXMLSentTime(1), False, tmGrf.lDollars(2)    'sent time
    
                            tmGrf.sGenDesc = ""         'EDS sent user field, it is encrypted
                            tlSrchKey0.iCode = tgChfCT.iXMLSentUrfCode
                            ilRet = btrGetEqual(hlUrf, tlUrf, Len(tlUrf), tlSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                            '   grfgentime - generation time for filter to crystal
                            '   grfgendate - generation date for filter to crystal
                            '   grfchfcode - internal contract code
                            '   grfdollars(1) - aquisition total gross $
                            '   grfdollars(2) - EDS sent date for sorting
                            '   grfdollars(3) - eds sent time for sorting
                            '   grfbkttype - status of eds (N = not sent, A = all sent, P = partially sent
                                tmGrf.sGenDesc = Trim$(gDecryptField(tlUrf.sRept))
                            End If
                            
                            'create the primary record to print
                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                        End If
                    End If                      'blprocessthis
                    
                    If blGetAllVersions Then        'all versions option
                        If tgChfCT.iCntRevNo > 0 Then       'this is R0, no other previous versions exist
                            tmChfSrchKey1.iCntRevNo = tgChfCT.iCntRevNo
                            tmChfSrchKey1.iPropVer = tgChfCT.iPropVer
                            tmChfSrchKey1.lCntrNo = tgChfCT.lCntrNo
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'set the key pointer to get previous version properly
                            If ilRet = BTRV_ERR_NONE Then
                                tmChfSrchKey1.iCntRevNo = tgChfCT.iCntRevNo - 1
                                tmChfSrchKey1.iPropVer = 32000
                                tmChfSrchKey1.lCntrNo = tgChfCT.lCntrNo
                                ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)       'get the previous version header
                                llContrCode = tmChf.lCode
                                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT(), False)      'get the entire contract
                            Else            'error on the read of previous version, exit this contract
                                Exit Do
                            End If
                        Else                'already version 0, no more to process; exit this contract
                            Exit Do
                        End If
                    Else                    'only process latest version, exit this contract
                        Exit Do
                    End If
                Loop            'while tgcntct.lcntrno = llCurrentCntrNo
            Next ilCurrentRecd
            
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmGrf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hlTxr)
            ilRet = btrClose(hlUrf)
            btrDestroy hmCff
            btrDestroy hmClf
            btrDestroy hmGrf
            btrDestroy hmCHF
            btrDestroy hlTxr
            btrDestroy hlUrf
            Exit Sub
End Sub

'TTP 10119 - Average 30 Rate Report - add option to export to CSV
Public Function mExportAvg30Rate(tlGrf As GRF, ilPeriods As Integer) As String
    mExportAvg30Rate = ""
    Dim slCSVString As String
    Dim slString As String
    Dim ilInt As Integer
    Dim illoop As Integer
    Dim llLong As Long
    Dim llDate As Long
    Dim iAvg30Spot As Integer
    'Dim tmPrfSrchKey As LONGKEY0
    'Dim slStamp As String
    'ReDim tlMnf(0 To 0) As MNF
    slCSVString = ""
    On Error GoTo ExportError
    '--------------
    '(Agency or Daypart or Daypart w/Overrides),Vehicle,Daypart,Advertiser,Product,Contract#,Ln,Len,30 " Unit Rate,30 " Unit Spots,Total,# Spots,Per1,[Per2],[Per3],[Per4],[Per5],[Per6],[Per7],[Per8],[Per9],[Per10],[Per11],[Per12],[Per13],[Per14]
    slCSVString = ""

    '--------------
    'Agency/Daypart/Daypart w/Overrides
    slCSVString = slCSVString & """" & Trim(tlGrf.sGenDesc) & """"
    '--------------
    'Vehicle
    ilInt = gBinarySearchVef(tlGrf.iVefCode)
    If ilInt = -1 Then
        slCSVString = slCSVString & ","
    Else
        slCSVString = slCSVString & ",""" & Trim(tgMVef(ilInt).sName) & """"
    End If
    
    '--------------
    'Advertiser
    ilInt = gBinarySearchAdf(tlGrf.iAdfCode)
    If ilInt = -1 Then
        slCSVString = slCSVString & ","
    Else
        slCSVString = slCSVString & ",""" & Trim(tgCommAdf(ilInt).sName) & """"
    End If
    
    '--------------
    'Product
    slCSVString = slCSVString & ",""" & Trim(tgChfCT.sProduct) & """"

    '--------------
    'Contract#
    slCSVString = slCSVString & "," & tgChfCT.lCntrNo
    
    '--------------
    'Ln
    slCSVString = slCSVString & "," & tlGrf.iSofCode
    
    '--------------
    'Len
    slCSVString = slCSVString & "," & tlGrf.iCode2 ' & Trim(tlGrf.sBktType)
    
    '--------------
    '30 " Unit Rate
    iAvg30Spot = mCalc30UnitSpots(tlGrf)
    If iAvg30Spot <> 0 Then
        slCSVString = slCSVString & ",$" & tlGrf.lDollars(0) / iAvg30Spot
    Else
        slCSVString = slCSVString & ",$0.00"
    End If

    '--------------
    '30 " Unit Spots
    slCSVString = slCSVString & "," & iAvg30Spot
    
    '--------------
    'Total
    slCSVString = slCSVString & ",$" & tlGrf.lDollars(0)
    
    '--------------
    '# Spots
    slCSVString = slCSVString & "," & tlGrf.iPerGenl(14)
    
    '--------------
    'Month1
    slCSVString = slCSVString & "," & tlGrf.iPerGenl(0)
    
    '--------------
    'Period2
    If ilPeriods >= 2 Then
        slCSVString = slCSVString & "," & tlGrf.iPerGenl(1)
    End If
    '--------------
    'Period3
    If ilPeriods >= 3 Then
        slCSVString = slCSVString & "," & tlGrf.iPerGenl(2)
    End If
    '--------------
    'Period4
    If ilPeriods >= 4 Then
        slCSVString = slCSVString & "," & tlGrf.iPerGenl(3)
    End If
    '--------------
    'Period5
    If ilPeriods >= 5 Then
        slCSVString = slCSVString & "," & tlGrf.iPerGenl(4)
    End If
    '--------------
    'Period6
    If ilPeriods >= 6 Then
        slCSVString = slCSVString & "," & tlGrf.iPerGenl(5)
    End If
    '--------------
    'Period7
    If ilPeriods >= 7 Then
        slCSVString = slCSVString & "," & tlGrf.iPerGenl(6)
    End If
    '--------------
    'Period8
    If ilPeriods >= 8 Then
        slCSVString = slCSVString & "," & tlGrf.iPerGenl(7)
    End If
    '--------------
    'Period9
    If ilPeriods >= 9 Then
        slCSVString = slCSVString & "," & tlGrf.iPerGenl(8)
    End If
    '--------------
    'Period10
    If ilPeriods >= 10 Then
        slCSVString = slCSVString & "," & tlGrf.iPerGenl(9)
    End If
    '--------------
    'Period11
    If ilPeriods >= 11 Then
        slCSVString = slCSVString & "," & tlGrf.iPerGenl(10)
    End If
    '--------------
    'Period12
    If ilPeriods >= 12 Then
        slCSVString = slCSVString & "," & tlGrf.iPerGenl(11)
    End If
    '--------------
    'Period13
    If ilPeriods >= 13 Then
        slCSVString = slCSVString & "," & tlGrf.iPerGenl(12)
    End If
    '--------------
    'Period14
    If ilPeriods >= 14 Then
        slCSVString = slCSVString & "," & tlGrf.iPerGenl(13)
    End If
    

    '--------------
    'Write File
    Print #hmExport, slCSVString
    'Show some status
    lmExportCount = lmExportCount + 1
    If lmExportCount / 10 - Int(lmExportCount / 10) = 0 Then
        'show every 10
        RptSelCt.lacExport.Caption = "Exported " & lmExportCount & " records..."
        RptSelCt.lacExport.Refresh
    End If
    mExportAvg30Rate = "Exported " & lmExportCount & " records..."
    Exit Function
    
ExportError:
    mExportAvg30Rate = "Error:" & err & "-" & Error(err)
End Function

Function mCalc30UnitSpots(tlGrf As GRF) As Integer
    Dim ilFound As Integer
    Dim slVPFSellout As String
    
    mCalc30UnitSpots = 0
    
    '@30UnitSpots=
    'if {VPF_Vehicle_Options.vpfSSellout} = "B" or {VPF_Vehicle_Options.vpfSSellout} = "T" then
    '(
    '   if Remainder ({GRF_Generic_Report.grfCode2},30 ) <> 0 then
    '   (
    '     if Remainder ({GRF_Generic_Report.grfCode2},30 ) < 15 then                                 // 0-14 seconds, make sure its 1 unit
    '        ( Round( ({GRF_Generic_Report.grfCode2} / 30)) +1) * {GRF_Generic_Report.grfPer15Genl}
    '     Else
    '         Round( ({GRF_Generic_Report.grfCode2} / 30)) * {GRF_Generic_Report.grfPer15Genl}       // 15 seconds it will round up one
    '    )
    ' Else
    '       Round({GRF_Generic_Report.grfCode2} / 30) * {GRF_Generic_Report.grfPer15Genl}            //get total 30" units
    ')
    'Else
    '(
    '    {GRF_Generic_Report.grfPer15Genl}
    ')
    
    ilFound = gVpfFindIndex(tlGrf.iVefCode)
    If ilFound < 0 Then
        
        mCalc30UnitSpots = tlGrf.iPerGenl(15)
    Else
        slVPFSellout = tgVpf(ilFound).sSSellOut
        If slVPFSellout = "B" Or slVPFSellout = "T" Then
            If tlGrf.iCode2 Mod 30 <> 0 Then
                If tlGrf.iCode2 Mod 30 < 15 Then
                    mCalc30UnitSpots = (Round((tlGrf.iCode2 / 30), 0) + 1) * tlGrf.iPerGenl(14)
                Else
                    mCalc30UnitSpots = Round((tlGrf.iCode2 / 30), 0) * tlGrf.iPerGenl(14)
                End If
            Else
                mCalc30UnitSpots = Round((tlGrf.iCode2 / 30), 0) * tlGrf.iPerGenl(14)
            End If
        Else
            mCalc30UnitSpots = tlGrf.iPerGenl(15)
        End If
    End If
    

End Function
