Attribute VB_Name = "RptcrNT"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrnt.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmItfSrchKey                  imItfRecLen                   tmItf                     *
'*  tmVsf                         imVsfRecLen                                             *
'******************************************************************************************

Option Explicit
Option Compare Text

'Agency for commission %
Dim hmAgf As Integer            'Agency file handle
Dim tmAgfSrchKey As INTKEY0      'AGF record image
Dim imAgfRecLen As Integer      'AGF record length
Dim tmAgf As AGF

'contract header
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0    'CHF record image
Dim tmChfSrchKey1 As CHFKEY1    'CHF record image
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChf As CHF

'IHF Inventory header
Dim hmIhf As Integer            'Inv header file handle
Dim tmIhfSrchKey As INTKEY0
Dim imIhfRecLen As Integer      'IHF record length
Dim tmIhf As IHF

'IHF Inventory Type
Dim hmItf As Integer            'Inv type file handle

'  Receivables File
Dim hmRvf As Integer            'receivables file handle
Dim tmRvf As RVF
Dim imRvfRecLen As Integer      'RVF record length

'  SBF (NTR) File
Dim hmSbf As Integer            'Special Billing file handle
Dim tmSbf As SBF
Dim tmSbfSrchKey1 As LONGKEY0    'SBF record image
Dim imSbfRecLen As Integer      'SBF record length

'  SLF File
Dim hmSlf As Integer            'Slsp file handle for Sales Source
Dim tmSlf As SLF
Dim imSlfRecLen As Integer      'SLF record length

'  SOF File
Dim hmSof As Integer            'Sales Office file handle for Sales Source
Dim tmSof As SOF
Dim imSofRecLen As Integer      'SOF record length

'  Vehicle File
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF
Dim tmVefSrchKey As INTKEY0     'VEF key record image
Dim imVefRecLen As Integer      'VEF record length

'  General Report File
Dim hmGrf As Integer            'Generic prepass file handle
Dim tmGrf As GRF
Dim imGrfRecLen As Integer      'GRF record length

'  Vehicle Slsp
Dim hmVsf As Integer            'Vehicle Slsp file handle

Dim lmLastBilled As Long                    'last Billed date
Dim imLastBilledInx As Integer              'index to date array of last month invoiced
'Dim lmStartDates(1 To 13) As Long      'start dates for 13 months
Dim lmStartDates(0 To 13) As Long      'start dates for 13 months. Index zero ignored
Dim imSortBy As Integer         'Sort by 0=adv, 1=agy, 2=ntr types, 3=owner, 4=slsp, 5=vehicles, 6=bill date
Dim imSortbyMinor As Integer    'for NTR B & B and NTR Multimedia : 'Sort by 0 = None, 1=adv, 2=agy, 3=ntr or multimedia types, 4=owner, 5=slsp, 6=vehicles
Dim imMinorVG As Integer      'minor vehicle group selected
Dim imMajorVG As Integer      'major vehicle group selected
Dim tmSofList() As SOFLIST
Dim tmSlfList() As SLFList
Dim smGrossNet As String * 1
Dim tmChfAdvtExt() As CHFADVTEXT    'array of contracts within the requested period to report
Dim lmContract As Long              'selective contr # for Billed & booked

Dim bmSplitSlsp As Boolean          '9-24-19 true to split slsp revenue

Dim tmMnf() As MNF          'sales sources to determine if participants are split at invoicing
Dim tmNTRMNF() As MNF       '3-17-05  NTR types to determine if Hard Cost item
Type SLFList
    iSlfCode As Integer     'salesperson code
    iSofCode As Integer     'office code
    iMnfSSCode As Integer       'sales source code
End Type
Dim tmPifKey() As PIFKEY          'array of vehicle codes and start/end indices pointing to the participant percentages
                                        'i.e Vehicle XYZ has 2 sales sources, each with 3 participants.  That will be a total of
                                        '6 entries.  Vehicle XYZ points to lo index equal to 1, and a hi index equal to 6; the
                                        'next vehicle will be a lo index of 7, etc.
Dim tmPifPct() As PIFPCT          'all vehicles and all percentages from PIF

'           mNoSplitRVF - calculate the gross and/or net $ for cash and/or trade category.
'           This routine is for all past (from receivables) transactions that have already been
'           split by participant in Invoicing (determined by Sales Source : Ask by Vehicle)
'           Write GRF record for each NTR record
'
'           <Input> ilMatchSSCode = Sales Source code
'                   slGross - Gross $ in string
'                   slNet - Net $ in string
'                   ilMonth - Month (1-12) to add $
Public Sub mNoSplitRVF(ilMatchSSCode As Integer, slGross As String, slNet As String, ilMonthNo As Integer)
'       GRF items used:
'       grfGenTime - Generation Time
'       grfGenDate - generation date
'       grfvefCode - vehicle code
'       grfChfcode - contract code
'       grfDateType - C = cash, T = trade
'       grfSofCode - Sales Source
'       grfrdfCode = item type code (itf)
'       GrfCode4 - SBF code
'       GrfCode2 - imSortBy
'       grfPerGenl(1) - Partipant (mnf code)
'       grfPerGenl(2) - NTR Type (mnf Code)
'       grfPerGenl(3) - Minor vehicle group code
'       grfPerGenl(4) - Major vehicle group code
'       grfPer(1-12) - Monthly $
    Dim ilHowMany As Integer
    Dim ilHowManyDefined As Integer
    Dim illoop As Integer
    Dim ilTemp As Integer
    Dim llProcessPct As Long
    Dim slDollar As String
    Dim llTransGross As Long
    Dim llTransNet As Long
    Dim slPct As String
    Dim llNetDollar As Long
    Dim llGrossDollar As Long
    Dim ilRet As Integer
    Dim ilRound As Integer
    'ReDim ilProdPct(1 To 1) As Integer            '5-22-07
    'ReDim ilMnfGroup(1 To 1) As Integer           '5-22-07
    'ReDim ilMnfSSCode(1 To 1) As Integer          '5-22-07
    'Index zero ignored in arrays below
    ReDim ilProdPct(0 To 1) As Integer            '5-22-07
    ReDim ilMnfGroup(0 To 1) As Integer           '5-22-07
    ReDim ilMnfSSCode(0 To 1) As Integer          '5-22-07
    Dim ilUse100pct As Integer                    '8-21-07 use 100% for participant split, do not look it

    llTransGross = gStrDecToLong(slGross, 2)
    llTransNet = gStrDecToLong(slNet, 2)
    ilUse100pct = True                      '8-21-07 if group exists in recv, set it as 100%
    gInitPartGroupAndPcts tmRvf.iAirVefCode, ilMatchSSCode, tmRvf.iMnfGroup, ilMnfSSCode(), ilMnfGroup(), ilProdPct(), tmRvf.iTranDate(), tmPifKey(), tmPifPct(), ilUse100pct

    'For Advt & Vehicle, create 1 record per transaction
    'For Owner (participant) , create as many as 3 records as there are particpants.
    mGetHowMany ilMatchSSCode, ilHowMany, ilHowManyDefined, ilMnfSSCode(), 0
    'For ilTemp = 1 To 18 Step 1               'init the years $ buckets
    For ilTemp = LBound(tmGrf.lDollars) To UBound(tmGrf.lDollars) Step 1               'init the years $ buckets
        tmGrf.lDollars(ilTemp) = 0
    Next ilTemp

    llProcessPct = 1000000          'everything is 100% since trans is already split by invoicing
    slPct = gLongToStrDec(llProcessPct, 4)
    slDollar = gMulStr(slPct, slGross)                 'slsp gross portion of possible split
    llGrossDollar = Val(gRoundStr(slDollar, "01.", 0))
    llTransGross = llTransGross - llGrossDollar
    If illoop = ilHowManyDefined - 1 Then      'last slsp or participant processed?  Handle
                                                 'extra pennies.
        llGrossDollar = llGrossDollar + llTransGross    'last recd gets left over pennies
    End If

    slDollar = gMulStr(slPct, slNet)                 'slsp net portion of possible split
    llNetDollar = Val(gRoundStr(slDollar, "01.", 0))
    llTransNet = llTransNet - llNetDollar
    If illoop = ilHowManyDefined - 1 Then      'last slsp or participant processed?  Handle
                                                 'extra pennies.
        llNetDollar = llNetDollar + llTransNet    'last recd gets left over pennies
    End If
    'bucket 1-12 contains slsp comm amount calc from his gross allocation
    If llNetDollar <> 0 Then    'did net amt have any value? if not, dont write out record
        'gGetVehGrpSets tmGrf.iVefCode, imMinorVG, imMajorVG, tmGrf.iPerGenl(3), tmGrf.iPerGenl(4)   'Genl(3) = minor sort code, genl(4) = major sort code
        gGetVehGrpSets tmGrf.iVefCode, imMinorVG, imMajorVG, tmGrf.iPerGenl(2), tmGrf.iPerGenl(3)   'Genl(3) = minor sort code, genl(4) = major sort code
        If imMinorVG = 0 Then
            'tmGrf.iPerGenl(3) = -1
            tmGrf.iPerGenl(2) = -1
        End If
        If imMajorVG = 0 Then
            'tmGrf.iPerGenl(4) = -1
            tmGrf.iPerGenl(3) = -1
        End If

         'round the values
        ilRound = 50
        If llGrossDollar < 0 Then
            ilRound = -50
        End If 'round the values
        ilRound = 50
        If llGrossDollar < 0 Then
            ilRound = -50
        End If
        If smGrossNet = "G" Then
            tmGrf.lDollars(ilMonthNo - 1) = ((llGrossDollar + ilRound) \ 100) * 100
        ElseIf smGrossNet = "N" Then  '2-26-01 Net or Net-Net option
            tmGrf.lDollars(ilMonthNo - 1) = ((llNetDollar + ilRound) \ 100) * 100
        End If
        tmGrf.iSofCode = ilMatchSSCode          'sales source
        'format remainder of record
        tmGrf.lChfCode = tmChf.lCode          'contract Code
        tmGrf.iVefCode = tmSbf.iAirVefCode
        'tmGrf.iPerGenl(1) = tmRvf.iMnfGroup 'participant
        'tmGrf.iPerGenl(2) = tmSbf.iMnfItem      'tran type
        tmGrf.iPerGenl(0) = tmRvf.iMnfGroup 'participant
        tmGrf.iPerGenl(1) = tmSbf.iMnfItem      'tran type
        tmGrf.iAdfCode = tmChf.iAdfCode
        tmGrf.iSlfCode = tmChf.iSlfCode(0)
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
     End If
End Sub

'                   NTR Billed & Booked - determine how many slsp to split $,
'                       or how many participants (owners to split $)
'
'                   mGetHowMany ilMatchSsCode, ilHowMany, ilHowManyDefined
'                   <input> ilMatchSSCode - valid only for owner option
'                   <output> ilHowMany - Total times to loop for splits
'                            ilHowManyDefined - # of elements having values to split
'                            (i.e. There may be only 2 slsp defined, but they are not, and the last element gets extra pennies
'                            in sequence in the arrays - element 1 used, element 2 unused, element 3 used)
'                            ilMnfSSCode() - array of the matching SS Entries
'                            ilMnfGroup() - array of participants with matching sales source
'                            ilProdPct() - array of particpant % with matching sales source
'                           ilMajorOrMinor - 0 = use major sort, 1 = use minor sort
'
'Sub mGetHowMany(ilMatchSSCode As Integer, ilHowMany As Integer, ilHowManyDefined As Integer, ilMnfSSCode() As Integer, ilMnfGroup() As Integer, ilProdPct() As Integer, ilMajorOrMinor As Integer)
Sub mGetHowMany(ilMatchSSCode As Integer, ilHowMany As Integer, ilHowManyDefined As Integer, ilMnfSSCode() As Integer, ilMajorOrMinor As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************
    Dim ilRet As Integer
    Dim illoop As Integer
    ilHowMany = 0
    ilHowManyDefined = 0
    
    If (ilMajorOrMinor = 0 And imSortBy = 3) Or (ilMajorOrMinor = 1 And imSortbyMinor = 4) Then
    '        If imSortBy = 3 Then                          'owner, get the vehicle groups associated with the vehicle
        'get the vehicle for this transaction
        If tmVef.iCode <> tmSbf.iAirVefCode Then
            tmVefSrchKey.iCode = tmSbf.iAirVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        End If
    
        ilHowMany = UBound(ilMnfSSCode)
        ilHowManyDefined = UBound(ilMnfSSCode)
    ElseIf (bmSplitSlsp) And ((ilMajorOrMinor = 0 And imSortBy = 4) Or (ilMajorOrMinor = 1 And imSortbyMinor = 5)) Then                  '9-24-19slsp, split  them?
        For illoop = 0 To 9 Step 1
            If tmChf.iSlfCode(illoop) > 0 Then
                ilHowMany = ilHowMany + 1
            End If
        Next illoop
        If ilHowMany = 1 And tmChf.lComm(0) = 0 Then    'theres only 1 slsp  w/o comm defined (force 100%)
            tmChf.lComm(0) = 1000000                'xxx.xxxx
        Else
            If ilHowMany > 1 Then
                ilHowMany = 10                      'more than 1 slsp, process all 10 because some in the
                                                'middle may not be used
            End If
        End If
        ilHowManyDefined = ilHowMany            'these % do not have to add up to 100%, so the actual
                                                '#  of slsp to process can be the same, whereas reporting
                                                'by owner, all $ must be accounted for and balance to the penny
    Else
        ilHowMany = 1
        ilHowManyDefined = 1
    End If
End Sub

'               mObtainCodes - get all codes to process or exclude
'               When selecting advt, agy or vehicles--make testing
'               of selection more efficient.  If more than half of
'               the entries are selected, create an array with entries
'               to exclude.  If less than half of entries are selected,
'               create an array with entries to include.
'               <input> ilSortBy - list box to test
'                       lbcListbox - array containing sort codes
'               <output> ilIncludeCodes - true if test to include the codes in array
'                                          false if test to exclude the codes in array
'                        ilUseCodes - array of advt, agy or vehicles codes to include/exclude
Sub mObtainCodes(ilSortBy As Integer, lbcListBox() As SORTCODE, ilIncludeCodes, ilUseCodes() As Integer)
    Dim ilHowManyDefined As Integer
    Dim ilHowMany As Integer
    Dim slNameCode As String
    Dim illoop As Integer
    Dim slCode As String
    Dim ilRet As Integer
    ilHowManyDefined = RptSelNT!lbcSelection(ilSortBy).ListCount
    'ilHowMany = RptSel!lbcSelection(ilSortBy).SelectCount
    ilHowMany = RptSelNT!lbcSelection(ilSortBy).SelCount
    If ilHowMany > ilHowManyDefined / 2 Then    'more than half selected
        ilIncludeCodes = False
    Else
        ilIncludeCodes = True
    End If
    For illoop = 0 To RptSelNT!lbcSelection(ilSortBy).ListCount - 1 Step 1
        slNameCode = lbcListBox(illoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If RptSelNT!lbcSelection(ilSortBy).Selected(illoop) And ilIncludeCodes Then               'selected ?
            ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
            ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
        Else        'exclude these
            If (Not RptSelNT!lbcSelection(ilSortBy).Selected(illoop)) And (Not ilIncludeCodes) Then
                ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
                ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
            End If
        End If
    Next illoop
End Sub

'       Determine if the SBF record is one to be printed or filter out
'
'       <input> ilListIndex = report type (NTR Recap or BB)
'               tlSbf - NTR record to test
'               ilincludeCodes - flag to indicate whether to include or exclude the codes sent
'               ilusecodes() - array of codes to either test for inclusion or exclusion
'               blTestDelete - optional flag to indicate to ignore testing the delete flag in sbf.  Coming from rvf/phf could be pointer to a contract that
'                               has revision, so the delete flag will be set
'
'       <Return> - true if OK
'
'       9-1-04 exclude proposals in NTR recap (include cntr statuses HOGN)
Function mFilterSBF(ilListIndex As Integer, tlSbf As SBF, ilIncludeCodes As Integer, ilUseCodes() As Integer, ilIncludeCodesMinor As Integer, ilUseCodesMinor() As Integer, Optional blTestDelete As Boolean = True) As Integer
'imsortby: 0 = advt, 1 = agy, 2 = ntr or multimedia type, 3 = owner,
'          4 = salesperson, 5 = vehicle, 6 = bill date
    Dim ilTestValue As Integer
    Dim ilOwner As Integer
    Dim ilOk As Integer
    Dim ilTemp As Integer
    Dim ilOKToSeeCnt As Integer
    Dim ilRet As Integer

    mFilterSBF = False
    '12-6-06 changed to see if user allowed to see contract
    ilOKToSeeCnt = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode())
    tmGrf.iRdfCode = 0              'Intialize Item Type for multimedia in case of multimedia report
    'OK if user allowed to see cntr, cntr is not deleted, and contr is an Order (sch or unsch)
    '12-8-17 if coming from receivables or history, do not check for a deleted NTR .  The contract may have been revised and the phf/rvf do have have any numbers swapped.
    'add an optional parameter when only coming from phf/rvf that the delete flag test be ignored
'    If (ilOKToSeeCnt) And (tmChf.sDelete = "N") And (tmChf.sStatus = "H" Or tmChf.sStatus = "O" Or tmChf.sStatus = "G" Or tmChf.sStatus = "N") Then             'this header is a current revision
     If (ilOKToSeeCnt) And (tmChf.sStatus = "H" Or tmChf.sStatus = "O" Or tmChf.sStatus = "G" Or tmChf.sStatus = "N") And ((tmChf.sDelete = "N" And blTestDelete = True) Or (blTestDelete = False)) Then               'this header is a current revision
        'if ntr recap and ntr date option, or not ntr date option and either NTR recap or NTR B & B reports with the ALL box selected
        'imsortby = 6 is option by bill date

        If (ilListIndex = CNT_NTRRECAP And imSortBy = 6) Or (imSortBy <> 6 And RptSelNT!ckcAll.Value = vbChecked) Then
            'if multimedia report, can only get the multimedia records,not the NTR records from SBF
            If ilListIndex = CNT_MULTIMBB Then
                'obtain the Inventory header to retrieve which type this NTR is assigned to
                tmIhfSrchKey.iCode = tmSbf.iIhfCode
                ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    mFilterSBF = False                   'invalid type
                    tmGrf.iRdfCode = 0
                Else
                    mFilterSBF = True
                    tmGrf.iRdfCode = tmIhf.iCode
                End If
            Else
                mFilterSBF = True
            End If
        Else

            If imSortBy = 0 Then             'Major adv
                ilTestValue = tmChf.iAdfCode
            ElseIf imSortBy = 1 Then         'agy
                ilTestValue = tmChf.iAgfCode
            ElseIf imSortBy = 2 Then         'ntr or multimedia types
                ilTestValue = tmSbf.iMnfItem
                If ilListIndex = CNT_MULTIMBB Then      '1-25-08
                    'obtain the Inventory header to retrieve which type this NTR is assigned to
                    tmIhfSrchKey.iCode = tmSbf.iIhfCode
                    ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilTestValue = -1                    'invalid type
                        tmGrf.iRdfCode = 0
                    Else
                        ilTestValue = tmIhf.iItfCode
                        tmGrf.iRdfCode = tmIhf.iCode
                    End If

                End If
            ElseIf imSortBy = 3 Then         'owners
                ilOwner = gBinarySearchVef(tmSbf.iAirVefCode)
                If ilOwner <> -1 Then
                    ilTestValue = tgMVef(ilOwner).iMnfGroup(0)
                End If
            ElseIf imSortBy = 4 Then         'slsp
                ilTestValue = tmChf.iSlfCode(0)
            ElseIf imSortBy = 5 Then        'vehicles
                ilTestValue = tmSbf.iAirVefCode
            End If
            ilOk = False
            If ilIncludeCodes Then
                For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                    If ilUseCodes(ilTemp) = ilTestValue Then
                        ilOk = True
                        Exit For
                    End If
                Next ilTemp
            Else
                ilOk = True        ' when more than half selected, selection fixed
                For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                    If ilUseCodes(ilTemp) = ilTestValue Then
                        ilOk = False            '6-28-05
                        Exit For
                    End If
                Next ilTemp
            End If
            
            If imSortbyMinor > 0 Then
                If imSortbyMinor = 1 Then              'Major adv, Minor adv
                    ilTestValue = tmChf.iAdfCode
                ElseIf imSortbyMinor = 2 Then         'agy
                    ilTestValue = tmChf.iAgfCode
                ElseIf imSortbyMinor = 3 Then         'ntr or multimedia types
                    ilTestValue = tmSbf.iMnfItem
                    If ilListIndex = CNT_MULTIMBB Then      '1-25-08
                        'obtain the Inventory header to retrieve which type this NTR is assigned to
                        tmIhfSrchKey.iCode = tmSbf.iIhfCode
                        ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilTestValue = -1                    'invalid type
                            tmGrf.iRdfCode = 0
                        Else
                            ilTestValue = tmIhf.iItfCode
                            tmGrf.iRdfCode = tmIhf.iCode
                        End If
    
                    End If
                ElseIf imSortbyMinor = 4 Then        'owners
                    ilOwner = gBinarySearchVef(tmSbf.iAirVefCode)
                    If ilOwner <> -1 Then
                        ilTestValue = tgMVef(ilOwner).iMnfGroup(0)
                    End If
                ElseIf imSortbyMinor = 5 Then         'slsp
                    ilTestValue = tmChf.iSlfCode(0)
                ElseIf imSortbyMinor = 6 Then         'vehicles
                    ilTestValue = tmSbf.iAirVefCode
                End If

                If ilOk Then
                    ilOk = False
                    If ilIncludeCodesMinor Then
                        For ilTemp = LBound(ilUseCodesMinor) To UBound(ilUseCodesMinor) - 1 Step 1
                            If ilUseCodesMinor(ilTemp) = ilTestValue Then
                                ilOk = True
                                Exit For
                            End If
                        Next ilTemp
                    Else
                        ilOk = True        ' when more than half selected, selection fixed
                        For ilTemp = LBound(ilUseCodesMinor) To UBound(ilUseCodesMinor) - 1 Step 1
                            If ilUseCodesMinor(ilTemp) = ilTestValue Then
                                ilOk = False            '6-28-05
                                Exit For
                            End If
                        Next ilTemp
                    End If
                End If
            End If
            
            If ilOk Then
                If ilListIndex = CNT_MULTIMBB Then
                    'obtain the Inventory header to retrieve which type this NTR is assigned to
                    tmIhfSrchKey.iCode = tmSbf.iIhfCode
                    ilRet = btrGetEqual(hmIhf, tmIhf, imIhfRecLen, tmIhfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilOk = False                  'invalid type
                        tmGrf.iRdfCode = 0
                        Exit Function
                    Else
                        ilOk = True
                        tmGrf.iRdfCode = tmIhf.iCode
                    End If
                End If
            End If
            mFilterSBF = ilOk
        End If
    End If
End Function

'       Create NTR reports: 1) NTR Recap and 2) NTR Billed & Booked
'       <input>  ilListIndex = report type (0 = NTR Recap, 1 = NTR B & B)
'                ilSortBy = Sort by 0=adv, 1=agy, 2=ntr or multimedia types, 3=owner, 4=slsp, 5=vehicles, 6=bill date
'
'   The Recap reports lists NTR items from SBF based on dates, and basically dumps the information
'       by selectivity of bill dates, advertiser, ntr types, salesperson, & vehicle
'   Billed & Booked shows revenue for a 12 month period by advertiser, agency NTR type,
'      owner, salesperson or vehicle.  Corporate & standard months are options for gathering$
'
'       4-3-03
'       2-11-05 Chg array of PHF/RVF array from integer to long (prevent overflow)
'       3-17-05 Option to Include Hard costs
'       12-14-06 add parm to gObtainRvfPhf to test on tran date (vs entry date)
Sub gCreateNTR(ilListIndex As Integer, ilSortBy As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilNTRLoop                                                                             *
'******************************************************************************************
    Dim ilRet As Integer
    Dim slStart As String
    Dim slEnd As String
    'TTP 10855 - NTR Recap: overflow due to too many NTR items
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim ilOk As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slStr As String
    Dim tlSBFTypes As SBFTypes
    Dim ilError As Integer
    'ReDim ilUseCodes(1 To 1) As Integer       'valid advt, agency or vehicles codes to process--
    ReDim ilUseCodes(0 To 0) As Integer       'valid advt, agency or vehicles codes to process--
                                                'or advt, agy or vehicles codes not to process
    Dim ilIncludeCodes As Integer
    ReDim ilUseCodesMinor(0 To 0) As Integer    'ntr B & B or Ntr Multimedia codes to include or exclude
    Dim ilIncludeCodesMinor As Integer
     
    Dim ilStd As Integer
    Dim ilTemp As Integer
    Dim tlTranType As TRANTYPES
    Dim illoop As Integer
    ReDim tlRvf(0 To 0) As RVF
    ReDim tlSbf(0 To 0) As SBF
    Dim ilBilled As Integer
    Dim ilUnbilled As Integer
    Dim slStamp As String
    ReDim tmMnf(0 To 0) As MNF
    ReDim tmNTRMNF(0 To 0) As MNF       '3-17-05

    hmAgf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imAgfRecLen = Len(tmAgf)

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

    hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSbfRecLen = Len(tmSbf)

    hmGrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imGrfRecLen = Len(tmGrf)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imVefRecLen = Len(tmVef)

    If ilListIndex = CNT_MULTIMBB Then          'multimedia B & B
        hmIhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmIhf, "", sgDBPath & "Ihf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilError = ilRet
        End If
        imIhfRecLen = Len(tmIhf)
    End If

    tlSBFTypes.iNTR = True          'include NTR billing
    tlSBFTypes.iInstallment = False      'exclude Installment billing
    tlSBFTypes.iImport = False           'exclude rep import billing
    imSortBy = ilSortBy                 'put sort option in common place

    bmSplitSlsp = False             '9-24-19
    
    If (ilListIndex = CNT_NTRRECAP And imSortBy = 6) Then        'there are no list box selectivity for Bill date selection
    Else
        imSortbyMinor = RptSelNT!cbcSet2.ListIndex          '0 implies no sort/selection
        If imSortBy = 0 Then            'advt
            mObtainCodes 0, tgAdvertiser(), ilIncludeCodes, ilUseCodes()
        ElseIf imSortBy = 1 Then        'agy
            mObtainCodes 1, tgAgency(), ilIncludeCodes, ilUseCodes()
        ElseIf imSortBy = 2 Then         'NTR
            If ilListIndex = CNT_MULTIMBB Then
'                mObtainCodes 2, tgItfCode(), ilIncludeCodes, ilUseCodes()
                mObtainCodes 6, tgItfCode(), ilIncludeCodes, ilUseCodes()           '11-14-18 add new list box so each type of option has its own list box
            Else
                mObtainCodes 2, tgMnfCodeCT(), ilIncludeCodes, ilUseCodes()
            End If
        ElseIf imSortBy = 3 Then        'owner
'            mObtainCodes 1, tgSalesperson(), ilIncludeCodes, ilUseCodes()
            mObtainCodes 3, tgTmpSort(), ilIncludeCodes, ilUseCodes()
        ElseIf imSortBy = 4 Then        'salesperson
            mObtainCodes 4, tgSalesperson(), ilIncludeCodes, ilUseCodes()
        ElseIf imSortBy = 5 Then        'vehicle
            mObtainCodes 5, tgVehicle(), ilIncludeCodes, ilUseCodes()
        End If
        '10-8-19 only split when either major or minor has been selected for slsp to split
        If (RptSelNT!ckcMajorSplit.Value = vbChecked And imSortBy = 4) Or (RptSelNT!ckcMinorSplit.Value = vbChecked And imSortbyMinor = 5) Then
            bmSplitSlsp = True
        End If

        If imSortbyMinor > 0 Then           'no minor sort selected
            If imSortbyMinor = 1 Then            'advt
                mObtainCodes 0, tgAdvertiser(), ilIncludeCodesMinor, ilUseCodesMinor()
            ElseIf imSortbyMinor = 2 Then        'agy
                mObtainCodes 1, tgAgency(), ilIncludeCodesMinor, ilUseCodesMinor()
            ElseIf imSortbyMinor = 3 Then         'NTR
                If ilListIndex = CNT_MULTIMBB Then
    '                mObtainCodes 2, tgItfCode(), ilIncludeCodes, ilUseCodes()
                    mObtainCodes 6, tgItfCode(), ilIncludeCodesMinor, ilUseCodesMinor()           '11-14-18 add new list box so each type of option has its own list box
                Else
                    mObtainCodes 2, tgMnfCodeCT(), ilIncludeCodesMinor, ilUseCodesMinor()
                End If
            ElseIf imSortbyMinor = 4 Then        'owner
'                mObtainCodes 1, tgSalesperson(), ilIncludeCodesMinor, ilUseCodesMinor()
                mObtainCodes 3, tgTmpSort(), ilIncludeCodes, ilUseCodes()
            ElseIf imSortbyMinor = 5 Then        'salesperson
                mObtainCodes 4, tgSalesperson(), ilIncludeCodesMinor, ilUseCodesMinor()
            ElseIf imSortbyMinor = 6 Then        'vehicle
                mObtainCodes 5, tgVehicle(), ilIncludeCodesMinor, ilUseCodesMinor()
            End If
        End If
    End If

    If ilListIndex = CNT_MULTIMBB Then
        'get all the multimedia types
    Else
        ilRet = gObtainMnfForType("I", slStamp, tmNTRMNF())        'get the NTR Types to see if Hard cost item
    End If

    If ilListIndex = CNT_NTRRECAP Then
        ilBilled = False
        ilUnbilled = False
        If RptSelNT!rbcTotalsBy(0).Value = True Or RptSelNT!rbcTotalsBy(2).Value = True Then       'include billed
            ilBilled = True
        End If
        If RptSelNT!rbcTotalsBy(1).Value = True Or RptSelNT!rbcTotalsBy(2).Value = True Then       'include billed
            ilUnbilled = True
        End If
        '12-17-19 use csi calendar control vs edit box
'        slStr = RptSelNT!edcDate1.Text                'Earliest date to retrieve  SBF Items
        slStr = RptSelNT!CSI_CalFrom.Text                'Earliest date to retrieve  SBF Items

        llStartDate = gDateValue(slStr)
        slStart = Format$(llStartDate, "m/d/yy")
'        slStr = RptSelNT!edcDate2.Text               'Latest date to retrieve from SBF Items
        slStr = RptSelNT!CSI_CalTo.Text               'Latest date to retrieve from SBF Items

        llEndDate = gDateValue(slStr)
        slEnd = Format$(llEndDate, "m/d/yy")
        'gPackDate slStart, tmGrf.iDateGenl(0, 1), tmGrf.iDateGenl(1, 1)
        gPackDate slStart, tmGrf.iDateGenl(0, 0), tmGrf.iDateGenl(1, 0)
        'gPackDate slEnd, tmGrf.iDateGenl(0, 2), tmGrf.iDateGenl(1, 2)
        gPackDate slEnd, tmGrf.iDateGenl(0, 1), tmGrf.iDateGenl(1, 1)

        'Create table of the vehicles and their participants with share of split
        gCreatePIFForRpts llStartDate, tmPifKey(), tmPifPct(), RptSelPP

        tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
        tmGrf.iGenDate(1) = igNowDate(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrf.lGenTime = lgNowTime

        ilRet = gObtainSBF(RptSelNT, hmSbf, 0, slStart, slEnd, tlSBFTypes, tlSbf(), 0)      '11-28-06 add last parm to indicate which key to use
        For llSbf = 0 To UBound(tlSbf)
            tmSbf = tlSbf(llSbf)
            tmChfSrchKey.lCode = tmSbf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                ilOk = False
            Else
                ilOk = mFilterSBF(ilListIndex, tmSbf, ilIncludeCodes, ilUseCodes(), ilIncludeCodesMinor, ilUseCodesMinor())    'filter advt, agy, ntr types, owners, slsp, vehicles
            End If

            If ilOk And ((ilBilled And tmSbf.sBilled = "Y") Or (ilUnbilled And tmSbf.sBilled = "N")) Then
                '3-17-05 determine if Hard cost and should it be included
                If RptSelNT!ckcInclHardCost = vbUnchecked Then      'if hard cost not included, need to check the MNF NTR type
                    ilRet = gIsItHardCost(tmSbf.iMnfItem, tmNTRMNF())
                    If ilRet = True Then                'its a Hard cost item
                        ilOk = False
                    End If
                End If
                If ilOk Then
                    '1-14-11 If using acq cost, must be non-zero to include; otherwise show the NTR cost
                    If (RptSelNT!ckcUseAcqCost.Value = vbChecked And tmSbf.lAcquisitionCost > 0) Or (RptSelNT!ckcUseAcqCost.Value = vbUnchecked) Then
                        tmGrf.lCode4 = tmSbf.lCode      'SBF code
                        tmGrf.iCode2 = imSortBy                 'Sort by flag for Crystal report heading
                        gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), tmGrf.lLong 'convert date to a number for sorting purposes
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    End If
                End If
            End If
        Next llSbf
        Erase tlSbf
    Else                'NTR Billed & Booked or Multimedia B & B
        illoop = RptSelNT!cbcVehGroup.ListIndex     '6-13-02
        imMajorVG = gFindVehGroupInx(illoop, tgVehicleSets1())
        imMinorVG = 0       'unused but common to subrooutine that gets the vehicle grous

        slStr = RptSelNT!edcContract.Text            'selective contrct #, if entered
        lmContract = Val(slStr)
        If RptSelNT!rbcGrossNet(0).Value = True Then        'gross
            smGrossNet = "G"
        Else
            smGrossNet = "N"
        End If
        If RptSelNT!rbcBillBy(0).Value = True Then     'Corporate?
            ilStd = 1                           'False
        ElseIf RptSelNT!rbcBillBy(1).Value = True Then     '8-13-19 std
            ilStd = 2                           'True
        Else                                    'calendar
            ilStd = 3
        End If
        'setup transaction types to retrieve from history and receivables
        tlTranType.iAdj = True              'adjustments
        tlTranType.iInv = True              'invoices
        tlTranType.iWriteOff = False
        tlTranType.iPymt = False
        tlTranType.iCash = True
        tlTranType.iTrade = True
        tlTranType.iMerch = False
        tlTranType.iPromo = False
        tlTranType.iNTR = True

        'use the entered #periods, otherwise force to 12 for all other options
        igPeriods = Val(RptSelNT!edcPeriods.Text)
        If igPeriods < 1 Or igPeriods > 12 Then
            igPeriods = 12              'if errors in input, get normal report of 12 months
        End If

        'build array of selling office codes and their sales sources.
        ilTemp = 0
        ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            ReDim Preserve tmSofList(0 To ilTemp) As SOFLIST
            tmSofList(ilTemp).iSofCode = tmSof.iCode
            tmSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
            ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            ilTemp = ilTemp + 1
        Loop

        'Gather the sales sources to determine how to update the vehicle (RVF/PHF).
        'Required to determine whether to automatically split transactions by revenue share in some reports
        ilRet = gObtainMnfForType("S", slStamp, tmMnf())

        'build array of salespeople and their offices/sales sources.
        ilTemp = 0
        ilRet = btrGetFirst(hmSlf, tmSlf, imSlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            ReDim Preserve tmSlfList(0 To ilTemp) As SLFList
            tmSlfList(ilTemp).iSlfCode = tmSlf.iCode
            tmSlfList(ilTemp).iSofCode = tmSlf.iSofCode
            For illoop = LBound(tmSofList) To UBound(tmSofList)
                If tmSlfList(ilTemp).iSofCode = tmSofList(illoop).iSofCode Then
                    tmSlfList(ilTemp).iMnfSSCode = tmSofList(illoop).iMnfSSCode
                    Exit For
                End If
            Next illoop
            ilRet = btrGetNext(hmSlf, tmSlf, imSlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            ilTemp = ilTemp + 1
        Loop
        Erase tmSofList
        'vehicle group selected, this report only uses the major sort
        imMajorVG = 0
        imMinorVG = 0
        illoop = RptSelNT!cbcVehGroup.ListIndex
        imMajorVG = tgVehicleSets1(illoop).iCode

'        If Not ilStd Then  'Corporate
        If ilStd <> 2 Then          '8-13-19 corp or cal
            '11-21-05 add pacing date to call (n/a to this report)

            'ilTemp = (igMonthOrQtr - 1) * 3 + 1
            ilTemp = RptSelNT!edcDate2.Text '7-3-08 change to use start month vs start qtr
'            gSetupBOBDates 1, lmStartDates(), lmLastBilled, imLastBilledInx, 0, ilTemp   'build array of corp start & end dates
            gSetupBOBDates ilStd, lmStartDates(), lmLastBilled, imLastBilledInx, 0, ilTemp   '8-13-19 add calendar option; build array of corp start & end dates
            tlTranType.iInv = False          'for corporate reporting, all data needs to be obtained from contracts since only billing by std
            'setup dates to gather in the past (RVF/PHF)

            imLastBilledInx = 12
            For illoop = 1 To 12 Step 1
                If lmLastBilled > lmStartDates(illoop) And lmLastBilled < lmStartDates(illoop + 1) Then
                    imLastBilledInx = illoop
                    Exit For
                End If
            Next illoop

            slStart = Format$(lmStartDates(1), "m/d/yy")
            slEnd = Format$(lmStartDates(imLastBilledInx + 1) - 1, "m/d/yy")

            'Create table of the vehicles and their participants with share of split
            gCreatePIFForRpts lmStartDates(1), tmPifKey(), tmPifPct(), RptSelPP

            'all contracts from future, and only adjustments from past
            ilRet = gObtainPhfRvf(RptSelNT, slStart, slEnd, tlTranType, tlRvf(), 0)
            If ilRet = 0 Then
                Exit Sub
            End If

            mNTRPast tlRvf(), ilListIndex, ilIncludeCodes, ilUseCodes(), ilIncludeCodesMinor, ilUseCodesMinor()

            imLastBilledInx = 0         'all contracts are from future (regardless of last billed date)
            ilRet = mObtainCntrForNTR(tlSBFTypes, ilListIndex, ilIncludeCodes, ilUseCodes(), ilIncludeCodesMinor, ilUseCodesMinor())
            If ilRet <> 0 Then
                Erase tmSlfList
                Erase lmStartDates
                Erase tlRvf
                Erase tmChfAdvtExt
                Erase tlRvf
                Erase tlSbf
                Erase tmMnf
                Exit Sub
            End If


        Else
            '11-21-05 add pacing date to call (n/a to this report)
            'ilTemp = (igMonthOrQtr - 1) * 3 + 1
            ilTemp = RptSelNT!edcDate2.Text '7-3-08 change to use start month vs start qtr
            gSetupBOBDates 2, lmStartDates(), lmLastBilled, imLastBilledInx, 0, ilTemp  'build array of std start & end dates

            'Create table of the vehicles and their participants with share of split
            gCreatePIFForRpts lmStartDates(1), tmPifKey(), tmPifPct(), RptSelPP

            'setup dates to gather in the past (RVF/PHF)
            slStart = Format$(lmStartDates(1), "m/d/yy")
            slEnd = Format$(lmStartDates(igPeriods + 1) - 1, "m/d/yy")

            If lmStartDates(1) > lmLastBilled Then          'everything in future only
                imLastBilledInx = 0             'everything is in the future
                ilRet = mObtainCntrForNTR(tlSBFTypes, ilListIndex, ilIncludeCodes, ilUseCodes(), ilIncludeCodesMinor, ilUseCodesMinor())
                If ilRet <> 0 Then
                    Erase tmSlfList
                    Erase lmStartDates
                    Erase tlRvf
                    Erase tmChfAdvtExt
                    Erase tlRvf
                    Erase tlSbf
                    Erase tmMnf
                    Exit Sub
                End If
            Else
                ilRet = gObtainPhfRvf(RptSelNT, slStart, slEnd, tlTranType, tlRvf(), 0)
                If ilRet = 0 Then
                    Exit Sub
                End If
                mNTRPast tlRvf(), ilListIndex, ilIncludeCodes, ilUseCodes(), ilIncludeCodesMinor, ilUseCodesMinor()
                 If igPeriods > imLastBilledInx Then
                    ilRet = mObtainCntrForNTR(tlSBFTypes, ilListIndex, ilIncludeCodes, ilUseCodes(), ilIncludeCodesMinor, ilUseCodesMinor())
                    If ilRet <> 0 Then
                        Erase tmSlfList
                        Erase lmStartDates
                        Erase tlRvf
                        Erase tmChfAdvtExt
                        Erase tlRvf
                        Erase tlSbf
                        Erase tmMnf
                        Exit Sub
                    End If
                End If
            End If
        End If

        Erase tmSlfList
        Erase lmStartDates
        Erase tlRvf
        Erase tmChfAdvtExt
        Erase tlRvf
        Erase tlSbf
        Erase tmMnf
        Erase tmNTRMNF
        sgCntrForDateStamp = ""         'clear to avoid re-entrant problem to rerun report without exiting selection screen
    End If

    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmRvf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmItf)
    ilRet = btrClose(hmIhf)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmSlf)
    btrDestroy hmCHF
    btrDestroy hmSbf
    btrDestroy hmGrf
    btrDestroy hmRvf
    btrDestroy hmSof
    btrDestroy hmVef
    btrDestroy hmItf
    btrDestroy hmIhf
    btrDestroy hmAgf
    btrDestroy hmSlf
    Exit Sub
End Sub

'          obtain NTR transactions prior to the last billed date
'
'          <input> - tlRvf() - array of transactions in the past
'                   ilIncludeCodes - true to include codes, else false to exclude codes
'                   ilusecodes() - array of advt, agy, slsp, etc, to include/exclude
'
'
Sub mNTRPast(tlRvf() As RVF, ilListIndex As Integer, ilIncludeCodes As Integer, ilUseCodes() As Integer, ilIncludeCodesMinor As Integer, ilUseCodesMinor() As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCurrentRecd                                                                         *
'******************************************************************************************
    Dim slCode As String
    Dim llDate As Long
    Dim ilFound As Integer
    Dim llAmt As Long
    Dim slStr As String
    Dim ilFoundMonth As Integer
    Dim ilMonthNo As Integer
    Dim ilRet As Integer
    Dim ilMatchSSCode As Integer
    Dim ilOk As Integer
    Dim slGross As String
    Dim slNet As String
    Dim ilAskforUpdate As Integer
    Dim ilLoopAsk As Integer
    Dim llRvfLoop As Long                   '2-11-05
    Dim ilIsItHardCost As Integer

    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    For llRvfLoop = LBound(tlRvf) To UBound(tlRvf) - 1 Step 1
        tmRvf = tlRvf(llRvfLoop)

        gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slCode
        llDate = gDateValue(slCode)
        ilFound = False
        gPDNToLong tmRvf.sGross, llAmt

        'if Tran date is after the first requested date then this trans is a possiblity;
        'the presence of rvf.imnfItem indicates an NTR, and it must be a $ amount other than zero
        'look at non-installment (tmrvf.stype = "") or the installment revenue records (type = "A"); ignore the billing records (type = "I")
        If llDate >= lmStartDates(1) And llDate < lmStartDates(imLastBilledInx + 1) And (tmRvf.iMnfItem > 0 And llAmt <> 0) And (tmRvf.sType = "A" Or Trim$(tmRvf.sType) = "") Then
            ilFound = True

            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slStr   'always use entered date for Corp report
            llDate = gDateValue(slStr)                  'convert trans date to test if within requested limits

            'get contract from history or rec file
            If tmChf.lCntrNo <> tmRvf.lCntrNo Then
                tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd

                Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")
                     ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop

                'if no contract reference, fake out a contract header so $ will be included
                gFakeChf tmRvf, tmChf
            End If
            
            'TTP 10863 - NTR Billed and Booked: Add Political Selectivity
            ilOk = True
            If ilListIndex = CNT_NTRBB Then
                If mCheckAdvPolitical(tmChf.iAdfCode) Then          'its a political, include this contract?
                     If RptSelNT.ckcInclPolit.Value = vbUnchecked Then
                        ilOk = False
                    End If
                Else                                                'not a political advt, include this contract?
                     If RptSelNT.ckcInclNonPolit.Value = vbUnchecked Then
                        ilOk = False
                    End If
                End If
            End If
            
            If ilOk Then
                'obtain the SBF record
                tmSbfSrchKey1.lCode = tmRvf.lSbfCode
                ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then  'sbf doesnt exist, may be a history item
                    'fake out entry for filtering
                    tmSbf.iMnfItem = tmRvf.iMnfItem
                    tmSbf.iAirVefCode = tmRvf.iAirVefCode
                End If
                '12-8-17 add optional parameter to ignore testing of delete flag.  Recv/History may have a pointer to sbf whose contract has been modified, making the sbf deleted
                ilOk = mFilterSBF(ilListIndex, tmSbf, ilIncludeCodes, ilUseCodes(), ilIncludeCodesMinor, ilUseCodesMinor(), False)   'filter advt, agy, ntr types, owners, slsp, vehicles
            End If
            
            '3-21-05 determine if Hard cost and should it be included
            ilIsItHardCost = gIsItHardCost(tmSbf.iMnfItem, tmNTRMNF())
            If RptSelNT!ckcInclHardCost = vbUnchecked Then      'if hard cost not included, need to check the MNF NTR type
                If ilIsItHardCost = True Then                'its a Hard cost item
                    ilOk = False
                End If
            End If

            If ilListIndex = CNT_MULTIMBB And tmSbf.iIhfCode <= 0 Then      'if multimedia b & B, it has to be an inventory item
                ilOk = False       'regular NTR item, ignore
            End If

            ilMatchSSCode = mGetSSCode()        'get the contracts sales source

            'determine the month that this transaction falls within
            ilFoundMonth = False
            For ilMonthNo = 1 To igPeriods + 1 Step 1       'loop thru months to find the match
                If llDate >= lmStartDates(ilMonthNo) And llDate < lmStartDates(ilMonthNo + 1) Then
                    ilFoundMonth = True
                    Exit For
                End If
            Next ilMonthNo

            If (ilFoundMonth) And (ilOk) And ((lmContract = tmChf.lCntrNo And lmContract <> 0) Or (lmContract = 0)) Then

                gPDNToStr tmRvf.sGross, 2, slGross
                gPDNToStr tmRvf.sNet, 2, slNet

                If ilIsItHardCost = True Then               'its a hard cost item, force to the end of report
                    tmGrf.sDateType = "Z"
                Else
                    tmGrf.sDateType = tmRvf.sCashTrade          'Type field used for C = Cash, T = Trade
                End If
                tmGrf.iVefCode = tmRvf.iAirVefCode
                tmGrf.iSofCode = ilMatchSSCode              'Sales Source code
                tmGrf.lCode4 = tmRvf.lCntrNo                'contract # incase no contr header exists
                'determine update method
                ilAskforUpdate = False      'assume to split the revenue share, everything goes into RVF
                For ilLoopAsk = LBound(tmMnf) To UBound(tmMnf) - 1
                    If tmMnf(ilLoopAsk).iCode = ilMatchSSCode Then
                        If Trim$(tmMnf(ilLoopAsk).sUnitType) = "A" Then
                            ilAskforUpdate = True       'dont split any revenue by participants, invoicing has already done that
                        End If
                        Exit For
                    End If
                Next ilLoopAsk

                If ilAskforUpdate Then      'trans has been split, dont split if option requires
'                    mNoSplitRVF ilMatchSSCode, slGross, slNet, ilMonthNo
                    mProcessMonth ilMatchSSCode, slGross, slNet, ilMonthNo, True
                Else
                    mProcessMonth ilMatchSSCode, slGross, slNet, ilMonthNo, False
                End If
            End If                                      'foundmonth
        End If                          'llDate >= llStdStartDates(1) And llDate < llStdStartDates(13)
    Next llRvfLoop                  'llRvfLoop = LBound(tlRvf) To UBound(tlRvf)
End Sub

'           mObtainCntrForNTR - obtain contracts and looked for NTR defined only
'           Gather all NTR data and create a Billed & Booked report
'
'       <INPUT> tlSBFTypes - included ntrs, installments, or import rep  from SBF
'               ilListIndex - 0 = NTR Recap, 1 = NTR Billed & Booked
'               ilIncludeCodes - includes array of selectivity codes (or exclude them)
'               ilUseCodes - array of selectivity codes to include or exclude
'      Return - true if contracts process

' imSortBy:  Sort by 0=adv, 1=agy, 2=ntr types, 3=owner, 4=slsp, 5=vehicles, 6=bill date
'  imSortByMinor (NTR B & B or NTR Multimedia):Sort by 0 = None, 1=adv, 2=agy, 3=ntr types, 4=owner, 5=slsp, 6=vehicles
Public Function mObtainCntrForNTR(tlSBFType As SBFTypes, ilListIndex As Integer, ilIncludeCodes As Integer, ilUseCodes() As Integer, ilIncludeCodesMinor As Integer, ilUseCodesMinor() As Integer)
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slCntrStatus As String
    Dim slCntrType As String
    Dim ilHOState As Integer
    Dim ilCurrentRecd As Integer
    ReDim tlSbf(0 To 0) As SBF
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim ilRet As Integer
    Dim ilOk As Integer
    Dim ilFoundMonth As Integer
    Dim ilMonthNo As Integer
    Dim slGross As String
    Dim slNet As String
    Dim ilMatchSSCode As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim llAmt As Long
    Dim slAgyCommPct As String
    Dim ilCorT As Integer           'loop for cash or trade
    Dim ilStartCorT As Integer
    Dim ilEndCorT As Integer
    Dim slPctTrade As String
    Dim slGrossPct As String
    Dim ilIsItHardCost As Integer

    mObtainCntrForNTR = 0
    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime

    'build table (into tlchfadvtext) of all contracts that fall within the dates required
    slCntrStatus = "HOGN"       'holds/orders, sch/unsch
    slCntrType = ""             'all types
    ilHOState = 2               'latest version
    slStartDate = Format$(lmStartDates(imLastBilledInx + 1), "m/d/yy")
    slEndDate = Format$((lmStartDates(igPeriods + 1)) - 1, "m/d/yy")  'get the start date of the last period to include, but
                                                                            'backit up by 1 day to get the end date of the previous period
    If lmContract > 0 Then
        tmChfSrchKey1.lCntrNo = lmContract
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd

        Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")
             ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If tmChf.lCntrNo = lmContract Then
            ReDim tmChfAdvtExt(0 To 1) As CHFADVTEXT
            tmChfAdvtExt(0).lCode = tmChf.lCode
        Else
            ReDim tmChfAdvtExt(0 To 0) As CHFADVTEXT
        End If
    Else
        ilRet = gObtainCntrForDate(RptSelNT, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOState, tmChfAdvtExt())
        If ilRet <> BTRV_ERR_NONE Then
            mObtainCntrForNTR = ilRet
            Exit Function
        End If
    End If

    For ilCurrentRecd = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1
        If lmContract = 0 Then          'single contract, already in memory
            tmChfSrchKey.lCode = tmChfAdvtExt(ilCurrentRecd).lCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                Exit Function
            End If
        End If
        
        'TTP 10863 - NTR Billed and Booked: Add Political Selectivity
        ilOk = True
        If ilListIndex = CNT_NTRBB Then
            If mCheckAdvPolitical(tmChf.iAdfCode) Then          'its a political, include this contract?
                 If RptSelNT.ckcInclPolit.Value = vbUnchecked Then
                    ilOk = False
                End If
            Else                                                'not a political advt, include this contract?
                 If RptSelNT.ckcInclNonPolit.Value = vbUnchecked Then
                    ilOk = False
                End If
            End If
        End If
            
        'If tmChf.sNTRDefined = "Y" Then          'thishas NTR billing
        If tmChf.sNTRDefined = "Y" And ilOk = True Then          'thishas NTR billing
            ilRet = gObtainSBF(RptSelNT, hmSbf, tmChf.lCode, slStartDate, slEndDate, tlSBFType, tlSbf(), 0) '11-28-06 add last parm to indicate which key to use
            For llSbf = LBound(tlSbf) To UBound(tlSbf) - 1
                tmSbf = tlSbf(llSbf)
                tmChfSrchKey.lCode = tmSbf.lChfCode
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Or ((lmContract <> tmChf.lCntrNo And lmContract <> 0) And (lmContract > 0)) Then
                    ilOk = False
                Else
                    ilOk = mFilterSBF(ilListIndex, tmSbf, ilIncludeCodes, ilUseCodes(), ilIncludeCodesMinor, ilUseCodesMinor())      'filter advt, agy, ntr types, owners, slsp, vehicles
                    'determine if this ntr is agy commissionable
                    tmAgf.iComm = 0         'default unless not direct account
                    If tmChf.iAgfCode > 0 Then
                        tmAgfSrchKey.iCode = tmChf.iAgfCode
                        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilOk = False
                        End If
                    End If

                End If

                '3-21-05 determine if Hard cost and should it be included
                ilIsItHardCost = gIsItHardCost(tmSbf.iMnfItem, tmNTRMNF())
                If RptSelNT!ckcInclHardCost = vbUnchecked Then      'if hard cost not included, need to check the MNF NTR type
                    If ilIsItHardCost = True Then                'its a Hard cost item
                        ilOk = False
                    End If
                End If

                If ilOk Then        'passed user selectivity (advt, agy, slsp, owner, vehicle)
                    ilMatchSSCode = mGetSSCode()        'get the contracts sales source
                    'send the gross & Net $

                    ilFoundMonth = False
                    gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), slDate
                    llDate = gDateValue(slDate)
                    For ilMonthNo = 1 To igPeriods + 1 Step 1       'loop thru months to find the match
                        If llDate >= lmStartDates(ilMonthNo) And llDate < lmStartDates(ilMonthNo + 1) Then
                            ilFoundMonth = True
                            Exit For
                        End If
                    Next ilMonthNo

                    If ilFoundMonth And ilOk Then
                        slPctTrade = gIntToStrDec(tmChf.iPctTrade, 0)
                        If tmChf.iPctTrade = 0 Then                     'setup loop to do cash & trade
                            ilStartCorT = 1
                            ilEndCorT = 1
                        ElseIf tmChf.iPctTrade = 100 Then
                            ilStartCorT = 2
                            ilEndCorT = 2
                        Else
                            ilStartCorT = 1
                            ilEndCorT = 2
                        End If

                        slAgyCommPct = gIntToStrDec(tmAgf.iComm, 2)
                        'determine agency commission
                        If tmSbf.sAgyComm = "N" Then
                            slAgyCommPct = ".00"
                        End If

                        For ilCorT = ilStartCorT To ilEndCorT
                            If ilCorT = 1 Then              'cash
                                slPctTrade = gSubStr("100.", gIntToStrDec(tmChf.iPctTrade, 0))
                                tmGrf.sDateType = "C"          'Type field used for C = Cash, T = Trade
                                'if hard cost item,force to different type
                                If ilIsItHardCost = True Then
                                    tmGrf.sDateType = "Z"       'sort to the end (after Cash & Trade)
                                End If
                            Else            'trade portion
                                slPctTrade = gIntToStrDec(tmChf.iPctTrade, 0)
                                tmGrf.sDateType = "T"          'Type field used for C = Cash, T = Trade
                            End If
                            'convert the $ to gross & net strings
                            llAmt = tmSbf.lGross * tmSbf.iNoItems
                            slGross = gLongToStrDec(llAmt, 2)       'convert to xxxx.xx
                            'determine agency commission
                            slAgyCommPct = gIntToStrDec(tmAgf.iComm, 2)
                            If tmSbf.sAgyComm = "N" Then
                                slAgyCommPct = ".00"
                            End If
                            slGrossPct = gSubStr("100.00", slAgyCommPct)        'determine  % to client (normally 85%)
                            slNet = gDivStr(gMulStr(slGrossPct, slGross), "100")    'net value

                            'calculate the new gross & net if split cash/trade
                            slNet = gDivStr(gMulStr(slNet, slPctTrade), "100")
                            slGross = gDivStr(gMulStr(slGross, slPctTrade), "100")

                            tmGrf.iSofCode = ilMatchSSCode              'Sales Source code
                            tmGrf.iVefCode = tmSbf.iAirVefCode
                            tmGrf.lCode4 = tmChf.lCntrNo                'contract # incase no contr header exists
                            tmGrf.iRdfCode = tmIhf.iCode                'needed in case of multimedia report

                            mProcessMonth ilMatchSSCode, slGross, slNet, ilMonthNo, False
                        Next ilCorT
                    End If

                End If
            Next llSbf
        End If
    Next ilCurrentRecd
    mObtainCntrForNTR = 0           'valid return
End Function

'       Obtain the Sales Source code from the contracts salesperson & office
'       Obtain from the common table created tmSlfList
'       Return - Sales source code
Public Function mGetSSCode() As Integer
    Dim ilTemp As Integer
    mGetSSCode = 0
    For ilTemp = LBound(tmSlfList) To UBound(tmSlfList)
        If tmSlfList(ilTemp).iSlfCode = tmChf.iSlfCode(0) Then
            mGetSSCode = tmSlfList(ilTemp).iMnfSSCode          'Sales source
            Exit For
        End If
    Next ilTemp
    If mGetSSCode = 0 Then
        ilTemp = ilTemp
    End If
End Function

'           mProcessMonth - calculate the gross and/or net $ for cash and/or trade category.
'           This routine is for all future (from contracts), along with  past from receivables with the
'           exception of those vehicles that are split by participant in Invoicing (determined by
'           Sales Source : Ask by Vehicle)
'           Write GRF record for each NTR record
'
'           <Input> ilMatchSSCode = Sales Source code
'                   slGross - Gross $ in string
'                   slNet - Net $ in string
'                   ilMonth - Month (1-12) to add $
Public Sub mProcessMonth(ilMatchSSCode As Integer, slGross As String, slNet As String, ilMonthNo As Integer, ilUse100pct As Integer)
'       GRF items used:
'       grfGenTime - Generation Time
'       grfGenDate - generation date
'       grfvefCode - vehicle code
'       grfChfcode - contract code
'       grfDateType - C = cash, T = trade
'       grfSofCode - Sales Source
'       grfrdfCode = item type code (itf)
'       GrfCode4 - SBF code
'       GrfCode2 - imSortBy
'       grfPerGenl(1) - Partipant (mnf code)
'       grfPerGenl(2) - NTR Type (mnf Code)
'       grfPerGenl(3) - Minor vehicle group code
'       grfPerGenl(4) - Major vehicle group code
'       grfPer(1-12) - Monthly $
    Dim ilHowMany As Integer
    Dim ilHowManyDefined As Integer
    Dim illoop As Integer
    Dim ilTemp As Integer
    Dim llProcessPct As Long
    Dim slAmount As String
    Dim slDollar As String
    Dim llTransGross As Long
    Dim llTransNet As Long
    Dim slPct As String
    Dim llNetDollar As Long
    Dim llGrossDollar As Long
    Dim ilRet As Integer
    Dim ilDidHowMany As Integer
    'Index zero ignored in arrays below
    ReDim ilProdPct(0 To 1) As Integer            '5-22-07
    ReDim ilMnfGroup(0 To 1) As Integer           '5-22-07
    ReDim ilMnfSSCode(0 To 1) As Integer          '5-22-07
    'Dim iluse100Pct As Integer                    '8-21-07 use 100% for participant split, dont look for the %
    Dim ilMajorOrMinor As Integer               '0 = major, 1 = minor
    Dim ilHowManyMinor As Integer
    Dim ilHowManyDefinedMinor As Integer
    Dim ilDidHowManyMinor As Integer
    Dim ilLoopMinor As Integer
    Dim llNetDollarMinor As Long
    Dim llGrossDollarMinor As Long
    Dim llTransGrossMinor As Long
    Dim llTransNetMinor As Long

    llTransGross = gStrDecToLong(slGross, 2)
    llTransNet = gStrDecToLong(slNet, 2)
    ilDidHowMany = 0
    'For Advt & Vehicle, create 1 record per transaction
    'For Owner (participant) , create as many as 3 records as there are particpants.
'    iluse100Pct = False          '8-21-07 call to participant isnt using recv, always find the participant share
    gInitPartGroupAndPcts tmSbf.iAirVefCode, ilMatchSSCode, 0, ilMnfSSCode(), ilMnfGroup(), ilProdPct(), tmSbf.iDate(), tmPifKey(), tmPifPct(), ilUse100pct

    mGetHowMany ilMatchSSCode, ilHowMany, ilHowManyDefined, ilMnfSSCode(), 0
    For illoop = 0 To ilHowMany - 1 Step 1      'loop based on report option and how to split revenue if necessary
        For ilTemp = LBound(tmGrf.lDollars) To UBound(tmGrf.lDollars) Step 1               'init the years $ buckets
            tmGrf.lDollars(ilTemp) = 0
        Next ilTemp
        If imSortBy = 3 Then       'sort by owner, split by the participants
            'First match on the sales source
            'If tmVef.iMnfSSCode(ilLoop + 1) = ilMatchSSCode Then
            If ilMnfSSCode(illoop + 1) = ilMatchSSCode Then
                'determine update method

                'llProcessPct = tmVef.iProdPct(ilLoop + 1)
                llProcessPct = ilProdPct(illoop + 1)
                llProcessPct = llProcessPct * 100      'make it xxx.xxxx
            Else
                llProcessPct = 0
            End If
        ElseIf bmSplitSlsp And imSortBy = 4 Then          '9-24-19
            llProcessPct = tmChf.lComm(illoop)
        Else
            llProcessPct = 1000000              'if advt or vehicle, just use slsp since there will always be one
                                                'and only one pass will be processed
        End If

        If llProcessPct > 0 Then                'process if percentage of revenue isnt zero
            ilDidHowMany = ilDidHowMany + 1
            slPct = gLongToStrDec(llProcessPct, 4)

            slDollar = gMulStr(slPct, slGross)                 'slsp gross portion of possible split
            llGrossDollar = Val(gRoundStr(slDollar, "01.", 0))
            llTransGross = llTransGross - llGrossDollar
            If (ilDidHowMany = ilHowManyDefined And RptSelNT!ckcAll.Value = vbChecked) Or llTransGross < 0 Then     'last participant gets extra pennies
                                                         'extra pennies.
                llGrossDollar = llGrossDollar + llTransGross    'last recd gets left over pennies
            End If

            slDollar = gMulStr(slPct, slNet)                 'slsp net portion of possible split
            llNetDollar = Val(gRoundStr(slDollar, "01.", 0))
            llTransNet = llTransNet - llNetDollar
            If (ilDidHowMany = ilHowManyDefined And RptSelNT!ckcAll.Value = vbChecked) Or llTransNet < 0 Then        'last participant gets extra pennies
                                                         'extra pennies.
                llNetDollar = llNetDollar + llTransNet    'last recd gets left over pennies
            End If
            'bucket 1-12 contains slsp comm amount calc from his gross allocation
            If llNetDollar <> 0 Then    'did net amt have any value? if not, dont write out record
                'gGetVehGrpSets tmGrf.iVefCode, imMinorVG, imMajorVG, tmGrf.iPerGenl(3), tmGrf.iPerGenl(4)   'Genl(3) = minor sort code, genl(4) = major sort code
                gGetVehGrpSets tmGrf.iVefCode, imMinorVG, imMajorVG, tmGrf.iPerGenl(2), tmGrf.iPerGenl(3)   'Genl(3) = minor sort code, genl(4) = major sort code
                If imMinorVG = 0 Then
                    'tmGrf.iPerGenl(3) = -1
                    tmGrf.iPerGenl(2) = -1
                End If
                If imMajorVG = 0 Then
                    'tmGrf.iPerGenl(4) = -1
                    tmGrf.iPerGenl(3) = -1
                End If
                If smGrossNet = "G" Then
                    tmGrf.lDollars(ilMonthNo - 1) = llGrossDollar     'gross record always created
                ElseIf smGrossNet = "N" Then  '2-26-01 Net or Net-Net option
                    tmGrf.lDollars(ilMonthNo - 1) = llNetDollar
                End If

                'round the 12 values so that monthly columns will crossfoot to quarter & yearly columns since pennies not shown
                slAmount = gLongToStrDec(tmGrf.lDollars(ilMonthNo - 1), 2)
                slAmount = gRoundStr(slAmount, "1", 0)
                slAmount = gMulStr("100", slAmount)                       ' gross portion of possible split
                tmGrf.lDollars(ilMonthNo - 1) = Val(slAmount)
                tmGrf.iSofCode = ilMatchSSCode          'sales source
                tmGrf.lChfCode = tmChf.lCode          'contract Code
                tmGrf.iVefCode = tmSbf.iAirVefCode
                tmGrf.iPerGenl(1) = tmSbf.iMnfItem      'tran type
                tmGrf.iAdfCode = tmChf.iAdfCode
                tmGrf.iSlfCode = tmChf.iSlfCode(0)          'default to primary slsp
                tmGrf.iPerGenl(0) = ilMnfGroup(1)           'default to primary sales source
                If imSortBy = 3 Or (bmSplitSlsp And imSortbyMinor = 5) Then            'owner
                    tmGrf.iPerGenl(0) = ilMnfGroup(illoop + 1)  '1-29-08 tmVef.iMnfGroup(ilLoop + 1) 'participant
                    If imSortbyMinor = 5 And bmSplitSlsp Then   'slsp, need to split?
                        mGetHowMany ilMatchSSCode, ilHowManyMinor, ilHowManyDefinedMinor, ilMnfSSCode(), 1
 
                        For ilLoopMinor = 0 To ilHowManyMinor - 1 Step 1      'loop based on report option and how to split revenue if necessary
                            llProcessPct = tmChf.lComm(ilLoopMinor)
                            tmGrf.iSlfCode = tmChf.iSlfCode(ilLoopMinor)
                            llTransGrossMinor = llGrossDollar
                            llTransNetMinor = llNetDollar
                            mProcessSplit llProcessPct, ilDidHowManyMinor, ilHowManyDefinedMinor, llGrossDollarMinor, llTransGrossMinor, llNetDollarMinor, llTransNetMinor, ilMonthNo, RptSelNT!ckcAllMinor
                        Next ilLoopMinor
                    Else
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    End If
                ElseIf imSortBy = 4 Or imSortbyMinor = 4 Then        'slsp
                    tmGrf.iSlfCode = tmChf.iSlfCode(illoop)
                    If imSortbyMinor = 4 Then       'owner
                        mGetHowMany ilMatchSSCode, ilHowManyMinor, ilHowManyDefinedMinor, ilMnfSSCode(), 1
                        For ilLoopMinor = 0 To ilHowManyMinor - 1 Step 1      'loop based on report option and how to split revenue if necessary
                            If ilMnfSSCode(ilLoopMinor + 1) = ilMatchSSCode Then
                                llProcessPct = ilProdPct(ilLoopMinor + 1)
                                llProcessPct = llProcessPct * 100      'make it xxx.xxxx
                            Else
                                llProcessPct = 0
                            End If
                            tmGrf.iPerGenl(0) = ilMnfGroup(ilLoopMinor + 1)
                            llTransGrossMinor = llGrossDollar
                            llTransNetMinor = llNetDollar
                            mProcessSplit llProcessPct, ilDidHowManyMinor, ilHowManyDefinedMinor, llGrossDollarMinor, llTransGrossMinor, llNetDollarMinor, llTransNetMinor, ilMonthNo, RptSelNT!ckcAllMinor
                        Next ilLoopMinor
                    Else
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    End If
                Else
                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                End If
            End If                 'llnetdollar <> 0
        End If                              'llProcessPct >
    Next illoop                             'loop for 10 slsp possible splits, or 3 possible owners, otherwise loop once
End Sub

Public Sub mProcessSplit(llProcessPct As Long, ilDidHowManyMinor As Integer, ilHowManyDefinedMinor As Integer, llGrossDollarMinor As Long, llTransGrossMinor As Long, llNetDollarMinor As Long, llTransNetMinor As Long, ilMonthNo As Integer, ckcAll As Control)
    Dim slDollar As String
    Dim slGross As String
    Dim slNet As String
    Dim slAmount As String
    Dim slPct As String
    Dim ilRet As Integer

    slGross = gLongToStrDec(llTransGrossMinor, 2)
    slNet = gLongToStrDec(llTransNetMinor, 2)
    If llProcessPct > 0 Then
        ilDidHowManyMinor = ilDidHowManyMinor + 1
        slPct = gLongToStrDec(llProcessPct, 4)

        slDollar = gMulStr(slPct, slGross)                 'slsp gross portion of possible split
        llGrossDollarMinor = Val(gRoundStr(slDollar, "01.", 0))
        llTransGrossMinor = llTransGrossMinor - llGrossDollarMinor

        slDollar = gMulStr(slPct, slNet)                 'slsp net portion of possible split
        llNetDollarMinor = Val(gRoundStr(slDollar, "01.", 0))
        llTransNetMinor = llTransNetMinor - llNetDollarMinor
        If (ilDidHowManyMinor = ilHowManyDefinedMinor And RptSelNT!ckcAll.Value = vbChecked) Or llTransNetMinor < 0 Then        'last participant gets extra pennies
                                                     'extra pennies.
            llNetDollarMinor = llNetDollarMinor + llTransNetMinor    'last recd gets left over pennies
        End If
        If llNetDollarMinor <> 0 Then
            If smGrossNet = "G" Then
                tmGrf.lDollars(ilMonthNo - 1) = llGrossDollarMinor     'gross record always created
            ElseIf smGrossNet = "N" Then  '2-26-01 Net or Net-Net option
                tmGrf.lDollars(ilMonthNo - 1) = llNetDollarMinor
            End If
            'round the 12 values so that monthly columns will crossfoot to quarter & yearly columns since pennies not shown
            slAmount = gLongToStrDec(tmGrf.lDollars(ilMonthNo - 1), 2)
            slAmount = gRoundStr(slAmount, "1", 0)
            slAmount = gMulStr("100", slAmount)                       ' gross portion of possible split
            tmGrf.lDollars(ilMonthNo - 1) = Val(slAmount)
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        End If
    End If
End Sub
