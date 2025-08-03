Attribute VB_Name = "RptCrSN"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptcrSN.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmChfSrchKey1                 tmSmfSrchKey                  tmSmfSrchKey2             *
'*  tmCxfSrchKey                  hmRdf                         imRdfRecLen               *
'*  tmRdfSrchKey                  tmRdf                                                   *
'******************************************************************************************

Option Explicit
Option Compare Text
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVsf As Integer            'Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim tmClfSrchKey As CLFKEY0     'CLF record image
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmSdf As Integer            'Spot detail file handle
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSdf As SDF
Dim tmSdfSrchKey3 As LONGKEY0     'SDF record image (SDF code as keyfield)
Dim hmSsf As Integer
Dim tmSsf As SSF                'SSF record image
Dim hmSsfRecLen As Integer
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim hmSmf As Integer            'MG file handle
Dim tmSmf As SMF                'SMF record image
Dim imSmfRecLen As Integer        'SMF record length
'Log Calendar
Dim hmLcf As Integer            'Log Calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim imLcfRecLen As Integer        'LCF record length
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GRF record length
Dim tmCxf As CXF
Dim hmCxf As Integer
Dim imCxfRecLen As Integer        'CXF record length
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF

Dim tmFsf As FSF
Dim hmFsf As Integer            'Feed Spots  file handle
Dim imFsfRecLen As Integer      'FSF record length

Dim tmAnf As ANF
Dim hmAnf As Integer            'Feed Spots  file handle
Dim imAnfRecLen As Integer

Dim tmRaf() As RAF
Dim hmRaf As Integer            'Region  file handle
Dim imRafRecLen As Integer

Dim tmSplitNetInv() As SPLITNETINV

Type SPLITNETINV
    sKey As String * 12                'avail time (string in form of number for string sorting)
    lTime As Long       'avail time
    lDate As Long       'avail date
    lInv As Long                 'inventory
End Type

'
'
'           Obtain Sdf record from the SSF
'
'       <Output> - ilSpotOK : False if error
'
Sub mGetSdfChf(ilSpotOK As Integer, tlCntTypes As CNTTYPES)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        ilFound                                                 *
'******************************************************************************************

Dim ilRet As Integer
Dim slShowOnInv As String * 1   '3-24-03

    tmSdfSrchKey3.lCode = tmSpot.lSdfCode
    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
    If tmSpot.lSdfCode = tmSdf.lCode And ilRet = BTRV_ERR_NONE Then
        If tmSdf.lChfCode = 0 Then                           'feed spot vs contr
            tmChfSrchKey.lCode = tmSdf.lFsfCode
            ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)

            If Not tlCntTypes.iNetwork Then
                ilSpotOK = False
            End If
        Else
            If tmSdf.lChfCode <> tmChf.lCode Then               'if already in mem, don't reread
                tmChfSrchKey.lCode = tmSdf.lChfCode
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            Else
                ilRet = BTRV_ERR_NONE
            End If
        End If

        If tmSdf.sSpotType = "X" Then
            slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)       '1-19-04 see if spot should be shown/not shown on inv
            If slShowOnInv = "N" And Not tlCntTypes.iFill Then
                ilSpotOK = False
            End If
            If slShowOnInv = "Y" And Not tlCntTypes.iXtra Then
                ilSpotOK = False
            End If

        End If
    End If

    If ilRet <> BTRV_ERR_NONE Then
        ilSpotOK = False
    End If
End Sub
Sub mFilterLine(tlCntTypes As CNTTYPES, ilSpotOK As Integer, ilCTypes As Integer, ilSpotTypes As Integer, slPrice As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        slStr                         slTempDays                *
'*  slStrip                       slXDay                                                  *
'******************************************************************************************

Dim ilRet As Integer
'Dim slDaysOfWk As String * 14

    If ilSpotOK Then
        If tmSdf.lChfCode = 0 Then
            If Not tlCntTypes.iFeedSpots Then
                ilSpotOK = False
            End If
        Else
            If tmChf.sStatus = "H" Then
                 If Not tlCntTypes.iHold Then
                     ilSpotOK = False
                 End If
             ElseIf tmChf.sStatus = "O" Then
    
                 If Not tlCntTypes.iOrder Then
                     ilSpotOK = False
                 End If
             End If
    
            If tmChf.sType = "C" Then           '3-16-10 wrong flag tested for standard, should be C, not S
                 ilCTypes = &H8
                 If Not tlCntTypes.iStandard Then       'include Standard types?
                     ilSpotOK = False
                 End If
    
             ElseIf tmChf.sType = "V" Then
                 ilCTypes = &H10
                 If Not tlCntTypes.iReserv Then      'include reservations ?
                     ilSpotOK = False
                 End If
    
             ElseIf tmChf.sType = "R" Then
                 ilCTypes = &H20
                 If Not tlCntTypes.iDR Then       'include DR?
                     ilSpotOK = False
                 End If
             End If
             '3-28-02 Reread line if either contract or line # is different.
             If (tmClf.lChfCode <> tmSdf.lChfCode) Or (tmClf.iLine <> tmSdf.iLineNo) Then     'got the matching line reference?
                 tmClfSrchKey.lChfCode = tmSdf.lChfCode
                 tmClfSrchKey.iLine = tmSdf.iLineNo
                 tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
                 tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
                 ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
             End If
             If (tmClf.lChfCode <> tmSdf.lChfCode) Then     'got the matching line reference?
                 ilSpotOK = False
             End If
    
             'Retrieve spot cost from flight ; flight not returned if spot type is Extra/Fill
             'otherwise flight returned in tgPriceCff
             ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
             tlCntTypes.lRate = 0                'init the spot rate until something found
             'look for inclusion of spot types
             If (InStr(slPrice, ".") <> 0) Then    'found spot cost
                 tlCntTypes.lRate = gStrDecToLong(slPrice, 2)    'get the actual spot value
                 'is it a .00?
                 If gCompNumberStr(slPrice, "0.00") = 0 Then     'its a .00 spot
                     ilSpotTypes = &H4
                     If Not tlCntTypes.iZero Then
                         ilSpotOK = False
                     End If
                 Else
                     ilSpotTypes = &H2
                     If Not tlCntTypes.iCharge Then           'exclude charged spots
                         ilSpotOK = False
                     End If
                 End If
             ElseIf Trim$(slPrice) = "ADU" Then
                 ilSpotTypes = &H8
                 If Not tlCntTypes.iADU Then
                     ilSpotOK = False
                 End If
             ElseIf Trim$(slPrice) = "Bonus" Then
                 ilSpotTypes = &H10
                 If Not tlCntTypes.iBonus Then
                     ilSpotOK = False
                 End If
             ElseIf Trim$(slPrice) = "+ Fill" Then
                 ilSpotTypes = &H20
                 If Not tlCntTypes.iXtra Then
                     ilSpotOK = False
                 End If
             ElseIf Trim$(slPrice) = "- Fill" Then
                 ilSpotTypes = &H40
                 If Not tlCntTypes.iFill Then
                     ilSpotOK = False
                 End If
             ElseIf Trim$(slPrice) = "N/C" Then
                 ilSpotTypes = &H80
                 If Not tlCntTypes.iNC Then
                     ilSpotOK = False
                 End If
             ElseIf Trim$(slPrice) = "Recapturable" Then
                 ilSpotTypes = &H100
                 If Not tlCntTypes.iRecapturable Then
                     ilSpotOK = False
                 End If
             ElseIf Trim$(slPrice) = "Spinoff" Then
                 ilSpotTypes = &H200
                 If Not tlCntTypes.iSpinoff Then
                     ilSpotOK = False
                 End If
            ElseIf Trim$(slPrice) = "MG" Then               '10-27-10
                ilSpotTypes = &H400
                If Not tlCntTypes.iMG Then
                    ilSpotOK = False
                End If
            End If
        
            '10-27-10  if excluding MG, that includes MG rate spot types & MG/outside spots
            'test for mg/outside (cant be a bonus spot) and if MG should be included
            If ((tmSdf.sSchStatus = "G" Or tmSdf.sSchStatus = "O") And tmSdf.sSpotType <> "X") And (Not tlCntTypes.iMG) Then
                ilSpotOK = False
            End If
        End If                              'tmsdf.lchfcode = 0
    End If                                  'ilspotOK
End Sub
'
'       gCreateSplitNetworkAvails - create prepass to generate the Split Network Avails report.
'       Report shows all spots within a break.  Each break is calculated for 30" avails, each
'       30" being sold as 100%.  Each spot shows the % of 30" audience sold, based on the audience
'       estimate stored with the region.
'       The report output shows 7 days across, with the advertisers down the left.
'       Each day is shown separately; i.e. all Mon spots, followed for Tue spots, etc.
'       where each day is aligned in its own column.
'
'
Public Sub gCreateSplitNetworkAvails()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVpfIndex                                                                            *
'******************************************************************************************


    Dim ilVefCode As Integer
    Dim tlCntTypes As CNTTYPES
    Dim ilLoopOnVehicle As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilError As Integer
    Dim ilNoWeeks As Integer
    Dim llStartDate As Long
    Dim ilLoop As Integer
    Dim slDate As String
    Dim slStr As String
    Dim ilInclRegCopy As Integer        'include regional copy
    Dim ilInclSplitNet As Integer       'include split network regions
    Dim ilInclSplitCopy As Integer      'include split coyp regions
    Dim ilDay As Integer

    ilError = mOpenSNFiles()
    If ilError Then
        Exit Sub            'at least 1 open error
    End If

    tlCntTypes.iHold = gSetCheck(RptSelSN!ckcCType(0).Value)
    tlCntTypes.iOrder = gSetCheck(RptSelSN!ckcCType(1).Value)
    tlCntTypes.iNetwork = gSetCheck(RptSelSN!ckcCType(2).Value)
    tlCntTypes.iStandard = gSetCheck(RptSelSN!ckcCType(3).Value)
    tlCntTypes.iReserv = gSetCheck(RptSelSN!ckcCType(4).Value)
    tlCntTypes.iRemnant = gSetCheck(RptSelSN!ckcCType(5).Value)
    tlCntTypes.iDR = gSetCheck(RptSelSN!ckcCType(6).Value)
    tlCntTypes.iPI = gSetCheck(RptSelSN!ckcCType(7).Value)
    tlCntTypes.iPSA = gSetCheck(RptSelSN!ckcCType(8).Value)
    tlCntTypes.iPromo = gSetCheck(RptSelSN!ckcCType(9).Value)
    tlCntTypes.iTrade = gSetCheck(RptSelSN!ckcCType(10).Value)
    tlCntTypes.iMissed = gSetCheck(RptSelSN!ckcSpots(0).Value)
    tlCntTypes.iCharge = gSetCheck(RptSelSN!ckcSpots(1).Value)
    tlCntTypes.iZero = gSetCheck(RptSelSN!ckcSpots(2).Value)
    tlCntTypes.iADU = gSetCheck(RptSelSN!ckcSpots(3).Value)
    tlCntTypes.iBonus = gSetCheck(RptSelSN!ckcSpots(4).Value)
    tlCntTypes.iXtra = gSetCheck(RptSelSN!ckcSpots(5).Value)
    tlCntTypes.iFill = gSetCheck(RptSelSN!ckcSpots(6).Value)
    tlCntTypes.iNC = gSetCheck(RptSelSN!ckcSpots(7).Value)
    tlCntTypes.iRecapturable = gSetCheck(RptSelSN!ckcSpots(8).Value)
    tlCntTypes.iSpinoff = gSetCheck(RptSelSN!ckcSpots(9).Value)
    tlCntTypes.iMG = gSetCheck(RptSelSN!ckcSpots(10).Value)         '10-29-10

    'ranks are hidden, all defaulted on
    tlCntTypes.iFixedTime = gSetCheck(RptSelSN!ckcRank(0).Value)        'ranks option hidden, all defaulted on
    tlCntTypes.iSponsor = gSetCheck(RptSelSN!ckcRank(1).Value)
    tlCntTypes.iDP = gSetCheck(RptSelSN!ckcRank(2).Value)
    tlCntTypes.iROS = gSetCheck(RptSelSN!ckcRank(3).Value)

    'days selectivity hidden, all days defaulted on
    For ilLoop = 0 To 6
        tlCntTypes.iValidDays(ilLoop) = True
        If Not gSetCheck(RptSelSN!ckcDays(ilLoop).Value) Then
            tlCntTypes.iValidDays(ilLoop) = False
        End If
    Next ilLoop

    'get all the dates needed to work with
    slDate = RptSelSN!edcSelCFrom.Text               'start date entred
    llStartDate = gDateValue(slDate)
    'backup until Monday start date
    ilDay = gWeekDayLong(llStartDate)
    Do While ilDay <> 0
        llStartDate = llStartDate - 1
        ilDay = gWeekDayLong(llStartDate)
    Loop

    slStr = RptSelSN!edcSelCFrom1.Text               'end date entered
    If Val(slStr) = 0 Then
        ilNoWeeks = 1
    Else
        ilNoWeeks = Val(slStr)
    End If

    ilInclRegCopy = False
    ilInclSplitNet = True
    ilInclSplitCopy = False
    ReDim tmRaf(0 To 0) As RAF
    ilRet = gObtainRAFByType(RptSelSN, hmRaf, tmRaf(), ilInclRegCopy, ilInclSplitNet, ilInclSplitCopy)

    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime

    'loop on vehicles
    For ilLoopOnVehicle = 0 To RptSelSN!lbcSelection.ListCount - 1
        If (RptSelSN!lbcSelection.Selected(ilLoopOnVehicle)) Then
            slNameCode = tgCSVNameCode(ilLoopOnVehicle).sKey
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            tmGrf.iVefCode = ilVefCode
            mGatherSplitNetwork ilVefCode, llStartDate, ilNoWeeks, tlCntTypes
        End If
    Next ilLoopOnVehicle
    ilError = 0         'no error
End Sub
'
'           Gather the split network data:
'           Cycle thru SSF and obtain inventory, create 1 inventory recd per day
'           loop thru spots and create 1 spot entry per day
'
'           <input> ilVefCode = vehicle code (conventional or selling)
'                   llStartDate - Monday start date requested
'                   ilNoWeeks - no weeks requested
'                   tlCntTypes() - parameters requested
'
Public Sub mGatherSplitNetwork(ilVefCode As Integer, llStartDate As Long, ilNoWeeks As Integer, tlCntTypes As CNTTYPES)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llCefCode                     ilEnfCode                     ilPass                    *
'*  ilEndTime                     slTime                        ilStartPass               *
'*  ilCreateAvail                 ilTempTime                    slProduct                 *
'*  llSpotRatio                                                                           *
'******************************************************************************************

Dim ilType As Integer
Dim ilDate0 As Integer
Dim ilDate1 As Integer
Dim slDate As String
Dim llDate As Long
Dim ilEvt As Integer
Dim ilRet As Integer
Dim ilSpot As Integer
Dim llTime As Long
Dim ilLoop As Integer
Dim ilSpotOK As Integer
Dim llLoopDate As Long
Dim ilWeekDay As Integer
Dim llLatestDate As Long
Dim ilIndex As Integer
Dim ilRemSec As Integer     'time in seconds of avail, each spot length subtracted to get remaining seconds
Dim ilRemUnits As Integer   '# of units of avail, each spot length subtracted to get remaining units
Dim ilCTypes As Integer       'bit map of cnt types to include starting lo order bit
                              '0 = unused, 1= unused, 2 = network, 3 = std, 4 = Reserved, 5 = remanant, 6 = DR
                              '7 = PI, 8 = psa, 9 = promo, 10 = trade
                              'bit 0 and 1 previously hold & order; but it conflicts with other contract types
Dim ilSpotTypes As Integer    'bit map of spot types to include starting lo order bit:
                              '0 = missed (not used in this report), 1 = charge, 2 = 0.00, 3 = adu, 4 = bonus, 5 = extra
                              '6 = fill, 7 = n/c 8 = recapturable, 9 = spinoff
Dim ilRanks As Integer        '0 = fixed time, 1 = sponsorship, 2 = DP, 3= ROS
Dim ilFilterDay As Integer
ReDim ilStartTime(0 To 1) As Integer
ReDim ilEvtType(0 To 14) As Integer
Dim slPrice As String
Dim ilLoopOnWeek As Integer
Dim ilVefIndex As Integer
Dim llSDate As Long
Dim llEDate As Long
Dim ilUpperInv As Integer
Dim llInv As Long
Dim ilLoopOnInv As Integer
Dim ilfirstTime As Integer
Dim slStr As String
Dim ilDay As Integer
Dim llRafIndex As Long

'       GRF descriptions:
'       grfGenDate - generation date for filtering to crystal
'       grfGenTime - generation time for filtering to crystal
'       grfvefcode - vehicle code
'       grfStartDate - start date of week
'       grfChfCode - contract code (spot record only)
'       grfCode4 - region code (spot record only)
'       grfDate - one record per day (date for spot)
'       grfTime - time of break
'       grfSofCode - 0 = Inventory record, 1 = spot record
'       grfDollars(1-7) - Inventory based on audience for Mon-Sun (inv record only)
'       grfDollars(8) - spot rate (spot record only)
'       grfGenDesc - spot rate in string if type is bonus, mg, or other type
'       grfDollars(9) - date (start of week) stored as long for sorting in crystal to avoid mixed types
'       grfPerGenl(1) - spot length (spot record only)
'       grfPerGenl(2) - Sequence #, to keep spots in same order as spot screen
'       grfPerGenl(3) -
'       GrfperGenl(4) - break #
'       grfperGenl(5) = region % of audience
'       Crystal reports will sort:  vehicle, start date of week, avail time, record type (inventory or spot),
'                                   day (date), sequence # (important on spots to keep in same order as spot screen)
    ilType = 0
    llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
    'set the type of events to get fro the day
    For ilLoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(ilLoop) = False
    Next ilLoop
    ilEvtType(2) = True     'retrieve contract avails
    imSdfRecLen = Len(tmSdf)
    imCHFRecLen = Len(tmChf)
    ilVefIndex = gBinarySearchVef(ilVefCode)
    'If ilVefIndex <= 0 Then
    If ilVefIndex < 0 Then
        Exit Sub
    End If

    llSDate = llStartDate
    llEDate = llStartDate + 6   'process 1 week at a time for the given # of requested weeks
    For ilLoopOnWeek = 1 To ilNoWeeks       'gather for requested # of weeks
        ReDim tmSplitNetInv(0 To 0) As SPLITNETINV
        ilUpperInv = 0
        For llLoopDate = llSDate To llEDate Step 1
            slDate = Format$(llLoopDate, "m/d/yy")
            gPackDate slDate, ilDate0, ilDate1
            If llLoopDate = llSDate Then        'save start date of week for report
                tmGrf.iStartDate(0) = ilDate0
                tmGrf.iStartDate(1) = ilDate1
                'tmGrf.lDollars(9) = gDateValue(slDate)      'save as long for sorting in Crystal due to different data types
                tmGrf.lDollars(8) = gDateValue(slDate)      'save as long for sorting in Crystal due to different data types
            End If

            ilFilterDay = gWeekDayLong(llLoopDate)
            'currently day selectivity is hidden, all days are defaulted on
            If tlCntTypes.iValidDays(ilFilterDay) Then      'Has this day of the week been selected?
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                'tmSsfSrchKey.sType = slType
                If tgMVef(ilVefIndex).sType <> "G" Then
                    tmSsfSrchKey.iType = ilType
                    tmSsfSrchKey.iVefCode = ilVefCode
                    tmSsfSrchKey.iDate(0) = ilDate0
                    tmSsfSrchKey.iDate(1) = ilDate1
                    tmSsfSrchKey.iStartTime(0) = 0
                    tmSsfSrchKey.iStartTime(1) = 0
                    ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                Else
                    tmSsfSrchKey2.iVefCode = ilVefCode
                    tmSsfSrchKey2.iDate(0) = ilDate0
                    tmSsfSrchKey2.iDate(1) = ilDate1
                    ilRet = gSSFGetGreaterOrEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
                    ilType = tmSsf.iType
                End If
                'Always build the events within the library so that the associated program
                'name and avail information can be retrieved (they are not in SSF)
                ReDim tlLLC(0 To 0) As LLC  'Merged library names
                If tgMVef(ilVefIndex).sType <> "G" Then
                    ilWeekDay = gWeekDayStr(slDate) + 1 '1=Monday TFN; 2=Tues,...7=Sunday TFN
                    If ilWeekDay = 1 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNMO", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 2 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTU", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 3 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNWE", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 4 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTH", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 5 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNFR", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 6 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSA", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 7 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSU", "12M", "12M", ilEvtType(), tlLLC())
                    End If
                End If
                If (llLoopDate > llLatestDate) Then     'if there are not spots scheduled this far in future, fake out an SSF entry
                                                        'so that empty avails can still be shown
                    'tmSsf.sType = "O"
                    tmSsf.iType = ilType
                    tmSsf.iVefCode = ilVefCode
                    tmSsf.iDate(0) = ilDate0
                    tmSsf.iDate(1) = ilDate1
                    gPackTime tlLLC(0).sStartTime, tmSsf.iStartTime(0), tmSsf.iStartTime(1)
                    tmSsf.iCount = 0
                    'tmSsf.iNextTime(0) = 1  'Time not defined
                    'tmSsf.iNextTime(1) = 0

                    For ilIndex = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1

                        tmAvail.iRecType = Val(tlLLC(ilIndex).sType)
                        gPackTime tlLLC(ilIndex).sStartTime, tmAvail.iTime(0), tmAvail.iTime(1)
                        tmAvail.iLtfCode = tlLLC(ilIndex).iLtfCode
                        tmAvail.iAvInfo = tlLLC(ilIndex).iAvailInfo Or tlLLC(ilIndex).iUnits
                        tmAvail.iLen = CInt(gLengthToCurrency(tlLLC(ilIndex).sLength))
                        tmAvail.ianfCode = Val(tlLLC(ilIndex).sName)
                        tmAvail.iNoSpotsThis = 0
                        tmAvail.iOrigUnit = 0
                        tmAvail.iOrigLen = 0
                        tmSsf.iCount = tmSsf.iCount + 1
                        tmSsf.tPas(ADJSSFPASBZ + tmSsf.iCount) = tmAvail
                    Next ilIndex
                End If
                ilRet = BTRV_ERR_NONE

                Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
                    gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llDate
                    ilDay = gWeekDayLong(llDate)   'obtain day of week
                    ilEvt = 1

                    Do While ilEvt <= tmSsf.iCount
                        tmGrf.iDate(0) = ilDate0  'spot date
                        tmGrf.iDate(1) = ilDate1
                       LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)

                        If (tmProg.iRecType >= 2) And (tmProg.iRecType <= 2) Then 'Contract Avails only
                           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                            tmGrf.iTime(0) = tmAvail.iTime(0)   'time of avail for the spot record
                            tmGrf.iTime(1) = tmAvail.iTime(1)
                            ilStartTime(0) = tmAvail.iTime(0)
                            ilStartTime(1) = tmAvail.iTime(1)

                            ilRemSec = tmAvail.iLen
                            ilRemUnits = tmAvail.iAvInfo And &H1F

                            'Create an entry in array for each avail and day to contain the time, avail, inventory and time sort string
                            'The array will be sorted by avail time and the inventory for 1 week at a time per avail will be created in
                            'prepass to show weekly inventory in crystal report
                            'calculate avails based on 30 seconds only, no units
                            llInv = (CLng(ilRemSec) * 1000) / 30 'get fraction of 30" avails; 45" will be 150%, store for xxx.x

                             tmSplitNetInv(ilUpperInv).lTime = llTime
                             tmSplitNetInv(ilUpperInv).lDate = llDate
                             tmSplitNetInv(ilUpperInv).lInv = llInv

                             slStr = Trim$(str$(llTime))
                             Do While Len(slStr) < 6
                                 slStr = "0" & slStr
                             Loop
                            tmSplitNetInv(ilUpperInv).sKey = Trim$(slStr)
                            ReDim Preserve tmSplitNetInv(0 To ilUpperInv + 1) As SPLITNETINV
                            ilUpperInv = ilUpperInv + 1

                            For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                               LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                                ilSpotOK = True                             'assume spot is OK to include
                                ilSpotTypes = 0
                                ilCTypes = 0
                                ilRanks = 0

                                gFilterSpotTypes ilCTypes, ilSpotTypes, ilSpotOK, tlCntTypes, tmSpot    'determine to include/exclude trades, psa,promo, remnants, PI, DR & extra spots
                                If ilSpotOK Then                            'continue testing other filters
                                    mGetSdfChf ilSpotOK, tlCntTypes   'get the sdf record from ssf

                                    mFilterLine tlCntTypes, ilSpotOK, ilCTypes, ilSpotTypes, slPrice
'                                        If ilSpotOK Then
'                                            If tmSdf.lChfCode > 0 Then          'contract spot vs feed
'                                                'contract spot
'                                                If (tmClf.lChfCode <> tmSdf.lChfCode) Then     'got the matching line reference?
'                                                    ilSpotOK = False
'                                                Else
'                                                    'Rank selectivity is hidden, but the defaults are on on
'                                                    'mFilterDPRanks ilRanks, ilSpotOK, tlCntTypes   'Filter out spot ranks (Fixed Time, Sponsorship, Daypart and ROS)
'                                                End If
'                                            End If
'                                        End If
                                End If

                                If ilSpotOK Then

                                    'create the spot record
                                    tmGrf.iSofCode = 1             'spot type for sorting
                                    tmGrf.lChfCode = tmSdf.lChfCode             'contr code
                                    'tmGrf.lDollars(8) = tlCntTypes.lRate                'spot rate
                                    tmGrf.lDollars(7) = tlCntTypes.lRate                'spot rate
                                    tmGrf.sGenDesc = Trim$(slPrice)             'rate in string in case bonus or other type of spot cost
                                    'tmGrf.iPerGenl(1) = tmSdf.iLen               'spot length
                                    'tmGrf.iPerGenl(2) = tmGrf.iPerGenl(2) + 1                'seq # for spots to show in same order as spot screen
                                    tmGrf.iPerGenl(0) = tmSdf.iLen               'spot length
                                    tmGrf.iPerGenl(1) = tmGrf.iPerGenl(1) + 1                'seq # for spots to show in same order as spot screen
                                    llRafIndex = mBinarySearchRaf(tmClf.lRafCode, tmRaf())
                                    'tmGrf.iPerGenl(5) = 10000              'region % of audience, assume full network buy @ 100.00%
                                    tmGrf.iPerGenl(4) = 10000              'region % of audience, assume full network buy @ 100.00%
                                    If llRafIndex >= 0 Then
                                        If tmRaf(llRafIndex).iAudPct > 0 Then
                                            'tmGrf.iPerGenl(5) = tmRaf(llRafIndex).iAudPct         'region % of audience
                                            tmGrf.iPerGenl(4) = tmRaf(llRafIndex).iAudPct         'region % of audience
                                        End If
                                    End If
                                    tmGrf.lCode4 = tmClf.lRafCode               'region code
                                    For ilLoop = 1 To 7
                                        tmGrf.lDollars(ilLoop - 1) = 0
                                    Next ilLoop
                                    'calc the amt of  network sold
                                    'tmGrf.lDollars(ilDay + 1) = (llInv * CLng(tmGrf.iPerGenl(5))) / 10000

                                    'determine the ratio of the spot length against a 30" spot to determine how much audience is left to sell
                                    ''tmGrf.lDollars(ilDay + 1) = (tmGrf.iPerGenl(5) * ((tmSdf.iLen * 100) / 30) / 100)
                                    'tmGrf.lDollars(ilDay) = (tmGrf.iPerGenl(5) * ((tmSdf.iLen * 100) / 30) / 100)
                                    tmGrf.lDollars(ilDay) = (tmGrf.iPerGenl(4) * ((tmSdf.iLen * 100) / 30) / 100)
                                    'tmGrf.iPerGenl(4) = 0           'break # determined by the avail
                                    tmGrf.iPerGenl(3) = 0           'break # determined by the avail
                                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)

                                    ilRemSec = ilRemSec - (tmSpot.iPosLen And &HFFF)  'keep running total of whats remaining in avail based on spots used in stats
                                    ilRemUnits = ilRemUnits - 1
                                End If
                            Next ilSpot                                 'loop from ssf file for # spots in avail

                            ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                        End If
                        ilEvt = ilEvt + 1   'Increment to next event
                    Loop                                                        'do while ilEvt <= tmSsf.iCount
                    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    If tgMVef(ilVefIndex).sType = "G" Then
                        ilType = tmSsf.iType
                    End If
                Loop
            End If
        Next llLoopDate
        If UBound(tmSplitNetInv) > 0 Then
            ArraySortTyp fnAV(tmSplitNetInv(), 0), UBound(tmSplitNetInv), 0, LenB(tmSplitNetInv(0)), 0, LenB(tmSplitNetInv(0).sKey), 0
        End If


        'array containing all the inventory by day and avail time has been sorted by avail time for the week.
        'Create 1 entry per week per avail time to show each breaks inventory
        ilfirstTime = True
        tmGrf.iVefCode = ilVefCode
        'tmGrf.iStartDate =set above for start week at start of week loop
        tmGrf.iSofCode = 0       'inv record (vs spot)
        'tmGrf.iPerGenl(4) = 1       'init break #s
        tmGrf.iPerGenl(3) = 1       'init break #s
        'set fields N/A to the avail record
        tmGrf.lChfCode = 0          'contract code
        'tmGrf.lDollars(8) = 0       'spot rate
        tmGrf.lDollars(7) = 0       'spot rate
        tmGrf.lCode4 = 0            'region code
        'tmGrf.iPerGenl(1) = 0       'spot length
        'tmGrf.iPerGenl(2) = 0       'seq # to keep spots in order same as spot screen
        tmGrf.iPerGenl(0) = 0       'spot length
        tmGrf.iPerGenl(1) = 0       'seq # to keep spots in order same as spot screen

        For ilLoopOnInv = 0 To UBound(tmSplitNetInv)
            If ilfirstTime Then         'if first time thru, init the inventory buckets
                For ilLoop = 1 To 7
                    tmGrf.lDollars(ilLoop - 1) = 0
                Next ilLoop

                ilfirstTime = False
                llTime = tmSplitNetInv(ilLoopOnInv).lTime
            End If
            If llTime = tmSplitNetInv(ilLoopOnInv).lTime Then
                'determine which day this entry belongs to
                ilDay = gWeekDayLong(tmSplitNetInv(ilLoopOnInv).lDate)   'obtain day of week
                'tmGrf.lDollars(ilDay + 1) = tmSplitNetInv(ilLoopOnInv).lInv     'plug in the days inventory
                tmGrf.lDollars(ilDay) = tmSplitNetInv(ilLoopOnInv).lInv     'plug in the days inventory
            Else
                'different avail, write out the weeks worth of inventory for this break
                gPackTimeLong llTime, tmGrf.iTime(0), tmGrf.iTime(1)
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                'initalize the next set of avails
                'tmGrf.iPerGenl(4) = tmGrf.iPerGenl(4) + 1
                tmGrf.iPerGenl(3) = tmGrf.iPerGenl(3) + 1
                For ilLoop = 1 To 7
                    tmGrf.lDollars(ilLoop - 1) = 0
                Next ilLoop
                llTime = tmSplitNetInv(ilLoopOnInv).lTime           'set current
                'determine which day this entry belongs to
                ilDay = gWeekDayLong(tmSplitNetInv(ilLoopOnInv).lDate)   'obtain day of week
                'tmGrf.lDollars(ilDay + 1) = tmSplitNetInv(ilLoopOnInv).lInv
                tmGrf.lDollars(ilDay) = tmSplitNetInv(ilLoopOnInv).lInv
            End If
        Next ilLoopOnInv
        'Loop thru the weeks avails and create 1 record for all 7 days per availtime

        llSDate = llEDate + 1
        llEDate = llSDate + 6
    Next ilLoopOnWeek
    Erase ilEvtType
    Erase tlLLC
    Exit Sub
End Sub
'
'           Open files required for Split Network Avails
'           Return - error flag = true for open error
'
Public Function mOpenSNFiles() As Integer
Dim ilRet As Integer
Dim slTable As String * 3
Dim ilError As Integer

    ilError = False
    On Error GoTo mOpenSNFilesErr

        slTable = "Chf"
    hmCHF = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Function
    End If
    imCHFRecLen = Len(tmChf)

    slTable = "Grf"
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Exit Function
    End If
    imGrfRecLen = Len(tmGrf)

    slTable = "Clf"
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        Exit Function
    End If
    imClfRecLen = Len(tmClf)

    slTable = "Cff"
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        Exit Function
    End If
    imCffRecLen = Len(tmCff)
    
    slTable = "Mnf"
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        Exit Function
    End If
    imMnfRecLen = Len(tmMnf)
    
    slTable = "Vef"
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Function
    End If
    imVefRecLen = Len(tmVef)
    
    slTable = "Sdf"
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmSdf)
        btrDestroy hmSdf
        Exit Function
    End If
    imSdfRecLen = Len(tmSdf)
    
    slTable = "Ssf"
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmSsf)
        btrDestroy hmSsf
        Exit Function
    End If
    hmSsfRecLen = Len(tmSsf)
    
    slTable = "Lcf"
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmLcf)
        btrDestroy hmLcf
        Exit Function
    End If
    imLcfRecLen = Len(tmLcf)
    
    slTable = "Smf"
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmSmf)
        btrDestroy hmSmf
        Exit Function
    End If
    imSmfRecLen = Len(tmSmf)
    
    slTable = "Vsf"
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmVsf)
        btrDestroy hmVsf
        Exit Function
    End If
    imVsfRecLen = Len(tmVsf)

    slTable = "Cxf"
    hmCxf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmCxf)
        btrDestroy hmCxf
        Exit Function
    End If
    imCxfRecLen = Len(tmCxf)
    
    slTable = "Fsf"
    hmFsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmFsf)
        btrDestroy hmFsf
        Exit Function
    End If
    imFsfRecLen = Len(tmFsf)

    slTable = "Anf"
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmAnf)
        btrDestroy hmAnf
        Exit Function
    End If
    imAnfRecLen = Len(tmAnf)
    
    slTable = "Raf"
    hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSNFiles = True
        ilRet = btrClose(hmRaf)
        btrDestroy hmRaf
        Exit Function
    End If
    imRafRecLen = Len(tmRaf(0))

    mOpenSNFiles = ilError
    Exit Function
    
mOpenSNFilesErr:
    ilError = Err.Number
    gBtrvErrorMsg ilRet, "mOpenSNFiles (OpenError) #" & str(ilError) & ": " & slTable, RptSelSN

    Resume Next
End Function
'
'       mBinarySearchRAF - search for the matching region
'       <input> RAFcode to match on
'               RAF array - records were read and created in array by code #
'
'       <return> - -1 if not found; else the index to matching entry
'
Public Function mBinarySearchRaf(llRafCode As Long, tlRaf() As RAF)
 Dim llMin As Long
    Dim llMax As Integer
    Dim llMiddle As Long
    llMin = LBound(tlRaf)
    llMax = UBound(tlRaf) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llRafCode = tlRaf(llMiddle).lCode Then
            'found the match
            mBinarySearchRaf = llMiddle
            Exit Function
        ElseIf llRafCode < tlRaf(llMiddle).lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchRaf = -1
End Function
