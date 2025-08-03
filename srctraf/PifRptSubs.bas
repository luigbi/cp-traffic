Attribute VB_Name = "PIFRptSub"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PifRptSubs.bas on Wed 6/17/09 @ 12:56
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  hmVsf                                                                                 *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PDFSubs.BAS
'
' Release: 5.6

Dim tmPif As PIF
Dim hmPif As Integer
Dim imPifRecLen As Integer
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
Dim tmVsfSrchKey As LONGKEY0

'this is a table of all vehicles that will point to the table of participant percentages
Type PIFKEY
    iVefCode As Integer
    iLoInx As Integer               'this indicates the starting index of PIFPCT table for 1 vehicle
    iHiInx As Integer               'this indicates the ending index of the PIFPCT table for 1 vevhicle
End Type

Type PIFPCT
    iSSMnfCode As Integer
    iMnfGroup As Integer
    lStartDate As Long
    lEndDate As Long
    iPct As Integer
    iOwnerSeq As Integer    '5-11-09 seq 1 indicates owner
    iOwnerByDate As Integer      '8-2-09 When % change by date, the participant
                                    'could also change owner, keep track of the first
                                    'entry for each date change
End Type

Type ONEPARTYEAR
    iVefCode As Integer
    iSSMnfCode As Integer
    iMnfGroup As Integer
    iOwnerSeq As Integer    '5-11-09 seq 1 indicates owner
    iOwnerByDate As Integer
    iPct(0 To 12) As Integer    'Index zero ignored
End Type

Type ALLPIFPCTYEAR
     AllYear As ONEPARTYEAR
End Type

'               mgInitPartGroupAndPcts - set up array of 8 participants and
'               participant percentages from vehicle table.  Create the list of
'               participants that match the sales source.  If the incoming
'               vehicle participant code from receivables (ilRvfMnfgroup) is
'               a non-zero, the share has already been split in the record, and
'               there is no need to split.  That participant gets 100%.  Alter
'               the table to use only the 1 participant for 100%
'               <input>  ilVefCode = vehicle code to find the matching sales source & participants
'                        ilSSCode - Sales Source code to match in vehicle table
'                        ilRvfMnfGroup - mnf group from receivables or history (rvfmnfgroup)
'                        ilMnfSSCode() - array of sales source codes (like tmvefMnfSSCode)
'                        ilMnfGroup() - array of participants (like tmvefmnfGroup)
'                        ilProdPct() - array of particpants revenue share percentages (like tmvefprodpct)
'                        ilUse100Pct - use 100% for the participant share or look for the participants %, true/false
'                                      if rvfmnfgroup exists, if true use 100%
'                                      if rvfmnfgroup exists, if false look for the % share
'               <return> ilMnfSSCode(), ilMnfGroup, ilProdPct() altered to 100% if the transactions mnfgroup is
'                       non-zero, which means the transaction has already been split
Public Sub gInitPartGroupAndPcts(ilVefCode As Integer, ilSSCode As Integer, ilRvfMnfGroup As Integer, ilMnfSSCode() As Integer, ilMnfGroup() As Integer, ilProdPct() As Integer, ilDate() As Integer, tlPIFKey() As PIFKEY, tlPifPct() As PIFPCT, ilUse100pct As Integer)
    Dim ilLoop As Integer
    Dim ilMatchingSSIndex As Integer
    Dim ilVefInx As Integer
    Dim ilLo As Integer
    Dim ilHi As Integer
    Dim llDate As Long
    Dim slStr As String
    Dim ilAdjustUBound As Integer
        
    ilAdjustUBound = False
    'find vehicle in participant key table
    ilVefInx = gBinarySearchPIFKey(ilVefCode, tlPIFKey())
    If ilVefInx < 0 Then                'no vehicle found, dont process it
        ilMnfSSCode(1) = ilSSCode
        ilMnfGroup(1) = ilRvfMnfGroup
        ilProdPct(1) = 10000
       Exit Sub
    Else
        gUnpackDate ilDate(0), ilDate(1), slStr
        llDate = gDateValue(slStr)

        ilLo = tlPIFKey(ilVefInx).iLoInx     'get the starting/ending index for all the participants that belong with this vehicle
        ilHi = tlPIFKey(ilVefInx).iHiInx
        ilMatchingSSIndex = 1
        For ilLoop = ilLo To ilHi
            If tlPifPct(ilLoop).iSSMnfCode = ilSSCode Then
                If ilRvfMnfGroup > 0 And ilUse100pct Then               'the transacction has already been split
                    ilMnfSSCode(1) = ilSSCode
                    ilMnfGroup(1) = ilRvfMnfGroup
                    ilProdPct(1) = 10000
                    ilLoop = ilHi                        'force exit
                Else                                    'not using split transactsions for this sales source, default
                                                        'the arrays to be the same as vehicle, using only the matching sources
                    'verify the effective date of the participant pct
                    '7-23-07 incorrect loop variable to test the dates
                    If llDate >= tlPifPct(ilLoop).lStartDate And llDate <= tlPifPct(ilLoop).lEndDate Then
                        ilMnfSSCode(ilMatchingSSIndex) = tlPifPct(ilLoop).iSSMnfCode
                        ilMnfGroup(ilMatchingSSIndex) = tlPifPct(ilLoop).iMnfGroup
                        ilProdPct(ilMatchingSSIndex) = tlPifPct(ilLoop).iPct
                        ilMatchingSSIndex = ilMatchingSSIndex + 1
                        ReDim Preserve ilMnfSSCode(0 To ilMatchingSSIndex) As Integer
                        ReDim Preserve ilMnfGroup(0 To ilMatchingSSIndex) As Integer
                        ReDim Preserve ilProdPct(0 To ilMatchingSSIndex) As Integer
                        ilAdjustUBound = True
                    End If
                End If
            End If
        Next ilLoop
    End If
    If ilAdjustUBound Then              'need to get the actual # of entries
        ReDim Preserve ilMnfSSCode(0 To ilMatchingSSIndex - 1) As Integer
        ReDim Preserve ilMnfGroup(0 To ilMatchingSSIndex - 1) As Integer
        ReDim Preserve ilProdPct(0 To ilMatchingSSIndex - 1) As Integer
    End If
    Exit Sub
gInitPartGroupAndPctsErr2:
    ilRet = 1
    Resume Next
End Sub

'          Build the Particpant percentages for all vehicles for a contract, matching
'          the contracts Sales Source.
Public Sub gInitCntPartYear(hlVsf As Integer, tlChf As CHF, ilSSCode As Integer, llStdStartDates() As Long, tlCntAllYear() As ALLPIFPCTYEAR, tlPIFKey() As PIFKEY, tlPifPct() As PIFPCT)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVeh                         ilDate                        slStr                     *
'*                                                                                        *
'******************************************************************************************
    Dim llLkVsfCode As Long
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVefInx As Integer
    Dim ilUpper As Integer
    Dim ilLo As Integer
    Dim ilHi As Integer
    Dim ilLoop12Months As Integer
    Dim llStartDate As Long
    Dim ilVef As Integer
    Dim llDate As Long
    Dim ilVefSpan As Integer
    Dim ilValidDate As Integer
    ReDim ilVefPartCodes(0 To 0) As Integer
    Dim ilLoopOnVehicle As Integer
    Dim ilFoundVehicle As Integer
    Dim ilLowLimit As Integer

    If PeekArray(tlCntAllYear).Ptr <> 0 Then
        ilLowLimit = LBound(tlCntAllYear)
    Else
        ilLowLimit = 0
    End If
    
    'ReDim tlCntAllYear(ilLowLimit To 1) As ALLPIFPCTYEAR
    ReDim tlCntAllYear(ilLowLimit To ilLowLimit) As ALLPIFPCTYEAR

    'Obtain all the vehicles of the contract (NTR and selling, conventional; no pckages are in vsf)
    imVsfRecLen = Len(tmVsf)
    If tlChf.lVefCode > 0 Then              'all vehicles are the same for this contract
        ilVefPartCodes(0) = tlChf.lVefCode
        ReDim Preserve ilVefPartCodes(0 To 1) As Integer
    ElseIf tlChf.lVefCode < 0 Then
        llLkVsfCode = -tlChf.lVefCode
        Do While llLkVsfCode > 0            'there could be multiple vsf records with all the vehicle references
            tmVsfSrchKey.lCode = llLkVsfCode
            ilRet = btrGetEqual(hlVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Sub
            End If
            For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                If tmVsf.iFSCode(ilLoop) > 0 Then
                    ilVefPartCodes(UBound(ilVefPartCodes)) = tmVsf.iFSCode(ilLoop)
                    ReDim Preserve ilVefPartCodes(0 To UBound(ilVefPartCodes) + 1) As Integer
                End If
            Next ilLoop
            llLkVsfCode = tmVsf.lLkVsfCode
        Loop
    End If
    ilUpper = ilLowLimit
    'loop thru the participants and create an entry for the percentages across the year (also matching sales source and vehicle)
    For ilLoop = LBound(ilVefPartCodes) To UBound(ilVefPartCodes) - 1
        ilVef = ilVefPartCodes(ilLoop)
        ilVefInx = gBinarySearchPIFKey(ilVef, tlPIFKey())
        If ilVefInx > 0 Then
            llDate = llStdStartDates(1)
            ilLo = tlPIFKey(ilVefInx).iLoInx     'get the starting/ending index for all the participants that belong with this vehicle
            ilHi = tlPIFKey(ilVefInx).iHiInx
            For ilVefSpan = ilLo To ilHi        'loop thru the particpants to gather their percentages by date
                If tlPifPct(ilVefSpan).iSSMnfCode = ilSSCode Then        'find only those participants that match the sales source of contract
                    ilValidDate = False
                    ilFoundVehicle = False

                    For ilLoopOnVehicle = LBound(tlCntAllYear) To UBound(tlCntAllYear) - 1
                        If tlCntAllYear(ilLoopOnVehicle).AllYear.iVefCode = ilVef And tlCntAllYear(ilLoopOnVehicle).AllYear.iMnfGroup = tlPifPct(ilVefSpan).iMnfGroup Then
                            If tlCntAllYear(ilLoopOnVehicle).AllYear.iOwnerByDate = tlPifPct(ilVefSpan).iOwnerByDate Then           '3-20-12 if participants have changed, create new record for each set.  Owners may have changed.
                                ilFoundVehicle = True
                                Exit For
                            End If
                        End If
                    Next ilLoopOnVehicle
                    If Not ilFoundVehicle Then
                        tlCntAllYear(ilUpper).AllYear.iVefCode = tlPIFKey(ilVefInx).iVefCode
                        tlCntAllYear(ilUpper).AllYear.iSSMnfCode = tlPifPct(ilVefSpan).iSSMnfCode
                        tlCntAllYear(ilUpper).AllYear.iMnfGroup = tlPifPct(ilVefSpan).iMnfGroup
                        tlCntAllYear(ilUpper).AllYear.iOwnerSeq = tlPifPct(ilVefSpan).iOwnerSeq     '5-11-09
                        tlCntAllYear(ilUpper).AllYear.iOwnerByDate = tlPifPct(ilVefSpan).iOwnerByDate   '8-2-09
                        ilLoopOnVehicle = ilUpper
                        ilUpper = ilUpper + 1
                        ReDim Preserve tlCntAllYear(ilLowLimit To ilUpper) As ALLPIFPCTYEAR
                    End If
                    For ilLoop12Months = 1 To 12
                        llStartDate = llStdStartDates(ilLoop12Months)
                        'verify the effective date of the participant pct
                        '7-23-07 incorrect date was tested for filtering; incorrect index tested for date span filtering
                        If llStartDate >= tlPifPct(ilVefSpan).lStartDate And llStartDate <= tlPifPct(ilVefSpan).lEndDate Then
                            If tlCntAllYear(ilLoopOnVehicle).AllYear.iPct(ilLoop12Months) = 0 Then
                                tlCntAllYear(ilLoopOnVehicle).AllYear.iPct(ilLoop12Months) = tlPifPct(ilVefSpan).iPct
                            End If
                        End If
                    Next ilLoop12Months
                End If
            Next ilVefSpan
        End If
    Next ilLoop
    Exit Sub
gInitCntPartYearErr2:
    ilRet = 1
    Resume Next
End Sub

'           mInitVehAllYearPct - build array of all participants for the one vehicle
'           that is being processed.  Each entry is a type record containing the
'           vehicle, sales source, participant, and 12 months of participant percentages
'           <Input>  ilMatchSSCode - sales source to retrieve for the vehicle
'                    ilVefCode - vehicle to process
'                    tlCntAllYear - table containing all the participants info for each vehicle in this contract
'           <output> tlOneVehAllYear - table containing all participants for 1 vehicle
'           Return - true if the vehicle is on this contract; otherwise need to get the vehicle participant info
Public Function gInitVehAllYearPcts(ilMatchSSCode As Integer, ilVefCode As Integer, tlCntAllYear() As ALLPIFPCTYEAR, tlOneVehAllYear() As ALLPIFPCTYEAR) As Integer
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim ilFound As Integer
    Dim ilLowLimit As Integer
    If PeekArray(tlOneVehAllYear).Ptr <> 0 Then
        ilLowLimit = LBound(tlOneVehAllYear)
    Else
        ilLowLimit = 0
    End If
    
    ilFound = False
    ReDim tlOneVehAllYear(ilLowLimit To ilLowLimit) As ALLPIFPCTYEAR
    ilUpper = ilLowLimit
    For ilLoop = LBound(tlCntAllYear) To UBound(tlCntAllYear)
        If tlCntAllYear(ilLoop).AllYear.iSSMnfCode = ilMatchSSCode And tlCntAllYear(ilLoop).AllYear.iVefCode = ilVefCode Then
            tlOneVehAllYear(ilUpper).AllYear = tlCntAllYear(ilLoop).AllYear
            ilUpper = ilUpper + 1
            ReDim Preserve tlOneVehAllYear(ilLowLimit To ilUpper) As ALLPIFPCTYEAR
            ilFound = True          'found at least 1 matching ss and vehicle
        End If
    Next ilLoop
    gInitVehAllYearPcts = ilFound
    Exit Function
gInitVehAllYearPctsErr2:
    ilRet = 1
    Resume Next
End Function

'           Obtain the Participant Percentages splits from PIF and build
'           array of the records by date.
'           2 arrays created:  1 array contains the vehicle.  This table has
'           a lo index and a hi index which points to the PIF date entries.
'           i.e. PIFKEY:  has array of vehicle, lo & hi index.  The lo &
'           hi index indicates the span of indices to use to check for valid
'           dates for receivables or contract line splits.
'
'           The participant date table contains an array of participant percentages
'           as many as there are defined beginning with a start date requested.
'           Each entry has the participant, sales source, start/end dates of split,
'           and pct of split.
'
'       <input> llStartDate :  earliest date to retrieve participant dates.  Get that one
'            and all future ones
'            blOwnerOnly -when building the table of participants, if by Vehicle Group Participants for Vehicle option,
'                   force to 100% to for owners share; optional parameters and default to false if not present
'       <output>  tlPIFKEY - array of vehicle keys which refernce the Pct table (tlPIFPCT)
'                 tlPIFPCT - array of entries containing the participants and their splits
Sub gCreatePIFForRpts(llStartDate As Long, tlPIFKey() As PIFKEY, tlPifPct() As PIFPCT, frm As Form, Optional blOwnerOnly As Boolean = False)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  gCreatePifForrptsErr2                                                                 *
'******************************************************************************************
    Dim ilLoopVef As Integer
    Dim ilRet As Integer
    Dim ilLoopPif As Integer
    Dim llPifSDate As Long
    Dim llPifEDate As Long
    Dim ilFirst As Integer
    Dim ilUpperKey As Integer
    Dim ilUpperPct As Integer
    Dim ilVefCode As Integer
    ReDim tlPif(0 To 0) As PIF
    Dim ilLo As Integer         '8-2-09
    Dim ilHi As Integer
    Dim ilSSMnfCode As Integer
    Dim llPifStartDate As Long
    ReDim tlPIFKey(0 To 0) As PIFKEY
    ReDim tlPifPct(0 To 0) As PIFPCT
    hmPif = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hmPif, "", sgDBPath & "Pif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCreatePifForRptsErr
    gBtrvErrorMsg ilRet, "gCreatePifForRpts (btrOpen):" & "Pif.Btr", frm
    On Error GoTo 0
    imPifRecLen = Len(tmPif)

    ilUpperKey = 0
    ilUpperPct = 0
    For ilLoopVef = LBound(tgMVef) To UBound(tgMVef) - 1
        'process only conventional, airing, selling and rep vehicles (dormant or active due to past data)
        If tgMVef(ilLoopVef).sType = "C" Or tgMVef(ilLoopVef).sType = "A" Or tgMVef(ilLoopVef).sType = "S" Or tgMVef(ilLoopVef).sType = "R" Or tgMVef(ilLoopVef).sType = "N" Or tgMVef(ilLoopVef).sType = "G" Then
            ilVefCode = tgMVef(ilLoopVef).iCode
            'get all participant records for a vehicle
            ilRet = gObtainPIF_ForVef(hmPif, ilVefCode, tlPif())
            'loop thru the participant records for the 1 vehicle and gather all records
            'whose start date is active and thru the future
            ilFirst = True
            ilSSMnfCode = -1
            llPifStartDate = -1
            For ilLoopPif = 0 To UBound(tlPif) - 1
                tmPif = tlPif(ilLoopPif)
                gUnpackDateLong tmPif.iStartDate(0), tmPif.iStartDate(1), llPifSDate
                gUnpackDateLong tmPif.iEndDate(0), tmPif.iEndDate(1), llPifEDate
                'find the closest participant date record.  if 1/1/1970 was sent as the earliest date, get all dates.  (1/1/70 is earliest system date)
                If (llPifEDate >= llStartDate And llStartDate <= llPifEDate) Or (Format(llStartDate, "m/d/yy") = "1/1/70") Then
                    If blOwnerOnly = False Then             'everything is split
                        If ilFirst Then
                            ilFirst = False
                            'create the key entry for the vehicle
                            tlPIFKey(ilUpperKey).iVefCode = ilVefCode
                            tlPIFKey(ilUpperKey).iLoInx = ilUpperPct
                            tlPIFKey(ilUpperKey).iHiInx = 0
                        End If
                        tlPifPct(ilUpperPct).iMnfGroup = tmPif.iMnfGroup     'participant
                        tlPifPct(ilUpperPct).iSSMnfCode = tmPif.iMnfSSCode   'sales source
                        tlPifPct(ilUpperPct).iPct = tmPif.iProdPct           'participant split pct
                       
                        tlPifPct(ilUpperPct).lStartDate = llPifSDate         'split effective start date
                        tlPifPct(ilUpperPct).lEndDate = llPifEDate           'split effective end date
                        tlPifPct(ilUpperPct).iOwnerSeq = tmPif.iSeqNo       '5-11-09
                        tlPifPct(ilUpperPct).iOwnerByDate = 0               '8-2-09 init, later to determine which is owner when date changes
                        tlPIFKey(ilUpperKey).iHiInx = ilUpperPct             'entry into the pct split table, vehicle table determines how many date based on this entry
                        ilUpperPct = ilUpperPct + 1
                        ReDim Preserve tlPifPct(0 To ilUpperPct) As PIFPCT
                    Else                                                'option is by vehicle with participant vehicle group; need to show 100%, no splits unless receivables is already split
                        If ilFirst Then
                            ilFirst = False
                            'create the key entry for the vehicle
                            tlPIFKey(ilUpperKey).iVefCode = ilVefCode
                            tlPIFKey(ilUpperKey).iLoInx = ilUpperPct
                            tlPIFKey(ilUpperKey).iHiInx = 0
                        End If
                        If ilSSMnfCode <> tmPif.iMnfSSCode Or llPifStartDate <> llPifSDate Then 'test change in sales source as well as date change
                            tlPifPct(ilUpperPct).iMnfGroup = tmPif.iMnfGroup     'participant
                            tlPifPct(ilUpperPct).iSSMnfCode = tmPif.iMnfSSCode   'sales source
                            tlPifPct(ilUpperPct).iPct = tmPif.iProdPct           'participant split pct
                            tlPifPct(ilUpperPct).iPct = 10000                    '3-31-16 force 100% for owner (report will have no splits)
                            tlPifPct(ilUpperPct).lStartDate = llPifSDate         'split effective start date
                            tlPifPct(ilUpperPct).lEndDate = llPifEDate           'split effective end date
                            tlPifPct(ilUpperPct).iOwnerSeq = tmPif.iSeqNo       '5-11-09
                            tlPifPct(ilUpperPct).iOwnerByDate = 0               '8-2-09 init, later to determine which is owner when date changes
                            tlPIFKey(ilUpperKey).iHiInx = ilUpperPct             'entry into the pct split table, vehicle table determines how many date based on this entry
                            ilSSMnfCode = tmPif.iMnfSSCode                      'use only the first participant of the different sales sources
                            llPifStartDate = llPifSDate
                            ilUpperPct = ilUpperPct + 1
                            ReDim Preserve tlPifPct(0 To ilUpperPct) As PIFPCT
                        End If
                    End If
                End If
            Next ilLoopPif
            'all the percent records have been processed for 1 vehicle
            If Not ilFirst Then              'if not the first time for this vehicle, then at least 1 PIF was found
                ilUpperKey = ilUpperKey + 1
                ReDim Preserve tlPIFKey(0 To ilUpperKey) As PIFKEY
            End If
        End If
    Next ilLoopVef
    
    '8-2-09  Loop thru the tables and determine which of the entries are owners.
    'When there is a date change in participant, the particpant (owner) may also change.
    'that participant needs to be processed
   
    For ilLoopVef = 0 To ilUpperKey
        ilLo = tlPIFKey(ilLoopVef).iLoInx
        ilHi = tlPIFKey(ilLoopVef).iHiInx
        ilFirst = True
         For ilLoopPif = ilLo To ilHi
            If ilFirst Then
                llPifStartDate = tlPifPct(ilLo).lStartDate
                ilSSMnfCode = tlPifPct(ilLo).iSSMnfCode
                tlPifPct(ilLo).iOwnerByDate = 1
                ilFirst = False
            End If
            If (llPifStartDate <> tlPifPct(ilLoopPif).lStartDate) Or (ilSSMnfCode <> tlPifPct(ilLoopPif).iSSMnfCode) Then
                llPifStartDate = tlPifPct(ilLoopPif).lStartDate
                ilSSMnfCode = tlPifPct(ilLoopPif).iSSMnfCode
                tlPifPct(ilLoopPif).iOwnerByDate = 1
            End If
        Next ilLoopPif
    Next ilLoopVef
    btrDestroy hmPif
    ilRet = btrClose(hmPif)
    Exit Sub
gCreatePifForRptsErr:
    btrDestroy hmPif
    gDbg_HandleError "RifRptSub: gCreatePIFForRpts"

gCreatePifForrptsErr2: 'VBC NR
    ilRet = 1
    Resume Next
End Sub

'           gBinarySearchPIFKey - find the matching vehicle code in array that contains
'           index references to participant pct array
'           <input> ilVefcode = vehicle code to match
'                   tlPIFKEY - vehicle table to search
'           return - index to matching vehicle entry
'                    -1 if not found
Public Function gBinarySearchPIFKey(ilVefCode As Integer, tlPIFKey() As PIFKEY) As Integer
    Dim ilMiddle As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    ilMin = LBound(tlPIFKey)
    ilMax = UBound(tlPIFKey)
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilVefCode = tlPIFKey(ilMiddle).iVefCode Then
            'found the match
            gBinarySearchPIFKey = ilMiddle
            Exit Function
        ElseIf ilVefCode < tlPIFKey(ilMiddle).iVefCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchPIFKey = -1
End Function

'           gGetOneVehAllyearForMG - obtain the participant info for a vehicle
'           that is not on the contract processing; add it to the list of contract
'           vehicle incase more mg/outside spots
'           <input> - start date for the mg/outside spot
Public Sub gGetOneVehAllYearForMG(ilVefCode As Integer, llStdStartDates() As Long, ilSSCode As Integer, tlPIFKey() As PIFKEY, tlPifPct() As PIFPCT, tlCntAllYear() As ALLPIFPCTYEAR)
    Dim ilVefInx As Integer
    Dim ilLo As Integer
    Dim ilHi As Integer
    Dim ilVefSpan As Integer
    Dim ilUpper As Integer
    Dim llDate As Long
    Dim ilLoop12Months As Integer
    Dim ilValidDate As Integer

    ilVefInx = gBinarySearchPIFKey(ilVefCode, tlPIFKey())
    If ilVefInx >= 0 Then
        llDate = llStdStartDates(1)

        ilLo = tlPIFKey(ilVefInx).iLoInx     'get the starting/ending index for all the participants that belong with this vehicle
        ilHi = tlPIFKey(ilVefInx).iHiInx
        ilUpper = UBound(tlCntAllYear)
        For ilVefSpan = ilLo To ilHi
            If tlPifPct(ilVefSpan).iSSMnfCode = ilSSCode Then        'find only those participants that match the sales source of contract
                ilValidDate = False
                For ilLoop12Months = 1 To 12
                    llStartDate = llStdStartDates(ilLoop12Months)
                    'verify the effective date of the participant pct
                    'If llDate >= tlPIFPCT(ilVefInx).lStartDate And llDate <= tlPIFPCT(ilVefInx).lEndDate Then
                    '8-28-09 fixed to get the dates from correct vehicle index
                    If llStartDate >= tlPifPct(ilVefSpan).lStartDate And llStartDate <= tlPifPct(ilVefSpan).lEndDate Then
                        tlCntAllYear(ilUpper).AllYear.iVefCode = tlPIFKey(ilVefInx).iVefCode
                        tlCntAllYear(ilUpper).AllYear.iSSMnfCode = tlPifPct(ilVefSpan).iSSMnfCode
                        tlCntAllYear(ilUpper).AllYear.iMnfGroup = tlPifPct(ilVefSpan).iMnfGroup
                        tlCntAllYear(ilUpper).AllYear.iOwnerSeq = tlPifPct(ilVefSpan).iOwnerSeq     '5-11-09
                        tlCntAllYear(ilUpper).AllYear.iOwnerByDate = tlPifPct(ilVefSpan).iOwnerByDate
                        tlCntAllYear(ilUpper).AllYear.iPct(ilLoop12Months) = tlPifPct(ilVefSpan).iPct
                        ilValidDate = True
                    End If
                Next ilLoop12Months
                If ilValidDate Then
                    ilUpper = ilUpper + 1
                    ReDim Preserve tlCntAllYear(LBound(tlCntAllYear) To ilUpper) As ALLPIFPCTYEAR
                End If
            End If
        Next ilVefSpan
    End If
    Exit Sub
End Sub

'           gObtainOwnerPctForSS - get the Owners share of the participant based on
'           the sales source.  Return the share %
'           <input>  ilVefCode - vehicle code
'                    ilSS - sales source to match
'                    tlPIFKey() - array of the keys containing the vehicle and pointers to
'                    the array containing PIF records (in date order)
'           <return> Share % from PIF for owner only
Public Function gObtainOwnerPctForSS(ilVefCode As Integer, ilSS As Integer, ilDate() As Integer, tlPIFKey() As PIFKEY, tlPifPct() As PIFPCT) As Integer
    Dim ilLoop As Integer
    Dim ilVefInx As Integer
    Dim ilLo As Integer
    Dim ilHi As Integer
    Dim llDate As Long
    Dim slStr As String
        
    gObtainOwnerPctForSS = 0
    'find vehicle in participant key table
    ilVefInx = gBinarySearchPIFKey(ilVefCode, tlPIFKey())
    If ilVefInx < 0 Then                'no vehicle found, dont process it
        Exit Function
    Else
        gUnpackDate ilDate(0), ilDate(1), slStr
        llDate = gDateValue(slStr)
    
        ilLo = tlPIFKey(ilVefInx).iLoInx     'get the starting/ending index for all the participants that belong with this vehicle
        ilHi = tlPIFKey(ilVefInx).iHiInx
        'ilMatchingSSIndex = 1
        For ilLoop = ilLo To ilHi
            If tlPifPct(ilLoop).iSSMnfCode = ilSS Then
                  'verify the effective date of the participant pct
                  'If llDate >= tlPIFPCT(ilVefInx).lStartDate And llDate <= tlPIFPCT(ilVefInx).lEndDate Then
                  '3-13-13 wrong index was used
                  If llDate >= tlPifPct(ilLoop).lStartDate And llDate <= tlPifPct(ilLoop).lEndDate Then
                      gObtainOwnerPctForSS = tlPifPct(ilLoop).iPct
                      Exit Function
                  End If
            End If
        Next ilLoop
    End If
    Exit Function
End Function

'           gObtainOwnerForSS - get the Owner of the vehicle based on sales source and participant date tables
'           since owners can change from one date to another.  Cannot use the Owner pointer in vehicle when that case occurs.
'           <input>  ilVefCode - vehicle code
'                    ilSS - sales source to match
'                    tlPIFKey() - array of the keys containing the vehicle and pointers to
'                    the array containing PIF records (in date order)
'           <output> None
'           <return> owner mnf
Public Function gObtainOwnerForSS(ilVefCode As Integer, ilSS As Integer, llDate As Long, tlPIFKey() As PIFKEY, tlPifPct() As PIFPCT) As Integer
    Dim ilLoop As Integer
    Dim ilVefInx As Integer
    Dim ilLo As Integer
    Dim ilHi As Integer
    Dim slStr As String

    gObtainOwnerForSS = 0
    'find vehicle in participant key table
    ilVefInx = gBinarySearchPIFKey(ilVefCode, tlPIFKey())
    If ilVefInx < 0 Then                'no vehicle found, dont process it
        Exit Function
    Else
        ilLo = tlPIFKey(ilVefInx).iLoInx     'get the starting/ending index for all the participants that belong with this vehicle
        ilHi = tlPIFKey(ilVefInx).iHiInx
        For ilLoop = ilLo To ilHi
            If tlPifPct(ilLoop).iSSMnfCode = ilSS Then
                  'verify the effective date of the participant pct
                  If llDate >= tlPifPct(ilLoop).lStartDate And llDate <= tlPifPct(ilLoop).lEndDate Then
                      gObtainOwnerForSS = tlPifPct(ilLoop).iMnfGroup
                      Exit Function
                  End If
            End If
        Next ilLoop
    End If
    Exit Function
End Function

