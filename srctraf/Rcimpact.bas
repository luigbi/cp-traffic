Attribute VB_Name = "RCIMPACTSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rcimpact.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RCImpact.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Date/Time subs and functions
Option Explicit
Option Compare Text
Type IMPACTREC
    sKey As String * 90 '
    sVehicle As String * 50
    sDaypart As String * 30
    iVefCode As Integer
    iRdfCode As Integer
    iRcfCode As Integer
    iPtDollarRec As Integer
    lPtRifRec(0 To 1) As Long
End Type
Type DOLLARREC
    iNoHits As Integer
    lRCPrice As Long    'Rate Card Price from Flight
    lSPrice As Long     'Average Spot Price from flights
    lTSPrice As Long     'Total Spot Price (# Spots * Price) from flights
    iTSpots As Integer  'Total Spots Sold from flights (as 30's)
    l30Inv As Long      '# 30 Inventory
    l30Sold As Long     '# 30's sold
    lDollarSold As Long 'Dollars Sold
    lBudget As Long     'Budget Dollars
    iAvailDefined As Integer    '1=Avails defined;0=Avail Missing
End Type
Type RCSPOTINFO
    iLen As Integer
    lPrice As Long
    iRank As Integer
    iRecType As Integer
End Type
Type DPBUDGETINFO
    lRifIndex As Long
    iRdfIndex As Integer
    iInv As Integer 'Inventory
    lRCPrice As Long  'Rate Card Daypart Price
    lPrice As Long
    lBudget As Long
    iAvailDefined As Integer    '1=Avails defined;0=Avail Missing
End Type
Public tgImpactRec() As IMPACTREC
Public tgDollarRec() As DOLLARREC
Dim tmDPBudgetInfo() As DPBUDGETINFO
'Spot
Dim imSdfRecLen As Integer  'Record length
Dim tmSdf As SDF
Dim tmSdfSrchKey1 As SDFKEY1
Dim tmSdfSrchKey3 As LONGKEY0
'Contract
Dim tmChf As CHF            'CHF record image
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer  'CHF record length
'Line
Dim tmClf As CLF            'CLF record image
Dim tmClfSrchKey As CLFKEY0
Dim imClfRecLen As Integer  'CLF record length
'Budget information
Dim tmBvfSrchKey As BVFKEY0    'Rcf key record image
Dim imBvfRecLen As Integer        'Rcf record length
Dim tmBvfVeh() As BVF   'Budget by vehicle
Dim tmBvfVeh2() As BVF   'Budget by vehicle
Dim tmBvf As BVF
Dim imVefCode() As Integer
'Spot Summary
Public tgRCSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Const LBONE = 1


'*******************************************************
'*                                                     *
'*      Procedure Name:mBdGetBudgetDollars             *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get budget dollars             *
'*                                                     *
'*            Note: Similar code in RateCard.Frm       *
'*                                                     *
'*******************************************************
Sub mBdGetBudgetDollars(ilSplitType As Integer, hlChf As Integer, hlClf As Integer, hlCff As Integer, hlSdf As Integer, hlSmf As Integer, hlVef As Integer, hlVsf As Integer, hlSsf As Integer, hlBvf As Integer, hlLcf As Integer, ilBdMnfCode() As Integer, ilBdYear() As Integer, llBdStartDate As Long, llBdEndDate As Long, tlRif() As RIF, tlRdf() As RDF)
'
'       ilSplitType(I) 0=No Split (Rate card and Budget match; 1=Split Rate Card (Year based on
'                      Budget year; 2=Split Budget (Year base on Rate Card)
'                      Matching
'                                 10  11  12   1   2   3   4   5   6   7   8   9
'                      Rate Card  **********************************************
'                      Budget     **********************************************
'                      Split Rate Card
'                                 10  11  12   1   2   3   4   5   6   7   8   9
'                      Budget     **********************************************
'                      Rate Card  ----------   *********************************
'                      Split Budget
'                                 1   2   3   4   5   6   7   8   9  10  11  12
'                      Rate Card  *********************************************
'                      Budget     *********************************  ++++++++++
'                      Note: * = Matching Year; - = Year - 1; + = Year + 1
'       hlChf(I)- CHF Handle
'       hlClf(I)- CLF Handle
'       hlCff(I)- CFF Handle
'       hlSdf(I)- SDF Handle
'       hlSmf(I)- SMF Handle
'       hlVef(I)- VEF Handle
'       hlVsf(I)- VSF Handle
'       hlSsf(I)- SSF Handle
'       hlBvf(I)- BVF Handle
'       ilMnfCode()(I)-Budget Name Code (0 To 0 or 0 to 1 if split budget)
'       ilYears()(I)-Year to retrieve (see ilMnfCode above)
'       llBdStartDate(I)- Start Date
'       llBdEndDate(I)- End date Note: start and end date must be within the same year
'       tlRif(I)- RIF records associated with the rate card
'       tlRdf(I)- Daypart records
'
    Dim ilLoop As Integer
    Dim ilVef As Integer
    Dim ilFound As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilWkIndex As Integer
    Dim ilRCWkNo As Integer
    Dim ilBdWkNo As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilFirstLastWk As Integer
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim ilBvf As Integer
    Dim ilIndex As Integer
    Dim ilNo30 As Integer
    Dim ilNo60 As Integer
    Dim ilLen As Integer
    Dim ilUnits As Integer
    Dim ilAvailOk As Integer
    Dim llSsfDate As Long
    Dim llTime As Long
    Dim llRdfStartTime As Long
    Dim llRdfEndTime As Long
    Dim ilRet As Integer
    Dim ilTime As Integer
    Dim ilDay As Integer
    Dim ilVpfIndex As Integer
    Dim ilDP As Integer
    Dim llPrice As Long
    Dim llBudget As Long
    Dim llDPInv As Long
    Dim ilDPNotZero As Integer
    Dim ilBdNoWks As Integer
    Dim llSplitDate As Long     'Start Date of next Budget or Rate Card if Split
    ReDim llStartDateBd(0 To 1) As Long
    ReDim llEndDateBd(0 To 1) As Long
    ReDim ilYearBd(0 To 1) As Integer
    Dim llLatestDate As Long
    Dim llLpDate As Long
    Dim ilType As Integer
    Dim ilVefIndex As Integer
    Dim ilSpot As Integer
    Dim ilImpact As Integer

    imSdfRecLen = Len(tmSdf)
    imCHFRecLen = Len(tmChf)
    imClfRecLen = Len(tmClf)
    imBvfRecLen = Len(tmBvf)
    ilBdNoWks = (llBdEndDate - llBdStartDate) / 7 + 1
    ReDim tgImpactRec(0 To 1) As IMPACTREC
    ReDim tgDollarRec(0 To ilBdNoWks, 0 To 1) As DOLLARREC
    ReDim tmBvfVeh2(0 To 0) As BVF
    ilYearBd(1) = -1
    If ilSplitType = 0 Then 'None
        ilYearBd(0) = ilBdYear(LBound(ilBdMnfCode))
        slDate = "12/15/" & Trim$(str$(ilBdYear(LBound(ilBdMnfCode))))
        slDate = gObtainEndStd(slDate)
        llSplitDate = gDateValue(slDate) + 1
        If Not mReadBvfRec(hlBvf, ilBdMnfCode(LBound(ilBdMnfCode)), ilBdYear(LBound(ilBdMnfCode)), tmBvfVeh()) Then
            Exit Sub
        End If
    ElseIf ilSplitType = 1 Then 'Rate Card
        ilYearBd(0) = ilBdYear(LBound(ilBdMnfCode))
        slDate = "1/15/" & Trim$(str$(ilBdYear(LBound(ilBdMnfCode))))
        slDate = gObtainStartStd(slDate)
        llSplitDate = gDateValue(slDate)
        If Not mReadBvfRec(hlBvf, ilBdMnfCode(LBound(ilBdMnfCode)), ilBdYear(LBound(ilBdMnfCode)), tmBvfVeh()) Then
            Exit Sub
        End If
    Else    'Budget
        'End end of year and add one
        llSplitDate = -1
        If LBound(ilBdMnfCode) < UBound(ilBdMnfCode) Then
            If ilBdYear(LBound(ilBdMnfCode)) < ilBdYear(LBound(ilBdMnfCode) + 1) Then
                ilYearBd(0) = ilBdYear(LBound(ilBdMnfCode))
                ilYearBd(1) = ilBdYear(LBound(ilBdMnfCode) + 1)
                If Not mReadBvfRec(hlBvf, ilBdMnfCode(LBound(ilBdMnfCode)), ilBdYear(LBound(ilBdMnfCode)), tmBvfVeh()) Then
                    Exit Sub
                End If
                If Not mReadBvfRec(hlBvf, ilBdMnfCode(LBound(ilBdMnfCode) + 1), ilBdYear(LBound(ilBdMnfCode) + 1), tmBvfVeh2()) Then
                    Exit Sub
                End If
                For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
                    If ilBdYear(LBound(ilBdMnfCode)) = tgMCof(ilLoop).iYear Then
                        gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iEndDate(1, 11), llSplitDate
                        llSplitDate = llSplitDate + 1
                        Exit For
                    End If
                Next ilLoop
            Else
                ilYearBd(0) = ilBdYear(LBound(ilBdMnfCode) + 1)
                ilYearBd(1) = ilBdYear(LBound(ilBdMnfCode))
                If Not mReadBvfRec(hlBvf, ilBdMnfCode(LBound(ilBdMnfCode)), ilBdYear(LBound(ilBdMnfCode)), tmBvfVeh2()) Then
                    Exit Sub
                End If
                If Not mReadBvfRec(hlBvf, ilBdMnfCode(LBound(ilBdMnfCode) + 1), ilBdYear(LBound(ilBdMnfCode) + 1), tmBvfVeh()) Then
                    Exit Sub
                End If
                For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
                    If ilBdYear(LBound(ilBdMnfCode) + 1) = tgMCof(ilLoop).iYear Then
                        gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iEndDate(1, 11), llSplitDate
                        llSplitDate = llSplitDate + 1
                        Exit For
                    End If
                Next ilLoop
            End If
            For ilLoop = LBound(tmBvfVeh2) To UBound(tmBvfVeh2) - 1 Step 1
                ilFound = False
                For ilBvf = LBound(tmBvfVeh) To UBound(tmBvfVeh) - 1 Step 1
                    If tmBvfVeh2(ilLoop).iVefCode = tmBvfVeh(ilBvf).iVefCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilBvf
                If Not ilFound Then
                    tmBvfVeh(UBound(tmBvfVeh)) = tmBvfVeh2(ilLoop)
                    ReDim Preserve tmBvfVeh(LBound(tmBvfVeh) To UBound(tmBvfVeh) + 1) As BVF
                End If
            Next ilLoop
        Else
            ilYearBd(0) = ilBdYear(LBound(ilBdMnfCode))
            If Not mReadBvfRec(hlBvf, ilBdMnfCode(LBound(ilBdMnfCode)), ilBdYear(LBound(ilBdMnfCode)), tmBvfVeh()) Then
                Exit Sub
            End If
            For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
                If ilBdYear(LBound(ilBdMnfCode)) = tgMCof(ilLoop).iYear Then
                    gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iEndDate(1, 11), llSplitDate
                    llSplitDate = llSplitDate + 1
                    Exit For
                End If
            Next ilLoop
        End If
        If llSplitDate = -1 Then
            Exit Sub
        End If
    End If
    If UBound(tmBvfVeh) = LBound(tmBvfVeh) Then
        Exit Sub
    End If
    For ilLoop = 0 To 1 Step 1
        If ilYearBd(ilLoop) <> -1 Then
            If tgSpf.sRUseCorpCal <> "Y" Then
                slDate = "1/15/" & Trim$(str$(ilYearBd(ilLoop)))
                slDate = gObtainYearStartDate(0, slDate)
                llStartDateBd(ilLoop) = gDateValue(slDate)
                slDate = "1/15/" & Trim$(str$(ilYearBd(ilLoop)))
                slDate = gObtainYearEndDate(0, slDate)
                llEndDateBd(ilLoop) = gDateValue(slDate)
            Else
                slDate = "1/15/" & Trim$(str$(ilYearBd(ilLoop)))
                slDate = gObtainYearStartDate(5, slDate)
                llStartDateBd(ilLoop) = gDateValue(slDate)
                slDate = "1/15/" & Trim$(str$(ilYearBd(ilLoop)))
                slDate = gObtainYearEndDate(5, slDate)
                llEndDateBd(ilLoop) = gDateValue(slDate)
            End If
        Else
            llStartDateBd(ilLoop) = -1
            llEndDateBd(ilLoop) = -1
        End If
    Next ilLoop
    For llRif = LBound(tlRif) To UBound(tlRif) - 1 Step 1
        For ilRdf = LBound(tlRdf) To UBound(tlRdf) - 1 Step 1
            If tlRif(llRif).iRdfCode = tlRdf(ilRdf).iCode Then
                'If tlRdf(ilRdf).sBase = "Y" Then
                If tlRif(llRif).sBase = "Y" Then
                    ilFound = False
                    If ilSplitType = 1 Then 'Rate Card
                        For ilLoop = LBONE To UBound(tgImpactRec) - 1 Step 1
                            If (tgImpactRec(ilLoop).iVefCode = tlRif(llRif).iVefCode) And (tgImpactRec(ilLoop).iRdfCode = tlRif(llRif).iRdfCode) Then
                                ilFound = True
                                If tlRif(llRif).iYear < tlRif(tgImpactRec(ilLoop).lPtRifRec(0)).iYear Then
                                    tgImpactRec(ilLoop).lPtRifRec(1) = tgImpactRec(ilLoop).lPtRifRec(0)
                                    tgImpactRec(ilLoop).lPtRifRec(0) = llRif
                                Else
                                    tgImpactRec(ilLoop).lPtRifRec(1) = llRif
                                End If
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If Not ilFound Then
                        'Exclude Sports for Now
                        ilVef = gBinarySearchVef(tlRif(llRif).iVefCode)
                        If ilVef <> -1 Then
                            If tgMVef(ilVef).sType = "G" Then
                                'ilFound = True
                            End If
                        End If
                    Else
                        ilFound = True
                    End If
                    If Not ilFound Then
                        ilIndex = UBound(tgImpactRec)
                        tgImpactRec(ilIndex).iVefCode = tlRif(llRif).iVefCode
                        tgImpactRec(ilIndex).iRdfCode = tlRif(llRif).iRdfCode
                        'tgImpactRec(ilIndex).iRcfCode = tlRif(ilRif).iRcfCode
                        tgImpactRec(ilIndex).iPtDollarRec = ilIndex
                        tgImpactRec(ilIndex).lPtRifRec(0) = llRif
                        ReDim Preserve tgImpactRec(0 To ilIndex + 1) As IMPACTREC
                        ReDim Preserve tgDollarRec(0 To ilBdNoWks, 0 To ilIndex + 1) As DOLLARREC
                    End If
                End If
            End If
        Next ilRdf
    Next llRif
    'Build array of vehicles
    'ReDim imVefCode(1 To 1) As Integer
    ReDim imVefCode(0 To 0) As Integer
    For ilVef = LBONE To UBound(tgImpactRec) - 1 Step 1
        ilFound = False
        For ilLoop = LBound(imVefCode) To UBound(imVefCode) - 1 Step 1
            If imVefCode(ilLoop) = tgImpactRec(ilVef).iVefCode Then
                ilFound = True
            End If
        Next ilLoop
        If Not ilFound Then
            imVefCode(UBound(imVefCode)) = tgImpactRec(ilVef).iVefCode
            'ReDim Preserve imVefCode(1 To UBound(imVefCode) + 1) As Integer
            ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
        End If
    Next ilVef
    'Count number of Base dayparts within the vehicle
    For ilVef = LBound(imVefCode) To UBound(imVefCode) - 1 Step 1
        ReDim tmDPBudgetInfo(0 To 0) As DPBUDGETINFO
        For llRif = LBound(tlRif) To UBound(tlRif) - 1 Step 1
            'If (tlRif(ilRif).iRcfCode = tgImpactRec(1).iRcfCode) And (tlRif(ilRif).iVefCode = imVefCode(ilVef)) And (tlRif(ilRif).iYear = ilBdYear) Then
            If (tlRif(llRif).iVefCode = imVefCode(ilVef)) Then
                'Test if daypart is base daypart
                'For ilRdf = 1 To UBound(tlRdf) - 1 Step 1
                For ilRdf = LBound(tlRdf) To UBound(tlRdf) - 1 Step 1
                    If tlRif(llRif).iRdfCode = tlRdf(ilRdf).iCode Then
                        'If tlRdf(ilRdf).sBase = "Y" Then
                        If tlRif(llRif).sBase = "Y" Then
                            tmDPBudgetInfo(UBound(tmDPBudgetInfo)).lRifIndex = llRif
                            tmDPBudgetInfo(UBound(tmDPBudgetInfo)).iRdfIndex = ilRdf
                            tmDPBudgetInfo(UBound(tmDPBudgetInfo)).iInv = 0
                            tmDPBudgetInfo(UBound(tmDPBudgetInfo)).lRCPrice = 0
                            tmDPBudgetInfo(UBound(tmDPBudgetInfo)).lPrice = 0
                            tmDPBudgetInfo(UBound(tmDPBudgetInfo)).lBudget = 0
                            tmDPBudgetInfo(UBound(tmDPBudgetInfo)).iAvailDefined = 0
                            ReDim Preserve tmDPBudgetInfo(0 To UBound(tmDPBudgetInfo) + 1) As DPBUDGETINFO
                        End If
                        Exit For
                    End If
                Next ilRdf
            End If
        Next llRif
        If UBound(tmDPBudgetInfo) > LBound(tmDPBudgetInfo) Then
            For ilBvf = LBound(tmBvfVeh) To UBound(tmBvfVeh) - 1 Step 1
                If tmBvfVeh(ilBvf).iVefCode = imVefCode(ilVef) Then
                    'ilVpfIndex = gVpfFind(RateCard, imVefCode(ilVef))
                    ilVpfIndex = -1 'gVpfFind(RateCard, imVefCode(ilVef))
                    'For ilLoop = 0 To UBound(tgVpf) Step 1
                    '    If imVefCode(ilVef) = tgVpf(ilLoop).iVefKCode Then
                        ilLoop = gBinarySearchVpf(imVefCode(ilVef))
                        If ilLoop <> -1 Then
                            ilVpfIndex = ilLoop
                    '        Exit For
                        End If
                    'Next ilLoop
                    ilVefIndex = gBinarySearchVef(imVefCode(ilVef))
                    llLatestDate = gGetLatestLCFDate(hlLcf, "C", imVefCode(ilVef))
                    For llDate = llBdStartDate To llBdEndDate Step 7
                        slDate = Format$(llDate, "m/d/yy")
                        gObtainMonthYear 0, slDate, ilMonth, ilYear
                        If ilSplitType = 0 Then 'None
                            gObtainWkNo 0, slDate, ilRCWkNo, ilFirstLastWk
                            gObtainWkNo 4, slDate, ilBdWkNo, ilFirstLastWk
                        ElseIf ilSplitType = 1 Then 'Rate Card
                            gObtainWkNo 0, slDate, ilRCWkNo, ilFirstLastWk
                            gObtainWkNo 5, slDate, ilBdWkNo, ilFirstLastWk
                        Else    'Budget
                            gObtainWkNo 0, slDate, ilRCWkNo, ilFirstLastWk
                            gObtainWkNo 5, slDate, ilBdWkNo, ilFirstLastWk
                        End If
                        ilWkIndex = (llDate - llBdStartDate) \ 7 + 1

                        For llLpDate = llDate To llDate + 6 Step 1
                            slDate = Format$(llLpDate, "m/d/yy")
                            If tgMVef(ilVefIndex).sType <> "G" Then
                                ilType = 0
                                tmSsfSrchKey.iType = 0
                                tmSsfSrchKey.iVefCode = imVefCode(ilVef)
                                gPackDate slDate, tmSsfSrchKey.iDate(0), tmSsfSrchKey.iDate(1)
                                tmSsfSrchKey.iStartTime(0) = 0
                                tmSsfSrchKey.iStartTime(1) = 0
                                imSsfRecLen = Len(tgRCSsf)
                                ilRet = gSSFGetGreaterOrEqual(hlSsf, tgRCSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                            Else
                                tmSsfSrchKey2.iVefCode = imVefCode(ilVef)
                                gPackDate slDate, tmSsfSrchKey2.iDate(0), tmSsfSrchKey2.iDate(1)
                                imSsfRecLen = Len(tgRCSsf)
                                ilRet = gSSFGetGreaterOrEqualKey2(hlSsf, tgRCSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
                                ilType = tgRCSsf.iType
                            End If
                            ilRet = gBuildAvails(ilRet, imVefCode(ilVef), llLpDate, llLatestDate, ilType)
                            Do While (ilRet = BTRV_ERR_NONE) And (tgRCSsf.iType = ilType) And (tgRCSsf.iVefCode = imVefCode(ilVef))
                                gUnpackDateLong tgRCSsf.iDate(0), tgRCSsf.iDate(1), llSsfDate
                                If llSsfDate <> llLpDate Then
                                    Exit Do
                                End If
                                ilDay = gWeekDayLong(llSsfDate)
                                For ilLoop = 1 To tgRCSsf.iCount Step 1
                                   LSet tmAvail = tgRCSsf.tPas(ADJSSFPASBZ + ilLoop)
                                    'If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                                    If tmAvail.iRecType = 2 Then    'Cmml Avails only
                                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                        For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                                            If ilYear = tlRif(tmDPBudgetInfo(ilDP).lRifIndex).iYear Then    'ilBdYear Then
                                                ilRdf = tmDPBudgetInfo(ilDP).iRdfIndex
                                                tmDPBudgetInfo(ilDP).iAvailDefined = 1
                                                ilAvailOk = True
                                                If (tlRdf(ilRdf).sInOut = "I") Then
                                                    If (tlRdf(ilRdf).ianfCode <> tmAvail.ianfCode) Then
                                                        ilAvailOk = False
                                                    End If
                                                End If
                                                If (tlRdf(ilRdf).sInOut = "O") Then
                                                    If tmAvail.ianfCode = tlRdf(ilRdf).ianfCode Then
                                                        ilAvailOk = False
                                                    End If
                                                End If
                                                If ilAvailOk Then
                                                    ilAvailOk = False
                                                    For ilTime = LBound(tlRdf(ilRdf).iStartTime, 2) To UBound(tlRdf(ilRdf).iStartTime, 2) Step 1
                                                        If (tlRdf(ilRdf).iStartTime(0, ilTime) <> 1) Or (tlRdf(ilRdf).iStartTime(1, ilTime) <> 0) Then
                                                            'If tlRdf(ilRdf).sWkDays(ilTime, ilDay + 1) = "Y" Then
                                                            If tlRdf(ilRdf).sWkDays(ilTime, ilDay) = "Y" Then
                                                                gUnpackTimeLong tlRdf(ilRdf).iStartTime(0, ilTime), tlRdf(ilRdf).iStartTime(1, ilTime), False, llRdfStartTime
                                                                gUnpackTimeLong tlRdf(ilRdf).iEndTime(0, ilTime), tlRdf(ilRdf).iEndTime(1, ilTime), True, llRdfEndTime
                                                                If (llTime >= llRdfStartTime) And (llTime < llRdfEndTime) Then
                                                                    ilAvailOk = True
                                                                End If
                                                            End If
                                                        End If
                                                    Next ilTime
                                                End If
                                                If ilAvailOk Then
                                                    ilLen = tmAvail.iLen
                                                    ilUnits = tmAvail.iAvInfo And &H1F
                                                    ilNo30 = 0
                                                    ilNo60 = 0
                                                    If ilLen >= 30 Then
                                                        If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                                            If (ilLen Mod 30) = 0 Then
                                                                Do While ilLen >= 30
                                                                    ilNo30 = ilNo30 + 1
                                                                    ilLen = ilLen - 30
                                                                Loop
                                                            End If
                                                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                                            If (ilLen Mod 30) = 0 Then
                                                                Do While ilLen >= 30
                                                                    ilNo30 = ilNo30 + 1
                                                                    ilLen = ilLen - 30
                                                                Loop
                                                            End If
                                                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                                            If (ilLen Mod 30) = 0 Then
                                                                Do While ilLen >= 30
                                                                    ilNo30 = ilNo30 + 1
                                                                    ilLen = ilLen - 30
                                                                Loop
                                                            End If
                                                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                                        End If
                                                    Else
                                                        ilNo30 = 1
                                                    End If
                                                    'Add
                                                    tmDPBudgetInfo(ilDP).iInv = tmDPBudgetInfo(ilDP).iInv + ilNo30

                                                    'Set Spots
                                                    For ilSpot = ilLoop + 1 To ilLoop + tmAvail.iNoSpotsThis Step 1
                                                       LSet tmSpot = tgRCSsf.tPas(ADJSSFPASBZ + ilSpot)
                                                        tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                        ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            ilNo30 = 0
                                                            ilNo60 = 0
                                                            'ilLen = tmSpot.iPosLen And &HFFF
                                                            ilLen = tmSdf.iLen
                                                            If ilLen >= 30 Then
                                                                If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                                                    If (ilLen Mod 30) = 0 Then
                                                                        Do While ilLen >= 30
                                                                            ilNo30 = ilNo30 + 1
                                                                            ilLen = ilLen - 30
                                                                        Loop
                                                                    End If
                                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                                                    If (ilLen Mod 30) = 0 Then
                                                                        Do While ilLen >= 30
                                                                            ilNo30 = ilNo30 + 1
                                                                            ilLen = ilLen - 30
                                                                        Loop
                                                                    End If
                                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                                                    If (ilLen Mod 30) = 0 Then
                                                                        Do While ilLen >= 30
                                                                            ilNo30 = ilNo30 + 1
                                                                            ilLen = ilLen - 30
                                                                        Loop
                                                                    End If
                                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                                                End If
                                                            Else
                                                                ilNo30 = 1
                                                            End If
                                                            tmChfSrchKey.lCode = tmSdf.lChfCode
                                                            ilRet = btrGetEqual(hlChf, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                            If (ilRet = BTRV_ERR_NONE) And (tmChf.sType <> "M") And (tmChf.sType <> "S") Then
                                                                For ilImpact = LBONE To UBound(tgImpactRec) - 1 Step 1
                                                                    ilRdf = tmDPBudgetInfo(ilDP).iRdfIndex
                                                                    If (tgImpactRec(ilImpact).iVefCode = imVefCode(ilVef)) And (tgImpactRec(ilImpact).iRdfCode = tlRdf(ilRdf).iCode) Then
                                                                        tgDollarRec(ilWkIndex, tgImpactRec(ilImpact).iPtDollarRec).l30Sold = tgDollarRec(ilWkIndex, tgImpactRec(ilImpact).iPtDollarRec).l30Sold + ilNo30
                                                                        tgDollarRec(ilWkIndex, tgImpactRec(ilImpact).iPtDollarRec).lDollarSold = tgDollarRec(ilWkIndex, tgImpactRec(ilImpact).iPtDollarRec).lDollarSold + mGetCost(tmSdf, hlClf, hlCff, hlSmf, hlVef, hlVsf) / 100
                                                                        'Exit For
                                                                    End If
                                                                Next ilImpact
                                                            End If
                                                        End If
                                                    Next ilSpot

                                                End If
                                            End If
                                        Next ilDP
                                    End If
                                Next ilLoop
                                imSsfRecLen = Len(tgRCSsf)
                                ilRet = gSSFGetNext(hlSsf, tgRCSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                If tgMVef(ilVefIndex).sType = "G" Then
                                    ilType = tgRCSsf.iType
                                End If
                            Loop
                        Next llLpDate
                        'Compute Budget by daypart
                        For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                            If ilYear = tlRif(tmDPBudgetInfo(ilDP).lRifIndex).iYear Then    'ilBdYear Then
                                tmDPBudgetInfo(ilDP).lRCPrice = tmDPBudgetInfo(ilDP).lRCPrice + tlRif(tmDPBudgetInfo(ilDP).lRifIndex).lRate(ilRCWkNo)
                            End If
                        Next ilDP
                        llBudget = 0
                        If ilSplitType = 0 Then 'None
                            If (llDate >= llStartDateBd(0)) And (llDate <= llEndDateBd(0)) Then
                                llBudget = tmBvfVeh(ilBvf).lGross(ilBdWkNo)
                            End If
                        ElseIf ilSplitType = 1 Then 'Rate Card
                            If (llDate >= llStartDateBd(0)) And (llDate <= llEndDateBd(0)) Then
                                llBudget = tmBvfVeh(ilBvf).lGross(ilBdWkNo)
                            End If
                        Else    'Budget
                            If llDate < llSplitDate Then
                                If (llDate >= llStartDateBd(0)) And (llDate <= llEndDateBd(0)) Then
                                    llBudget = tmBvfVeh(ilBvf).lGross(ilBdWkNo)
                                End If
                            Else
                                'Scan tmBvfVeh2 for matching tmBvfVeh
                                For ilLoop = LBound(tmBvfVeh2) To UBound(tmBvfVeh2) - 1 Step 1
                                    If tmBvfVeh(ilBvf).iVefCode = tmBvfVeh2(ilLoop).iVefCode Then
                                        If (llDate >= llStartDateBd(1)) And (llDate <= llEndDateBd(1)) Then
                                            llBudget = tmBvfVeh2(ilBvf).lGross(ilBdWkNo)
                                        End If
                                        Exit For
                                    End If
                                Next ilLoop
                            End If
                        End If
                        ilDPNotZero = -1
                        For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                            If tmDPBudgetInfo(ilDP).lRCPrice > 0 Then
                                ilDPNotZero = ilDP
                                Exit For
                            End If
                        Next ilDP
                        'Distribute Budget to each Daypart
                        If ilDPNotZero >= LBound(tmDPBudgetInfo) Then
                            llDPInv = 0
                            For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                                '4-12-11 maintain better accuracy
                                llDPInv = llDPInv + (100 * CSng(tmDPBudgetInfo(ilDP).lRCPrice) * tmDPBudgetInfo(ilDP).iInv) / tmDPBudgetInfo(ilDPNotZero).lRCPrice
                            Next ilDP
                            If llDPInv > 0 Then
                                llPrice = (10000 * CSng(llBudget)) / llDPInv
                            Else
                                llPrice = 0
                            End If
                            For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                                tmDPBudgetInfo(ilDP).lPrice = (CSng(tmDPBudgetInfo(ilDP).lRCPrice) * CSng(llPrice)) / tmDPBudgetInfo(ilDPNotZero).lRCPrice
                            Next ilDP
                            For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                                tmDPBudgetInfo(ilDP).lBudget = (tmDPBudgetInfo(ilDP).lPrice * tmDPBudgetInfo(ilDP).iInv) / 100
                                tmDPBudgetInfo(ilDP).lPrice = tmDPBudgetInfo(ilDP).lPrice / 100
                            Next ilDP
                        Else
                            For ilDP = LBound(tmDPBudgetInfo) + 1 To UBound(tmDPBudgetInfo) - 1 Step 1
                                tmDPBudgetInfo(ilDP).lBudget = 0
                            Next ilDP
                        End If
                        ilWkIndex = (llDate - llBdStartDate) \ 7 + 1
                        For ilLoop = LBONE To UBound(tgImpactRec) - 1 Step 1
                            For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                                ilRdf = tmDPBudgetInfo(ilDP).iRdfIndex
                                If (tgImpactRec(ilLoop).iVefCode = imVefCode(ilVef)) And (tgImpactRec(ilLoop).iRdfCode = tlRdf(ilRdf).iCode) Then
                                    tgDollarRec(ilWkIndex, tgImpactRec(ilLoop).iPtDollarRec).lBudget = tgDollarRec(ilWkIndex, tgImpactRec(ilLoop).iPtDollarRec).lBudget + tmDPBudgetInfo(ilDP).lBudget 'Get Pennies
                                    tgDollarRec(ilWkIndex, tgImpactRec(ilLoop).iPtDollarRec).l30Inv = tgDollarRec(ilWkIndex, tgImpactRec(ilLoop).iPtDollarRec).l30Inv + tmDPBudgetInfo(ilDP).iInv
                                    If (tmDPBudgetInfo(ilDP).iAvailDefined = 1) And (ilYear = tlRif(tmDPBudgetInfo(ilDP).lRifIndex).iYear) Then
                                        tgDollarRec(ilWkIndex, tgImpactRec(ilLoop).iPtDollarRec).iAvailDefined = 1
                                    End If
                                    'Exit For
                                End If
                            Next ilDP
                        Next ilLoop
                        For ilDP = LBound(tmDPBudgetInfo) To UBound(tmDPBudgetInfo) - 1 Step 1
                            tmDPBudgetInfo(ilDP).iInv = 0
                            tmDPBudgetInfo(ilDP).lRCPrice = 0
                            tmDPBudgetInfo(ilDP).lPrice = 0
                            tmDPBudgetInfo(ilDP).lBudget = 0
                            tmDPBudgetInfo(ilDP).iAvailDefined = 0
                        Next ilDP
                    Next llDate
                    Exit For
                End If
            Next ilBvf
        End If
        mGetSpots hlChf, hlClf, hlCff, hlSdf, hlSmf, hlVef, hlVsf, imVefCode(ilVef), llBdStartDate, llBdEndDate, tlRdf()
    Next ilVef
    Erase tmBvfVeh
    Erase tmBvfVeh2
    Erase imVefCode
    Erase tmDPBudgetInfo
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gBuildAvails                    *
'*                                                     *
'*             Created:7/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Build avails from TFN if Ssf   *
'*                      does not exist                 *
'*                                                     *
'*******************************************************
Public Function gBuildAvails(ilRetIn As Integer, ilVefCode As Integer, llDate As Long, llLatestDate As Long, ilType As Integer) As Integer
    Dim llTstDate As Long
    Dim ilWeekDay As Integer
    Dim ilRet As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    'Dim ilType As Integer
    ReDim ilEvtType(0 To 14) As Integer
    gBuildAvails = ilRetIn
    'ilType = tgRCSsf.iType
    gUnpackDateLong tgRCSsf.iDate(0), tgRCSsf.iDate(1), llTstDate
    If (ilRet <> BTRV_ERR_NONE) Or (tgRCSsf.iType <> ilType) Or (tgRCSsf.iVefCode <> ilVefCode) Or (llTstDate <> llDate) Then
        If (llDate > llLatestDate) Then
            For ilLoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
                ilEvtType(ilLoop) = False
            Next ilLoop
            ilEvtType(2) = True
            gPackDateLong llDate, ilDate0, ilDate1
            ReDim tlLLC(0 To 0) As LLC  'Merged library names
            If ilType = 0 Then
                ilWeekDay = gWeekDayLong(llDate) + 1 '1=Monday TFN; 2=Tues,...7=Sunday TFN
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
            tgRCSsf.iType = ilType
            tgRCSsf.iVefCode = ilVefCode
            tgRCSsf.iDate(0) = ilDate0
            tgRCSsf.iDate(1) = ilDate1
            gPackTime tlLLC(0).sStartTime, tgRCSsf.iStartTime(0), tgRCSsf.iStartTime(1)
            tgRCSsf.iCount = 0
            'tgRCSsf.iNextTime(0) = 1  'Time not defined
            'tgRCSsf.iNextTime(1) = 0

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
                tgRCSsf.iCount = tgRCSsf.iCount + 1
                tgRCSsf.tPas(ADJSSFPASBZ + tgRCSsf.iCount) = tmAvail
            Next ilIndex
            gBuildAvails = BTRV_ERR_NONE
            Erase ilEvtType
            Erase tlLLC
        End If
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetCost                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Cost from line             *
'*                      Note: The flight while be      *
'*                            obtain for timing test   *
'*                                                     *
'*******************************************************
Function mGetCost(tlSdf As SDF, hlClf As Integer, hlCff As Integer, hlSmf As Integer, hlVef As Integer, hlVsf As Integer) As Long
    Dim ilRet As Integer
    Dim slPrice As String
    mGetCost = 0
    If tlSdf.sSpotType = "X" Then
        Exit Function
    End If
    imClfRecLen = Len(tmClf)
    tmClfSrchKey.lChfCode = tlSdf.lChfCode
    tmClfSrchKey.iLine = tlSdf.iLineNo
    tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest Revision
    tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
    ilRet = btrGetGreaterOrEqual(hlClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tlSdf.lChfCode) And (tmClf.iLine = tlSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))
        ilRet = btrGetNext(hlClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tlSdf.lChfCode) And (tmClf.iLine = tlSdf.iLineNo) And ((tmClf.sSchStatus = "M") Or (tmClf.sSchStatus = "F")) Then
        ilRet = gGetSpotPrice(tlSdf, tmClf, hlCff, hlSmf, hlVef, hlVsf, slPrice)
        If InStr(slPrice, ".") > 0 Then
            mGetCost = gStrDecToLong(slPrice, 2)
        End If
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetSpots                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Missed Information         *
'*                                                     *
'*******************************************************
Sub mGetSpots(hlChf As Integer, hlClf As Integer, hlCff As Integer, hlSdf As Integer, hlSmf As Integer, hlVef As Integer, hlVsf As Integer, ilVefCode As Integer, llStartDate As Long, llEndDate As Long, tlRdf() As RDF)
    Dim ilRet As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilIndex As Integer
    Dim ilRdf As Integer
    Dim ilAvailOk As Integer
    Dim ilVpfIndex As Integer
    Dim ilWkIndex As Integer
    Dim ilDay As Integer
    Dim llTime As Long
    Dim ilTime As Integer
    Dim llRdfStartTime As Long
    Dim llRdfEndTime As Long
    Dim ilLen As Integer
    Dim ilNo30 As Integer
    Dim ilNo60 As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilLoop As Integer
    Dim llPrice As Long
    Dim ilLnRdf As Integer
    Dim ilRdfAnfCode As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    ilVpfIndex = -1 'gVpfFind(RateCard, ilVefCode)
    'For ilLoop = 0 To UBound(tgVpf) Step 1
    '    If ilVefCode = tgVpf(ilLoop).iVefKCode Then
        ilLoop = gBinarySearchVpf(ilVefCode)
        If ilLoop <> -1 Then
            ilVpfIndex = ilLoop
    '        Exit For
        End If
    'Next ilLoop
    btrExtClear hlSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)  'Extract operation record size
    tmSdfSrchKey1.iVefCode = ilVefCode
    slDate = Format$(llStartDate, "m/d/yy")
    gPackDate slDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
    tmSdfSrchKey1.iTime(0) = 0
    tmSdfSrchKey1.iTime(1) = 0
    tmSdfSrchKey1.sSchStatus = ""   'slType
    imSdfRecLen = Len(tmSdf)
    imCHFRecLen = Len(tmChf)
    ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
    If (tmSdf.iVefCode = ilVefCode) And (ilRet <> BTRV_ERR_END_OF_FILE) Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC", "SDF", "") '"EG") 'Set extract limits (all records)
        tlIntTypeBuff.iType = ilVefCode   'Type field record
        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        slDate = Format$(llEndDate, "m/d/yy")
        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        ilRet = btrExtAddField(hlSdf, 0, ilExtLen) 'Extract the whole record
        'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
        ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            ilExtLen = Len(tmSdf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                'If (tmSdf.sSchStatus = "M") Or (tmSdf.sSchStatus = "U") Or (tmSdf.sSchStatus = "R") Or (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                If (tmSdf.sSchStatus = "M") Or (tmSdf.sSchStatus = "U") Or (tmSdf.sSchStatus = "R") Then
                    If (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                        ilWkIndex = (llDate - llStartDate) \ 7 + 1
                        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                        ilDay = gWeekDayStr(slDate)
                        gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                        'Increment inventory
                        For ilIndex = LBONE To UBound(tgImpactRec) - 1 Step 1
                            If tgImpactRec(ilIndex).iVefCode = ilVefCode Then
                                For ilRdf = LBound(tlRdf) To UBound(tlRdf) - 1 Step 1
                                    If tlRdf(ilRdf).iCode = tgImpactRec(ilIndex).iRdfCode Then
                                        ilAvailOk = False
                                        For ilTime = LBound(tlRdf(ilRdf).iStartTime, 2) To UBound(tlRdf(ilRdf).iStartTime, 2) Step 1
                                            If (tlRdf(ilRdf).iStartTime(0, ilTime) <> 1) Or (tlRdf(ilRdf).iStartTime(1, ilTime) <> 0) Then
                                                'If tlRdf(ilRdf).sWkDays(ilTime, ilDay + 1) = "Y" Then
                                                If tlRdf(ilRdf).sWkDays(ilTime, ilDay) = "Y" Then
                                                    gUnpackTimeLong tlRdf(ilRdf).iStartTime(0, ilTime), tlRdf(ilRdf).iStartTime(1, ilTime), False, llRdfStartTime
                                                    gUnpackTimeLong tlRdf(ilRdf).iEndTime(0, ilTime), tlRdf(ilRdf).iEndTime(1, ilTime), True, llRdfEndTime
                                                    If (llTime >= llRdfStartTime) And (llTime < llRdfEndTime) Then
                                                        ilAvailOk = True
                                                    End If
                                                End If
                                            End If
                                        Next ilTime

                                        If ilAvailOk Then
                                            llPrice = mGetCost(tmSdf, hlClf, hlCff, hlSmf, hlVef, hlVsf)
                                            ilRdfAnfCode = 0
                                            ilLnRdf = gBinarySearchRdf(tmClf.iRdfCode)
                                            If ilLnRdf <> -1 Then
                                                ilRdfAnfCode = tgMRdf(ilLnRdf).ianfCode
                                            End If
                                            If (tlRdf(ilRdf).sInOut = "I") Then
                                                If (tlRdf(ilRdf).ianfCode <> ilRdfAnfCode) Then
                                                    ilAvailOk = False
                                                End If
                                            End If
                                            If (tlRdf(ilRdf).sInOut = "O") Then
                                                If ilRdfAnfCode = tlRdf(ilRdf).ianfCode Then
                                                    ilAvailOk = False
                                                End If
                                            End If
                                        End If

                                        If ilAvailOk Then
                                            ilNo30 = 0
                                            ilNo60 = 0
                                            'ilLen = tmSpot.iPosLen And &HFFF
                                            ilLen = tmSdf.iLen
                                            If ilLen >= 30 Then
                                                If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                                    If (ilLen Mod 30) = 0 Then
                                                        Do While ilLen >= 30
                                                            ilNo30 = ilNo30 + 1
                                                            ilLen = ilLen - 30
                                                        Loop
                                                    End If
                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                                    If (ilLen Mod 30) = 0 Then
                                                        Do While ilLen >= 30
                                                            ilNo30 = ilNo30 + 1
                                                            ilLen = ilLen - 30
                                                        Loop
                                                    End If
                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                                    If (ilLen Mod 30) = 0 Then
                                                        Do While ilLen >= 30
                                                            ilNo30 = ilNo30 + 1
                                                            ilLen = ilLen - 30
                                                        Loop
                                                    End If
                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                                End If
                                            Else
                                                ilNo30 = 1
                                            End If
                                            tmChfSrchKey.lCode = tmSdf.lChfCode
                                            ilRet = btrGetEqual(hlChf, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If (ilRet = BTRV_ERR_NONE) And (tmChf.sType <> "M") And (tmChf.sType <> "S") Then
                                                tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold + ilNo30
                                                'tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold + mGetCost(tmSdf, hlClf, hlCff, hlSmf, hlVef, hlVsf) / 100
                                                tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold + llPrice / 100
                                            End If
                                            'If (tmSpot.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE Then
                                            'Else
                                            'End If
                                            'If tmSpot.iRank <= 1000 Then
                                            '    tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).l30Sold + ilNo30
                                            '    tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold = tgDollarRec(ilWkIndex, tgImpactRec(ilIndex).iPtDollarRec).lDollarSold + mGetCost()
                                            'ElseIf (tmSpot.iRank = 1020) Then   'Remnant
                                            'ElseIf (tmSpot.iRank = 1010) Or (tmSpot.iRank = 1030) Then   'Direct Response or per Inquiry
                                            'ElseIf (tmSpot.iRank = 1040) Then   'Trade
                                            'ElseIf (tmSpot.iRank = 1050) Then 'Promo
                                            'ElseIf (tmSpot.iRank = 1060) Then    'PSA
                                            'End If
                                        End If
                                        Exit For
                                    End If
                                Next ilRdf
                            End If
                        Next ilIndex
                    End If
                End If
                ilExtLen = Len(tmSdf)
                ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If

End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadBvfRec                     *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Function mReadBvfRec(hlBvf As Integer, ilMnfCode As Integer, ilYear As Integer, tlBvfVeh() As BVF) As Integer
'
'   iRet = mReadBvfRec (hlBvf As Integer, iMnfCode, ilYears)
'   Where:
'       ilMnfCode(I)-Budget Name Code
'       ilYears(I)-Year to retrieve
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilRecOK As Integer
    Dim ilFound As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record

    ReDim tlBvfVeh(0 To 0) As BVF

    ilUpper = UBound(tlBvfVeh)
    btrExtClear hlBvf   'Clear any previous extend operation
    ilExtLen = Len(tlBvfVeh(0))  'Extract operation record size
    imBvfRecLen = Len(tmBvf)
    tmBvfSrchKey.iYear = ilYear
    tmBvfSrchKey.iSeqNo = 1
    tmBvfSrchKey.iMnfBudget = ilMnfCode
    ilRet = btrGetGreaterOrEqual(hlBvf, tmBvf, imBvfRecLen, tmBvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    'ilRet = btrGetFirst(hlBvf, tgBvfRec(1).tBvf, imBvfRecLen, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hlBvf, llNoRec, -1, "UC", "BVF", "") '"EG") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("Bvf", "BvfMnfBudget")
        tlIntTypeBuff.iType = ilMnfCode
        ilRet = btrExtAddLogicConst(hlBvf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        'On Error GoTo mReadBvfRecErr
        'gBtrvErrorMsg ilRet, "mReadBvfRec (btrExtAddLogicConst):" & "Bvf.Btr", RCImpact
        'On Error GoTo 0
        ilOffSet = gFieldOffset("Bvf", "BvfYear")
        tlIntTypeBuff.iType = ilYear
        ilRet = btrExtAddLogicConst(hlBvf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        'On Error GoTo mReadBvfRecErr
        'gBtrvErrorMsg ilRet, "mReadBvfRec (btrExtAddLogicConst):" & "Bvf.Btr", RCImpact
        'On Error GoTo 0
        ilRet = btrExtAddField(hlBvf, 0, ilExtLen) 'Extract the whole record
        'On Error GoTo mReadBvfRecErr
        'gBtrvErrorMsg ilRet, "mReadBvfRec (btrExtAddField):" & "Bvf.Btr", RCImpact
        'On Error GoTo 0
        ilRet = btrExtGetNext(hlBvf, tmBvf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            'On Error GoTo mReadBvfRecErr
            'gBtrvErrorMsg ilRet, "mReadBvfRec (btrExtGetNextExt):" & "Bvf.Btr", RCImpact
            'On Error GoTo 0
            ilExtLen = Len(tmBvf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlBvf, tmBvf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                'User allow to see vehicle
                ilRecOK = True  'False
                'For ilVeh = LBound(tmUserVeh) To UBound(tmUserVeh) - 1 Step 1
                '    If tmUserVeh(ilVeh).iCode = tgBvfRec(ilUpper).tBvf.iVefCode Then
                '        ilRecOk = True
                '        If tmVef.iCode <> tmUserVeh(ilVeh).iCode Then
                '            tmVefSrchKey.iCode = tmUserVeh(ilVeh).iCode
                '            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                '            If ilRet <> BTRV_ERR_NONE Then
                '                ilRecOk = False
                '            End If
                '        End If
                '        Exit For
                '    End If
                'Next ilVeh
                'If ilRecOk Then
                '    slStr = ""
                '    slStr = Trim$(Str$(tmVef.iSort))
                '    Do While Len(slStr) < 5
                '        slStr = "0" & slStr
                '    Loop
                '    tgBvfRec(ilUpper).sVehSort = slStr
                '    tgBvfRec(ilUpper).sVehicle = tmVef.sName
                '    ilRecOk = False
                '    For ilSaleOffice = LBound(tmSaleOffice) To UBound(tmSaleOffice) - 1 Step 1
                '        If tmSaleOffice(ilSaleOffice).iCode = tgBvfRec(ilUpper).tBvf.iSofCode Then
                '            ilRecOk = True
                '            tgBvfRec(ilUpper).sOffice = tmSaleOffice(ilSaleOffice).sName
                '            Exit For
                '        End If
                '    Next ilSaleOffice
                'End If
                If ilRecOK Then
                    ilFound = False
                    For ilLoop = 0 To ilUpper - 1 Step 1
                        If tlBvfVeh(ilLoop).iVefCode = tmBvf.iVefCode Then
                            For ilIndex = LBound(tmBvf.lGross) To UBound(tmBvf.lGross) Step 1
                                tlBvfVeh(ilLoop).lGross(ilIndex) = tlBvfVeh(ilLoop).lGross(ilIndex) + tmBvf.lGross(ilIndex)
                            Next ilIndex
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        tlBvfVeh(ilUpper) = tmBvf
                        ilUpper = ilUpper + 1
                        ReDim Preserve tlBvfVeh(0 To ilUpper) As BVF
                    End If
                End If
                ilRet = btrExtGetNext(hlBvf, tmBvf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlBvf, tmBvf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    mReadBvfRec = True
    Exit Function

    On Error GoTo 0
    mReadBvfRec = False
    Exit Function
End Function
