Attribute VB_Name = "RPTCRIV"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcriv.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptCRGet.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report Get Data for Crystal screen code
Option Explicit
Option Compare Text
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal x%, ByVal y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
'Public Const TA_LEFT = 0
'Public Const TA_RIGHT = 2
'Public Const TA_CENTER = 6
'Public Const TA_TOP = 0
'Public Const TA_BOTTOM = 8
'Public Const TA_BASELINE = 24
'Inventory Valuation
Dim imQNo As Integer
'Public igYear As Integer
'Public igPdStartDate(0 To 1) As Integer
'Public sgPdType As String * 1
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
'Public igMonthOrQtr As Integer          'entered month or qtr
Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey2 As SDFKEY2            'SDF record image (key 2)
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSdf As SDF
Dim tmSdfSrchKey3 As LONGKEY0     'SDF record image (SDF code as keyfield)
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmRdf As Integer            'Dayparts file handle
Dim tmAvRdf() As RDF            'array of dayparts
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmSsf As Integer            'Spot Summary file handle
Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS

'7-19-04 new field in ANF indicating if avail allows only Nework, local or both types of spots
Dim hmAnf As Integer            'Avail name file handle
Dim tmAnf As ANF                'ANF record image
Dim tmAnfSrchKey As INTKEY0            'ANF record image
Dim imAnfRecLen As Integer        'ANF record length
Dim tmRcf As RCF
Dim hmRcf As Integer            'Rate Card file handle
Dim tmRif As RIF
Dim hmRif As Integer            'Rate Card items file handle
Dim imRifRecLen As Integer      'RIF record length
Dim tmRifRate() As RIF
'Same as in Quarterly Avails
Dim hmAvr As Integer            'Quarterly Avails file handle
Dim tmAvr() As AVR                'AVR record image
Dim imAvrRecLen As Integer        'AVR record length
'Dim lmSAvailsDates(1 To 14) As Long   '7-5-00 Start Dates of avail week
Dim lmSAvailsDates(0 To 14) As Long   '7-5-00 Start Dates of avail week. Index zero ignored
'Dim lmEAvailsDates(1 To 14) As Long   '7-5-00 End dates of avail week
Dim lmEAvailsDates(0 To 14) As Long   '7-5-00 End dates of avail week. Index zero ignored
Dim smBucketType As String 'I=Inventory; A=Avail; S=Sold
Dim imMissed As Integer 'True = Include Missed
Dim imStandard As Integer
Dim imRemnant As Integer    'True=Include Remnant
Dim imReserv As Integer  'true = include reservations
Dim imDR As Integer     'True =Include Direct Response
Dim imPI As Integer     'True=Include per Inquiry
Dim imPSA As Integer    'True=Include PSA
Dim imPromo As Integer  'True=Include Promo
Dim imXtra As Integer   'true = include xtra bonus spots
Dim imNC As Integer     'true = include NC spots
Dim imHold As Integer   'true = include hold contracts
Dim imTrade As Integer  'true = include trade contracts
Dim imOrder As Integer  'true = include Complete order contracts
'Log Calendar
Dim hmLcf As Integer            'Log Calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim imLcfRecLen As Integer        'LCF record length
'*******************************************************
'*                                                     *
'*      Procedure Name:gCRQAvailsClearIV                 *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear Quarterly avails Data     *
'*                     for Crystal report              *
'*                                                     *
'*******************************************************
Sub gCRQAvailsClearIV()
    Dim ilRet As Integer
    hmAvr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAvr, "", sgDBPath & "Avr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAvr)
        btrDestroy hmAvr
        Exit Sub
    End If
    ReDim tmAvr(0 To 0) As AVR
    imAvrRecLen = Len(tmAvr(0))
    'tmAvrSrchKey.iGenDate(0) = igNowDate(0)
    'tmAvrSrchKey.iGenDate(1) = igNowDate(1)
    'tmAvrSrchKey.iGenTime(0) = igNowTime(0)
    'tmAvrSrchKey.iGenTime(1) = igNowTime(1)
    'ilRet = btrGetGreaterOrEqual(hmAvr, tmAvr(0), imAvrRecLen, tmAvrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)

    'The preceeding code was bypassed to allow two reports to be called from
    'the same prepass file and still be able to erase the data from Avr.btr when
    'the second report was finished
    ilRet = btrGetLast(hmAvr, tmAvr(0), imAvrRecLen, 0, 0, BTRV_OP_GET_KEY)
    ilRet = btrGetGreaterOrEqual(hmAvr, tmAvr(0), imAvrRecLen, tmAvr(0), INDEXKEY0, BTRV_LOCK_NONE)
    'end of change to procedure
    Do While (ilRet = BTRV_ERR_NONE) And (tmAvr(0).iGenDate(0) = igNowDate(0)) And (tmAvr(0).iGenDate(1) = igNowDate(1)) And (tmAvr(0).lGenTime = lgNowTime)
        ilRet = btrDelete(hmAvr)
        ilRet = btrGetNext(hmAvr, tmAvr(0), imAvrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    Erase tmAvr
    ilRet = btrClose(hmAvr)
    btrDestroy hmAvr
End Sub
'********************************************************************
'*                                                                  *
'*      Procedure Name:gCRQAvailsGenIV - Inventory Valuation        *
'*                                                                  *
'*             Created:10/09/93      By:D. LeVine                   *
'*            Modified:              By:                            *
'*                                                                  *
'*            Comments:Generate Quarterly avails Data               *
'*                     for Crystal report                           *
'*            3/11/98 - change to look at BAse DP only              *
'*                     (screen out stuff from seeing                *
'*                       twice)                                     *
'*                                                                  *
'*          7-5-00 Expand the quarterly 13 buckets to               *
'                   14 weeks                                        *
'           7-2-00 fix calculation of std qtrs                      *
'           3-6-01 Vehicle Groups codes were not setup properly     *
'           7-19-04 Exclude Network spots
'********************************************************************

Sub gCRQAvailsGenIV()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilVehicle As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilNoQuarters As Integer
    Dim slQSDate As String
    Dim sl1WkDate As String
    Dim ll1WkDate As Long
    Dim ilFirstQ As Integer
    Dim ilMonth As Integer
    Dim ilRec As Integer
    Dim ilVpfIndex As Integer
    Dim ilUpper As Integer
    Dim ilDateOk As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilRcf As Integer
    Dim ilFound As Integer
    Dim llEndStd As Long
    Dim slSaveBase As String
    Dim ilSaveSort As Integer
    Dim ilMinorSet As Integer
    Dim ilMajorSet As Integer
    Dim ilmnfMinorCode As Integer           'field used to sort the minor sort with
    Dim ilMnfMajorCode As Integer           'field used to sort the major sort with
    Dim ilNoWks As Integer                  '7-5-00
    ilLoop = RptSelIv!cbcSet1.ListIndex
    ilMajorSet = gFindVehGroupInx(ilLoop, tgVehicleSets1())
    ilLoop = RptSelIv!cbcSet2.ListIndex
    ilMinorSet = gFindVehGroupInx(ilLoop, tgVehicleSets2())
    ilNoQuarters = Val(RptSelIv!edcSelCFrom1.Text)
    If RptSelIv!rbcSelC7(0).Value Then           'avails
        smBucketType = "A"
    'ElseIf (rptseliv!rbcSelCSelect(1).Value) Then       'sold in minutes
    '    smBucketType = "S"
    'ElseIf (rptseliv!rbcSelCSelect(3).Value) Then         'sold in percent
    '    smBucketType = "P"
    Else
        smBucketType = "I"
    End If
    If smBucketType = "I" Then
        imMissed = False
        imRemnant = False
        imDR = False
        imPI = False
        imPSA = False
        imPromo = False
        imXtra = False
        imReserv = False
        imStandard = False
        imTrade = False
        imNC = False
        imHold = False
        imOrder = False

    Else
        imHold = gSetCheck(RptSelIv!ckcSelC3(0).Value)
        imOrder = gSetCheck(RptSelIv!ckcSelC3(1).Value)
        imStandard = gSetCheck(RptSelIv!ckcSelC5(0).Value)
        imReserv = gSetCheck(RptSelIv!ckcSelC5(1).Value)
        imRemnant = gSetCheck(RptSelIv!ckcSelC5(2).Value)
        imDR = gSetCheck(RptSelIv!ckcSelC6(0).Value)
        imPI = gSetCheck(RptSelIv!ckcSelC6(1).Value)
        imPSA = gSetCheck(RptSelIv!ckcSelC6(2).Value)
        imPromo = gSetCheck(RptSelIv!ckcSelC6(3).Value)
        imTrade = gSetCheck(RptSelIv!ckcSelC8(0).Value)
        imMissed = gSetCheck(RptSelIv!ckcSelC8(1).Value)
        'imNC = rptseliv!ckcSelC6(2).Value
        imXtra = gSetCheck(RptSelIv!ckcSelC8(2).Value)
    End If
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)
    hmAvr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAvr, "", sgDBPath & "Avr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAvr)
        btrDestroy hmAvr
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    ReDim tmAvr(0 To 0) As AVR
    imAvrRecLen = Len(tmAvr(0))
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAvr)
        btrDestroy hmSsf
        btrDestroy hmAvr
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAvr)
        btrDestroy hmLcf
        btrDestroy hmSsf
        btrDestroy hmAvr
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    imLcfRecLen = Len(tmLcf)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAvr)
        btrDestroy hmCHF
        btrDestroy hmLcf
        btrDestroy hmSsf
        btrDestroy hmAvr
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)
    hmRif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRif, "", sgDBPath & "Rif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRif)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAvr)
        btrDestroy hmRif
        btrDestroy hmCHF
        btrDestroy hmLcf
        btrDestroy hmSsf
        btrDestroy hmAvr
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
   End If
    imRifRecLen = Len(tmRif)

    '7-19-04
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmRif)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAvr)
        btrDestroy hmRif
        btrDestroy hmCHF
        btrDestroy hmLcf
        btrDestroy hmSsf
        btrDestroy hmAvr
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmAnf
        Exit Sub
    End If
    imAnfRecLen = Len(tmAnf)

'    slDate = RptSelIv!edcSelCFrom.Text
    slDate = RptSelIv!CSI_CalFrom.Text      '9-6-19 use csi calendar control vs edit box
    ilRet = gObtainRcfRifRdf()          'get the rate cards and assoc dayparts
    'slDate = gObtainStartStd(slDate)    'get the standard start date from eff. date
    'ilMonth = Month(Format$(gDateValue(slDate), "m/d/yy"))
    'ilDay = Day(Format$(gDateValue(slDate), "m/d/yy"))
    'Do While (ilMonth <> 12) And ((ilMonth <> 1) Or (ilDay <> 1))
    '    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
    '    ilMonth = Month(Format$(gDateValue(slDate), "m/d/yy"))
    '    ilDay = Day(Format$(gDateValue(slDate), "m/d/yy"))
    'Loop

    slDate = gObtainYearStartDate(0, slDate)       'find the start of the std bdcst year
'    sl1WkDate = RptSelIv!edcSelCFrom.Text          'start date requested from user
    sl1WkDate = RptSelIv!CSI_CalFrom.Text          'start date requested from user
    ll1WkDate = gDateValue(sl1WkDate)
    llDate = gDateValue(slDate)                     'begin with start of std year

    'lldate needs each std bdcst start date
    llEndStd = gDateValue(gObtainEndStd(slDate))    '7-5-00 get the current std months end date to calculate the next one

    For ilLoop = 1 To 4         '8-2-00
        'llendstd = end of the std, retrieve std month # from that date
        ilMonth = Month(Format$(llEndStd, "m/d/yy"))
        ilNoWks = Year(Format$(llEndStd, "mm/dd/yyyy"))
        slName = str$(ilMonth + 2) + "/15/" + str$(ilNoWks)
        llEndStd = gDateValue(gObtainEndStd(slName))
        If ll1WkDate >= llDate And ll1WkDate <= llEndStd Then
            'found the quarter
            ilLoop = 4
            Exit For
        Else
            'get next start quarter
            llDate = llEndStd + 1
            slDate = Format$(llDate, "m/d/yy")
            llEndStd = gDateValue(gObtainEndStd(slDate))    '7-5-00 get the current std months end date to calculate the next one
        End If
    Next ilLoop
    'Do While (ll1WkDate < llDate) Or (ll1WkDate > llEndStd)   '7-5-00 increase from start of std year until its equal or greater to the requested user start date
    '    llDate = llEndStd + 1                   '7-5-00 bump up to next std month
    '    slDate = Format$(llDate, "m/d/yy")       '7-5-00 get end date of next std month
    '    llEndStd = gDateValue(gObtainEndStd(slDate))    '7-5-00 convert the next months end date to a long for comparison
    'Loop
    slQSDate = Format$(llDate, "m/d/yy")     'true start date to begin showing data, anyweek from the start of the quarter to the entered date should be shown blank
    tmVef.iCode = 0
    'determine # of weeks in this quarter
    llEndStd = llDate + 83         '8-2-00 add minimum # of weeks in a quarter (12 weeks)
    llEndStd = gDateValue(gObtainEndStd(Format$(llEndStd, "m/d/yy")))
    ilNoWks = (llEndStd - llDate) / 7
    slQSDate = Format$(llDate, "m/d/yy")
    For imQNo = 1 To ilNoQuarters Step 1
        llDate = gDateValue(slQSDate)
        For ilLoop = 1 To ilNoWks Step 1    '7-5-00
            If (ll1WkDate >= llDate) And (ll1WkDate <= llDate + 6) Then
                ilFirstQ = ilLoop
            End If
            lmSAvailsDates(ilLoop) = llDate
            lmEAvailsDates(ilLoop) = llDate + 6
            llDate = llDate + 7
        Next ilLoop
        For ilVehicle = 0 To RptSelIv!lbcSelection(0).ListCount - 1 Step 1
            If (RptSelIv!lbcSelection(0).Selected(ilVehicle)) Then
                slNameCode = tgCSVNameCode(ilVehicle).sKey 'rptseliv!lbcCSVNameCode.List(ilVehicle)
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slName, 3, "|", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                ilVpfIndex = -1
                'For ilLoop = 0 To UBound(tgVpf) Step 1
                '    If ilVefCode = tgVpf(ilLoop).iVefKCode Then
                    ilLoop = gBinarySearchVpf(ilVefCode)
                    If ilLoop <> -1 Then
                        ilVpfIndex = ilLoop
                '        Exit For
                    End If
                'Next ilLoop
                If ilVpfIndex >= 0 Then
                    For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
                        tmRcf = tgMRcf(ilRcf)
                        ilDateOk = False
'                        For ilLoop = 0 To RptSelIv!lbcSelection(12).ListCount - 1 Step 1
                        For ilLoop = 0 To RptSelIv!lbcSelection(1).ListCount - 1 Step 1         '9-6-19 extra list boxes removed, chg from index 12 to 1
                            slNameCode = tgRateCardCode(ilLoop).sKey
                            ilRet = gParseItem(slNameCode, 3, "\", slCode)
                            If Val(slCode) = tgMRcf(ilRcf).iCode Then
'                                If (RptSelIv!lbcSelection(12).Selected(ilLoop)) Then
                                If (RptSelIv!lbcSelection(1).Selected(ilLoop)) Then
                                    ilDateOk = True
                                End If
                                Exit For
                            End If
                        Next ilLoop

                        If ilDateOk Then
                            ReDim tmAvr(0 To 0) As AVR
                            ReDim tmAvRdf(0 To 0) As RDF
                            ReDim tmRifRate(0 To 0) As RIF
                            ilUpper = 0
                            For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                                If tgMRif(llRif).iRcfCode = tgMRcf(ilRcf).iCode And tgMRif(llRif).iVefCode = ilVefCode Then 'is this Rif record belong to the rate card
                                    'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1 'is this a daypart record
                                    ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
                                    If ilRdf <> -1 Then
                                        'If tgMRdf(ilRdf).iCode = tgMRif(ilRif).iRdfCode And tgMRdf(ilRdf).sReport <> "N" And tgMRdf(ilRdf).sBase = "Y" And tgMRdf(ilRdf).sState = "A" And tgMRif(ilRif).iVefCode = ilVefCode Then 'is this a report daypart and for the selected vehicle
                                        'If tgMRdf(ilRdf).iCode = tgMRif(ilRif).iRdfCode And tgMRdf(ilRdf).sReport <> "N" And tgMRdf(ilRdf).sBase = "Y" And tgMRif(ilRif).iVefCode = ilVefCode Then  'is this a report daypart and for the selected vehicle
                                        'Determine if Rate Card items have been entered, if not, use Daypart (RDF) fields
                                        If tgMRif(llRif).sRpt <> "Y" And tgMRif(llRif).sRpt <> "N" Then
                                            slSaveBase = tgMRdf(ilRdf).sBase
                                        Else
                                            slSaveBase = tgMRif(llRif).sBase
                                        End If
                                        If tgMRif(llRif).iSort = 0 Then
                                            ilSaveSort = tgMRdf(ilRdf).iSortCode
                                        Else
                                            ilSaveSort = tgMRif(llRif).iSort
                                        End If

                                        If tgMRdf(ilRdf).iCode = tgMRif(llRif).iRdfCode And Trim$(slSaveBase) <> "N" And tgMRdf(ilRdf).sState <> "D" And tgMRif(llRif).iVefCode = ilVefCode Then   'is this a base daypart  for the selected vehicle
                                        'If tgMRdf(ilRdf).iCode = tgMRif(ilRif).iRdfCode And tgMRdf(ilRdf).sBase = "Y" And tgMRdf(ilRdf).sState <> "D" And tgMRif(ilRif).iVefCode = ilvefCode Then   'is this a base daypart  for the selected vehicle
                                        ilFound = False
                                            For ilLoop = LBound(tmAvRdf) To ilUpper - 1 Step 1  'has record already been record
                                                If tmAvRdf(ilLoop).iCode = tgMRdf(ilRdf).iCode Then  'add record
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilLoop
                                            If Not ilFound Then
                                                tmAvRdf(ilUpper) = tgMRdf(ilRdf)
                                                tmRifRate(ilUpper) = tgMRif(llRif)
                                                tmRifRate(ilUpper).iSort = ilSaveSort
                                                ilUpper = ilUpper + 1
                                                ReDim Preserve tmAvRdf(0 To ilUpper) As RDF
                                                ReDim Preserve tmRifRate(0 To ilUpper) As RIF
                                            End If
                                    '        Exit For
                                        End If
                                    End If
                                    'Next ilRdf
                                End If
                            Next llRif
                            mGetAvailCounts ilVefCode, ilVpfIndex, ilFirstQ, lmSAvailsDates(1), lmEAvailsDates(ilNoWks)      '7-5-00
                            'Output records
                            For ilRec = 0 To UBound(tmAvr) - 1 Step 1

                                '******  the Vehicle Group Sets 1 and 2 is placed in avr.Day & avr.ianfCode. The values
                                '        were gathered for these 2 fields in mGetAvail Count but won't be used in
                                '        the output of the Inventory Valuation report
                                gGetVehGrpSets ilVefCode, ilMinorSet, ilMajorSet, ilmnfMinorCode, ilMnfMajorCode
                                tmAvr(ilRec).iDay = ilmnfMinorCode
                                tmAvr(ilRec).ianfCode = ilMnfMajorCode
                                tmAvr(ilRec).iWksInQtr = ilNoWks        '7-5-00
                                ilRet = btrInsert(hmAvr, tmAvr(ilRec), imAvrRecLen, INDEXKEY0)
                            Next ilRec
                        End If
                    Next ilRcf
                End If
            End If
        Next ilVehicle
        '7-5-00 llDate = gDateValue(slQSDate) + 13 * 7
        '7-5-00 slQSDate = Format$(llDate, "m/d/yy")
        '7-5-00 ll1WkDate = llDate
        llDate = llEndStd + 1                   '7-5-00 bump up to next std month
        slDate = Format$(llDate, "m/d/yy")       '7-5-00 get end date of next std month
        llEndStd = gDateValue(gObtainEndStd(slDate))    '7-5-00 convert the next months end date to a long for comparison
        llEndStd = llDate + 83         '7-5-00 add minimum # of weeks in a quarter (12 weeks)
        llEndStd = gDateValue(gObtainEndStd(Format$(llEndStd, "m/d/yy")))  '7-5-00
        ilNoWks = (llEndStd - llDate) / 7                                 '7-5-00
        slQSDate = Format$(llDate, "m/d/yy")
        ilFirstQ = 1                'all weeks need to be gathered for next quarters
    Next imQNo
    Erase tmAvr
    Erase tmAvRdf
    Erase tmRifRate
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmRcf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAvr)
    btrDestroy hmRdf
    btrDestroy hmRcf
    btrDestroy hmAvr
    btrDestroy hmSdf
    btrDestroy hmVef
    btrDestroy hmSsf
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetAvailCounts                 *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the Avail counts         *
'*                                                     *
'*          7-19-04 Always exclude network spots       *
'                   (scheduled & missed spots)         *
'*******************************************************
Sub mGetAvailCounts(ilVefCode As Integer, ilVpfIndex As Integer, ilFirstQ As Integer, llStartDate As Long, llEndDate As Long)
'
'   Where:
'
'   smBucketType(I): A=Avail; S=Sold; I=Inventory  , P = Percent sellout
'   lmSAvailsDates(I)- Array of bucket start dates
'   lmEAvailsDates(I)- Array of bucket end dates
'   imHold(I)- True = include hold contracts
'   imOrder(I)- True= include complete order contracts
'   imMissed(I)- True=Include missed
'   imXtra(I)- True=Include Xtra bonus spots
'   imTrade(I)- True = include trade contracts
'   imNC(I)- True = include NC spots
'   imReserv(I) - True = include Reservations spots
'   imRemnant(I)- True=Include Remnant
'   imStandard(I)- true = include std contracts
'   imDR(I)- True=Include Direct Response
'   imPI(I)- True=Include per Inquiry
'   imPSA(I)- True=Include PSA
'   imPromo(I)- True=Include Promo
'
'   Note: Remnants; Direct Response; per Inquiry; PSA and Promos are not
'         saved with a miss status
'         For scheduled spots the rank is used to determine if it is one
'         of the above (Direct reponse=1010; Remnant=1020; per Inquiry= 1030;
'         PSA=1060; Promo=1050.
'
'   3-24-03 change way to test for exclusion/inclusion of fill spots

    Dim slType As String
    Dim ilType As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilEvt As Integer
    Dim ilRet As Integer
    Dim ilSpot As Integer
    Dim llTime As Long
    Dim ilRdf As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilRec As Integer
    Dim ilRecIndex As Integer
    Dim ilLen As Integer
    Dim ilUnits As Integer
    Dim ilNo30 As Integer
    Dim ilNo60 As Integer
    Dim ilDay As Integer
    Dim ilSaveDay As Integer
    Dim slDays As String
    Dim ilLtfCode As Integer
    Dim ilAvailOk As Integer
    Dim ilPass As Integer
    Dim ilDayIndex As Integer
    Dim ilLoopIndex As Integer
    Dim ilBucketIndex As Integer
    Dim ilBucketIndexMinusOne As Integer
    Dim ilSpotOK As Integer
    Dim llLoopDate As Long
    Dim ilWeekDay As Integer
    Dim llLatestDate As Long
    Dim ilIndex As Integer
    Dim slStr As String
    Dim ilAdjAdd As Integer
    Dim ilAdjSub As Integer
    Dim ilWkNo As Integer
    Dim ilFirstWk As Integer
    Dim llMonthStart As Long
    Dim llMonthEnd As Long
    Dim ilWkInx As Integer
    Dim ilLo As Integer
    Dim ilHi As Integer
    Dim ilMonth As Integer
    Dim ilNoWks As Integer      '# weeks in quarter could vary from 12 - 14 weeks
    Dim ilVefIndex As Integer
    Dim slChfType As String * 1     '4-12-18
    Dim ilPctTrade As Integer       '4-12-18

    ReDim ilSAvailsDates(0 To 1) As Integer
    ReDim ilEvtType(0 To 14) As Integer
    ReDim ilRdfCodes(0 To 1) As Integer
    slDate = Format$(lmSAvailsDates(1), "m/d/yy")
    gPackDate slDate, ilSAvailsDates(0), ilSAvailsDates(1)
    slType = "O"
    ilType = 0

    'Obtain end of the first month
'    slStr = RptSelIv!edcSelCFrom.Text
    slStr = RptSelIv!CSI_CalFrom.Text           '9-6-19 use csi calendar vs edit box
    slStr = gObtainYearStartDate(0, slStr)
    slDate = gObtainEndStd(slStr)
    'If gDateValue(slMonth) > gDateValue(slStr) Then
    If Mid$(slDate, 3, 1) = "/" Then
        ilMonth = Val(Left$(slDate, 2))
    Else
       ilMonth = Val(Left$(slDate, 1))
    End If
    igYear = Val(right$(slDate, 2))

    'ReDim ilWksInMonth(1 To 3) As Integer
    ReDim ilWksInMonth(0 To 3) As Integer       'Index zero ignored
    slStr = slDate
    llMonthEnd = gDateValue(gObtainEndStd(Format$(llStartDate, "m/d/yy")))
    llMonthStart = llStartDate
    For ilLoop = 1 To 3 Step 1
        ilWksInMonth(ilLoop) = ((llMonthEnd - llMonthStart) / 7)
        llMonthStart = llMonthEnd + 1
        llMonthEnd = gDateValue(gObtainEndStd(Format$(llMonthStart, "m/d/yy")))
    Next ilLoop

    ilNoWks = (llEndDate - llStartDate) / 7
    llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
    'set the type of events to get fro the day (only Contract avails)
    For ilLoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(ilLoop) = False
    Next ilLoop
    ilEvtType(2) = True
    If tgVpf(ilVpfIndex).sSSellOut = "B" Then           'if units & seconds - add 2 to 30 sec unit and take away 1 fro 60
        ilAdjAdd = 2
        ilAdjSub = 1
    ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then       'if units only - take 1 away from 60 count and add 1 to 30 count
        ilAdjAdd = 1
        ilAdjSub = 1
    End If
    ilVefIndex = gBinarySearchVef(ilVefCode)
    For llLoopDate = llStartDate To llEndDate Step 1
        slDate = Format$(llLoopDate, "m/d/yy")
        gPackDate slDate, ilDate0, ilDate1
        gObtainWkNo 0, slDate, ilWkNo, ilFirstWk        'obtain the week bucket number
        imSsfRecLen = Len(tmSsf)                        'Max size of variable length record
        If tgMVef(ilVefIndex).sType <> "G" Then
        'tmSsfSrchKey.sType = slType
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
        'If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.sType <> slType) Or (tmSsf.iVefcode <> ilVefCode Or (tmSsf.iDate(0) <> ilDate0) And (tmSsf.iDate(1) = ilDate1)) Then
        If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iType <> ilType) Or (tmSsf.iVefCode <> ilVefCode Or (tmSsf.iDate(0) <> ilDate0) And (tmSsf.iDate(1) = ilDate1)) Then
            If (llLoopDate > llLatestDate) Then
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
                ilRet = BTRV_ERR_NONE
            End If
        End If

        'Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slType) And (tmSsf.iVefcode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
        Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
            gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llDate
            ilBucketIndex = -1
            For ilLoop = ilFirstQ To ilNoWks Step 1    '7-5-00
                If (llDate >= lmSAvailsDates(ilLoop)) And (llDate <= lmEAvailsDates(ilLoop)) Then
                    ilBucketIndex = ilLoop
                    Exit For
                End If
            Next ilLoop
            If ilBucketIndex > 0 Then
                ilBucketIndexMinusOne = ilBucketIndex - 1
                ilDay = gWeekDayLong(llDate)
                ilEvt = 1
                Do While ilEvt <= tmSsf.iCount
                   LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    If tmProg.iRecType = 1 Then    'Program (not working for nested prog)
                        ilLtfCode = tmProg.iLtfCode
                    ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 2) Then 'Contract Avails only
                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                        'Determine which rate card program this is associated with
                        For ilRdf = LBound(tmAvRdf) To UBound(tmAvRdf) - 1 Step 1
                            ilAvailOk = False
                            If (tmAvRdf(ilRdf).iLtfCode(0) <> 0) Or (tmAvRdf(ilRdf).iLtfCode(1) <> 0) Or (tmAvRdf(ilRdf).iLtfCode(2) <> 0) Then
                                If (ilLtfCode = tmAvRdf(ilRdf).iLtfCode(0)) Or (ilLtfCode = tmAvRdf(ilRdf).iLtfCode(1)) Or (ilLtfCode = tmAvRdf(ilRdf).iLtfCode(1)) Then
                                    ilAvailOk = False    'True- code later
                                End If
                            Else

                                For ilLoop = LBound(tmAvRdf(ilRdf).iStartTime, 2) To UBound(tmAvRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                    If (tmAvRdf(ilRdf).iStartTime(0, ilLoop) <> 1) Or (tmAvRdf(ilRdf).iStartTime(1, ilLoop) <> 0) Then
                                        gUnpackTimeLong tmAvRdf(ilRdf).iStartTime(0, ilLoop), tmAvRdf(ilRdf).iStartTime(1, ilLoop), False, llStartTime
                                        gUnpackTimeLong tmAvRdf(ilRdf).iEndTime(0, ilLoop), tmAvRdf(ilRdf).iEndTime(1, ilLoop), True, llEndTime
                                        'If (llTime >= llStartTime) And (llTime < llEndTime) And (tmAvRdf(ilRdf).sWkDays(ilLoop, ilDay + 1) = "Y") Then
                                        If (llTime >= llStartTime) And (llTime < llEndTime) And (tmAvRdf(ilRdf).sWkDays(ilLoop, ilDay) = "Y") Then
                                            ilAvailOk = True
                                            ilLoopIndex = ilLoop
                                            slDays = ""
                                            For ilDayIndex = 1 To 7 Step 1
                                                If (tmAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "Y") Or (tmAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "N") Then
                                                    slDays = slDays & tmAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1)
                                                Else
                                                    slDays = slDays & "N"
                                                End If
                                            Next ilDayIndex
                                            Exit For
                                        End If
                                    End If
                                Next ilLoop
                            End If
                            If ilAvailOk Then
                                If tmAvRdf(ilRdf).sInOut = "I" Then   'Book into
                                    If tmAvail.ianfCode <> tmAvRdf(ilRdf).ianfCode Then
                                        ilAvailOk = False
                                    End If
                                ElseIf tmAvRdf(ilRdf).sInOut = "O" Then   'Exclude
                                    If tmAvail.ianfCode = tmAvRdf(ilRdf).ianfCode Then
                                        ilAvailOk = False
                                    End If
                                End If
                            End If

                            '7-19-04 the Named avail property must allow local spots to be included
                            tmAnfSrchKey.iCode = tmAvail.ianfCode
                            ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If (ilRet = BTRV_ERR_NONE) Then
                                If tmAnf.sBookLocalFeed = "F" Then      'Network (feed spots) only avail
                                    ilAvailOk = False
                                End If
                            End If

                            If ilAvailOk Then
                                'Determine if Avr created
                                ilFound = False
                                ilSaveDay = ilDay
                                ilDay = 0                                       'force all data in same day of week
                                For ilRec = 0 To UBound(tmAvr) - 1 Step 1
                                    'If (tmAvr(ilRec).iRdfCode = tmAvRdf(ilRdf).iCode) And (tmAvr(ilRec).iFirstBucket = ilFirstQ) And (tmAvr(ilRec).iDay = ilDay) Then
                                    If (ilRdfCodes(ilRec) = tmAvRdf(ilRdf).iCode) And (tmAvr(ilRec).iFirstBucket = ilFirstQ) And (tmAvr(ilRec).iDay = ilDay) Then
                                        ilFound = True
                                        ilRecIndex = ilRec
                                        Exit For
                                    End If
                                Next ilRec
                                If Not ilFound Then
                                    ilRecIndex = UBound(tmAvr)
                                    tmAvr(ilRecIndex).iGenDate(0) = igNowDate(0)
                                    tmAvr(ilRecIndex).iGenDate(1) = igNowDate(1)
                                    'tmAvr(ilRecIndex).iGenTime(0) = igNowTime(0)
                                    'tmAvr(ilRecIndex).iGenTime(1) = igNowTime(1)
                                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                    tmAvr(ilRecIndex).lGenTime = lgNowTime
                                    tmAvr(ilRecIndex).iVefCode = ilVefCode
                                    tmAvr(ilRecIndex).iDay = ilDay
                                    tmAvr(ilRecIndex).iQStartDate(0) = ilSAvailsDates(0)
                                    tmAvr(ilRecIndex).iQStartDate(1) = ilSAvailsDates(1)
                                    tmAvr(ilRecIndex).iFirstBucket = ilFirstQ
                                    tmAvr(ilRecIndex).sBucketType = smBucketType
                                    'tmAvr(ilRecIndex).iRdfCode = tmAvRdf(ilRdf).iCode
                                    ilRdfCodes(ilRecIndex) = tmAvRdf(ilRdf).iCode
                                    tmAvr(ilRecIndex).iRdfCode = tmAvRdf(ilRdf).iSortCode
                                    tmAvr(ilRecIndex).sInOut = tmAvRdf(ilRdf).sInOut
                                    tmAvr(ilRecIndex).ianfCode = tmAvRdf(ilRdf).ianfCode
                                    tmAvr(ilRecIndex).iDPStartTime(0) = tmAvRdf(ilRdf).iStartTime(0, ilLoopIndex)
                                    tmAvr(ilRecIndex).iDPStartTime(1) = tmAvRdf(ilRdf).iStartTime(1, ilLoopIndex)
                                    tmAvr(ilRecIndex).iDPEndTime(0) = tmAvRdf(ilRdf).iEndTime(0, ilLoopIndex)
                                    tmAvr(ilRecIndex).iDPEndTime(1) = tmAvRdf(ilRdf).iEndTime(1, ilLoopIndex)
                                    tmAvr(ilRecIndex).sDPDays = slDays
                                    tmAvr(ilRecIndex).sNot30Or60 = "N"
                                    tmAvr(ilRecIndex).iRdfSortCode = ilMonth
                                    ReDim Preserve tmAvr(0 To ilRecIndex + 1) As AVR
                                    ReDim Preserve ilRdfCodes(0 To ilRecIndex + 1)
                                End If

                                'Darlene
                                tmAvr(ilRecIndex).lRate(ilBucketIndexMinusOne) = tmRifRate(ilRdf).lRate(ilWkNo)
                                'End Darlene
                                ilDay = ilSaveDay
                                '****
                                If (smBucketType <> "S") Or RptSelIv!rbcSelC7(0).Value Then   'not sold (min ) or it's the qtrly detail which
                                    'needs to gather inventory
                                    ilLen = tmAvail.iLen
                                    ilUnits = tmAvail.iAvInfo And &H1F
                                    ilNo30 = 0
                                    ilNo60 = 0
                                    If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                        'Convert inventory to number of 30's and 60's
                                        Do While ilLen >= 60
                                            ilNo60 = ilNo60 + 1
                                            ilLen = ilLen - 60
                                        Loop
                                        Do While ilLen >= 30
                                            ilNo30 = ilNo30 + 1
                                            ilLen = ilLen - 30
                                        Loop
                                        If ilLen < 30 And ilLen > 0 Then    '7-6-00 assume anything under 30" is 1-30" unit availability
                                            ilNo30 = ilNo30 + 1
                                            ilLen = 0
                                        End If

                                        If (smBucketType <> "S") Or (smBucketType <> "P") Then    'sellout by min or pcts, don't update these yet
                                            'place raw counts into Count fields in Avr
                                            'tmAvr(ilRecIndex).i30Count(ilBucketIndex) = tmAvr(ilRecIndex).i30Count(ilBucketIndex) + ilNo30
                                            'tmAvr(ilRecIndex).i60Count(ilBucketIndex) = tmAvr(ilRecIndex).i60Count(ilBucketIndex) + ilNo60
                                        End If
                                        'always put total inventory into record and avail bucket (avail bucket for qtrly detail)
                                        'tmAvr(ilRecIndex).i30InvCount(ilBucketIndex) = tmAvr(ilRecIndex).i30InvCount(ilBucketIndex) + ilNo30
                                        'tmAvr(ilRecIndex).i60InvCount(ilBucketIndex) = tmAvr(ilRecIndex).i60InvCount(ilBucketIndex) + ilNo60
                                        'tmAvr(ilRecIndex).i30Avail(ilBucketIndex) = tmAvr(ilRecIndex).i30Avail(ilBucketIndex) + ilNo30
                                        'tmAvr(ilRecIndex).i60Avail(ilBucketIndex) = tmAvr(ilRecIndex).i60Avail(ilBucketIndex) + ilNo60
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                        'Count 30 or 60 and set flag if neither
                                        If ilLen = 60 Then
                                            ilNo60 = 1
                                        ElseIf ilLen = 30 Then
                                            ilNo30 = 1
                                        Else
                                            tmAvr(ilRecIndex).sNot30Or60 = "Y"
                                            If ilLen <= 30 Then
                                                ilNo30 = 1
                                            Else
                                                ilNo60 = 1
                                            End If
                                        End If
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                        'Count 30 or 60 and set flag if neither
                                        If ilLen = 60 Then
                                            ilNo60 = 1
                                        ElseIf ilLen = 30 Then
                                            ilNo30 = 1
                                        Else
                                            tmAvr(ilRecIndex).sNot30Or60 = "Y"
                                        End If
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                    End If
                                    If (smBucketType <> "S") Or (smBucketType <> "P") Then    'sellout by min or pcts, don't update these yet
                                        'place raw counts into Count fields in Avr


                                        'remove 3/4/98 - these counts are accumulated twice
                                        'tmAvr(ilRecIndex).i30Count(ilBucketIndex) = tmAvr(ilRecIndex).i30Count(ilBucketIndex) + ilNo30
                                        'tmAvr(ilRecIndex).i60Count(ilBucketIndex) = tmAvr(ilRecIndex).i60Count(ilBucketIndex) + ilNo60



                                    End If
                                    'always put total inventory into record and avail bucket (avail bucket for qtrly detail)
                                            tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                            tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                    tmAvr(ilRecIndex).i30InvCount(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30InvCount(ilBucketIndexMinusOne) + ilNo30
                                    tmAvr(ilRecIndex).i60InvCount(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60InvCount(ilBucketIndexMinusOne) + ilNo60
                                    tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) + ilNo30
                                    tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) + ilNo60
                                End If
                                If smBucketType <> "I" Then                         'avails, sellout, sellout %
                                    For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                                        ilSpotOK = True                             'assume spot is OK to include

                                        If tmSpot.iRecType = 11 Then                'network spot - always exclude
                                            ilSpotOK = False
                                        End If
                                        If ((tmSpot.iRank And RANKMASK) = REMNANTRANK) And (Not imRemnant) Then
                                            ilSpotOK = False
                                        End If
                                        If ((tmSpot.iRank And RANKMASK) = PERINQUIRYRANK) And (Not imPI) Then
                                            ilSpotOK = False
                                        End If
'                                        'Added 4/1/18
'                                        If ((tmSpot.iRank And RANKMASK) = 1010) Then      'DR
'                                            If (Asc(tgSaf(0).sFeatures4) And AVAILINCLDEDIRECTRESPONSES) <> AVAILINCLDEDIRECTRESPONSES Then
'                                                ilSpotOK = False
'                                            End If
'                                        End If
'                                        'end of add

                                        '4-12-18 DR previously not tested
                                        If ((tmSpot.iRank And RANKMASK) = DIRECTRESPONSERANK) And (Not imDR) Then
                                            ilSpotOK = False
                                        End If
                                        If ((tmSpot.iRank And RANKMASK) = TRADERANK) And (Not imTrade) Then
                                            ilSpotOK = False
                                        End If
                                        
                                        '3-24-05 test SDF for fill instead
                                        'If tmSpot.iRank = 1045 And Not imXtra Then
                                        '    ilSpotOK = False
                                        'End If
                                        If ((tmSpot.iRank And RANKMASK) = PROMORANK) And (Not imPromo) Then
                                            ilSpotOK = False
                                        End If
                                        If ((tmSpot.iRank And RANKMASK) = PSARANK) And (Not imPSA) Then
                                            ilSpotOK = False
                                        End If
                                        
'                                        'Added 4/1/18
'                                        If (tmSpot.iRank And RANKMASK) = RESERVATION Then     'Reservation
'                                            If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDERESERVATION) <> AVAILINCLUDERESERVATION Then
'                                                ilSpotOK = False
'                                            End If
'                                        End If
'                                        'end of add

                                        '4-12-18 Reservation previously not tested
                                        If ((tmSpot.iRank And RANKMASK) = RESERVATIONRANK) And (Not imReserv) Then
                                            ilSpotOK = False
                                        End If
                                        If (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                            ilSpotOK = False
                                        End If
                                        ilLen = tmSpot.iPosLen And &HFFF

                                        If ilSpotOK Then                            'continue testing other filters
                                        'If Not ilSpotOK Then
                                            tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                            'ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
                                            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If tmSpot.lSdfCode = tmSdf.lCode And ilRet = BTRV_ERR_NONE Then
                                                If tmSdf.lChfCode <> tmChf.lCode Then               'if already in mem, don't reread
                                                    tmChfSrchKey.lCode = tmSdf.lChfCode
                                                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                                Else
                                                    ilRet = BTRV_ERR_NONE
                                                End If
                                            End If
                                            If ilRet <> BTRV_ERR_NONE Then
                                                ilSpotOK = False
                                            Else
                                                ilLen = tmSdf.iLen

                                                 '3-24-03, check for exclusion of fills & extras
                                                If tmSdf.sSpotType = "X" And Not imXtra Then
                                                    ilSpotOK = False
                                                End If
                                                If tmChf.sStatus = "H" Then
                                                    If Not imHold Then
                                                        ilSpotOK = False
                                                    End If
                                                ElseIf tmChf.sStatus = "O" Then
                                                    If Not imOrder Then
                                                        ilSpotOK = False
                                                    End If
                                                Else
                                                    ilSpotOK = False
                                                End If
                                                '3-16-10 wrong code tested for standard, S---->C
                                                If tmChf.sType = "C" And Not imStandard Then      'include Standard types?
                                                    ilSpotOK = False
                                                End If
                                                If tmChf.sType = "V" And Not imReserv Then      'include reservations ?
                                                    ilSpotOK = False
                                                End If
                                                If tmChf.sType = "R" And Not imDR Then      'include DR?
                                                    ilSpotOK = False
                                                End If
                                                
                                                If (tmChf.sType = "T") And (Not imRemnant) Then
                                                    ilSpotOK = False
                                                End If
                                                If (tmChf.sType = "Q") And (Not imPI) Then
                                                    ilSpotOK = False
                                                End If
                                                                                                
                                                If (tmChf.sType = "M") And (Not imPromo) Then
                                                    ilSpotOK = False
                                                End If
                                                If (tmChf.sType = "S") And (Not imPSA) Then
                                                    ilSpotOK = False
                                                End If
                                                
                                                If (tmChf.iPctTrade = 100) And (Not imTrade) Then       'exclude only if 100% trade
                                                    ilSpotOK = False
                                                End If
                                            End If
                                            If ilSpotOK Then
                                                ilNo30 = 0
                                                ilNo60 = 0
                                                'ilLen = tmSpot.iPosLen And &HFFF
                                                If tgVpf(ilVpfIndex).sSSellOut = "B" Then                   'both units and seconds
                                                    'Convert inventory to number of 30's and 60's
                                                    Do While ilLen >= 60
                                                        ilNo60 = ilNo60 + 1
                                                        ilLen = ilLen - 60
                                                    Loop
                                                    Do While ilLen >= 30
                                                        ilNo30 = ilNo30 + 1
                                                        ilLen = ilLen - 30
                                                    Loop
                                                    If ilLen < 30 And ilLen > 0 Then    '7-6-00 assume anything under 30" is 1-30" unit availability
                                                        ilNo30 = ilNo30 + 1
                                                        ilLen = 0
                                                    End If

                                                    'Count 30 or 60 and set flag if neither
                                                    If (ilNo60 <> 0) Or (ilNo30 <> 0) Then
                                                        If (smBucketType = "S") Or (smBucketType = "P") Then    'sellout or %sellout, accum sold
                                                            If RptSelIv!rbcSelC7(0).Value Then
                                                                If tmChf.sType = "V" Then
                                                                    If Not (RptSelIv!ckcSelC5(1).Value = vbChecked) Then
                                                                        tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                                        tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                                    Else 'avails: accum available
                                                                        tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) + ilNo30
                                                                        tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) + ilNo60
                                                                    End If
                                                                ElseIf tmChf.sStatus = "H" Then
                                                                    tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) - ilNo60
                                                                    tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) - ilNo30
                                                                Else
                                                                    tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                                    tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                                End If
                                                            Else
                                                                tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                                tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                            End If
                                                        Else
                                                            tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) - ilNo30
                                                            tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) - ilNo60
                                                        End If
                                                        'adjust the available buckets (used for qtrly detail report only)
                                                        tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                                        tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                                    End If
                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then               'units sold
                                                    'Count 30 or 60 and set flag if neither
                                                    If ilLen = 60 Then
                                                        ilNo60 = 1
                                                    ElseIf ilLen = 30 Then
                                                        ilNo30 = 1
                                                    Else
                                                        tmAvr(ilRecIndex).sNot30Or60 = "Y"
                                                        If ilLen <= 30 Then
                                                            ilNo30 = 1
                                                        Else
                                                            ilNo60 = 1
                                                        End If
                                                    End If
                                                    If (ilNo60 <> 0) Or (ilNo30 <> 0) Then
                                                        If (smBucketType = "S") Or (smBucketType = "P") Then
                                                            If RptSelIv!rbcSelC7(0).Value Then
                                                                If tmChf.sType = "V" Then
                                                                    If Not RptSelIv!ckcSelC5(1).Value = vbChecked Then
                                                                        tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                                        tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                                    Else 'staus hold or reserve n/a for other qtrly summary options
                                                                        tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) + ilNo30
                                                                        tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) + ilNo60
                                                                    End If
                                                                ElseIf tmChf.sStatus = "H" Then
                                                                    tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) + ilNo30
                                                                    tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) + ilNo60
                                                                Else
                                                                    tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                                    tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                                End If
                                                            Else
                                                                tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                                tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                            End If
                                                        Else
                                                            If ilNo60 > 0 Then                     'spot found a 60?
                                                                tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) - ilNo60
                                                            Else
                                                                If tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) > 0 Then
                                                                    tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) - ilNo30
                                                                Else
                                                                    If tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) > 0 Then
                                                                        tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) - ilNo30
                                                                    Else                        'oversold units
                                                                        tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) - ilNo30
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    'adjust the available buckets (used for qtrly detail report only)
                                                    tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                                    tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then               'matching units
                                                    'Count 30 or 60 and set flag if neither
                                                    If ilLen = 60 Then
                                                        ilNo60 = 1
                                                    ElseIf ilLen = 30 Then
                                                        ilNo30 = 1
                                                    Else
                                                        tmAvr(ilRecIndex).sNot30Or60 = "Y"
                                                    End If

                                                    If (smBucketType = "S") Or (smBucketType = "P") Then  'if Sellout or % sellout, accum the seconds sold
                                                        If RptSelIv!rbcSelC7(0).Value Then
                                                            If tmChf.sType = "V" Then
                                                                If RptSelIv!ckcSelC5(1).Value = vbChecked Then
                                                                    tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                                    tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                                Else 'staus hold or reserve n/a for other qtrly summary options
                                                                    tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) + ilNo30
                                                                    tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) + ilNo60
                                                                End If
                                                            ElseIf tmChf.sStatus = "H" Then
                                                                tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) + ilNo30
                                                                tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) + ilNo60
                                                            Else
                                                                tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                                tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                            End If
                                                        Else
                                                            tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                            tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                        End If

                                                    Else
                                                        tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) - ilNo30
                                                        tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) - ilNo60
                                                    End If
                                                    'adjust the available bucket (used for qrtrly detail report only)
                                                    tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                                    tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                                End If                      'B,M,U,Y
                                            End If                          'ilSpotOK
                                        End If                              'If ilSpotOK
                                    Next ilSpot                             'loop from ssf file for # spots in avail
                                End If                                      'if <> "I"
                            End If                                          'Avail OK
                        Next ilRdf                                          'ilRdf = lBound(tmAvRdf)
                        ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                    End If
                ilEvt = ilEvt + 1   'Increment to next event
            Loop                                                        'do while ilEvt <= tmSsf.iCount
            End If                                                              'ilBucketIndex > 0
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            If tgMVef(ilVefIndex).sType = "G" Then
                ilType = tmSsf.iType
            End If
        Loop
    Next llLoopDate

    'Get missed
    If (imMissed) And (smBucketType <> "I") Then
        'Key 2: VefCode; SchStatus; AdfCode; Date, Time
        For ilPass = 0 To 2 Step 1
            tmSdfSrchKey2.iVefCode = ilVefCode
            If ilPass = 0 Then
                slType = "M"
            ElseIf ilPass = 1 Then
                slType = "R"
            ElseIf ilPass = 2 Then
                slType = "U"
            End If
            tmSdfSrchKey2.sSchStatus = slType
            tmSdfSrchKey2.iAdfCode = 0
            tmSdfSrchKey2.iDate(0) = ilSAvailsDates(0)
            tmSdfSrchKey2.iDate(1) = ilSAvailsDates(1)
            tmSdfSrchKey2.iTime(0) = 0
            tmSdfSrchKey2.iTime(1) = 0
            ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
            'This code added as replacement for Ext operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.sSchStatus = slType)
                '4/12/18: Add spot type test
                ilSpotOK = True                             'assume spot is OK to include

                'How to handle this test? - ignore for now
                'If tmSpot.iRecType = 11 Then                'network spot - always exclude
                '    ilSpotOK = False
                'End If
                
                '4-12-18 implement testing of contract types for missed spots
                'gGetContractParameters tmSdf.lChfCode, slChfType, ilPctTrade
                If tmSdf.lChfCode <> tmChf.lCode Then               'if already in mem, don't reread
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    slChfType = tmChf.sType
                    ilPctTrade = tmChf.iPctTrade
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                If ilRet <> BTRV_ERR_NONE Then
                    ilSpotOK = False
                Else
                
                    If (slChfType = "C") And (Not imStandard) Then
                        ilSpotOK = False
                    End If
                    If (slChfType = "V") And (Not imReserv) Then
                        ilSpotOK = False
                    End If
                    If (slChfType = "R") And (Not imDR) Then
                        ilSpotOK = False
                    End If
                    If (slChfType = "T") And (Not imRemnant) Then
                        ilSpotOK = False
                    End If
                    If (slChfType = "Q") And (Not imPI) Then
                        ilSpotOK = False
                    End If
                    If (slChfType = "M") And (Not imPromo) Then
                        ilSpotOK = False
                    End If
                    If (slChfType = "S") And (Not imPSA) Then
                        ilSpotOK = False
                    End If
                    'ignore trade if partial or full trade
                    If (ilPctTrade = 100) And (Not imTrade) Then
                        ilSpotOK = False
                    End If
                End If
                If ilSpotOK Then
                
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                    If (llDate >= lmSAvailsDates(ilFirstQ)) And (llDate <= lmEAvailsDates(ilNoWks)) And tmSdf.lChfCode > 0 Then    '7-19-04 if chfcode has value, its a normal spot (vs network).  OK to include
                        ilBucketIndex = -1
                        For ilLoop = ilFirstQ To ilNoWks Step 1   '7-5-00
                            If (llDate >= lmSAvailsDates(ilLoop)) And (llDate <= lmEAvailsDates(ilLoop)) Then
                                ilBucketIndex = ilLoop
                                Exit For
                            End If
                        Next ilLoop
                        If ilBucketIndex > 0 Then
                            ilBucketIndexMinusOne = ilBucketIndex - 1
                            slDate = Format$(llDate, "m/d/yy")
                            gPackDate slDate, ilDate0, ilDate1
                            gObtainWkNo 0, slDate, ilWkNo, ilFirstWk        'obtain the week bucket number
                            ilDay = gWeekDayLong(llDate)
                            gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                            For ilRdf = LBound(tmAvRdf) To UBound(tmAvRdf) - 1 Step 1
                                ilAvailOk = False
                                If (tmAvRdf(ilRdf).iLtfCode(0) <> 0) Or (tmAvRdf(ilRdf).iLtfCode(1) <> 0) Or (tmAvRdf(ilRdf).iLtfCode(2) <> 0) Then
                                    If (ilLtfCode = tmAvRdf(ilRdf).iLtfCode(0)) Or (ilLtfCode = tmAvRdf(ilRdf).iLtfCode(1)) Or (ilLtfCode = tmAvRdf(ilRdf).iLtfCode(1)) Then
                                        ilAvailOk = False    'True- code later
                                    End If
                                Else
                                    For ilLoop = LBound(tmAvRdf(ilRdf).iStartTime, 2) To UBound(tmAvRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                        If (tmAvRdf(ilRdf).iStartTime(0, ilLoop) <> 1) Or (tmAvRdf(ilRdf).iStartTime(1, ilLoop) <> 0) Then
                                            gUnpackTimeLong tmAvRdf(ilRdf).iStartTime(0, ilLoop), tmAvRdf(ilRdf).iStartTime(1, ilLoop), False, llStartTime
                                            gUnpackTimeLong tmAvRdf(ilRdf).iEndTime(0, ilLoop), tmAvRdf(ilRdf).iEndTime(1, ilLoop), True, llEndTime
                                            If UBound(tmAvRdf) - 1 = LBound(tmAvRdf) Then   'could be a conv bumped spot sched in
                                                                                        'in conven veh.  The VV has DP times different than the
                                                                                        'conven veh.
                                                llStartTime = llTime
                                                llEndTime = llTime + 1              'actual time of spot
                                            End If
                                            'Don't include the end time i.e. 10a-3p is 10a thru 2:59:59p
                                            'If (llTime >= llStartTime) And (llTime < llEndTime) And (tmAvRdf(ilRdf).sWkDays(ilLoop, ilDay + 1) = "Y") Then
                                            If (llTime >= llStartTime) And (llTime < llEndTime) And (tmAvRdf(ilRdf).sWkDays(ilLoop, ilDay) = "Y") Then
                                                ilAvailOk = True
                                                ilLoopIndex = ilLoop
                                                slDays = ""
                                                For ilDayIndex = 1 To 7 Step 1
                                                    If (tmAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "Y") Or (tmAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "N") Then
                                                        slDays = slDays & tmAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1)
                                                    Else
                                                        slDays = slDays & "N"
                                                    End If
                                                Next ilDayIndex
                                                Exit For
                                            End If
                                        End If
                                    Next ilLoop
                                End If
                                If ilAvailOk Then
                                    'Determine if Avr created
                                    ilFound = False
                                    ilSaveDay = ilDay
                                    ilDay = 0                                       'force all data in same day of week
                                    For ilRec = 0 To UBound(tmAvr) - 1 Step 1
                                        'If (tmAvr(ilRec).iRdfCode = tmAvRdf(ilRdf).iCode) And (tmAvr(ilRec).iFirstBucket = ilFirstQ) And (tmAvr(ilRec).iDay = ilDay) Then
                                        If (ilRdfCodes(ilRec) = tmAvRdf(ilRdf).iCode) And (tmAvr(ilRec).iFirstBucket = ilFirstQ) And (tmAvr(ilRec).iDay = ilDay) Then
                                            ilFound = True
                                            ilRecIndex = ilRec
                                            Exit For
                                        End If
                                    Next ilRec
                                    If Not ilFound Then
                                        ilRecIndex = UBound(tmAvr)
                                        tmAvr(ilRecIndex).iGenDate(0) = igNowDate(0)
                                        tmAvr(ilRecIndex).iGenDate(1) = igNowDate(1)
                                        'tmAvr(ilRecIndex).iGenTime(0) = igNowTime(0)
                                        'tmAvr(ilRecIndex).iGenTime(1) = igNowTime(1)
                                        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                        tmAvr(ilRecIndex).lGenTime = lgNowTime
                                        tmAvr(ilRecIndex).iDay = ilDay
                                        tmAvr(ilRecIndex).iQStartDate(0) = ilSAvailsDates(0)
                                        tmAvr(ilRecIndex).iQStartDate(1) = ilSAvailsDates(1)
                                        tmAvr(ilRecIndex).iFirstBucket = ilFirstQ
                                        tmAvr(ilRecIndex).sBucketType = smBucketType
                                        tmAvr(ilRecIndex).iDPStartTime(0) = tmAvRdf(ilRdf).iStartTime(0, ilLoopIndex)
                                        tmAvr(ilRecIndex).iDPStartTime(1) = tmAvRdf(ilRdf).iStartTime(1, ilLoopIndex)
                                        tmAvr(ilRecIndex).iDPEndTime(0) = tmAvRdf(ilRdf).iEndTime(0, ilLoopIndex)
                                        tmAvr(ilRecIndex).iDPEndTime(1) = tmAvRdf(ilRdf).iEndTime(1, ilLoopIndex)
                                        tmAvr(ilRecIndex).sDPDays = slDays
                                        tmAvr(ilRecIndex).sNot30Or60 = "N"
                                        tmAvr(ilRecIndex).iVefCode = ilVefCode
                                        'tmAvr(ilRecIndex).iRdfCode = tmAvRdf(ilRdf).iCode
                                        tmAvr(ilRecIndex).iRdfCode = tmAvRdf(ilRdf).iSortCode
                                        ilRdfCodes(ilRecIndex) = tmAvRdf(ilRdf).iCode
                                        tmAvr(ilRecIndex).sInOut = tmAvRdf(ilRdf).sInOut
                                        tmAvr(ilRecIndex).ianfCode = tmAvRdf(ilRdf).ianfCode
    
                                        ReDim Preserve tmAvr(0 To ilRecIndex + 1) As AVR
                                        ReDim Preserve ilRdfCodes(0 To ilRecIndex + 1)
                                    End If
    
    
                                    tmAvr(ilRecIndex).lRate(ilBucketIndexMinusOne) = tmRifRate(ilRdf).lRate(ilWkNo)
                                    ilDay = ilSaveDay
                                    ilNo30 = 0
                                    ilNo60 = 0
                                    ilLen = tmSdf.iLen
                                    If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                    'Convert inventory to number of 30's and 60's
                                        Do While ilLen >= 60
                                            ilNo60 = ilNo60 + 1
                                            ilLen = ilLen - 60
                                        Loop
                                        Do While ilLen >= 30
                                            ilNo30 = ilNo30 + 1
                                            ilLen = ilLen - 30
                                        Loop
                                        If ilLen < 30 And ilLen > 0 Then    '7-6-00 assume anything under 30" is 1-30" unit availability
                                            ilNo30 = ilNo30 + 1
                                            ilLen = 0
                                        End If
    
                                        If (smBucketType = "S") Or (smBucketType = "P") Then
                                            tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                            tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                        Else
                                            tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) - ilNo60
                                            tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) - ilNo30
                                        End If
                                        'adjust the available bucket (used for qrtrly detail report only)
                                        tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                        tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                        'Count 30 or 60 and set flag if neither
                                        If ilLen = 60 Then
                                            ilNo60 = 1
                                        ElseIf ilLen = 30 Then
                                            ilNo30 = 1
                                        Else
                                            tmAvr(ilRecIndex).sNot30Or60 = "Y"
                                            If ilLen <= 30 Then
                                                ilNo30 = 1
                                            Else
                                                ilNo60 = 1
                                            End If
                                        End If
                                        If (smBucketType = "S") Or (smBucketType = "P") Then
                                            tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                            tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                        Else
                                            tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) - ilNo60
                                            tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) - ilNo30
                                        End If
                                        If ilNo60 > 0 Then                      'spot found a 60?
                                            tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                        Else
                                            If tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) > 0 Then
                                                tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                            Else
                                                If tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) > 0 Then
                                                    tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo30
                                                Else                        'oversold units
                                                    tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                                End If
                                            End If
                                        End If
    
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                        'Count 30 or 60 and set flag if neither
                                        If ilLen = 60 Then
                                            ilNo60 = 1
                                        ElseIf ilLen = 30 Then
                                            ilNo30 = 1
                                        Else
                                            tmAvr(ilRecIndex).sNot30Or60 = "Y"
                                        End If
                                        If (smBucketType = "S") Or (smBucketType = "P") Then
                                            tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                            tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                        Else
                                            tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) - ilNo30
                                            tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) - ilNo60
                                        End If
                                        'adjust the available bucket (used for qrtrly detail report only)
                                        tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                        tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                    End If
                                End If
                            Next ilRdf
                        End If
                    End If
                End If
                ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Next ilPass

    End If

    'Adjust counts
    'If (smBucketType = "A") And (tgVpf(ilVpfIndex).sSSellOut = "B" Or tgVpf(ilVpfIndex).sSSellOut = "U") Then
    If (smBucketType = "A" And tgVpf(ilVpfIndex).sSSellOut = "B") Then
        For ilRec = 0 To UBound(tmAvr) - 1 Step 1
            'For ilLoop = 1 To ilNoWks Step 1    '7-5-00
            For ilLoop = 0 To ilNoWks - 1 Step 1  '7-5-00
                If tmAvr(ilRec).i30Count(ilLoop) < 0 Then
                    Do While (tmAvr(ilRec).i60Count(ilLoop) > 0) And (tmAvr(ilRec).i30Count(ilLoop) < 0)
                        tmAvr(ilRec).i60Count(ilLoop) = tmAvr(ilRec).i60Count(ilLoop) - ilAdjSub    '1
                        tmAvr(ilRec).i30Count(ilLoop) = tmAvr(ilRec).i30Count(ilLoop) + ilAdjAdd    '2
                    Loop
                ElseIf (tmAvr(ilRec).i60Count(ilLoop) < 0) Then
                End If
            Next ilLoop
        Next ilRec
    End If
    'Adjust counts for qtrly detail availability
    'If (smBucketType = "A") And (tgVpf(ilVpfIndex).sSSellOut = "B" Or tgVpf(ilVpfIndex).sSSellOut = "U") Then
    If (smBucketType = "A" And tgVpf(ilVpfIndex).sSSellOut = "B") Then
        For ilRec = 0 To UBound(tmAvr) - 1 Step 1
            'For ilLoop = 1 To ilNoWks Step 1    '7-5-00
            For ilLoop = 0 To ilNoWks - 1 Step 1  '7-5-00
                If tmAvr(ilRec).i30Avail(ilLoop) < 0 Then
                   Do While (tmAvr(ilRec).i60Avail(ilLoop) > 0) And (tmAvr(ilRec).i30Avail(ilLoop) < 0)
                        tmAvr(ilRec).i60Avail(ilLoop) = tmAvr(ilRec).i60Avail(ilLoop) - 1
                        tmAvr(ilRec).i30Avail(ilLoop) = tmAvr(ilRec).i30Avail(ilLoop) + 2
                    Loop
                ElseIf (tmAvr(ilRec).i60Avail(ilLoop) < 0) Then
                End If
            Next ilLoop
        Next ilRec
    End If

    'Combines weeks into the proper months for the weekly and monthly report
    For ilRec = 0 To UBound(tmAvr) - 1 Step 1  'next daypart
        For ilBucketIndex = 1 To ilNoWks Step 1 '7-5-00
            ilBucketIndexMinusOne = ilBucketIndex - 1
            tmAvr(ilRec).i60Hold(ilBucketIndexMinusOne) = tmAvr(ilRec).i60Count(ilBucketIndexMinusOne)
            tmAvr(ilRec).i30Hold(ilBucketIndexMinusOne) = tmAvr(ilRec).i30Count(ilBucketIndexMinusOne)
        Next ilBucketIndex
        For ilLoop = 1 To 3 Step 1
            If ilLoop = 1 Then
                ilLo = 1
                ilHi = ilWksInMonth(1)
            Else
                ilLo = ilHi + 1
                ilHi = ilHi + ilWksInMonth(ilLoop)
            End If
            For ilWkInx = ilLo To ilHi Step 1
                If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                    tmAvr(ilRec).lMonth(ilLoop - 1) = tmAvr(ilRec).lMonth(ilLoop - 1) + (((tmAvr(ilRec).i60Count(ilWkInx - 1) * 2) + tmAvr(ilRec).i30Count(ilWkInx - 1)) * tmAvr(ilRec).lRate(ilWkInx - 1))
                ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Or tgVpf(ilVpfIndex).sSSellOut = "U" Then
                    tmAvr(ilRec).lMonth(ilLoop - 1) = tmAvr(ilRec).lMonth(ilLoop - 1) + ((tmAvr(ilRec).i60Count(ilWkInx - 1) + tmAvr(ilRec).i30Count(ilWkInx - 1)) * tmAvr(ilRec).lRate(ilWkInx - 1))
                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                End If
            Next ilWkInx
        Next ilLoop
    Next ilRec
Erase ilSAvailsDates
Erase ilEvtType
Erase ilRdfCodes
Erase tlLLC
End Sub
