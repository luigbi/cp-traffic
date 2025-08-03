Attribute VB_Name = "RPTSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptsubs.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  hmRaf                         imShfRecLen                   tmShf                     *
'*  hmShf                         tmShfSrchKey                  imMktRecLen               *
'*  tmMkt                         hmMkt                         tmMktSrchKey              *
'*                                                                                        *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  gGetMonthsForYr               gYearForCorpStartMonth                                  *
'******************************************************************************************

Option Explicit
Option Compare Text

Global Const FEED_BYCONVERTED = &H2
Global Const FEED_BYNEEDSCONVERT = &H4
Global Const FEED_BYINSERT = &H1

'Margin Acquisition
'Global Const SORT1_ADVCNT = 0
'Global Const SORT1_ADVCNTVEH = 1
'Global Const SORT1_SLSP = 2
'Global Const SORT1_VEHCNT = 3
'Global Const SORT1_GROUP = 4
'
'Global Const SORT2_NONE = 0
'Global Const SORT2_ADVCNT = 1
'Global Const SORT2_ADVCNTVEH = 2
'Global Const SORT2_SLSP = 3
'Global Const SORT2_VEHCNT = 4
'Global Const SORT2_GROUP = 5
'
'Global Const SORT3_NONE = 0
'Global Const SORT3_ADVCNT = 1
'Global Const SORT3_ADVCNTVEH = 2
'Global Const SORT3_VEHCNT = 3

Global Const SORT1_NONE = 0
Global Const SORT1_SLSP = 1
Global Const SORT1_GROUP = 2

Global Const SORT2_NONE = 0
Global Const SORT2_SLSP = 1
Global Const SORT2_GROUP = 2

Global Const SORT3_ADVCNT = 0
Global Const SORT3_ADVCNTVEH = 1
Global Const SORT3_VEHCNT = 2
' End Martin Acquisition

Dim tmVsf As VSF
Dim tmClf As CLF
Dim imRafRecLen As Integer
Dim tmRaf As RAF
Dim tmRafSrchKey As LONGKEY0


Dim tmTxr As TXR
'
'                   gIsCntrRep - determine if there is at least 1 rep vehicle on the order
'
Function gIsCntrRep(llVsfCode As Long, hlVsf As Integer, ilRepVehicles() As Integer) As Integer
'
'   Return:  False if No vehicle is REP  True if at least 1 vehicle is REP
'
    Dim ilVeh As Integer
    Dim ilVefCode As Integer
    Dim llLkVsfCode As Long
    Dim ilVsfReclen As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim tlVsfSrchKey As LONGKEY0

    gIsCntrRep = False
    If (Asc(tgSpf.sUsingFeatures) And USINGREP) <> USINGREP Then
        Exit Function
    End If
    If (UBound(ilRepVehicles) <= 0) Then
        Exit Function
    End If
    ilVsfReclen = Len(tmVsf)
    If llVsfCode > 0 Then
        For ilVeh = LBound(ilRepVehicles) To UBound(ilRepVehicles) - 1 Step 1
            ilVefCode = ilRepVehicles(ilVeh)
            If llVsfCode = ilVefCode Then
                gIsCntrRep = True
                Exit Function
            End If
        Next ilVeh
    ElseIf llVsfCode < 0 Then
        llLkVsfCode = -llVsfCode
        Do While llLkVsfCode > 0
            tlVsfSrchKey.lCode = llLkVsfCode
            ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Function
            End If
            For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                If tmVsf.iFSCode(ilLoop) > 0 Then
                    For ilVeh = LBound(ilRepVehicles) To UBound(ilRepVehicles) - 1 Step 1
                        ilVefCode = ilRepVehicles(ilVeh)
                        If tmVsf.iFSCode(ilLoop) = ilVefCode Then
                            gIsCntrRep = True
                            Exit Function
                        End If
                    Next ilVeh
                End If
            Next ilLoop
            llLkVsfCode = tmVsf.lLkVsfCode
        Loop
    End If
End Function
Public Function gSetCheckStr(lCntrl As Integer) As String
    'Returns N (no) if the control equals vbUnChecked.
    'Returns Y (yes) if the control equals vbChecked or vbGray.

    If lCntrl = vbUnchecked Then
        gSetCheckStr = "N"
    Else
        gSetCheckStr = "Y"
    End If
End Function
'
'
'           For User selection, use the list of Traffic & AFfiliate users.
'           Need special code to handle the traffic users vs affiliate users
'
'           <input> ilincludecodes - true if use the array of codes to include
'                                    false if use the array of codes to exclude
'                   ilusecodes() - array of codes to include or exclude
'                   ilIndex - index of which lbcselection box (10-9-03)
Public Sub gObtainUrfUstCodes(ilIncludeCodes As Integer, ilUseCodes() As Integer, ilIndex As Integer, Form As Form)
Dim ilHowManyDefined As Integer
Dim ilHowMany As Integer
Dim ilLoop As Integer
Dim slNameCode As String
Dim ilRet As Integer
Dim slCode As String

    ilHowManyDefined = Form!lbcSelection(ilIndex).ListCount
    ilHowMany = Form!lbcSelection(ilIndex).SelCount
    If ilHowMany > ilHowManyDefined / 2 Then    'more than half selected
        ilIncludeCodes = False
    Else
        ilIncludeCodes = True
    End If
    For ilLoop = 0 To Form!lbcSelection(ilIndex).ListCount - 1 Step 1
        slNameCode = tgUserSortCode(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If Form!lbcSelection(ilIndex).Selected(ilLoop) And ilIncludeCodes Then               'selected
            If InStr(slNameCode, "/Traffic") > 0 Then           '
                ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
            Else
                ilUseCodes(UBound(ilUseCodes)) = -Val(slCode)       'negate the affiliate user codes
            End If
            ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
        Else        'exclude these
            If (Not Form!lbcSelection(ilIndex).Selected(ilLoop)) And (Not ilIncludeCodes) Then
                If InStr(slNameCode, "/Traffic") > 0 Then      '
                    ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
                Else
                    ilUseCodes(UBound(ilUseCodes)) = -Val(slCode)
                End If
                ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
            End If
        End If
    Next ilLoop
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
Public Sub gObtainAgyAdvCodes(ilIncludeCodes As Integer, ilUseCodes() As Integer, ilIndex As Integer, Form As Form)
Dim ilHowManyDefined As Integer
Dim ilHowMany As Integer
Dim ilLoop As Integer
Dim slNameCode As String
Dim ilRet As Integer
Dim slCode As String
'ReDim ilUseCodes(1 To 1) As Integer
ReDim ilUseCodes(0 To 0) As Integer
    ilHowManyDefined = Form!lbcSelection(ilIndex).ListCount
    ilHowMany = Form!lbcSelection(ilIndex).SelCount
    If ilHowMany > ilHowManyDefined / 2 Then    'more than half selected
        ilIncludeCodes = False
    Else
        ilIncludeCodes = True
    End If
    For ilLoop = 0 To Form!lbcSelection(ilIndex).ListCount - 1 Step 1
        slNameCode = Form!lbcAgyAdvtCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If Form!lbcSelection(ilIndex).Selected(ilLoop) And ilIncludeCodes Then               'selected ?
            If InStr(slNameCode, "/Direct") = 0 And InStr(slNameCode, "/Non-") = 0 Then         'not a direct, and not a direct that was changed to reg agency, this is reg agency
                ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
            Else
                ilUseCodes(UBound(ilUseCodes)) = -Val(slCode)
            End If
            ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
        Else        'exclude these
            If (Not Form!lbcSelection(ilIndex).Selected(ilLoop)) And (Not ilIncludeCodes) Then
                If InStr(slNameCode, "/Direct") = 0 And InStr(slNameCode, "/Non-") = 0 Then     'not a direct, and not a direct that was changed to reg agency, this is reg agency
                    ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
                Else
                    ilUseCodes(UBound(ilUseCodes)) = -Val(slCode)
                End If
                ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
            End If
        End If
    Next ilLoop
End Sub
'       Filter the SpotRanks for Remnant, Per Inquiry, trade, Extra,
'               PSA & Promo spots
'
'       <input> tlCntTypes - structure of elements to be included/excluded
'       <output> - set ilCtypes bit for the type of spot rank
'                  set ilSpotTypes bit for the type of spot
'                  set ilSpotOK = false if spot should be excluded
'   ilCTypes        'bit map of cnt types to include starting lo order bit
'                    '0 = unused, 1= unused, 2 = network, 3 = std, 4 = Reserved, 5 = remanant, 6 = DR
'                    '7 = PI, 8 = psa, 9 = promo, 10 = trade
'                    'bit 0 and 1 previously hold & order; but it conflicts with other contract types
'
'   ilSpotTypes     'bit map of spot types to include starting lo order bit:
                    '0 = missed (not used in this report), 1 = charge, 2 = 0.00, 3 = adu, 4 = bonus, 5 = extra
                    '6 = fill, 7 = n/c 8 = recapturable, 9 = spinoff
'   tlSpot - SSF image of spot entry
Sub gFilterSpotTypes(ilCTypes As Integer, ilSpotTypes As Integer, ilSpotOK As Integer, tlCntTypes As CNTTYPES, tlSpot As CSPOTSS)
    If (tlSpot.iRank And RANKMASK) = REMNANTRANK Then
        ilCTypes = &H10
        If Not tlCntTypes.iRemnant Then
            ilSpotOK = False
        End If

    ElseIf (tlSpot.iRank And RANKMASK) = PERINQUIRYRANK Then
        ilCTypes = &H800
        If Not tlCntTypes.iPI Then
            ilSpotOK = False
        End If

    ElseIf (tlSpot.iRank And RANKMASK) = TRADERANK Then
        ilCTypes = &H400
        If Not tlCntTypes.iTrade Then
            ilSpotOK = False
        End If

    'ElseIf tlSpot.iRank = 1045 Then
    '    ilSpotTypes = &H20
    '    If Not tlCntTypes.iXtra Then
    '        ilSpotOK = False
    '    End If

    ElseIf (tlSpot.iRank And RANKMASK) = PROMORANK Then
        ilCTypes = &H200
        If Not tlCntTypes.iPromo Then
            ilSpotOK = False
        End If

    ElseIf (tlSpot.iRank And RANKMASK) = PSARANK Then
        ilCTypes = &H100
        If Not tlCntTypes.iPSA Then
            ilSpotOK = False
        End If
    End If
    'always include all spots (including the primary/seconday split network spots)
    'If (tlSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
    '    ilSpotOK = False
    'End If
End Sub
'
'       gGetNTRTrfCode - obtain the auto code for the Tax record used
'       to compute the taxes for NTR
'       <input>
'       <output> none
'       <return> trf record code used to to obtain the tax rates, else 0
'
Public Function gGetNTRTrfCode(ilNTRTrfCode As Integer) As Integer
    Dim ilTrf As Integer

    gGetNTRTrfCode = 0
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) <> TAXONNTR) Then
        Exit Function
    End If
    If ilNTRTrfCode <= 0 Then
        Exit Function
    End If
    ilTrf = gBinarySearchTrf(ilNTRTrfCode)
    If ilTrf <> -1 Then
        gGetNTRTrfCode = tgTrf(ilTrf).iCode
    End If
End Function

'
'       gGetAirTimeTrfCode - obtain the auto code for the Tax record used
'       to compute the taxes for Air Time
'       <input> iladfcode - advertiser code
'               ilagfcode - agency code
'               ilvefcode = vehicle code
'       <output> none
'       <return> trf record code used to to obtain the tax rates, else 0
'
Public Function gGetAirTimeTrfCode(ilAdfCode As Integer, ilAgfCode As Integer, ilVefCode As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVef                                                                                 *
'******************************************************************************************

    Dim ilTrfVef As Integer
    Dim ilTrfAgyAdvt As Integer

    gGetAirTimeTrfCode = 0
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) <> TAXONAIRTIME) Then
        Exit Function
    End If
    ilTrfAgyAdvt = gGetTrfIndexForAgyAdvt(ilAdfCode, ilAgfCode)
    If ilTrfAgyAdvt < 0 Then
        Exit Function
    End If
    ilTrfVef = gGetTrfIndexForVeh(ilVefCode)
    If ilTrfVef <> -1 Then
        If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
            'Match State
            If StrComp(Trim$(tgTrf(ilTrfAgyAdvt).sTax1Name), Trim$(tgTrf(ilTrfVef).sTax1Name), vbTextCompare) = 0 Then
                gGetAirTimeTrfCode = tgTrf(ilTrfAgyAdvt).iCode

            End If
        Else
            gGetAirTimeTrfCode = tgTrf(ilTrfVef).iCode

        End If
    End If
End Function

'
'       gFilterLists - check the option and which list boxes to test
'       for inclusion/exclusion
'
'       <input>
'               'ilWhichField - 0 = advt, 1 = vehicle
'               ilWhichField = 12-14-05 value of field to compare for inclusion/exclusion
'               ilIncludeCodes = true to include codes in array;
'                                false to exclude codes in array
'               ilUseCodes()- array of codes to include/exclude
'       <return> true = include transaction, else false to exclude
'
'       12-14-05 change the parameters.  Send the field to compare rathern
'       than sending a flag and the buffer for record to retrieve item to compare
'
Public Function gFilterLists(ilWhichField As Integer, ilIncludeCodes As Integer, ilUseCodes() As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************

Dim ilCompare As Integer
Dim ilTemp As Integer
Dim ilFoundOption As Integer

    ilFoundOption = False
    ilCompare = ilWhichField            '12-14-05
    'If ilWhichField = 0 Then            'test advt
     '   ilCompare = tlCrf.iAdfCode
    'ElseIf ilWhichField = 1 Then
    '    ilCompare = tlCrf.iVefCode      'test vehicle
    'End If

     If ilIncludeCodes Then
        For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
            If ilUseCodes(ilTemp) = ilCompare Then
                ilFoundOption = True
                Exit For
            End If
        Next ilTemp
    Else
        ilFoundOption = True
        For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
            If ilUseCodes(ilTemp) = ilCompare Then
                ilFoundOption = False
                Exit For
            End If
        Next ilTemp
    End If

    gFilterLists = ilFoundOption
End Function
'
'
'           Format Spot Time as HH:MM:SS
'
'       <input> llSpotTime - time of spot
'       <output> Return string of HH:MM:SS
'
Public Function gFormatSpotTime(llSpotTime As Long) As String
Dim llHour As Long
Dim llMin As Long
Dim llSec As Long
Dim slStr As String
Dim slWholeTime As String
    'Time (HH:MM:ss)

    slWholeTime = ""
    llHour = llSpotTime \ 3600
    llMin = llSpotTime Mod 3600
    llSec = llMin Mod 60
    llMin = llMin \ 60
    slStr = Trim$(str$(llHour))
    Do While Len(slStr) < 2
        slStr = "0" & slStr
    Loop
    slWholeTime = slWholeTime & slStr & ":"
    'Minutes
    slStr = Trim$(str$(llMin))
    Do While Len(slStr) < 2
        slStr = "0" & slStr
    Loop
    slWholeTime = slWholeTime & slStr & ":"
    'Seconds
    slStr = Trim$(str$(llSec))
    Do While Len(slStr) < 2
        slStr = "0" & slStr
    Loop
    slWholeTime = slWholeTime & slStr
    gFormatSpotTime = slWholeTime
End Function
'
'
'
'               gObtainCodes - get all codes to process or exclude
'               When selecting advt, agy or vehicles (anthing in list box)--make testing
'               of selection more efficient.  If more than half of
'               the entries are selected, create an array with entries
'               to exclude.  If less than half of entries are selected,
'               create an array with entries to include.
'               <input> ilListIndex - list box to test
'                       lbcListbox - array containing sort codes
'               <output> ilIncludeCodes - true if test to include the codes in array
'                                          false if test to exclude the codes in array
'                        ilUseCodes - array of advt, agy or vehicles codes to include/exclude
Sub gObtainCodes(lbcSelection As Control, lbcListBox() As SORTCODE, ilIncludeCodes, ilUseCodes() As Integer, Form As Form)
Dim ilHowManyDefined As Integer
Dim ilHowMany As Integer
Dim slNameCode As String
Dim ilLoop As Integer
Dim slCode As String
Dim ilRet As Integer
'ReDim ilUseCodes(1 To 1) As Integer
ReDim ilUseCodes(0 To 0) As Integer
    ilHowManyDefined = Form!lbcSelection.ListCount
    'ilHowMany = RptSel!lbcSelection(ilListIndex).SelectCount
    ilHowMany = Form!lbcSelection.SelCount
    If ilHowMany > (ilHowManyDefined / 2) + 1 Then   'more than half selected
        ilIncludeCodes = False
    Else
        ilIncludeCodes = True
    End If

    For ilLoop = 0 To Form!lbcSelection.ListCount - 1 Step 1
        slNameCode = lbcListBox(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If Form!lbcSelection.Selected(ilLoop) And ilIncludeCodes Then               'selected ?
            ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
            ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
        Else        'exclude these
            If (Not Form!lbcSelection.Selected(ilLoop)) And (Not ilIncludeCodes) Then
                ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
                ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
            End If
        End If
    Next ilLoop
End Sub
'
'                   sub gTestCostType - test Spot type for inclusion/exclusion
'                   Include different types of spots test
'                   <input> ilCosttype - bit string based on user request of
'                           types of spots to include
'                           slStrCost - string defining cost of spot ($ value as string
'                                       or text such as ADU , bonus, etc.
'                   <Return>  false if not a spot to report
'
Function gTestCostType(ilCostType As Integer, slStrCost As String) As Integer
Dim ilOk As Integer

    ilOk = True
    'look for inclusion of charge spots with
    If (InStr(slStrCost, ".") <> 0) Then        'found spot cost
        'is it a .00?
        If gCompNumberStr(slStrCost, "0.00") = 0 Then       'its a .00 spot
            If (ilCostType And SPOT_00) <> SPOT_00 Then      'include .00?
                ilOk = False
            End If
        Else
            If (ilCostType And SPOT_CHARGE) <> SPOT_CHARGE Then    'include charged spots?
                ilOk = False                                            'no
            End If
        End If
    ElseIf Trim$(slStrCost) = "ADU" And (ilCostType And SPOT_ADU) <> SPOT_ADU Then
            ilOk = False
    ElseIf Trim$(slStrCost) = "Bonus" And (ilCostType And SPOT_BONUS) <> SPOT_BONUS Then
            ilOk = False
    'ElseIf Trim$(slStrCost) = "Extra" And (ilCostType And SPOT_EXTRA) <> SPOT_EXTRA Then
        ElseIf Trim$(slStrCost) = "+ Fill" And (ilCostType And SPOT_EXTRA) <> SPOT_EXTRA Then
            ilOk = False
    'ElseIf Trim$(slStrCost) = "Fill" And (ilCostType And SPOT_FILL) <> SPOT_FILL Then
    ElseIf Trim$(slStrCost) = "- Fill" And (ilCostType And SPOT_FILL) <> SPOT_FILL Then

            ilOk = False
    ElseIf Trim$(slStrCost) = "N/C" And (ilCostType And SPOT_NC) <> SPOT_NC Then
            ilOk = False
    ElseIf Trim$(slStrCost) = "MG" And (ilCostType And SPOT_MG) <> SPOT_MG Then
            ilOk = False
    ElseIf Trim$(slStrCost) = "Recapturable" And (ilCostType And SPOT_RECAP) <> SPOT_RECAP Then
            ilOk = False
    ElseIf Trim$(slStrCost) = "Spinoff" And (ilCostType And SPOT_SPINOFF) <> SPOT_SPINOFF Then
            ilOk = False
    End If

    gTestCostType = ilOk
End Function





'
'                   gRptMnfPop - Populate list box with MNF records
'                           slType = Mnf type to match (i.e. "H", "A")
'                           lbcLocal  - local list box to fill
'                           lbcMster - master list box with codes
'                   Created: DH 9/12/96
'
Function gRptMnfPop(frm As Form, slType As String, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer   'lbcMster As Control)
ReDim ilfilter(0) As Integer
ReDim slFilter(0) As String
ReDim ilOffSet(0) As Integer
Dim ilRet As Integer
    ilfilter(0) = CHARFILTER
    slFilter(0) = slType
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType")

    'ilRet = gIMoveListBox(RptSelCt, lbcLocal, lbcMster, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(frm, lbcLocal, tlSortCode(), slSortCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo gRptMnfPopErr
        gCPErrorMsg ilRet, "gRptMnfPop (gImoveListBox)", frm
        On Error GoTo 0
    End If
    Exit Function
gRptMnfPopErr:
 On Error GoTo 0
    gRptMnfPop = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gRptSPersonPop                  *
'*                                                     *
'*             Created:6/4-2-03       By:D. Hosaka     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Salespeople  list box *
'*                      if required for reports        *
'*                                                     *
'*******************************************************
Function gRptSPersonPop(frm As Form, lbcSelection As Control) As Integer
'
'   gRptSPersonPop
'   Where:
'       Frm: Form calling the populate
'       lbcSelection: list box to populate
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopSalespersonBox(Frm, 0, True, True, lbcSelection, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    '7-8-13 show the slsp allowed to see
    ilRet = gPopSalespersonBox(frm, 5, True, True, lbcSelection, tgSalesperson(), sgSalespersonTag, igSlfFirstNameFirst)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo gRptSPersonPopErr
        gCPErrorMsg ilRet, "gRptSPersonPop (gPopSalespersonBox)", frm
        On Error GoTo 0
    End If
    Exit Function
gRptSPersonPopErr:
    On Error GoTo 0
    gRptSPersonPop = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gRptSellConvVehPop              *
'*                                                     *
'*             Created: 4-2-03       By:D. Hoska       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate selling, convention,  *
'*            rep vehiclesbox                           *
'*                                                     *
'*******************************************************
Function gRptSellConvVehPop(frm As Form, lbcSelection As Control) As Integer
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(frm, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHNTR + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + ACTIVEVEH, lbcSelection, tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo gRptSellConvVehPopErr
        gCPErrorMsg ilRet, "gRptSellConvVehPop (gPopUserVehicleBox: Vehicle)", frm
        On Error GoTo 0
    End If
    Exit Function
gRptSellConvVehPopErr:
    On Error GoTo 0
    gRptSellConvVehPop = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gRptAgencyPop                   *
'*                                                     *
'*             Created: 4-2-03         By:D. Hosaka    *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Agency list box       *
'*                      if required for Reports        *
'*                                                     *
'*******************************************************
Function gRptAgencyPop(frm As Form, lbcSelection As Control) As Integer
'
'   gRptAgencyPop
'   Where:
'       Frm: Form calling this populate
'       lbcSelection: list box to fill
'
    Dim ilRet As Integer

    ilRet = gPopAgyBox(frm, lbcSelection, tgAgency(), sgAgencyTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo gRptAgencyPopErr
        gCPErrorMsg ilRet, "gRptAgencyPop (gPopAgyBox)", frm
        On Error GoTo 0
    End If
    Exit Function
gRptAgencyPopErr:
 On Error GoTo 0
    gRptAgencyPop = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name: gRptAdvtPop                    *
'*                                                     *
'*             Created: 4-02-03      By:D. Hosaka      *
'*                                                     *
'*            Comments: Populate Advertiser list box   *
'*                      if required for Reports        *
'*                                                     *
'*******************************************************
Function gRptAdvtPop(frm As Form, lbcSelection As Control) As Integer
'
'   gRptAdvtPop
'   Where: lbcSelection = list box to populate
'
    Dim ilRet As Integer

    ilRet = gPopAdvtBox(frm, lbcSelection, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo gRptAdvtPopErr
        gCPErrorMsg ilRet, "gRptAdvtPop (gPopAdvtBox)", frm
        On Error GoTo 0
    End If
    Exit Function
gRptAdvtPopErr:
 On Error GoTo 0
    'imTerminate = True
    gRptAdvtPop = True
    Exit Function
End Function

'           gGetMonthsForYr - Setup array of standard broadcast start dates for the year
'                   based on the Corp or STd calendar.  Array must be 13 buckets for the start
'                   date of the 13th month.
'
'           <input> ilCorpStd - 1 = Corp, 2 = std
'                   ilQtr - start quarter requested
'                   ilYear - start year requested
'           <output>llStdStartDates()- array of monthly start dates
'                   llLastBilled - last date invoices from Site Pref
'                   ilLastBilledInx - last month (as a month #) invoiced
'
'
Sub gGetMonthsForYr(ilCorpStd As Integer, ilQtr As Integer, ilYear As Integer, llStdStartDates() As Long, llLastBilled As Long, ilLastBilledInx As Integer) 'VBC NR
Dim slStr As String 'VBC NR
Dim ilLoop As Integer 'VBC NR
Dim llDate As Long 'VBC NR
Dim ilYearInx As Integer 'VBC NR
Dim ilStartMonth As Integer 'VBC NR
Dim ilNextDate As Integer 'VBC NR
Dim llSaveLastDate As Long 'VBC NR

    'Determine last month invoiced for past & future calculations
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr       'convert last bdcst billing date to string 'VBC NR
    llLastBilled = gDateValue(slStr)            'convert last month billed to long 'VBC NR

    ilLoop = (ilQtr - 1) * 3 + 1 'VBC NR

    If ilCorpStd = 1 Then                'build array of corp start months 'VBC NR
        'Determine what month the fiscal year starts
        ilYearInx = gGetCorpCalIndex(ilYear) 'VBC NR
        ilNextDate = 1 'VBC NR
        For ilStartMonth = ilLoop To 12             'start from requested qtr of corp cal, may wrap around to next year if not starting at 1st qtr 'VBC NR
            gUnpackDateLong tgMCof(ilYearInx).iStartDate(0, ilStartMonth - 1), tgMCof(ilYearInx).iStartDate(1, ilStartMonth - 1), llStdStartDates(ilNextDate) 'VBC NR
            gUnpackDateLong tgMCof(ilYearInx).iEndDate(0, ilStartMonth - 1), tgMCof(ilYearInx).iEndDate(1, ilStartMonth - 1), llSaveLastDate 'VBC NR
            ilNextDate = ilNextDate + 1 'VBC NR
        Next ilStartMonth 'VBC NR
        If ilNextDate <> 12 Then                'corp year wrap around; started from qtr other than 1 'VBC NR
            ilYearInx = gGetCorpCalIndex(ilYear + 1) 'VBC NR
            For ilStartMonth = 1 To (12 - ilNextDate) + 1       'do the next qtr from the following year on wrap-around 'VBC NR
            gUnpackDateLong tgMCof(ilYearInx).iStartDate(0, ilStartMonth - 1), tgMCof(ilYearInx).iStartDate(1, ilStartMonth - 1), llStdStartDates(ilNextDate) 'VBC NR
            gUnpackDateLong tgMCof(ilYearInx).iEndDate(0, ilStartMonth - 1), tgMCof(ilYearInx).iEndDate(1, ilStartMonth - 1), llSaveLastDate 'VBC NR
            ilNextDate = ilNextDate + 1 'VBC NR
            Next ilStartMonth 'VBC NR
        End If 'VBC NR
        'get the 13th month start date.  Increment the last saved end date by 1 day
        llStdStartDates(13) = llSaveLastDate + 1 'VBC NR
    Else                           'build array of std start months 'VBC NR
        slStr = Trim$(str$(ilLoop)) & "/15/" & Trim$(str$(ilYear))      'format xx/xx/xxxx 'VBC NR
        For ilLoop = 1 To 13 Step 1 'VBC NR
            slStr = gObtainStartStd(slStr) 'VBC NR
            llStdStartDates(ilLoop) = gDateValue(slStr) 'VBC NR
            slStr = gObtainEndStd(slStr) 'VBC NR
            llDate = gDateValue(slStr) + 1                      'increment for next month 'VBC NR
            slStr = Format$(llDate, "m/d/yy") 'VBC NR
        Next ilLoop 'VBC NR
        End If 'VBC NR
        'determine what month index the actual is (versus the future dates)
        'assume everything in the future if by std; if by corp it wont use
        'history files since billing is done all by std
        If ilCorpStd = 1 Then                           'corp 'VBC NR
            ilLastBilledInx = 1                         'retrieve everything from contracts 'VBC NR
        Else 'VBC NR
            ilLastBilledInx = 1 'VBC NR
        For ilLoop = 1 To 12 Step 1 'VBC NR
            If llLastBilled > llStdStartDates(ilLoop) And llLastBilled < llStdStartDates(ilLoop + 1) Then 'VBC NR
                ilLastBilledInx = ilLoop 'VBC NR
                Exit For 'VBC NR
            End If 'VBC NR
        Next ilLoop 'VBC NR
    End If 'VBC NR
End Sub 'VBC NR
'
'               gVerifyInt - verify input.  Value must be between two arguments provided
'               <input>     slStr - user input
'                           ilLowInt - lowest value allowed
'                           ilHiInt - highest value allowed
'               <output>    Return - converted integer
'                                    -1 if invalid
Function gVerifyInt(slStr As String, ilLowInt As Integer, ilHiInt As Integer) As Integer
Dim ilInput As Integer
    gVerifyInt = 0
    ilInput = Val(slStr)
    If (ilInput < ilLowInt) Or (ilInput > ilHiInt) Then
        gVerifyInt = -1
    Else
        gVerifyInt = ilInput
    End If
End Function

'
'               gVerifyYear - Verify that the year entered is valid
'                             If not 4 digit year, add 1900 or 2000
'                             Valid year must be > than 1950 and < 2050
'                             Input - string containing input
'                             Output - Integer containing year else 0
'
Function gVerifyYear(slStr As String) As Integer
Dim ilInput As Integer
    gVerifyYear = 0
    If IsNumeric(slStr) Then
        ilInput = Val(slStr)
        If ilInput < 100 Then           'only 2 digit year input ie.  96, 95,
            If ilInput < 50 Then        'adjust for year 1900 or 2000
                ilInput = 2000 + ilInput
            Else
                ilInput = 1900 + ilInput
            End If
        End If
        If (ilInput < 1950) Or (ilInput > 2050) Then
            gVerifyYear = 0
        Else
            gVerifyYear = ilInput
        End If
    End If
End Function


'******************************************************************************
'
'               gCorpStdDates(slCalType, slEffDate, ilresult, llDates)
'               Send a date and what calendar to use for conversion (corp or std)
'               and return an array of 13 start or end dates.  The first date in the
'               array will be the start or end  date of the current quarter from the
'               Effective date .
'               <input> slCalType - C = corporate, S = std
'                       slEffDate - XX/XX/XXXX - date to calc the first months
'                                                and end dates (whether corp or std)
'               <output> llStartDates(1 to 13) - array of start dates of 13 months
'                        llEndDates(1 to 13) - array of end dates for 13 months
'                        ilresult - starting month # (1-12) of array
'               Created 7/18/96
'
'*********************************************************************************
'
Sub gCorpStdDates(slCalType As String, slEffDate As String, ilresult As Integer, llStartDates() As Long, llEndDates() As Long)
Dim ilLoop As Integer
Dim llAdjustDate As Long
Dim slAdjustDate As String
Dim slStartDate As String
Dim slEndDate As String
Dim slYear As String
Dim slMonth As String
Dim slDay As String
    If slCalType = "S" Then                         'use std calendar
        slStartDate = gObtainStartStd(slEffDate)
        slEndDate = gObtainEndStd(slEffDate)
    Else                                            'use corp calendar
        slStartDate = gObtainStartCorp(slEffDate, True)
        slEndDate = gObtainEndCorp(slEffDate, True)
    End If
    llAdjustDate = gDateValue(slStartDate) + 14     'get to the middle of the month to get the actual month name
                                                    'since the corporate cal does not start or end on the
                                                    'actual month it denotes  ie. the corporate month could start
                                                    'in the previous or end in the following month
    slAdjustDate = Format$(llAdjustDate, "m/d/yy")
    gObtainYearMonthDayStr slAdjustDate, True, slYear, slMonth, slDay

    Do While (Val(slMonth) <> 1 And Val(slMonth) <> 4 And Val(slMonth) <> 7 And Val(slMonth) <> 10)
        slMonth = str$(Val(slMonth) - 1)                    'decrement month until start of qtr
    Loop
    ilresult = Val(slMonth) 'return starting month of the month array
    'If igRptCallType = 98 Then
    '    Exit Sub
    'Else
        For ilLoop = 1 To 13 Step 1
            If slCalType = "S" Then                         'use std calendar
                slStartDate = gObtainStartStd(slMonth & "/" & slDay & "/" & slYear)
                slEndDate = gObtainEndStd(slStartDate)
            Else                                            'use corp calendar
                slStartDate = gObtainStartCorp(slMonth & "/" & slDay & "/" & slYear, True)
                slEndDate = gObtainEndCorp(slStartDate, True)
            End If
            llStartDates(ilLoop) = gDateValue(slStartDate)
            llEndDates(ilLoop) = gDateValue(slEndDate)
            slStartDate = Format$(gDateValue(slEndDate) + 1, "m/d/yy")
            gObtainYearMonthDayStr slStartDate, True, slYear, slMonth, slDay
        Next ilLoop
    'End If
End Sub
'
'
'           gIncludeExclude - Setup string to pass to Crystal report for
'               report heading to show requested Inclusions/exclusions
'           <input>  ilckc - radio control button index to test
'           <output> slinput - list of types/statuses of contracts or spots to include
'                   sloutput - list of types/statuses of contracts/spots to exclude
Sub gIncludeExcludeCkc(ilckc As Control, slInclude As String, slExclude As String, slStr As String)
    If ilckc = vbChecked Then
        If Len(slInclude) = 0 Then
            slInclude = slStr
        Else
            slInclude = slInclude & ", " & slStr
        End If
    Else
        If Len(slExclude) = 0 Then
            slExclude = slStr
        Else
            slExclude = slExclude & ", " & slStr
        End If
    End If
End Sub

'********************************************************************
'
'           gGetMonthNoFromString
'           pass month in text format (jan, feb...), and return
'           the month # (1-12)
'           <input> slStr - jan, FEB, etc.
'           <output> ilMonthNo - 1-12
'
'********************************************************************
Sub gGetMonthNoFromString(slStr As String, ilMonthNo As Integer)
    ilMonthNo = 0
    If Trim$(UCase(slStr)) = "JAN" Then
        ilMonthNo = 1
    ElseIf Trim$(UCase(slStr)) = "FEB" Then
        ilMonthNo = 2
    ElseIf Trim$(UCase(slStr)) = "MAR" Then
        ilMonthNo = 3
    ElseIf Trim$(UCase(slStr)) = "APR" Then
        ilMonthNo = 4
    ElseIf Trim$(UCase(slStr)) = "MAY" Then
        ilMonthNo = 5
    ElseIf Trim$(UCase(slStr)) = "JUN" Then
        ilMonthNo = 6
    ElseIf Trim$(UCase(slStr)) = "JUL" Then
        ilMonthNo = 7
    ElseIf Trim$(UCase(slStr)) = "AUG" Then
        ilMonthNo = 8
    ElseIf Trim$(UCase(slStr)) = "SEP" Then
        ilMonthNo = 9
    ElseIf Trim$(UCase(slStr)) = "OCT" Then
        ilMonthNo = 10
    ElseIf Trim$(UCase(slStr)) = "NOV" Then
        ilMonthNo = 11
    ElseIf Trim$(UCase(slStr)) = "DEC" Then
        ilMonthNo = 12
    End If
End Sub

'
'
'               gGetQtrHeader - given Starting year and quarter, send
'                   text to  Crystal as formula for report
'                   header:  ie.   3rd quarter 1997
'                   Return 0 if ok, -1 if NG
'
Function gGetQtrHeader(ilStartYr As Integer, ilStartQtr As Integer) As Integer
Dim slStr As String
    gGetQtrHeader = 0
    slStr = ""
    If ilStartQtr = 1 Then
        slStr = "1st Quarter "
    ElseIf ilStartQtr = 2 Then
        slStr = "2nd Quarter "
    ElseIf ilStartQtr = 3 Then
        slStr = "3rd Quarter "
    Else
        slStr = "4th Quarter "
    End If
    slStr = slStr & ilStartYr
    If Not gSetFormula("QtrHeader", "'" & slStr & "'") Then
        gGetQtrHeader = -1
        Exit Function
    End If
End Function


'**********************************************************************
'*
'*           gBuildCntTypes - Build array of contract
'*                       Types that user is allowed to see
'                        Holds, Orders, Remnants, PI, DR,
'                        PSA, Promos
'*          <Output> slCntrTypes
'*
'*          Created: 8/6/97         D.Hosaka
'*
'***********************************************************************
'*
Function gBuildCntTypes() As String
Dim slCntrType As String
    slCntrType = "C"                        'everyone gets to orders
    If tgUrf(0).sResvType <> "H" Then
        slCntrType = slCntrType & "V"
    End If
    If tgUrf(0).sRemType <> "H" Then
        slCntrType = slCntrType & "T"
    End If
    If tgUrf(0).sDRType <> "H" Then
        slCntrType = slCntrType & "R"
    End If
    If tgUrf(0).sPIType <> "H" Then
        slCntrType = slCntrType & "Q"
    End If
    'If tgUrf(0).sPSAType <> "H" Then
    '    slCntrType = slCntrType & "S"
    'End If
    'If tgUrf(0).sPromoType <> "H" Then
     '   slCntrType = slCntrType & "M"
    'End If
    If slCntrType = "CVTRQSM" Then
        slCntrType = ""                 'all types
    End If
    gBuildCntTypes = slCntrType
End Function

'
'
'                   gFakeChf: Billed & booked subroutine to
'                   create a header with fields from the
'                   Receivables record so that the transaction
'                   is included in report.  Header is not found
'                   when transactions are added without references
'                   to the contract (no contract entered) or
'                   the contract # is zero
'
'                   dh copied and modified from mFakeChf 4/8/99
'
'           3-5-07  if a contract hdr doesnt exist, set the slsp comm to zero
Sub gFakeChf(tlRvf As RVF, tlChf As CHF)
Dim ilLoop As Integer
    If tlRvf.lCntrNo <> tlChf.lCntrNo Or tlRvf.lCntrNo = 0 Then     '8/26/99 multiple receivables without cntr #, force the fake header
        tlChf.lCntrNo = tlRvf.lCntrNo
        tlChf.lCode = 0
        tlChf.sSchStatus = "F"              'assume fully scheduled
        tlChf.sStatus = "O"                 'assume order
        tlChf.sType = "C"                   'standard contract
        tlChf.lVefCode = 0
        If tlRvf.sCashTrade = "T" Then
            tlChf.iPctTrade = 100
        Else                                'could be merchndising, promotions or cash
            tlChf.iPctTrade = 0
        End If
        For ilLoop = 0 To 9 Step 1
            If ilLoop = 0 Then
                tlChf.iSlfCode(ilLoop) = tlRvf.iSlfCode
                tlChf.lComm(ilLoop) = 1000000
                tlChf.iSlspCommPct(ilLoop) = 0      '3-5-07 10000
                tlChf.iMnfSubCmpy(ilLoop) = 0
            Else
                tlChf.iSlfCode(ilLoop) = 0
                tlChf.lComm(ilLoop) = 0
                tlChf.iSlspCommPct(ilLoop) = 0
                tlChf.iMnfSubCmpy(ilLoop) = 0
            End If
        Next ilLoop
        tlChf.iAgfCode = tlRvf.iAgfCode
        tlChf.iAdfCode = tlRvf.iAdfCode
        tlChf.lVefCode = tlRvf.iBillVefCode     'reqd for contrct user is allowed to see
        '1-22-12
        tlChf.iMnfComp(0) = 0
        tlChf.iMnfComp(1) = 0
        tlChf.iMnfBus = 0                          '1-9-18
    End If
End Sub
'
'
'           gGetRollOverDate - obtain the closest Rollover DAte
'               from Pjf based on the user entered date .
'               Each slsp is retrieved find its earliest rollover
'               date.  Whatever rollover date that is closest
'               to the date user requested will be returned.
'
'           <input> Form indicating the Form used as input
'                   lbcSelection array index containing the slsp
'                   slEnterDate - User date entered
'           <output> llClosestDate - closest rollover date found
'
'           1-5-05 Invalid procedure call or argument generated when
'               converting a date from PJF after it hit an EOF.  DAte
'               stored in the PJF record couldn't be converted.
Sub gGetRollOverDate(Form As Form, ilIndex As Integer, slEnterDate As String, llClosestDate As Long)
Dim ilLoop As Integer
Dim slNameCode As String
Dim slCode As String
Dim llEarliestRO As Long            'earliest date to use
Dim llLatestRO As Long              'latest date to use (1 week span)
Dim llPjfRO As Long                 'rollover date from projection record converted to long
ReDim ilEarliest(0 To 1) As Integer    'btrieve format user date entered   (for key reads)
Dim ilRet As Integer
Dim hlPjf As Integer
Dim tlPjf As PJF
Dim tlSrchKey As PJFKEY0
    hlPjf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlPjf, "", sgDBPath & "Pjf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlPjf)
        btrDestroy hlPjf
        Exit Sub
    End If
    gPackDate slEnterDate, ilEarliest(0), ilEarliest(1)
    llEarliestRO = gDateValue(slEnterDate)      'gather for 1 week span only
    llLatestRO = llEarliestRO + 6
    For ilLoop = 0 To Form!lbcSelection(ilIndex).ListCount - 1 Step 1
        'If Rptselpj!lbcSelection(ilIndex).Selected(ilLoop) Then
            slNameCode = tgSalesperson(ilLoop).sKey    'Traffic!lbcSalesperson.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
            'get the closest rollover date for this slsp
            tlSrchKey.iSlfCode = Val(slCode)
            tlSrchKey.iRolloverDate(0) = ilEarliest(0)
            tlSrchKey.iRolloverDate(1) = ilEarliest(1)
            'If ilRet = BTRV_ERR_NONE Then       '1-5-05 if end of file, convert date to long generate error due to an invalid date stored in record
            ilRet = btrGetGreaterOrEqual(hlPjf, tlPjf, Len(tlPjf), tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE) 'get matching projection recd
            '3-27-07 move test for ilRet to after the PJF has been read
            If ilRet = BTRV_ERR_NONE Then       '1-5-05 if end of file, convert date to long generate error due to an invalid date stored in record

                gUnpackDateLong tlPjf.iRolloverDate(0), tlPjf.iRolloverDate(1), llPjfRO
                Do While (ilRet = BTRV_ERR_NONE) And (llPjfRO >= llEarliestRO And llPjfRO <= llLatestRO) And Val(slCode) = tlPjf.iSlfCode
                    If llClosestDate = 0 Then
                        gUnpackDateLong tlPjf.iRolloverDate(0), tlPjf.iRolloverDate(1), llClosestDate
                    End If
                    If (llPjfRO < llClosestDate) Then   'save new closest date to the user date entered
                        gUnpackDateLong tlPjf.iRolloverDate(0), tlPjf.iRolloverDate(1), llClosestDate
                        Exit For
                    End If

                    ilRet = btrGetNext(hlPjf, tlPjf, Len(tlPjf), BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        gUnpackDateLong tlPjf.iRolloverDate(0), tlPjf.iRolloverDate(1), llPjfRO
                    Else
                        llPjfRO = 0
                    End If
                Loop
            End If
        'End If
    Next ilLoop
    If llClosestDate = 0 Then               'if no date found, just use the date entered
        llClosestDate = llEarliestRO
    End If
    ilRet = btrClose(hlPjf)
End Sub
'
'
'               gGetStartEndQtr - get the start date and end date of the quarter
'               for STandard or Corporate year.
'               <input> ilCalType - 1 = Corp , 2 = STd
'                       ilYear - Year to obtain the start date
'                       ilQtrInx - quarter index # to obtain start & end dates
'               <output> slStartOfQtr- Start date of the corp or std quarter
'                        slEndOfQtr - end date of corp or std quarter
'               Assume that the Corporate calendar is in memory in (tmMCof)
'               if using Corporate calendar.
'
'               Created:  DH    10/6/97
'
Sub gGetStartEndQtr(ilCalType As Integer, ilYear As Integer, ilQtrInx As Integer, slStartOfQtr As String, slEndOfQtr As String)
Dim ilIndex As Integer
Dim llDate As Long
Dim ilTemp As Integer
    ilTemp = (ilQtrInx - 1) * 3 + 1
    If ilCalType = 1 Then                       'corporate
        ilIndex = gGetCorpCalIndex(ilYear)
        'gUnpackDateLong tgMCof(ilIndex).iStartDate(0, ilTemp), tgMCof(ilIndex).iStartDate(1, ilTemp), llDate
        gUnpackDateLong tgMCof(ilIndex).iStartDate(0, ilTemp - 1), tgMCof(ilIndex).iStartDate(1, ilTemp - 1), llDate
        slStartOfQtr = Format(llDate, "m/d/yy")
        'gUnpackDateLong tgMCof(ilIndex).iEndDate(0, ilTemp + 2), tgMCof(ilIndex).iEndDate(1, ilTemp + 2), llDate
        gUnpackDateLong tgMCof(ilIndex).iEndDate(0, ilTemp + 1), tgMCof(ilIndex).iEndDate(1, ilTemp + 1), llDate
        'slEndOfQtr = Format(llDate - 1, "m/d/yy")     'returned start date of next qtr, butneed end date of prev qtr
        slEndOfQtr = Format(llDate, "m/d/yy")      'returned start date of next qtr, butneed end date of prev qtr
    Else
        slStartOfQtr = Trim$(str$(ilTemp)) & "/15/" & Trim$(str$(ilYear))      '12/15/year entered
        slStartOfQtr = gObtainStartStd(slStartOfQtr)      'get the stnd years start date

        slEndOfQtr = Trim$(str$(ilTemp + 2)) & "/15/" & Trim$(str$(ilYear))    '12/15/year entered
        slEndOfQtr = gObtainEndStd(slEndOfQtr)
    End If
End Sub
'
'
'               gGetStartEndYear - get the start date and end date of the year
'               for STandard or Corporate year.
'               <input> ilCalType - 1 = Corp , 2 = STd
'                       ilYear - Year to obtain the start date
'               <output> slStartOfYear - Start date of the corp or std year
'                        slEndOfYear - end date of corp or std year
'               Assume that the Corporate calendar is in memory in (tmMCof)
'               if using Corporate calendar.
'
'               Created:  DH    10/6/97
'
Sub gGetStartEndYear(ilCalType As Integer, ilYear As Integer, slStartOfYear As String, slEndofYear As String)
Dim ilIndex As Integer
Dim llDate As Long
    If ilCalType = 1 Then                       'corporate
        ilIndex = gGetCorpCalIndex(ilYear)
        'gUnpackDateLong tgMCof(ilIndex).iStartDate(0, 1), tgMCof(ilIndex).iStartDate(1, 1), llDate
        gUnpackDateLong tgMCof(ilIndex).iStartDate(0, 0), tgMCof(ilIndex).iStartDate(1, 0), llDate
        slStartOfYear = Format(llDate, "m/d/yy")
        'gUnpackDateLong tgMCof(ilIndex).iEndDate(0, 12), tgMCof(ilIndex).iEndDate(1, 12), llDate
        gUnpackDateLong tgMCof(ilIndex).iEndDate(0, 11), tgMCof(ilIndex).iEndDate(1, 11), llDate
        slEndofYear = Format(llDate, "m/d/yy")
    Else
        slStartOfYear = "1/15/" & Trim$(str$(ilYear))      '12/15/year entered
        slStartOfYear = gObtainStartStd(slStartOfYear)      'get the stnd years start date
        slEndofYear = "12/15/" & Trim$(str$(ilYear))               '12/15/year entered
        slEndofYear = gObtainEndStd(slEndofYear)            'get the stnd years end date
    End If
End Sub
'
'
'           gIncludeExclude - Setup string to pass to Crystal report for
'               report heading to show requested Inclusions/exclusions
'           <input>  ilckc - radio control button index to test
'           <output> slinput - list of types/statuses of contracts or spots to include
'                   sloutput - list of types/statuses of contracts/spots to exclude
Sub gIncludeExcludeRbc(ilckc As Control, slInclude As String, slExclude As String, slStr As String)
    If ilckc Then
        If Len(slInclude) = 0 Then
            slInclude = slStr
        Else
            slInclude = slInclude & ", " & slStr
        End If
    Else
        If Len(slExclude) = 0 Then
            slExclude = slStr
        Else
            slExclude = slExclude & ", " & slStr
        End If
    End If
End Sub
'
'           gSpanDates - determine if given start/end dates span
'               another set of start/end dates
'               <input>  llStartSpan1 & llEndSpan1 - 1st set of start&
'                         end dates to test spanning of period
'                        llStartSpan2 & llEndSpan2 - 2nd set of start &
'                         end dates to test spanning of period
'               return - true if 1st set of dates span 2nd set of dates
Function gSpanDates(llStartSpan1 As Long, llEndSpan1 As Long, llStartSpan2 As Long, llEndSpan2 As Long) As Integer
    gSpanDates = False
    'if 1st set of spans prior to span2 earliest or 1st set of spans later than span2 end--get out
    If llEndSpan1 < llStartSpan2 Or llStartSpan1 > llEndSpan2 Then
        Exit Function
    Else
        gSpanDates = True
    End If
End Function

Public Function gProducerProviderPop(frm As Form, slPPType As String, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gProducerProviderPop (MainForm, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       slPPType (I) - retrieve producers (P) or providers (C) from ARF
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'Adf date/time stamp
    Dim hlArf As Integer        'Adf handle
    Dim ilRecLen As Integer     'Record length
    Dim tlArf As ARF
    Dim slName As String
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim tlCharTypeBuff As POPCHARTYPE
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer


    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Arf.Btr") & Trim$(slPPType)

    On Error GoTo gProducerProviderPopErr2
    ilRet = 0

    If ilRet <> 0 Then
        slSortCodeTag = ""
    End If
    On Error GoTo 0

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gProducerProviderPop = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gProducerProviderPop = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlArf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlArf, "", sgDBPath & "Arf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gProducerProviderPopErr
        gBtrvErrorMsg ilRet, "gProducerProviderPop (btrOpen): Arf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlArf) 'btrRecordLength(hlArf)  'Get and save record length
        ilSortCode = 0
        ReDim tlSortCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlArf)
        llNoRec = gExtNoRec(ilExtLen)
        btrExtClear hlArf
        Call btrExtSetBounds(hlArf, llNoRec, -1, "UC", "Arf", "") 'Set extract limits (all records)

        ilRet = btrGetFirst(hlArf, tlArf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)    'Get first record as starting point
        If (ilRet <> BTRV_ERR_NONE) Then
            ilRet = btrClose(hlArf)
            On Error GoTo gProducerProviderPopErr
            gBtrvErrorMsg ilRet, "gProducerProviderPop (btrGetFirst):" & "Arf.Btr", frm
            On Error GoTo 0
            btrDestroy hlArf
            Exit Function
        End If
        tlCharTypeBuff.sType = slPPType     'Producer (P) or Provider (C) type
        ilOffSet = gFieldOffset("Arf", "ArfType")
        ilRet = btrExtAddLogicConst(hlArf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        'Call btrExtSetBounds(hlArf, llNoRec, -1, "UC", "Arf", "") 'Set extract limits (all records)
        ilOffSet = 0
        ilRet = btrExtAddField(hlArf, ilOffSet, ilRecLen)  'Extract iCode field
        On Error GoTo gProducerProviderPopErr
        gBtrvErrorMsg ilRet, "gProducerProviderPop (btrExtAddField):" & "Arf.Btr", frm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlArf, tlArf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gProducerProviderPopErr
            gBtrvErrorMsg ilRet, "gProducerProviderPop (btrExtGetNextExt):" & "Arf.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlArf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlArf, tlArf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slName = tlArf.sName
                slName = slName & "\" & Trim$(str$(tlArf.iCode))
                'If Not gOkAddStrToListBox(slName, llLen, True) Then
                '    Exit Do
                'End If
                'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                tlSortCode(ilSortCode).sKey = slName
                If ilSortCode >= UBound(tlSortCode) Then
                    ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 100) As SORTCODE
                End If
                ilSortCode = ilSortCode + 1
                ilRet = btrExtGetNext(hlArf, tlArf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlArf, tlArf, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(0 To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - 1 > 0 Then
                ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
            End If
        End If



        ilRet = btrClose(hlArf)
        On Error GoTo gProducerProviderPopErr
        gBtrvErrorMsg ilRet, "gProducerProviderPop (btrReset):" & "Arf.Btr", frm
        On Error GoTo 0
        btrDestroy hlArf
    End If
    llLen = 0
    For ilLoop = 0 To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gProducerProviderPop = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next ilLoop
    Exit Function
gProducerProviderPopErr:
    ilRet = btrClose(hlArf)
    btrDestroy hlArf
    gDbg_HandleError "RptSubs: gProducerProviderPop"

gProducerProviderPopErr2:
    ilRet = 1
    Resume Next
End Function
'
'
'       Test the flag in spot (PriceType) to determine if spot should
'       be shown on invoice or not.  If the spot doesnt have a "- or +" stored
'       in pricetype field, then it has been overridden in spot screen to
'       answer on-demand; test the spot instead of advertiser to determine
'       how its to be shown
'       <input>  slPriceType = price type flag from spot (SDF)
'                ilAdfCode - advertiser code
'       return - Y = show on invoice (+), N = dont show on invoice (-)
'       1-19-04

'
'           gFeedNamesPop - populate a list box with the list of feed names from FNF
'
Public Function gFeedNamesPop(frm As Form, ilFeedTypes As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gFeedNamesPop (MainForm, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilFeedTypes (I) - &H1= Insertion Order, &H2 = PreConverted, &H4 = Log Needs conversion
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String       'Adf date/time stamp
    Dim hlFnf As Integer        'Adf handle
    Dim ilRecLen As Integer     'Record length
    Dim tlFnf As FNF
    Dim slName As String
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilOk As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Fnf.btr") & Trim$(str$(ilFeedTypes))

    On Error GoTo gFeedNamesPopErr2
    ilRet = 0

    If ilRet <> 0 Then
        slSortCodeTag = ""
    End If
    On Error GoTo 0

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gFeedNamesPop = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gFeedNamesPop = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlFnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlFnf, "", sgDBPath & "Fnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gFeedNamesPopErr
        gBtrvErrorMsg ilRet, "gFeedNamesPop (btrOpen): Fnf.btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlFnf) 'btrRecordLength(hlFnf)  'Get and save record length
        ilSortCode = 0
        ReDim tlSortCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlFnf)
        llNoRec = gExtNoRec(ilExtLen)
        btrExtClear hlFnf
        Call btrExtSetBounds(hlFnf, llNoRec, -1, "UC", "Fnf", "") 'Set extract limits (all records)

        ilRet = btrGetFirst(hlFnf, tlFnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)    'Get first record as starting point
        If (ilRet <> BTRV_ERR_NONE) Then
            ilRet = btrClose(hlFnf)
            On Error GoTo gFeedNamesPopErr
            gBtrvErrorMsg ilRet, "gFeedNamesPop (btrGetFirst):" & "Fnf.btr", frm
            On Error GoTo 0
            btrDestroy hlFnf
            Exit Function
        End If

        ilOffSet = 0
        ilRet = btrExtAddField(hlFnf, ilOffSet, ilRecLen)  'Extract iCode field
        On Error GoTo gFeedNamesPopErr
        gBtrvErrorMsg ilRet, "gFeedNamesPop (btrExtAddField):" & "Fnf.btr", frm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlFnf, tlFnf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gFeedNamesPopErr
            gBtrvErrorMsg ilRet, "gFeedNamesPop (btrExtGetNextExt):" & "Fnf.btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlFnf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlFnf, tlFnf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE

                'test for Feed type (P = preconverted, L = needs conversion, I = insertion order)
                ilOk = mTestFeedType(ilFeedTypes, tlFnf)
                If ilOk Then
                    slName = tlFnf.sName
                    slName = slName & "\" & Trim$(str$(tlFnf.iCode))

                    tlSortCode(ilSortCode).sKey = slName
                    If ilSortCode >= UBound(tlSortCode) Then
                        ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 100) As SORTCODE
                    End If
                    ilSortCode = ilSortCode + 1
                End If
                ilRet = btrExtGetNext(hlFnf, tlFnf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlFnf, tlFnf, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(0 To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - 1 > 0 Then
                ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
            End If
        End If



        ilRet = btrClose(hlFnf)
        On Error GoTo gFeedNamesPopErr
        gBtrvErrorMsg ilRet, "gFeedNamesPop (btrReset):" & "Fnf.btr", frm
        On Error GoTo 0
        btrDestroy hlFnf
    End If
    llLen = 0
    For ilLoop = 0 To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gFeedNamesPop = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next ilLoop
    Exit Function
gFeedNamesPopErr:
    ilRet = btrClose(hlFnf)
    btrDestroy hlFnf
    gDbg_HandleError "RptSubs: gFeedNamesPop"

gFeedNamesPopErr2:
    ilRet = 1
    Resume Next
End Function
'
'
'           mTestFeedType - test if the FNF record is a valid feed name to populate
'
'           <input> ilFeedType : &H1= Insertion Order, &H2 = PreConverted, &H4 = Log Needs conversion
'                   tlFnf - feed name buffer
'           <return> true - valid feed name to populate, else false
'
Private Function mTestFeedType(ilFeedType As Integer, tlFnf As FNF) As Integer
Dim ilOk As Integer
    ilOk = False
    Select Case Trim$(tlFnf.sPledgeTime)
        Case "P"            'preconverted
            If (ilFeedType And FEED_BYCONVERTED) Then
                ilOk = True
            End If
        Case "L"            'log needs conversion
            If (ilFeedType And FEED_BYNEEDSCONVERT) Then
                ilOk = True
            End If
        Case "I"            'Insertion Order
            If (ilFeedType And FEED_BYINSERT) Then
               ilOk = True
            End If
    End Select
    mTestFeedType = ilOk

End Function
'               gObtainCodes - get all codes to process or exclude
'               When selecting advt, agy or vehicles (anthing in list box)--make testing
'               of selection more efficient.  If more than half of
'               the entries are selected, create an array with entries
'               to exclude.  If less than half of entries are selected,
'               create an array with entries to include.
'               <input> ilListIndex - list box to test
'                       ilIndex - index to list box when multiple list boxes exists
'                       lbcListbox - array containing sort codes
'               <output> ilIncludeCodes - true if test to include the codes in array
'                                          false if test to exclude the codes in array
'                        ilUseCodes - array of advt, agy or vehicles codes to include/exclude
'
Public Sub gObtainCodesForMultipleLists(ilIndex As Integer, lbcListBox() As SORTCODE, ilIncludeCodes As Integer, ilUseCodes() As Integer, Form As Form)
Dim ilHowManyDefined As Integer
Dim ilHowMany As Integer
Dim slNameCode As String
Dim ilLoop As Integer
Dim slCode As String
Dim ilRet As Integer
'ReDim ilUseCodes(1 To 1) As Integer
ReDim ilUseCodes(0 To 0) As Integer
    ilHowManyDefined = Form!lbcSelection(ilIndex).ListCount
    ilHowMany = Form!lbcSelection(ilIndex).SelCount
    If ilHowMany > (ilHowManyDefined / 2) + 1 Then  'more than half selected
        ilIncludeCodes = False
    Else
        ilIncludeCodes = True
    End If

    For ilLoop = 0 To Form!lbcSelection(ilIndex).ListCount - 1 Step 1
        slNameCode = lbcListBox(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If Form!lbcSelection(ilIndex).Selected(ilLoop) And ilIncludeCodes Then               'selected ?
            ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
            ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
        Else        'exclude these
            If (Not Form!lbcSelection(ilIndex).Selected(ilLoop)) And (Not ilIncludeCodes) Then
                ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
                ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
            End If
        End If
    Next ilLoop
    Exit Sub
End Sub
'
'               gVerifyLong - verify input of long variable.  Value must be between two arguments provided
'               <input>     slStr - user input
'                           llLowValue - lowest value allowed
'                           llHiValue - highest value allowed
'               <output>    Return - converted integer
'                                    -1 if invalid
Public Function gVerifyLong(slStr As String, llLowValue As Long, llHiValue As Long) As Long
Dim llInput As Long
    gVerifyLong = 0
    llInput = Val(slStr)
    If (llInput < llLowValue) Or (llInput > llHiValue) Then
        gVerifyLong = -1
    Else
        gVerifyLong = llInput
    End If
    Exit Function
End Function
'
'               Test site to determine if using Acquisition or Barters
'           <return> true if using feature
'
Public Function gUsingBarters() As Integer
    gUsingBarters = False
    '6/7/15: replaced acquisition from site override with Barter in system options
    If (tgSpf.sUsingNTR = "Y") And ((Asc(tgSpf.sOverrideOptions) And SPNTRACQUISITION) = SPNTRACQUISITION) Then
        gUsingBarters = True
        Exit Function
    End If
    If (Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER Then
        gUsingBarters = True
        Exit Function
    End If
End Function


'       Set check box as unchecked
'       <input> form & control name - form and control to uncheck
'               ilAl- flag to set/clear the list
'       <output> form & control name - set as unchecked
'               ilSetAll - flag to set/clear the list
Public Sub gUncheckAll(ilCkcAll As Control, ilSetAll As Integer)
        ilSetAll = False
        ilCkcAll = vbUnchecked
        ilSetAll = True
        Exit Sub
End Sub

'
'           Convert contract header start/end dates to long and strings
'           <input>  Contract start date
'                    contract EndDate
'           <output> llDate - cnt start date as long
'                     llDate2 - cnt end date as long
'                    slSTartDate - cnt start date as string
'                    slEndDate - cnt end date as string
Sub gChfDatesToLong(ChfStartDate() As Integer, ChfEndDate() As Integer, llDate As Long, llDate2 As Long, slStartDate As String, slEndDate As String)

    gUnpackDate ChfStartDate(0), ChfStartDate(1), slStartDate
    If slStartDate = "" Then
        llDate = 0
    Else
        llDate = gDateValue(slStartDate)
    End If
    gUnpackDate ChfEndDate(0), ChfEndDate(1), slEndDate
    If slEndDate = "" Then
        llDate2 = 0
    Else
        llDate2 = gDateValue(slEndDate)
    End If
End Sub

'
'           gFilterSpotRateType - determine whether spot should be included/excluded
'           based on its schedule line rate type entered; or if its a Bonus spot that
'           has been +/- Fill created
'           <input> tlCntTypes() - array of spot rate types to be included/excluded
'                   slprice - Spot rate from spot
'           return - true if OK, else false and ignore spot
Public Function gFilterSpotRateType(tlCntTypes As CNTTYPES, slPrice As String) As Integer
Dim ilSpotOK As Integer
    ilSpotOK = True
    If (InStr(slPrice, ".") <> 0) Then    'found spot cost
        tlCntTypes.lRate = gStrDecToLong(slPrice, 2)    'get the actual spot value
        'is it a .00?
        If gCompNumberStr(slPrice, "0.00") = 0 Then     'its a .00 spot
            If Not tlCntTypes.iZero Then
                ilSpotOK = False
            End If
        Else
            If Not tlCntTypes.iCharge Then           'exclude charged spots
                ilSpotOK = False
            End If
        End If
    ElseIf Trim$(slPrice) = "ADU" Then
        If Not tlCntTypes.iADU Then
            ilSpotOK = False
        End If
    ElseIf Trim$(slPrice) = "Bonus" Then
        If Not tlCntTypes.iBonus Then
            ilSpotOK = False
        End If
    ElseIf Trim$(slPrice) = "+ Fill" Then
        If Not tlCntTypes.iXtra Then
            ilSpotOK = False
        End If
    ElseIf Trim$(slPrice) = "- Fill" Then
        If Not tlCntTypes.iFill Then
            ilSpotOK = False
        End If
    ElseIf Trim$(slPrice) = "N/C" Then
        If Not tlCntTypes.iNC Then
            ilSpotOK = False
        End If
    ElseIf Trim$(slPrice) = "Recapturable" Then
        If Not tlCntTypes.iRecapturable Then
            ilSpotOK = False
        End If
    ElseIf Trim$(slPrice) = "Spinoff" Then
        If Not tlCntTypes.iSpinoff Then
            ilSpotOK = False
        End If
    ElseIf Trim$(slPrice) = "MG" Then
        If Not tlCntTypes.iMG Then
            ilSpotOK = False
        End If
    End If
    gFilterSpotRateType = ilSpotOK      'return if OK to include/exclude
End Function
'
'                   gBuildFlightInfo - Loop through the flights of the schedule line
'                           and build the projections dollars into llproject array,
'                           and build projection # of spots into llprojectspots array
'                           Build projection of R/C $ from cffpropprice
'                           Build projection $ of Acquisition $
'                   <input> ilclf = sched line index into tlClfInp
'                           llStdStartDates() - array of dates to build $ from flights
'                           ilFirstProjInx - index of 1st month/week to start projecting
'                           ilMaxInx - max # of buckets to loop thru
'                           ilWkOrMonth - 1 = Month, 2 = Week
'                   <output> llProject() = array of $ buckets corresponding to array of dates
'                           llProjectSpots() array of spot count buckets corresponding to array of dates
'                           llProjectRC() array of $ buckets from rate card price (proposal price stored in flight)
'                           llProjectAcq() array of $ buckets from acquisition $ stored in line
'                   General routine to build flight $/cpot count into week, month, qtr buckets
'            Created : 7-12-05
'
Public Sub gNtrByContract(llCurrentChfCode As Long, llStartDate As Long, llEndDate As Long, tlNTRInfo() As NTRPacing, tlMnf() As MNF, hmSbf As Integer, blIncludeNTR As Boolean, blIncludeHardCost As Boolean, ReportName As Form)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llLineAmount                                                                          *
'******************************************************************************************
    
    '********************* gNtrByContract*****************
    ' Dan M. 6-23-08  NTR/HardCost to pacing reports
    '   llCurrentChfCode (I)
    '   llStartDate(I)
    '   llEndDate(I)
    '   tlNTRInfo() (O)
    '   tlMnf() (I) array of mnf records
    '   hmSBF (I)
    '   blIncludeNTR (I)
    '   blIncludeHardCost(I)
    '   ReportName (I) for error messages in gObtainSbf
    Dim ilRet As Integer
    Dim tlSbf() As SBF
    Dim tlSBFTypes As SBFTypes
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilLowEndArray As Integer
    Dim ilHighEndArray As Integer
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSbfCounter As Integer
    Dim llSbfCounter As Long
    Dim ilMnfItem As Integer
    
    Dim ilUpperBoundInfo As Integer
    Dim ilLowerBoundInfo As Integer
    Const HardCost = -1
    
    tlSBFTypes.iNTR = True
    tlSBFTypes.iImport = False
    tlSBFTypes.iInstallment = False
    slStartDate = Format(llStartDate, "m/d/yy")
    slEndDate = Format(llEndDate, "m/d/yy")
    ilRet = gObtainSBF(ReportName, hmSbf, llCurrentChfCode, slStartDate, slEndDate, tlSBFTypes, tlSbf(), 0)
    ilLowEndArray = LBound(tlSbf)
    ilHighEndArray = UBound(tlSbf)
    ReDim tlNTRInfo(ilLowEndArray To ilHighEndArray) As NTRPacing
    ReDim tlNTRInfo(0 To 0) As NTRPacing
    ilUpperBoundInfo = UBound(tlNTRInfo)
    ilLowerBoundInfo = LBound(tlNTRInfo)
    'both ntr and hard cost
    If blIncludeHardCost And blIncludeNTR Then
        ReDim tlNTRInfo(ilLowEndArray To ilHighEndArray) As NTRPacing
        For llSbfCounter = ilLowEndArray To ilHighEndArray - 1 Step 1
            gUnpackDateLong tlSbf(llSbfCounter).iDate(0), tlSbf(llSbfCounter).iDate(1), tlNTRInfo(llSbfCounter).lSbfDate
            tlNTRInfo(llSbfCounter).lSBFTotal = tlSbf(llSbfCounter).lGross * tlSbf(llSbfCounter).iNoItems
            tlNTRInfo(llSbfCounter).iVefCode = tlSbf(llSbfCounter).iBillVefCode         'used bill vef code
        Next llSbfCounter
    Else
        For llSbfCounter = ilLowEndArray To ilHighEndArray - 1 Step 1
            ilMnfItem = tlSbf(llSbfCounter).iMnfItem
            ilRet = gIsItHardCost(ilMnfItem, tlMnf())
            'if chose Ntr only, and not hard cost, or chose hard cost only, and is hard cost
            If ((blIncludeNTR) And (ilRet <> HardCost)) Or ((blIncludeHardCost) And (ilRet = HardCost)) Then
                gUnpackDateLong tlSbf(llSbfCounter).iDate(0), tlSbf(llSbfCounter).iDate(1), tlNTRInfo(ilUpperBoundInfo).lSbfDate
                tlNTRInfo(ilUpperBoundInfo).lSBFTotal = tlSbf(llSbfCounter).lGross * tlSbf(llSbfCounter).iNoItems
                tlNTRInfo(ilUpperBoundInfo).iVefCode = tlSbf(llSbfCounter).iBillVefCode         'used bill vef code
                ilUpperBoundInfo = ilUpperBoundInfo + 1
                ReDim Preserve tlNTRInfo(ilLowerBoundInfo To ilUpperBoundInfo) As NTRPacing
           End If
        Next llSbfCounter
    End If
End Sub
'
'           Figure out what the starting quarter is for column headings
'           Billed & Booked, B &B Recap, NTR B & B, B & B Multimedia
'           These reports enter a starting month for input, but need to
'           determine what quarter it belongs in
'           gGetQtrForColumns
'               <input> ilCorp :  true if corp calendar; otherwise false for all
'                                 other calendar types
'                       ilMonthInput - month # entered (1-12)
'               <return> Qtr # (1-4)
'
'Public Function gGetQtrForColumns(ilCorp As Integer, ilWhatMonth As Integer, ilCorpStartMonth As Integer) As Integer
Public Function gGetQtrForColumns(ilWhatMonth As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        ilWhichQtr                    ilTemp                    *
'*                                                                                        *
'******************************************************************************************

Dim ilTempMonth As Integer
Dim ilStartQtr As Integer
Dim ilMonthInput As Integer

    gGetQtrForColumns = 0
    ilMonthInput = ilWhatMonth
    'calculate qtr header for columns (Q1, Q2, Q3, Q4)
    'this calc would be for Std or calendar month reporting; corp needs to check the corp calendar to determine which
    'qtr the month belongs in
    ilTempMonth = ilMonthInput
    ilStartQtr = ilTempMonth Mod 3
    Do While ilStartQtr <> 0     'get the starting quarter
        ilMonthInput = ilMonthInput + 1
        ilStartQtr = ilMonthInput Mod 3
    Loop
    ilStartQtr = ilMonthInput \ 3

'    If ilCorp Then              'determine corp quarter from the month entered
'        'build array of start months for each qtr of corporate calendars
'        For ilLoop = 1 To 4
'            If ilLoop = 1 Then
'                ilWhichQtr(1) = ilCorpStartMonth
'                ilTemp = ilWhichQtr(1)
'            Else
'                ilTemp = ilTemp + 3
'                If ilTemp > 12 Then
'                    ilWhichQtr(ilLoop) = ilTemp - 12
'                    ilTemp = ilWhichQtr(ilLoop)
'                Else
'                    ilWhichQtr(ilLoop) = ilTemp
'                End If
'            End If
'        Next ilLoop
'        ilTemp = ilMonthInput
'        Do While ilTemp <> 1 And ilTemp <> 4 And ilTemp <> 7 And ilTemp <> 10
'            ilTemp = ilTemp - 1
'        Loop
'
'        For ilLoop = 1 To 4
'            If ilTemp = ilWhichQtr(ilLoop) Then
'                Exit For
'            End If
'        Next ilLoop
'        ilStartQtr = ilLoop
'    End If
    gGetQtrForColumns = ilStartQtr
End Function
'
'                   gGetCorpMonthString - return a string of 3-char months
'                   starting from the month of the corp calendar
'                   i.e:  if corp year starts in Oct, return "OctNov....Sep"
'                   <input> ilYear - corporate year to retrieve the month string
'
Public Function gGetCorpMonthString(ilYear As Integer) As String
Dim ilRet As Integer
Dim ilMonth As Integer
Dim slTempCalString As String * 36
Dim ilLoop As Integer
Dim slCorpMonthString As String * 36

        slTempCalString = "JanFebMarAprMayJunJulAugSepOctNovDec"
        slCorpMonthString = ""
        ilRet = gGetCorpCalIndex(ilYear)
        'If ilRet <= 0 Then                  'year not defined
        If ilRet < 0 Then                  'year not defined
            gGetCorpMonthString = slTempCalString
            Exit Function
        End If
        ilMonth = tgMCof(ilRet).iStartMnthNo
        For ilLoop = 1 To 12
            slCorpMonthString = Trim$(slCorpMonthString) & Trim$(Mid$(slTempCalString, (ilMonth - 1) * 3 + 1, 3))
            ilMonth = ilMonth + 1
            If ilMonth > 12 Then
                ilMonth = 1
            End If
        Next ilLoop
        gGetCorpMonthString = slCorpMonthString
        Exit Function
End Function

Public Function gYearForCorpStartMonth() 'VBC NR

End Function 'VBC NR
'
'               gGetYearofCorpMonth - look at corporate calendar and determine
'               what the year is for a given corporate month & year
'               i.e. corp year 2007, month 1 for ABC corporate calendar
'               starts in Oct of 2006.  Return 2006 for given month and year
'               of dec/2007.
'               <input> ilMonth - month # within corporate calendar (1 indicates
'                       the month relative to the defined start month of the fiscal year)
'                       ilyear - corporate year
Public Function gGetYearofCorpMonth(ilMonth As Integer, ilYear As Integer) As Integer
Dim ilRet As Integer
Dim llDate As Long
Dim slDate As String
Dim slMonth As String
Dim slDay As String
Dim slYear As String

        ilRet = gGetCorpCalIndex(ilYear)
        'If ilRet <= 0 Then                  'year not defined
        If ilRet < 0 Then                  'year not defined
            gGetYearofCorpMonth = ilYear
            Exit Function
        End If
        gUnpackDateLong tgMCof(ilRet).iStartDate(0, ilMonth - 1), tgMCof(ilRet).iStartDate(1, ilMonth - 1), llDate
        llDate = llDate + 15        'get to the middle of the month
        slDate = Format$(llDate, "m/d/yy")
        gObtainYearMonthDayStr slDate, False, slYear, slMonth, slDay
        gGetYearofCorpMonth = Val(slYear)
        Exit Function
End Function
'
'               gGetCorpMonthNoFromMonthName - return the starting month relative to start of corp year
'                       from 3 month text input
'               ie. Oct input, return the month # relative to start of corporate year
'               <input> slCorpMonths = string of months starting with start month of corp calenar
'                       (i.e. octNovDec.....Sep)
'                       slInputMonth (Jan, feb....dec)
'               return - Month # (0 if invalid)
Public Function gGetCorpMonthNoFromMonthName(slCorpMonthString As String, slInputMonth As String) As Integer
Dim ilLoop As Integer
Dim ilMonth As Integer

        ilMonth = 0
        For ilLoop = 1 To 12
            If Trim$(UCase(slInputMonth)) = Mid$(UCase(slCorpMonthString), (ilLoop - 1) * 3 + 1, 3) Then
                ilMonth = ilLoop
                Exit For
            End If
        Next ilLoop
        gGetCorpMonthNoFromMonthName = ilMonth
        Exit Function
End Function
'
'           gFilterAgyCodes - Find a matching direct advertiser or agency code
'           <input>  ilAgfCode = value of agency code; negated if direct advt
'                    ilIncludeCodes - true to test for inclusion of code; otherwise test for exclusion of code
'                    ilUseCodes() - array of codes to include or exclude
'            <return> true = include transaction, else false to exclude
'
Public Function gFilterAgyAdvCodes(ilAgfCode As Integer, ilIncludeCodes As Integer, ilUseCodes() As Integer) As Integer
Dim ilCompare As Integer
Dim ilTemp As Integer
Dim ilFoundOption As Integer
Dim ilRet As Integer
        
        ilFoundOption = False
        ilCompare = ilAgfCode
        If ilIncludeCodes Then
            For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
'                If ilUseCodes(ilTemp) < 0 Then      '11-11-03 do not test tmrvf.iagfcode, because there are some transactions that do have agencies
'                                                    'as well as the advt having been changed to non-direct
'                'If tmRvf.iAgfCode = 0 Then              'direct, codes has been negated to know to test advertisr code
'                    If ilUseCodes(ilTemp) = ilAgfCode Then
'                        ilFoundOption = True
'                        Exit For
'                    End If
'                Else
                    If ilUseCodes(ilTemp) = ilAgfCode Then
                        ilFoundOption = True
                        Exit For
                    End If
'                End If
            Next ilTemp
        Else
            ilFoundOption = True        '8/23/99 when more than half selected, selection fixed
            For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
'                If ilUseCodes(ilTemp) < 0 Then          '11-11-03 do not test tmrvf.iagfcode, because there are some transactions that do have agencies
'                                                        'as well as the advt having been changed to non-direct
'                'If tmRvf.iAgfCode = 0 Then              'direct, codes has been negated to know to test advertisr code
'                    If ilUseCodes(ilTemp) = -ilAgfCode Then
'                        ilFoundOption = False
'                        Exit For
'                    End If
'                Else
                    If ilUseCodes(ilTemp) = ilAgfCode Then
                        ilFoundOption = False
                        Exit For
                    End If
'                End If
            Next ilTemp
        End If
        gFilterAgyAdvCodes = ilFoundOption
End Function
'
'
'               gGetMonthFromInx - determine month name from the month # entered
'               For std and calendar, the month index is relative to Jan
'               For corp, its relative to the start month of the corporate calendar
'               i.e. corp calendar start month = Oct.  If starting month input on
'               report is 4, the month returned would be Jan for the report header.
'               <input> rbcCorp = true if corporate
'                       rbcStd = true if standard
'                       rbcCal = true if standard
'               <return> string of type of Month (corp, std, cal), starting month (Jan - dec), & Year
'                       i.e. Corp Oct 2009
Public Function gGetMonthFromInx(rbcCorp As Control, rbcStd As Control, rbcCal As Control) As String
Dim slStr As String
Dim ilLoop As Integer
Dim ilRet As Integer
Dim ilSaveMonth As Integer
Dim ilTemp As Integer
Dim slMonthInYear As String * 36
Dim slMonth As String
Dim ilStartQtr As Integer
'Dim ilWhichQtr(1 To 4) As Integer
Dim ilYear As Integer
Dim slTypeOfMonth As String

    'igMonthOrQtr is a month # ; used to be starting qtr but has changed to starting month
    slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"
    gGetMonthFromInx = ""
    ilSaveMonth = igMonthOrQtr
    If rbcCorp.Value Then          'corp option
        slMonthInYear = gGetCorpMonthString(igYear)     'get the string for the months of the corp cal:  i.e. OctNoveDecJan....Sep
        slMonth = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)    'get month text from input month # relative to start of the corp year
        slTypeOfMonth = "Corp"
        gGetMonthNoFromString slMonth, ilSaveMonth         'getmonth index for the first column header (actual month user wants report to start)

    ElseIf rbcStd.Value Then     'std
        slMonth = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)    'get month text from input month #
        slTypeOfMonth = "Std"
    Else                                'calendar (calc by day)
        slMonth = Mid$(slMonthInYear, (igMonthOrQtr - 1) * 3 + 1, 3)    'get month text from input month #
        slTypeOfMonth = "Cal"
    End If
 
    If Not gSetFormula("StartingMonth", ilSaveMonth) Then        'pass starting month of the starting std qtr for report column headings
        Exit Function
    End If
    gGetMonthFromInx = Trim$(slTypeOfMonth) & " " & slMonth & " " & str$(igYear)
End Function
'
'           Format spot time (long) for 12hr time or 24hr
'           <input> time of spot
'                   ilType = 12 for 12 hr time, 24 for 24 hr time
'           return - string of converted time
Public Function gFormatSpotTimeByType(llSpotTime As Long, ilType As Integer)
Dim llHour As Long
Dim llMin As Long
Dim llSec As Long
Dim slStr As String
Dim slWholeTime As String
Dim slMeridiem As String * 2

    slWholeTime = ""
    llHour = llSpotTime \ 3600
    llMin = llSpotTime Mod 3600
    llSec = llMin Mod 60
    llMin = llMin \ 60
    If ilType = 24 Then                'military time (24 hr time)
        slStr = Trim$(str$(llHour))
        Do While Len(slStr) < 2
            slStr = "0" & slStr
        Loop
        slWholeTime = slWholeTime & slStr & ":"
        'Minutes
        slStr = Trim$(str$(llMin))
        Do While Len(slStr) < 2
            slStr = "0" & slStr
        Loop
        slWholeTime = slWholeTime & slStr & ":"
        'Seconds
        slStr = Trim$(str$(llSec))
        Do While Len(slStr) < 2
            slStr = "0" & slStr
        Loop
        slWholeTime = slWholeTime & slStr
    Else            'am/pm
        If llHour < 12 Then         'am
            slMeridiem = "AM"
            
        Else
            slMeridiem = "PM"
            llHour = llHour - 12
        End If
        slStr = Trim$(str$(llHour))
        Do While Len(slStr) < 2
            slStr = "0" & slStr
        Loop
        slWholeTime = slWholeTime & slStr & ":"
        'Minutes
        slStr = Trim$(str$(llMin))
        Do While Len(slStr) < 2
            slStr = "0" & slStr
        Loop
        slWholeTime = slWholeTime & slStr & ":"
        'Seconds
        slStr = Trim$(str$(llSec))
        Do While Len(slStr) < 2
            slStr = "0" & slStr
        Loop
        slWholeTime = slWholeTime & slStr & slMeridiem
    
    End If
    gFormatSpotTimeByType = slWholeTime
End Function
'           gVerifyDate - Verify date entered in report input selectivity
'           <input>  Form - form name
'                    control name - text input field
'           return - true if valid date
Public Function gVerifyDate(Form As Form, InpDate As Control) As Integer
Dim slDate As String
        gVerifyDate = True
        slDate = InpDate.Text
        If slDate <> "" Then
            If Not gValidDate(slDate) Then
                gReset Form
                InpDate.SetFocus
                gVerifyDate = False
                Exit Function
            End If
        End If
End Function
'           gReset - used when input parameter is invalid data
'           Put focus on invalid input field , beep
Public Sub gReset(Form As Form)
    igGenRpt = False
    Form!frcOutput.Enabled = igOutput
    Form!frcCopies.Enabled = igCopies
    Form!frcFile.Enabled = igFile
    Form!frcOption.Enabled = igOption
    Beep
End Sub
'
'           gRptVehPop: populate vehicles by vehicle type
'           <input>  Form name
'                    lbcSelection - list box to contain the vehicles
'                    llVehicleTypes - vehicles type constants
'                        ie VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHNTR + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + ACTIVEVEH
'          <output>  List box filled with requested vehicles
'
Function gRptVehPop(frm As Form, lbcSelection As Control, llVehicletypes As Long) As Integer
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(frm, llVehicletypes, lbcSelection, tgVehicle(), sgVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo gRptVehPopErr
        gCPErrorMsg ilRet, "gRptVehPop (gPopUserVehicleBox: Vehicle)", frm
        On Error GoTo 0
    End If
    Exit Function
gRptVehPopErr:
    On Error GoTo 0
    gRptVehPop = True
    Exit Function
End Function

'
'           Determine the gross, net or T-net value from spot price
'           gGetGrossNetTnetFromPrice
'           <input> ilGrossNetTnet - 0 = return gross, 1 = return net, 2 = return tnet
'                   llSpotPrice (Spot price from cff as long)
'                   ilAgfCode (agy index to determine % of commission)
'           <return>  gross, net or tnet value (long)
'
Public Function gGetGrossNetTNetFromPrice(ilGrossNetTNet As Integer, llSpotPrice As Long, llAcquisitionCost As Long, ilChfAgfCode As Integer) As Long
Dim ilAgfInx As Integer
Dim ilCommPct As Integer
Dim slAmount As String
Dim slSharePct As String
Dim slStr As String
Dim llTempActPrice As Long

        llTempActPrice = llSpotPrice
        If ilGrossNetTNet > 0 Then          'net or tnet; 0 = gross, return same value
            If ilChfAgfCode > 0 Then                'agy exists
                ilAgfInx = gBinarySearchAgf(ilChfAgfCode)
                
                ilCommPct = 8500         'default to commissionable if no agency found
                'see what the agency comm is defined as
                If ilAgfInx <> -1 Then
                    ilCommPct = (10000 - tgCommAgf(ilAgfInx).iCommPct)
                End If
    
                slAmount = gLongToStrDec(llSpotPrice, 2)
                slSharePct = gIntToStrDec(ilCommPct, 4)
                slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
                slStr = gRoundStr(slStr, ".01", 2)
                llTempActPrice = gStrDecToLong(slStr, 2) 'adjusted net
                If ilGrossNetTNet = 2 Then          'tnet, adjust for acquisition cost
                    llTempActPrice = llTempActPrice - llAcquisitionCost
                End If
                gGetGrossNetTNetFromPrice = llTempActPrice
            End If
        End If
        gGetGrossNetTNetFromPrice = llTempActPrice
        Exit Function
End Function
'
'       Test the check box to see if checked or unchecked
Public Function gSetIncludeExcludeCkc(ilckc As Control) As Boolean
   If ilckc = vbChecked Then
        gSetIncludeExcludeCkc = True
    Else
        gSetIncludeExcludeCkc = False
    End If
    Exit Function
End Function
Public Function gVerifyMonthYrPeriods(slPeriodType As String, slYear As String, slMonthDate As String, slPeriods As String, ilLoLimit As Integer, ilHiLimit As Integer, ilInvalidField As Integer) As Boolean
'******************************************************************************************
'       <input>  slPeriodType : W = weekly (date entered), C = Calendar, U = Corporate, user defined, S = std
'                slYear - starting year, not used if weekly period type
'                slMonth - Month text, or month #, or if weekly, its the start date
'                slPeriods - # periods to process
'                ilLoLimit - lo limit of valid #
'                ilHiLimit - hi limit of valid #
'        <output>  ilInvalidField - 1 = year, 2 = month or startdate, 3 = # periods
'******************************************************************************************

Dim slStr As String
Dim ilSaveMonth As Integer
Dim ilRet As Integer
Dim slMonthInYear As String * 36
Dim slDateFrom As String
Dim llDate As Long
Dim ilYear As Integer

        ilInvalidField = 0
        gVerifyMonthYrPeriods = True
        If slPeriodType = "W" Then      'weekly
            'check for valid start
            slDateFrom = Trim$(slMonthDate)
            If Not gValidDate(slDateFrom) Then
                gVerifyMonthYrPeriods = False
                ilInvalidField = 2
                Exit Function
            End If
            llDate = gDateValue(slDateFrom)
            slDateFrom = Format$(llDate, "m/d/yy")
            slStr = Trim$(slPeriods)           '#weeks
        Else                                                'monthly
            slMonthInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"
  
            'verify user input dates
            igYear = gVerifyYear(slYear)
            If igYear = 0 Then
                gVerifyMonthYrPeriods = False
                ilInvalidField = 1
                Exit Function
            End If
            slStr = Trim$(slMonthDate)                'month in text form (jan..dec), or just a month # could have been entered
            gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
            If ilSaveMonth = 0 Then                                 'input isn't text month name, try month #
                ilSaveMonth = Val(slStr)
            Else
                slStr = str$(ilSaveMonth)
            End If

            ilRet = gVerifyInt(slStr, ilLoLimit, ilHiLimit)        'if month number came back 0, its invalid
            If ilRet = -1 Then
                gVerifyMonthYrPeriods = False
                ilInvalidField = 2
                Exit Function
            End If

            'also test the converted month
            If ilSaveMonth < ilLoLimit Or ilSaveMonth > ilHiLimit Then
                gVerifyMonthYrPeriods = False
                ilInvalidField = 2
                Exit Function
            End If
            
            If slPeriodType = "U" Then          'corporate
                 'convert the month name to the correct relative month # of the corp calendar
                'i.e. if 10 entered and corp calendar starts with oct, the result will be july (10th month of corp cal)
                ilYear = gGetCorpCalIndex(igYear)
                If ilYear <= 0 Then                  'year not defined
                    MsgBox "Corporate Year not Defined"
                    gVerifyMonthYrPeriods = False         'invalid corporate calendar
                    ilInvalidField = 1
                    Exit Function
                End If
                slStr = slMonthDate                 'month in text form (jan..dec), or just a month # could have been entered
                gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
                If ilSaveMonth <> 0 Then                           'input is text month name,
                    slMonthInYear = gGetCorpMonthString(igYear)     'get the string for the months of the corp cal:  i.e. OctNoveDecJan....Sep
                    igMonthOrQtr = gGetCorpMonthNoFromMonthName(slMonthInYear, slStr)         'getmonth # relative to start of corp cal
                Else
                    igMonthOrQtr = Val(slStr)
                End If
            Else                                        'std or calendar
                igMonthOrQtr = Val(slStr)
            End If
           
            slStr = Trim$(slPeriods)           '#periods
 
        End If

        ilRet = gVerifyInt(slStr, ilLoLimit, ilHiLimit)             '#periods  validity check
        If ilRet = -1 Then
            gVerifyMonthYrPeriods = False
            ilInvalidField = 3
            Exit Function
        End If

        
    Exit Function
End Function
'
'                   get the spot lengths and their index ratios for 30" spot lengths
'                   return back each spot length and index into array
'           <Input> slspotlen - string of 10 spot lengths
'                   slIndex - string of 10 spot length indices
'           <output> tlSpotLenRatio - array of spot lengths and their ratios in spot length order, lo to hi
Sub gBuildSpotLenAndIndexTable(slSpotLen() As String, slIndex() As String, tlSpotLenRatio As SPOTLENRATIO)

Dim ilLoop As Integer
Dim slStr As String
Dim ilDone As Integer
Dim ilRet As Integer

        'build the spot lengths lo to hi order, along with its associated index value
        For ilLoop = 0 To 9
            tlSpotLenRatio.iLen(ilLoop) = Val(slSpotLen(ilLoop))
            slStr = slIndex(ilLoop)
            gFormatStr slStr, 0, 2, slStr
            tlSpotLenRatio.iRatio(ilLoop) = gStrDecToInt(slStr, 2)
        Next ilLoop
        ilDone = False
        Do While Not ilDone
            ilDone = True
            For ilLoop = 1 To 9
                If tlSpotLenRatio.iLen(ilLoop - 1) > tlSpotLenRatio.iLen(ilLoop) And tlSpotLenRatio.iLen(ilLoop) > 0 Then
                    'swap the two lengths
                    ilRet = tlSpotLenRatio.iLen(ilLoop - 1)
                    tlSpotLenRatio.iLen(ilLoop - 1) = tlSpotLenRatio.iLen(ilLoop)
                    tlSpotLenRatio.iLen(ilLoop) = ilRet
                    
                    'swap the two index values
                    ilRet = tlSpotLenRatio.iRatio(ilLoop - 1)
                    tlSpotLenRatio.iRatio(ilLoop - 1) = tlSpotLenRatio.iRatio(ilLoop)
                    tlSpotLenRatio.iRatio(ilLoop) = ilRet
    
                    ilDone = False
                End If
            Next ilLoop
        Loop
    Exit Sub
    End Sub
'
'               gFilterContract Types - test the contract type (order, hold, standard, remnant) for inclusion
'           <input> tlChf - contract header image
'                   tlCntTypes - contract header types to include/exclude
'                   blIncludeProposalCheck
'            return - true is OK to process, passes filter
Public Function gFilterContractType(tlChf As CHF, tlCntTypes As CNTTYPES, blIncludeProposalCheck As Boolean) As Boolean

            'assume valid contract type
            gFilterContractType = True
            'test contract types
            If (tlChf.sStatus = "H" Or tlChf.sStatus = "G") And Not (tlCntTypes.iHold) Then     'sch & unsch holds
                gFilterContractType = False
            End If
            If (tlChf.sStatus = "O" Or tlChf.sStatus = "N") And Not (tlCntTypes.iOrder) Then        'sch & unsch orders
                gFilterContractType = False
            End If
    
            If tlChf.sType = "T" And Not tlCntTypes.iRemnant Then
                gFilterContractType = False
            End If
            If tlChf.sType = "Q" And Not tlCntTypes.iPI Then
                gFilterContractType = False
            End If
            If tlChf.iPctTrade = 100 And Not tlCntTypes.iTrade Then
                gFilterContractType = False
            End If
            If tlChf.sType = "M" And Not tlCntTypes.iPromo Then
                gFilterContractType = False
            End If
            If tlChf.sType = "S" And Not tlCntTypes.iPSA Then
                gFilterContractType = False
            End If
             If tlChf.sType = "C" And Not tlCntTypes.iStandard Then      'include Standard types?
                gFilterContractType = False
            End If
            If tlChf.sType = "V" And Not tlCntTypes.iReserv Then      'include reservations ?
                gFilterContractType = False
            End If
            If tlChf.sType = "R" And Not tlCntTypes.iDR Then      'include DR?
                gFilterContractType = False
            End If
            
            If blIncludeProposalCheck Then              'continue to test the proposal types
                If tlChf.sStatus = "W" And Not tlCntTypes.iWorking Then
                    gFilterContractType = False
                End If
                If tlChf.sStatus = "C" And Not tlCntTypes.iComplete Then
                    gFilterContractType = False
                End If
                If tlChf.sStatus = "I" And Not tlCntTypes.iIncomplete Then      'Incomplete/unapproved
                    gFilterContractType = False
                End If
           End If

    Exit Function
End Function
'********************************************************
'*                                                      *
'*   Procedure Name:gVerifyMonth                        *
'*                                                      *
'*   Created:2/5/01       By:D. Smith                   *
'*   Modified:            By:                           *
'*                                                      *
'*   Comments: Test if entered month is between 1-12    *
'*             or Jan-Dec.  If so return the numeric    *
'*             string representation of the month       *
'*             otherwise return "0"                     *
'*                                                      *
'********************************************************
Function gVerifyMonth(slInput As String) As String
Dim ilIntVal As Integer
    'Did they enter a number
    If IsNumeric(slInput) Then
        ilIntVal = Val(slInput)
        If ((ilIntVal > 0) And (ilIntVal < 13)) Then
            gVerifyMonth = slInput
        Else
            gVerifyMonth = "0"
        End If
    Else
        Select Case UCase(slInput)
            Case "Jan", "January"
                gVerifyMonth = "1"
            Case "FEB", "FEBRUARY"
                gVerifyMonth = "2"
            Case "MAR", "MARCH"
                gVerifyMonth = "3"
            Case "APR", "APRIL"
                gVerifyMonth = "4"
            Case "MAY"
                gVerifyMonth = "5"
            Case "JUN", "JUNE"
                gVerifyMonth = "6"
            Case "JUL", "JULY"
                gVerifyMonth = "7"
            Case "AUG", "AUGUST"
                gVerifyMonth = "8"
            Case "SEP", "SEPT", "SEPTEMBER"
                gVerifyMonth = "9"
            Case "OCT", "OCTOBER"
                gVerifyMonth = "10"
            Case "NOV", "NOVEMBER"
                gVerifyMonth = "11"
            Case "DEC", "DECEMBER"
                gVerifyMonth = "12"
            Case Else
                gVerifyMonth = "0"
        End Select
    End If
End Function

'
'               gSingleContract - determine if single contract # entered
'                                 & retrieve it
'               <input>  contract #
'               <return> Single selection contract code
Public Function gSingleContract(hlChf As Integer, tlChf As CHF, llSingleCntr As Long) As Long
Dim llChfCode As Long
Dim ilFoundCntr As Integer
Dim ilRet As Integer
Dim tlChfSrchKey1 As CHFKEY1

        llChfCode = -1
        'determine if there is a single contract to retrieve
        ilFoundCntr = False
        If llSingleCntr > 0 Then            'get the contracts internal code
            tlChfSrchKey1.lCntrNo = llSingleCntr
            tlChfSrchKey1.iCntRevNo = 32000
            tlChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hlChf, tlChf, Len(tlChf), tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tlChf.lCntrNo = llSingleCntr)
                If ((tlChf.sSchStatus = "F") Or (tlChf.sSchStatus = "M")) And (tlChf.sDelete <> "Y") Then
                    ilFoundCntr = True
                    llChfCode = tlChf.lCode
                    Exit Do
                End If
                ilRet = btrGetNext(hlChf, tlChf, Len(tlChf), BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If Not ilFoundCntr Then
                gSingleContract = 0
                Exit Function
            End If
        Else
            llChfCode = 0
        End If
        gSingleContract = llChfCode
End Function
'
'               convert generation date and time to string for Email PDF filename
'               Remove slash and replace with nothing in date, remove colon in time
'           gFormatGenDateTime()
'           <input>  ilInputdate(0 to 1)
'                    llInputTime
'           <output>  slDate
'                     slTime
Public Sub gFormatGenDateTimeToStr(ilInputDate() As Integer, slDate As String, llInputTime As Long, slTime As String)
Dim slTemp As String
Dim ilLoop As Integer
Dim slStr As String
Dim llDate As Long
       gUnpackDateLong ilInputDate(0), ilInputDate(1), llDate
        slStr = Format$(llDate, "ddddd")               'Now date as string
       'replace slash with nothing in date
        slDate = ""
        For ilLoop = 1 To Len(slStr) Step 1
            slTemp = Mid$(slStr, ilLoop, 1)
            If slTemp <> "/" Then
                slDate = Trim$(slDate) & Trim$(slTemp)
            End If
        Next ilLoop
        
        slStr = gFormatTimeLong(llInputTime, "A", "1")
        'remove colons from time
        slTime = ""
        For ilLoop = 1 To Len(slStr) Step 1
            slTemp = Mid$(slStr, ilLoop, 1)
            If slTemp <> ":" Then
                slTime = Trim$(slTime) & Trim$(slTemp)
            End If
        Next ilLoop
        Do While Len(slTime) < 8
            slTime = "0" & slTime
        Loop
        
        
        Exit Sub
End Sub
'
'Place a decimal in a string amt
'
Public Function gFormatStrDec(slAmt As String, ilNoDecPlaces As Integer) As String
Dim slTemp As String
        slTemp = Trim$(slAmt)
        If Len(slTemp) >= ilNoDecPlaces Then
            slTemp = Left$(slTemp, Len(slTemp) - ilNoDecPlaces) & "." & right$(slTemp, ilNoDecPlaces)
        Else
            Do While Len(slTemp) < ilNoDecPlaces
                slTemp = "0" & slTemp
            Loop
            slTemp = "." & slTemp
        End If
        gFormatStrDec = slTemp
        Exit Function
End Function
