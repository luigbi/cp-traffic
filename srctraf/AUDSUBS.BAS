Attribute VB_Name = "AUDSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of AUDSUBS.BAS on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: AudSubs.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the File subs and functions
Option Explicit
Option Compare Text
'10256 changed to show its global
Public sgAvgAudStamp As String
Dim lmAvgAud As Long
Dim imAudFromSource As Integer
Dim lmAudFromCode As Long
Dim smPopStamp As String
Dim lmPop As Long
Dim lmPopEst As Long
Dim hmDnf As Integer
Dim tmDnf As DNF
Dim imDnfRecLen As Integer
Dim tmRaf As RAF
Dim imRafRecLen As Integer
Dim tmRafSrchKey0 As LONGKEY0
'10256
Public sgActLineup As String
Public igActTime As Integer
Public igActSpots As Integer
Public igActPct As Integer
Public igActDaypart As Integer


'*******************************************************
'*                                                     *
'*      Procedure Name:gAvgAudToLnResearch             *
'*                                                     *
'*             Created:10/22/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute Rating, Gross Impression*
'*                     Gross Rating points             *
'*                     Rating = 100%*AvgAud/Population *
'*                     AvgRating = Sum(Rating)/# Summed*
'*                     GrImp = WkNoSpots*AvgAud        *
'*                     GrImpTotal = Sum of GrImp       *
'*                     GRP = 100%*GrImp/Population     *
'*                     GRPTotal = Sum of GRP           *
'*                                                     *
'*******************************************************
'Sub gAvgAudToLnResearch(sl1or2PlaceRating As String, ilAllowNeg As Integer, llLnPop As Long, llWkPopEst() As Long, llWkNoSpots() As Long, llSpotPrice() As Long, llAvgAud() As Long, llTotalCost As Long, llAvgAudAvg As Long, ilRating() As Integer, ilAvgRating As Integer, llGrImp() As Long, llTotalGrImp As Long, llGRP() As Long, llTotalGRP As Long, llCPP As Long, llCPM As Long, llPopEst As Long, Optional flTotalCost As Single)
Sub gAvgAudToLnResearch(sl1or2PlaceRating As String, ilAllowNeg As Integer, llLnPop As Long, llWkPopEst() As Long, llWkNoSpots() As Long, llSpotPrice() As Long, llAvgAud() As Long, dlTotalCost As Double, llAvgAudAvg As Long, ilRating() As Integer, ilAvgRating As Integer, llGrImp() As Long, llTotalGrImp As Long, llGRP() As Long, llTotalGRP As Long, llCPP As Long, llCPM As Long, llPopEst As Long, Optional flTotalCost As Single) 'TTP 10439 - Rerate 21,000,000
'
'   gAvgAudToLnResearch sl1or2PlaceRating, ilAllowNeg, llLnPop, llWkPopEst(), llWkNoSpots(), llSpotPrice(), llAvgAud(), llTotalCost, llAvgAudAvg, ilRating(), ilAvgRating, llGrImp(), llTotalGrImp, llGRP(), llTotalGRP, llCPP, llCPM
'   Where:
'       sl1or2PlaceRating(I)- 1 or 2 place rating. Test for 2.
'       ilAllowNeg(I)- True if neg spots and dollars allowed
'       llLnPop(I)- Poulation value (for tgSpf.sDemoEstAllowed = "Y", llEkPopEst() used and llPopEst is set)
'       llWkPopEst()(I)- Array of population, used instead of llPop if  tgSpf.sDemoEstAllowed = "Y" since it can vary by week
'       llWkNoSpots()(I)- Array of number of spots/week.  Size of the array should be exactly the number of weeks that values are to be commputed for
'       llSpotPrice()(I)- Array of spot prices (two place dec number 123.45 stored as 12345)
'       llAvgAud()(I)- Array of average audience values for each week.  Size of the array should be exactly the number of weeks that values are to be commputed for
'                      Audience values obtained from gGetDemoAvgAud
'       dlTotalCost(O)- Total Cost.  Two decimal places
'       llAvgAudAvg(O)- Average of the average audience values. Whole Number (no decimal point).
'       10/27/14: changed to two places
'       ilRating()(O)- Return rating for week. One or two place after decimal point (One place: 0.8 stored as 8; Two places 0.82 stored as 82). Size of array is = to size of llWkNoSpots
'       ilAvgRating(O)- Average of the Ratings.  Two place after decimal point (0.8 stored as 8).
'       llGrImp()(O)- Return Gross Impression for week .  Whole Number (no decimal point). Size of array is = to size of llWkNoSpots
'       llTotalGrImp(O)- Total Gross Impressions. Whole number
'       llGRP()(O)- Return Gross Rating Points for week.  One place after decimal point (6.7 stored as 67).  Size of array is = to size of llWkNoSpots
'       llTotalGRP(O)- Total Gross Rating Points. One place after decimal point (6.7 stored as 67)
'       llCPP(O)- Cost per Point. Whole number (no decimal place)
'       llCPM(O)- Cost per thousand. Two place after decimal point (6.77 stored as 677)
'       llPopEst(O)- Estimate Population by line if tgSpf.sDemEstAllowed = "Y"
'
'   Note: All input and output array should be allocated in the calling program
'
    Dim illoop As Integer
    Dim llSumAvgAud As Long
    Dim ilNoAirWeeks As Integer
    Dim llCPMAdj As Long
    Dim llTotalRating As Long
    Dim llTempGRP As Long
    Dim llPop As Long
    '11/22/04- fix overflow error with llPopEst
    Dim clPopEst As Currency
    'ReDim ilRating(LBound(llWkNoSpots) To UBound(llWkNoSpots)) As Integer
    'ReDim llGRP(LBound(llWkNoSpots) To UBound(llWkNoSpots)) As Long
    'ReDim llGrImp(LBound(llWkNoSpots) To UBound(llWkNoSpots)) As Long

    On Error GoTo TotalErr:
    llSumAvgAud = 0
    llTotalGrImp = 0
    llTotalRating = 0
    llTotalGRP = 0
    dlTotalCost = 0 'TTP 10439 - Rerate 21,000,000
    'flTotalCost = 0# 'TTP 10439 - Rerate 21,000,000
    ilNoAirWeeks = 0
'    If tgSpf.sSAudData <> "H" Then
'        llCPMAdj = 1
'    Else
'        llCPMAdj = 10
'    End If
    If tgSpf.sSAudData = "H" Then
        llCPMAdj = 10
    ElseIf tgSpf.sSAudData = "N" Then
        llCPMAdj = 100
    ElseIf tgSpf.sSAudData = "U" Then
        llCPMAdj = 1000
    Else
        llCPMAdj = 1
    End If
    llPopEst = 0
    clPopEst = 0
    For illoop = LBound(llWkNoSpots) To UBound(llWkNoSpots) Step 1
        If (llWkNoSpots(illoop) > 0) Or (ilAllowNeg And (llWkNoSpots(illoop) < 0)) Then
            If tgSpf.sDemoEstAllowed <> "Y" Then
                llPop = llLnPop
            Else
                llPop = llWkPopEst(illoop)
            End If
            ilNoAirWeeks = ilNoAirWeeks + 1
            llSumAvgAud = llSumAvgAud + llAvgAud(illoop)
            llGrImp(illoop) = llWkNoSpots(illoop) * llAvgAud(illoop)
            llTotalGrImp = llTotalGrImp + llGrImp(illoop)
            If llPop > 0 Then
                '10/27/14: 2 places
                If (sl1or2PlaceRating <> "2") Or (tgSpf.sSGRPCPPCal <> "A") Then
                    ilRating(illoop) = (1000 * CSng(llAvgAud(illoop))) / llPop    'Use 1000 instead of 100 to get One dec place
                Else
                    ilRating(illoop) = (10000 * CSng(llAvgAud(illoop))) / llPop    'Use 10000 instead of 100 to get Two dec place
                End If
                If tgSpf.sSGRPCPPCal = "R" Then
                    '10/27/14: reduce to one place
                    'llGRP(ilLoop) = llWkNoSpots(ilLoop) * ilRating(ilLoop)
                    llGRP(illoop) = llWkNoSpots(illoop) * CLng(ilRating(illoop))
                ElseIf tgSpf.sSGRPCPPCal = "G" Then
                    'llGRP(ilLoop) = llWkNoSpots(ilLoop) * ilRating(ilLoop)
                    llGRP(illoop) = (1000 * CSng(llGrImp(illoop))) / llPop    'Use 1000 instead of 100 to get One dec place
                Else
                    llGRP(illoop) = (1000 * CSng(llGrImp(illoop))) / llPop    'Use 1000 instead of 100 to get One dec place
                End If
            Else
                ilRating(illoop) = 0
                llGRP(illoop) = 0
            End If
            llTotalRating = llTotalRating + ilRating(illoop)
            llTotalGRP = llTotalGRP + llGRP(illoop)
            If llSpotPrice(illoop) > 0 Then
                If CSng(dlTotalCost) + llWkNoSpots(illoop) * CSng(llSpotPrice(illoop)) > 2000000000 Or dlTotalCost < 0 Then
                    dlTotalCost = -1 'TTP 10439 - Rerate 21,000,000
                    'flTotalCost = flTotalCost + CSng(llWkNoSpots(illoop)) * llSpotPrice(illoop)
                Else
                    dlTotalCost = dlTotalCost + llWkNoSpots(illoop) * llSpotPrice(illoop) 'TTP 10439 - Rerate 21,000,000
                    'flTotalCost = flTotalCost + llWkNoSpots(illoop) * llSpotPrice(illoop)
                End If
            ElseIf ilAllowNeg And (llSpotPrice(illoop) < 0) Then
                If llWkNoSpots(illoop) > 0 Then
                    dlTotalCost = dlTotalCost + llWkNoSpots(illoop) * llSpotPrice(illoop) 'TTP 10439 - Rerate 21,000,000
                    'flTotalCost = flTotalCost + llWkNoSpots(illoop) * llSpotPrice(illoop)
                Else
                    dlTotalCost = dlTotalCost - llWkNoSpots(illoop) * llSpotPrice(illoop) 'TTP 10439 - Rerate 21,000,000
                    'flTotalCost = flTotalCost - llWkNoSpots(illoop) * llSpotPrice(illoop)
                End If
            End If
        Else
            llGrImp(illoop) = 0
            ilRating(illoop) = 0
            llGRP(illoop) = 0
        End If
    Next illoop
    If ilNoAirWeeks > 0 Then
        llAvgAudAvg = llSumAvgAud / ilNoAirWeeks
        ilAvgRating = llTotalRating / ilNoAirWeeks
        llPopEst = llPopEst / ilNoAirWeeks
    Else
        llAvgAudAvg = 0
        ilAvgRating = 0
    End If
    If tgSpf.sDemoEstAllowed = "Y" Then
        'Might want to average the population instead of just testing if not equal
        llPopEst = 0
        clPopEst = 0
        ilNoAirWeeks = 0
'        llPop = 0
        For illoop = LBound(llWkNoSpots) To UBound(llWkNoSpots) Step 1
            If (llWkNoSpots(illoop) > 0) Or (ilAllowNeg And (llWkNoSpots(illoop) < 0)) Then
'                If ilLoop = LBound(llWkNoSpots) Then
'                    llPop = llWkPopEst(ilLoop)
'                Else
'                    If llPop <> llWkPopEst(ilLoop) Then
'                        llPop = 0
'                        Exit For
'                    End If
'                End If
                'llPopEst = llPopEst + llWkPopEst(ilLoop)
                clPopEst = clPopEst + CCur(llWkPopEst(illoop))
                ilNoAirWeeks = ilNoAirWeeks + 1
            End If
        Next illoop
        If ilNoAirWeeks > 0 Then
            'llPopEst = llPopEst / ilNoAirWeeks
            llPopEst = CLng(clPopEst / ilNoAirWeeks)
            llPop = llPopEst
        Else
            llPopEst = 0
            llPop = 0
        End If
    Else
        llPopEst = llLnPop
    End If
    If llPop > 0 Then
        If tgSpf.sSGRPCPPCal = "R" Then
            'Computed above
        ElseIf tgSpf.sSGRPCPPCal = "G" Then
            llTempGRP = (10000 * CSng(llTotalGrImp)) / llPop  'Use 10000 instead of 100 to get two dec place
            'llTotalGRP = llTempGRP / 10
            llTotalGRP = (1000 * CSng(llTotalGrImp)) / llPop  'Use 1000 instead of 100 to get One dec place
        Else
            '10/27/14: 2 places
            If (sl1or2PlaceRating <> "2") Or (tgSpf.sSGRPCPPCal <> "A") Then
                ilAvgRating = (1000 * CSng(llAvgAudAvg)) / llPop  'Use 1000 instead of 100 to get One dec place
            Else
                ilAvgRating = (10000 * CSng(llAvgAudAvg)) / llPop  'Use 10000 instead of 100 to get Two dec place
            End If
            llTotalGRP = (1000 * CSng(llTotalGrImp)) / llPop  'Use 1000 instead of 100 to get One dec place
        End If
        'If tgSpf.sSGRPCPPCal <> "R" Then
        '    ilAvgRating = (1000 * CSng(llAvgAudAvg)) / llPop  'Use 1000 instead of 100 to get One dec place
        '    llTotalGRP = (1000 * llTotalGrImp) / llPop  'Use 1000 instead of 100 to get One dec place
        'Else
        '    llTempGRP = (10000 * CSng(llTotalGrImp)) / llPop  'Use 10000 instead of 100 to get two dec place
        '    llTotalGRP = (llTempGRP + 5) / 10
        'End If
        If (llTotalGrImp > 0) Then
            'llCPP = ((CSng(llTotalCost) / 100) * llPop) / llTotalGrImp
            'llCPM = (CSng(llTotalCost) / 100) / llTotalGrImp
            If tgSpf.sSGRPCPPCal = "R" Then
                If llTotalGRP > 0 Then
                    'llCPP = (CSng(llTotalCost) / 10) / llTotalGRP
                    'llCPP = ((flTotalCost) / 10) / llTotalGRP
                    llCPP = ((dlTotalCost) / 10) / llTotalGRP 'TTP 10439 - Rerate 21,000,000
                Else
                    llCPP = 0
                End If
            ElseIf tgSpf.sSGRPCPPCal = "G" Then
                If llTempGRP > 0 Then  'llGRP > 0 Then
                    ''llCPP = (CSng(llTotalCost) / 10) / llTotalGRP
                    'llCPP = (CSng(llTotalCost)) / llTempGRP
                    'llCPP = (CSng(flTotalCost)) / llTempGRP
                    llCPP = (CSng(dlTotalCost)) / llTempGRP 'TTP 10439 - Rerate 21,000,000
                Else
                    llCPP = 0
                End If
            Else
                'llCPP = ((CSng(llTotalCost) / 10000) * llPop) / llTotalGrImp
                'llCPP = (((flTotalCost) / 10000) * llPop) / llTotalGrImp
                llCPP = (((dlTotalCost) / 10000) * llPop) / llTotalGrImp 'TTP 10439 - Rerate 21,000,000
            End If
            'llCPM = (llCPMAdj * CSng(llTotalCost)) / llTotalGrImp
            'llCPM = (llCPMAdj * CSng(flTotalCost)) / llTotalGrImp
            llCPM = (llCPMAdj * dlTotalCost) / llTotalGrImp 'TTP 10439 - Rerate 21,000,000
        ElseIf ilAllowNeg And (llTotalGrImp < 0) Then
            'If llTotalCost > 0 Then
            'If flTotalCost > 0 Then
            If dlTotalCost > 0 Then 'TTP 10439 - Rerate 21,000,000
                If tgSpf.sSGRPCPPCal = "R" Then
                    If llTotalGRP > 0 Then
                        'llCPP = (CSng(llTotalCost) / 10) / llTotalGRP
                        llCPP = ((dlTotalCost) / 10) / llTotalGRP 'TTP 10439 - Rerate 21,000,000
                    Else
                        llCPP = 0
                    End If
                ElseIf tgSpf.sSGRPCPPCal = "G" Then
                    If llTempGRP > 0 Then  'llGRP > 0 Then
                        ''llCPP = (CSng(llTotalCost) / 10) / llTotalGRP
                        'llCPP = (CSng(llTotalCost)) / llTempGRP
                        llCPP = ((dlTotalCost)) / llTempGRP 'TTP 10439 - Rerate 21,000,000
                    Else
                        llCPP = 0
                    End If
                Else
                    'llCPP = ((CSng(llTotalCost) / 10000) * llPop) / llTotalGrImp
                    llCPP = (((dlTotalCost) / 10000) * llPop) / llTotalGrImp 'TTP 10439 - Rerate 21,000,000
                End If
                'llCPM = (llCPMAdj * CSng(llTotalCost)) / llTotalGrImp
                llCPM = (llCPMAdj * (dlTotalCost)) / llTotalGrImp 'TTP 10439 - Rerate 21,000,000
            Else
                If tgSpf.sSGRPCPPCal = "R" Then
                    If llTotalGRP > 0 Then
                        'llCPP = (CSng(-llTotalCost) / 10) / llTotalGRP
                        llCPP = ((-dlTotalCost) / 10) / llTotalGRP 'TTP 10439 - Rerate 21,000,000
                    Else
                        llCPP = 0
                    End If
                ElseIf tgSpf.sSGRPCPPCal = "G" Then
                    If llTempGRP > 0 Then  'llGRP > 0 Then
                        ''llCPP = (CSng(-llTotalCost) / 10) / llTotalGRP
                        'llCPP = (CSng(-llTotalCost)) / llTempGRP
                        llCPP = ((-dlTotalCost)) / llTempGRP 'TTP 10439 - Rerate 21,000,000
                    Else
                        llCPP = 0
                    End If
                Else
                    'llCPP = ((CSng(-llTotalCost) / 10000) * llPop) / llTotalGrImp
                    llCPP = (((-dlTotalCost) / 10000) * llPop) / llTotalGrImp 'TTP 10439 - Rerate 21,000,000
                End If
                'llCPM = (-llCPMAdj * CSng(llTotalCost)) / llTotalGrImp
                llCPM = (-llCPMAdj * (dlTotalCost)) / llTotalGrImp 'TTP 10439 - Rerate 21,000,000
            End If
        Else
            llCPP = 0
            llCPM = 0
        End If
    Else
        ilAvgRating = 0
'        llTotalGRP = 0
'        llCPP = 0
        If llTotalGRP > 0 Then
            'llCPP = (CSng(llTotalCost) / 10) / llTotalGRP   '(llTotalCost / llGRP) / 10
            llCPP = ((flTotalCost) / 10) / llTotalGRP   '(llTotalCost / llGRP) / 10
            llCPP = ((dlTotalCost) / 10) / llTotalGRP   'TTP 10439 - Rerate 21,000,000
        ElseIf ilAllowNeg And (llTotalGRP < 0) Then
            'If llTotalCost > 0 Then
            If flTotalCost > 0 Then
                'llCPP = (CSng(llTotalCost) / 10) / llTotalGRP   '(llTotalCost / llGRP) / 10
                llCPP = ((flTotalCost) / 10) / llTotalGRP   '(llTotalCost / llGRP) / 10
                llCPP = ((dlTotalCost) / 10) / llTotalGRP   'TTP 10439 - Rerate 21,000,000
            Else
                'llCPP = (CSng(-llTotalCost) / 10) / llTotalGRP   '(-llTotalCost / llGRP) / 10
                llCPP = ((-flTotalCost) / 10) / llTotalGRP   '(-llTotalCost / llGRP) / 10
                llCPP = ((-dlTotalCost) / 10) / llTotalGRP 'TTP 10439 - Rerate 21,000,000
            End If
        Else
            llCPP = 0
        End If
        If llTotalGrImp > 0 Then
            'llCPM = (llCPMAdj * CSng(llTotalCost)) / llTotalGrImp
            'llCPM = (llCPMAdj * (flTotalCost)) / llTotalGrImp
            llCPM = (llCPMAdj * (dlTotalCost)) / llTotalGrImp 'TTP 10439 - Rerate 21,000,000
        ElseIf ilAllowNeg And (llTotalGrImp < 0) Then
            'If llTotalCost > 0 Then
            'If flTotalCost > 0 Then
            If dlTotalCost > 0 Then 'TTP 10439 - Rerate 21,000,000
                'llCPM = (llCPMAdj * CSng(llTotalCost)) / llTotalGrImp
                'llCPM = (llCPMAdj * (flTotalCost)) / llTotalGrImp
                llCPM = (llCPMAdj * (dlTotalCost)) / llTotalGrImp 'TTP 10439 - Rerate 21,000,000
            Else
                'llCPM = (-llCPMAdj * CSng(llTotalCost)) / llTotalGrImp
                'llCPM = (-llCPMAdj * (flTotalCost)) / llTotalGrImp
                llCPM = (-llCPMAdj * (dlTotalCost)) / llTotalGrImp 'TTP 10439 - Rerate 21,000,000
            End If
        Else
            llCPM = 0
        End If
    End If
    
    flTotalCost = dlTotalCost
    Exit Sub
TotalErr:
#If programmatic <> 1 Then
    MsgBox "Combination of Price, Spots and Audience Exceeded 21,000,000", vbOKOnly + vbExclamation, "Warning"
#End If
    For illoop = LBound(llWkNoSpots) To UBound(llWkNoSpots) Step 1
        llGrImp(illoop) = 0
        ilRating(illoop) = 0
        llGRP(illoop) = 0
    Next illoop
    llPopEst = 0
    llPop = 0
    llSumAvgAud = 0
    llTotalGrImp = 0
    llTotalRating = 0
    llTotalGRP = 0
    dlTotalCost = 0
    'flTotalCost = 0#
    ilNoAirWeeks = 0
    ilAvgRating = 0
    llCPP = 0
    llCPM = 0
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gGetDemoAvgAud                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get AQF for CFF Record         *
'*                                                     *
'*******************************************************
Function gGetDemoAvgAud(hlDrf As Integer, hlMnf As Integer, hlDpf As Integer, hlDef As Integer, hlRaf As Integer, ilDnfCode As Integer, ilVefCode As Integer, ilMnfSocEco As Integer, ilMnfDemo As Integer, llStartDateForEst As Long, llEndDateForEst As Long, ilRdfCode As Integer, llInOvStartTime As Long, llInOvEndTime As Long, ilInDay() As Integer, slLineType As String, llRafCode As Long, llAvgAud As Long, llPopForEst As Long, ilAudFromSource As Integer, llAudFromCode As Long) As Long
'Debug.Print "gGetDemoAvgAud: DnfCode:" & ilDnfCode & " ,VefCode:" & ilVefCode & ", RdfCode:" & ilRdfCode & " , AvgAud=" & llAvgAud
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPass                        ilSvInputDay                                            *
'******************************************************************************************

'
'   ilRet = gGetDemoAvgAud(hlDrf, hlMnf, hlDpf, hlDef, ilDnfCode, ilVefCode, ilMnfSocEco, ilMnfDemo, ilRdfCode, llOvStartTime, llOvEndTime, ilInputDay(), slLineType, llAvgAud, llPopForEst, ilAudFromSource, llAudFromCode)
'   Where:
'       hlDrf(I)- Handle open to DRF
'       hlMnf(I)- Handle to open MNF
'       hlDpf(I)- Handle open to DPF
'       hlDef(I)- Handle open to DEF (required if tgSpf.sDemoEstAllowed = "Y")
'       ilDnfCode(I)- Rating book code
'       ilVefCode(I)- Vehicle Code
'       ilMnfSocEco(I)- Social Ecomony MultiName Code or Zero
'       ilMnfDemo(I)- Demo MultiName Code (M18-34)
'       llStartDateForEst(I)- Week start date used only if tgSpf.sDemoEstAllowed = "Y"
'       llEndDateForEst(I)- Week end date used only if tgSpf.sDemoEstAllowed = "Y"
'                           Note:  No off weeks allowed between Start Date and End Date
'       ilRdfCode(I)- Daypart or zero (only if override times defined)
'       llOvStartTime(I)- Override start time (if specified, then RdfCode is ignored)
'       llOvEndTime(I)- Override end time (if specified, then RdfCode is ignored)
'       ilInputDay()(I)- Days to air (Index: 0-Monday, 1=Tuesday,...,6=Sunday); True = air day; False = non air day
'       slLineType(I)- H for Hidden, any other letter for non-hidden
'       llAvgAud(O)- AvgAud values
'       llPopForEst(O)- Population (returned only if tgSpf.sDemoEstAllowed = "Y")
'       ilAudFromSource (O)- 0=Sold Daypart; 1= Extra Daypart; 2= Time; 3=Vehicle; 4=Best Fit Sold Daypart; 5=Best Fit Extra Daypart
'       llAudFromCode (O)- 0 if ilAudFromSource is 2 or 3; RdfCode if ilAudFromSource is 0 or 4; DrfCode if ilAudFromSource is 1 or 5
'
'       Note:  If calling gAvgAudToLnResearch after this call, then this routine should be called for one week only if tgSpf.sDemoEstAllowed = "Y"
'              since popuation can very by week
'
    Dim ilDrfRecLen As Integer     'Record length
    Dim tlDrf As DRF
    Dim tlDrfSrchKey As DRFKEY0
    Dim tlDrfSrchKey2 As LONGKEY0
    Dim ilDpfRecLen As Integer     'Record length
    Dim tlDpf As DPF
    Dim tlDpfSrchKey1 As DPFKEY1
    Dim ilDefRecLen As Integer     'Record length
    Dim tlDef As DEF
    Dim tlDefSrchKey1 As DEFKEY1
    Dim ilMnfRecLen As Integer     'Record length
    Dim tlMnf As MNF
    Dim tlMnfSrchKey As INTKEY0
    Dim slDataType As String    'A=Broadcast data for standard Demo Data; B=Broadcast data for custom demos(1-18); Chr(127) for 19-36; Chr(128) for 37-54;....
    Dim ilCustomIndex As Integer
    Dim ilDemoSex As Integer    'Demo Sex: 0=Man; 1=Women; 2= Both
    Dim ilDemoSAge As Integer   'Demo Start Age
    Dim ilDemoEAge As Integer   'Demo End Age
    Dim ilStartDemoIndex As Integer
    Dim ilEndDemoIndex As Integer
    Dim illoop As Integer
    Dim ilRet As Integer
    Dim ilTestRdf As Integer
    Dim ilDay As Integer
    Dim ilRdf As Integer
    Dim ilIndex As Integer
    Dim ilOk As Integer
    Dim ilNoTimes As Integer
    Dim ilTimeCount As Integer
    Dim llSTime As Long
    Dim llETime As Long
    Dim llTime As Long
    Dim llSRdfTime As Long
    Dim llERdfTime As Long
    Dim ilSvRdf As Integer
    Dim ilQH1 As Integer
    Dim ilDayIndex As Integer
    Dim llMDAvgAud As Long 'MultiDay AvgAud
    Dim ilMDNoSummed As Integer 'MultiDay Number of value to be averaged
    Dim llSDAvgAud As Long 'Single Day AvgAud
    Dim ilSDNoSummed As Integer 'Single Day Number of value to be averaged
    Dim llBFTAvgAud As Long 'Best fit time AvgAud
    Dim ilBFTNoSummed As Integer 'Best Fit Time Day Number of value to be averaged
    Dim slBFTDays As String
    Dim slDays As String
    Dim slStamp As String
                            '3=replace part m-f with m-f and remove override times
    Dim ilDayMatch As Integer
    Dim llTimeMatch As Long
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim ilOvStartTimeQH(0 To 1) As Integer
    Dim ilOvEndTimeQH(0 To 1) As Integer
    Dim ilAnyDayOff As Integer
    ReDim ilInputDay(0 To 6) As Integer
    ReDim ilSDDay(0 To 6) As Integer
    Dim llTAvgAud As Long
    Dim llDate As Long
    Dim llDemoPop As Long
    Dim slEstAud As String
    Dim ilEstNoWks As Integer
    Dim slEstPop As String
    Dim slEstLOrU As String
    Dim ilBestDayCount As Integer
    Dim llBestTime As Long
    Dim llBestDrfCode As Long
    Dim ilCrossMid As Integer   '0=No Start or End time at 12M; 1=Either start or end time at 12m but not both
                                '2=End Time of one segment ends a 12m and another segment starts at 12m
                                '3=Start time after end time (10p-2a)
                                'For case 2 and 3 ignore times between 12m and end time
                                'Use times only with a start time prior to 12m
    Dim llSMidTime As Long
    Dim llEMidTime As Long


    If ((Asc(tgSpf.sUsingFeatures) And HIDDENOVERRIDE) = HIDDENOVERRIDE) And (slLineType = "H") Then
        llOvStartTime = 0
        llOvEndTime = 0
        ilOvStartTimeQH(0) = 1
        ilOvEndTimeQH(0) = 1
        ilOvStartTimeQH(1) = -1
        ilOvEndTimeQH(1) = -1
        For ilDay = LBound(ilInputDay) To UBound(ilInputDay) Step 1
            ilInputDay(ilDay) = False
        Next ilDay
        ilRet = gObtainRcfRifRdf()
        ilRdf = gBinarySearchRdf(ilRdfCode)
        ilAnyDayOff = False
        If ilRdf <> -1 Then
            For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                    ilDayIndex = LBound(ilInputDay)
                    For ilDay = LBound(tgMRdf(ilRdf).sWkDays, 2) To UBound(tgMRdf(ilRdf).sWkDays, 2) Step 1
                        If (tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) = "Y") Then
                            ilInputDay(ilDayIndex) = True
                        End If
                        ilDayIndex = ilDayIndex + 1
                    Next ilDay
                End If
            Next ilIndex
            For ilDay = LBound(ilInputDay) To UBound(ilInputDay) Step 1
                If (ilInputDay(ilDay) = False) Then
                    ilAnyDayOff = True
                End If
            Next ilDay
        Else
            ilDayIndex = LBound(ilInputDay)
            ilAnyDayOff = False
            For illoop = LBound(ilInDay) To UBound(ilInDay) Step 1
                ilInputDay(ilDayIndex) = ilInDay(illoop)
                'If (ilInputDay(ilDayIndex) = True) And (ilDayIndex <= 4) Then
                If (ilInputDay(ilDayIndex) = False) Then
                    ilAnyDayOff = True
                End If
                ilDayIndex = ilDayIndex + 1
            Next illoop
        End If
    Else
        llOvStartTime = llInOvStartTime
        'If (llInOvEndTime > 0) And (llInOvStartTime < llInOvEndTime) Then
        If (llInOvEndTime > 0) And ((llInOvStartTime < llInOvEndTime) Or (llInOvStartTime > llInOvEndTime)) Then
            llOvEndTime = llInOvEndTime - 1 'Don't include end time
        Else
            llOvEndTime = llInOvEndTime
        End If
        'ilOvStartTimeQH = (llOvStartTime \ 900) + 1
        ''This should not be required, it will take extra second off as corrected above- but leave it in
        'If llOvStartTime <> llOvEndTime Then
        '    ilOvEndTimeQH = ((llOvEndTime - 1) \ 900) + 1
        'Else
        '    ilOvEndTimeQH = (llOvEndTime \ 900) + 1
        'End If
        ilOvStartTimeQH(1) = -1
        ilOvEndTimeQH(1) = -1
        ilOvStartTimeQH(0) = (llOvStartTime \ 900) + 1
        'This should not be required, it will take extra second off as corrected above- but leave it in
        If llOvStartTime <> llOvEndTime Then
            If llOvStartTime <= llOvEndTime Then
                ilOvEndTimeQH(0) = ((llOvEndTime - 1) \ 900) + 1
            Else
                ilOvEndTimeQH(0) = ((86400 - 1) \ 900) + 1
                ilOvStartTimeQH(1) = (0 \ 900) + 1
                ilOvEndTimeQH(1) = ((llOvEndTime - 1) \ 900) + 1
            End If
        Else
            ilOvEndTimeQH(0) = (llOvEndTime \ 900) + 1
        End If
        ilDayIndex = LBound(ilInputDay)
        ilAnyDayOff = False
        For illoop = LBound(ilInDay) To UBound(ilInDay) Step 1
            ilInputDay(ilDayIndex) = ilInDay(illoop)
            'If (ilInputDay(ilDayIndex) = True) And (ilDayIndex <= 4) Then
            If (ilInputDay(ilDayIndex) = False) Then
                ilAnyDayOff = True
            End If
            ilDayIndex = ilDayIndex + 1
        Next illoop
    End If
    ilAudFromSource = -1
    llAudFromCode = 0
    ilDayIndex = LBound(ilInputDay)
    If tgSpf.sDemoEstAllowed <> "Y" Then
        slStamp = gFileDateTime(sgDBPath & "Drf.btr") & Trim$(str$(ilDnfCode)) & Trim$(str$(ilVefCode)) & Trim$(str$(ilMnfSocEco)) & Trim$(str$(ilMnfDemo)) & Trim$(str$(ilRdfCode)) & Trim$(str$(llOvStartTime)) & Trim$(str$(llOvEndTime))
    Else
        slStamp = gFileDateTime(sgDBPath & "Drf.btr") & Trim$(str$(ilDnfCode)) & Trim$(str$(ilVefCode)) & Trim$(str$(ilMnfSocEco)) & Trim$(str$(ilMnfDemo)) & Trim$(str$(llStartDateForEst)) & Trim$(str$(llEndDateForEst)) & Trim$(str$(ilRdfCode)) & Trim$(str$(llOvStartTime)) & Trim$(str$(llOvEndTime))
    End If
    slStamp = slStamp & slLineType
    For ilDayIndex = LBound(ilInputDay) To UBound(ilInputDay) Step 1
        If ilInputDay(ilDayIndex) Then
            slStamp = slStamp & "T"
        Else
            slStamp = slStamp & "F"
        End If
    Next ilDayIndex
    If slStamp = sgAvgAudStamp Then
        llAvgAud = lmAvgAud
        llPopForEst = lmPopEst
        ilAudFromSource = imAudFromSource
        llAudFromCode = lmAudFromCode
        gGetDemoAvgAud = True
        '10256 should be already defined, so don't need to set
'Debug.Print " -gGetDemoExit1"
        Exit Function
    End If
    llAvgAud = 0
    ilDrfRecLen = Len(tlDrf) 'btrRecordLength(hlDrf)  'Get and save record length
    ilDpfRecLen = Len(tlDpf) 'btrRecordLength(hlDpf)  'Get and save record length
    ilDefRecLen = Len(tlDef) 'btrRecordLength(hlDef)  'Get and save record length
    ilMnfRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
    tlMnfSrchKey.iCode = ilMnfDemo
    ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        gGetDemoAvgAud = False
'Debug.Print " -gGetDemoExit2"
        '10256
        gSetAct1ToBlank
        Exit Function
    End If
    'Determine which demos make-up demos
    'Standard 16(M12-17;M18-24;M25-34;M35-44;M45-49;M50-54;M55-64;M65+
    '            W12-17;W18-24;W25-34;W35-44;W45-49;W50-54;W55-64;W65+)
    'MnfDemo names: Mxx-xx; Pxx-xx; Axx-xx; Wxx-xx
    'Parse fields
    If tlMnf.iGroupNo = 0 Then
        slDataType = "A"
        If Left$(tlMnf.sName, 1) = "M" Then
            ilDemoSex = 0
        ElseIf Left$(tlMnf.sName, 1) = "W" Then
            ilDemoSex = 1
        Else
            ilDemoSex = 2
        End If
        ilDemoSAge = Val(Mid$(tlMnf.sName, 2, 2))
        If Mid$(tlMnf.sName, 4, 1) = "+" Then
            ilDemoEAge = 99
        Else
            ilDemoEAge = Val(Mid$(tlMnf.sName, 5, 2))
        End If
        ilStartDemoIndex = -1
    Else
        gConvCustomGroup tlMnf.iGroupNo, slDataType, ilCustomIndex
    End If
'Jim: 3/30/05:  Replace the multi-pass with looking for best fit daypart
'               Only one pass is required and addition code added prior to looking for vehicle
'               The addition code will look for best fit dayparts
'               Rule:  Use Days within Times
'                      Buying 6-10 M-F
'                             Result 1: Found 5-10 Mo-Su and 5-11 Mo-Su, use 5-10 Mo-Su
'
'                             Result 2: Found 6-10 Mo-Su and 5-10 Mo-Sa, use 6-10 Mo-Su
'
'    ilPass = 1
'    Do
        ilTestRdf = True
        ilTimeCount = 0
        ilSvRdf = -1
        If ilRdfCode <> 0 Then
            'If override times bypass daypart- use times
            If (llOvStartTime = 0) And (llOvEndTime = 0) Then
                'Check that days specified match daypart definition of days- if not use times
                ilRet = gObtainRcfRifRdf()
                'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) Step 1
                '    If ilRdfCode = tgMRdf(ilRdf).iCode Then
                    ilRdf = gBinarySearchRdf(ilRdfCode)
                    If ilRdf <> -1 Then
                        ilSvRdf = ilRdf
                        For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                            If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                ilTimeCount = ilTimeCount + 1
                                ilDayIndex = LBound(ilInputDay)
                                For ilDay = LBound(tgMRdf(ilRdf).sWkDays, 2) To UBound(tgMRdf(ilRdf).sWkDays, 2) Step 1
                                    If (tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) = "Y") And (ilInputDay(ilDayIndex) <> True) Then
                                        ilTestRdf = False
                                        Exit For
                                    ElseIf (tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) <> "Y") And (ilInputDay(ilDayIndex) = True) Then
                                        ilTestRdf = False
                                        Exit For
                                    End If
                                    ilDayIndex = ilDayIndex + 1
                                Next ilDay
                            End If
                        Next ilIndex
                '        Exit For
                    End If
                'Next ilRdf
                If ilSvRdf = -1 Then
                    gGetDemoAvgAud = False
                    '10256
                    gSetAct1ToBlank
'Debug.Print " -gGetDemoExit3"
                    Exit Function
                End If
            Else
                ilTestRdf = False
            End If
        Else
            ilTestRdf = False
        End If


        'Determine if Daypart record exist and if so get values from it
        If ilTestRdf Then
            'Test if Drf exist for vehicle specified
'Debug.Print " LOAD DEMO DATA1: (DD) DnfCode=" & ilDnfCode & " , MnfSocEco=" & ilMnfSocEco & " , RdfCode=" & ilRdfCode & " , VefCode=" & ilVefCode
            tlDrfSrchKey.iDnfCode = ilDnfCode
            tlDrfSrchKey.sDemoDataType = "D"    'Demo Data
            tlDrfSrchKey.iMnfSocEco = ilMnfSocEco
            tlDrfSrchKey.iVefCode = ilVefCode
            tlDrfSrchKey.sInfoType = "D"        'Daypart
            tlDrfSrchKey.iRdfCode = ilRdfCode
            ilRet = btrGetGreaterOrEqual(hlDrf, tlDrf, ilDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            'Bypass Cume
            Do While (tlDrf.sDataType <> slDataType) And (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = ilDnfCode) And (tlDrf.iVefCode = ilVefCode) And (tlDrf.iMnfSocEco = ilMnfSocEco) And (tlDrf.sDemoDataType = "D") And (tlDrf.sInfoType = "D") And (tlDrf.iRdfCode = ilRdfCode)
                ilRet = btrGetNext(hlDrf, tlDrf, ilDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = ilDnfCode) And (tlDrf.iVefCode = ilVefCode) And (tlDrf.iMnfSocEco = ilMnfSocEco) And (tlDrf.sDemoDataType = "D") And (tlDrf.sInfoType = "D") And (tlDrf.iRdfCode = ilRdfCode) And (tlDrf.iCount >= 0) Then
                llTAvgAud = llAvgAud
                '6/4/16: Replaced GoSub
                'GoSub GetAvgAud
                mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                llAvgAud = llTAvgAud
                '6/4/16: Replaced GoSub
                'GoSub GetEstAdj
                mGetEstAdj ilDnfCode, slEstLOrU, slEstPop, slEstAud, ilEstNoWks, llStartDateForEst, llEndDateForEst, hlDrf, hlMnf, hlDpf, ilMnfSocEco, ilMnfDemo, llDemoPop, hlDef, tlDef, ilDefRecLen, llAvgAud, llPopForEst
                llAvgAud = mRegionAudAdj(hlRaf, llRafCode, llAvgAud)
                lmAvgAud = llAvgAud
                lmPopEst = llPopForEst
                sgAvgAudStamp = slStamp
                ilAudFromSource = 0
                llAudFromCode = ilRdfCode
                imAudFromSource = ilAudFromSource
                lmAudFromCode = llAudFromCode
                gGetDemoAvgAud = True
                '10256
                mSetActOneInfoFromDemoCall tlDrf
'Debug.Print " -gGetDemoExit4"
                Exit Function
            End If
        End If
        If ((tlDrf.iStartTime2(0) = 1) And (tlDrf.iStartTime2(1) = 0)) Or (Not ilTestRdf) Then
            ilNoTimes = 1
        Else
            ilNoTimes = 2
        End If
        'Test for matching daypart records
'Debug.Print " LOAD DEMO DATA2: (DD) DnfCode=" & ilDnfCode & " , MnfSocEco=" & ilMnfSocEco & " , RdfCode=0 , VefCode=" & ilVefCode
        tlDrfSrchKey.iDnfCode = ilDnfCode
        tlDrfSrchKey.sDemoDataType = "D"    'Demo Data
        tlDrfSrchKey.iMnfSocEco = ilMnfSocEco
        tlDrfSrchKey.iVefCode = ilVefCode
        tlDrfSrchKey.sInfoType = "D"        'Daypart Or Time
        tlDrfSrchKey.iRdfCode = 0
        ilRet = btrGetGreaterOrEqual(hlDrf, tlDrf, ilDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = ilDnfCode) And (tlDrf.iVefCode = ilVefCode) And (tlDrf.iMnfSocEco = ilMnfSocEco) And (tlDrf.sDemoDataType = "D") And (tlDrf.sInfoType = "D")
            'Bypass Cume
            If (tlDrf.sDataType = slDataType) And (tlDrf.iCount >= 0) Then
                ilOk = True
                'Test if times and days match- if so use record
                ilDayIndex = LBound(ilInputDay)
                For ilDay = LBound(tlDrf.sDay) To UBound(tlDrf.sDay) Step 1
                    If (tlDrf.sDay(ilDay) = "Y") And (ilInputDay(ilDayIndex) <> True) Then
                        ilOk = False
                        Exit For
                    ElseIf (tlDrf.sDay(ilDay) <> "Y") And (ilInputDay(ilDayIndex) = True) Then
                        ilOk = False
                        Exit For
                    End If
                    ilDayIndex = ilDayIndex + 1
                Next ilDay
                If ilOk Then
                    If (ilRdfCode > 0) And (llOvStartTime = 0) And (llOvEndTime = 0) Then
                        'Check times from rate card
                        If ilTimeCount = ilNoTimes Then
                            For ilIndex = LBound(tgMRdf(ilSvRdf).iStartTime, 2) To UBound(tgMRdf(ilSvRdf).iStartTime, 2) Step 1
                                If (tgMRdf(ilSvRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilSvRdf).iStartTime(1, ilIndex) <> 0) Then
                                    If (tgMRdf(ilSvRdf).iStartTime(0, ilIndex) <> tlDrf.iStartTime(0)) Or (tgMRdf(ilSvRdf).iStartTime(1, ilIndex) <> tlDrf.iStartTime(1)) Then
                                        If (tgMRdf(ilSvRdf).iStartTime(0, ilIndex) <> tlDrf.iStartTime2(0)) Or (tgMRdf(ilSvRdf).iStartTime(1, ilIndex) <> tlDrf.iStartTime2(1)) Then
                                            ilOk = False
                                            Exit For
                                        Else
                                            If (tgMRdf(ilSvRdf).iEndTime(0, ilIndex) <> tlDrf.iEndTime2(0)) Or (tgMRdf(ilSvRdf).iEndTime(1, ilIndex) <> tlDrf.iEndTime2(1)) Then
                                                ilOk = False
                                                Exit For
                                            End If
                                        End If
                                    Else
                                        If (tgMRdf(ilSvRdf).iEndTime(0, ilIndex) <> tlDrf.iEndTime(0)) Or (tgMRdf(ilSvRdf).iEndTime(1, ilIndex) <> tlDrf.iEndTime(1)) Then
                                            If (tgMRdf(ilSvRdf).iStartTime(0, ilIndex) <> tlDrf.iStartTime2(0)) Or (tgMRdf(ilSvRdf).iStartTime(1, ilIndex) <> tlDrf.iStartTime2(1)) Then
                                                ilOk = False
                                                Exit For
                                            Else
                                                If (tgMRdf(ilSvRdf).iEndTime(0, ilIndex) <> tlDrf.iEndTime2(0)) Or (tgMRdf(ilSvRdf).iEndTime(1, ilIndex) <> tlDrf.iEndTime2(1)) Then
                                                    ilOk = False
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next ilIndex
                        Else
                            ilOk = False
                        End If
                    Else
                        'Test override times
                        If ilNoTimes = 1 Then
                            gUnpackTimeLong tlDrf.iStartTime(0), tlDrf.iStartTime(1), False, llSTime
                            If llSTime = llOvStartTime Then
                                gUnpackTimeLong tlDrf.iEndTime(0), tlDrf.iEndTime(1), True, llETime
                                If llOvStartTime <> llOvEndTime Then
                                    If llETime - 1 <> llOvEndTime Then
                                        ilOk = False
                                    End If
                                Else
                                    If llETime <> llOvEndTime Then
                                        ilOk = False
                                    End If
                                End If
                            Else
                                ilOk = False
                            End If
                        Else
                            ilOk = False
                        End If
                    End If
                    If ilOk Then
                        llTAvgAud = llAvgAud
                        '6/4/16: Replaced GoSub
                        'GoSub GetAvgAud
                        mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                        llAvgAud = llTAvgAud
                        '6/4/16: Replaced GoSub
                        'GoSub GetEstAdj
                        mGetEstAdj ilDnfCode, slEstLOrU, slEstPop, slEstAud, ilEstNoWks, llStartDateForEst, llEndDateForEst, hlDrf, hlMnf, hlDpf, ilMnfSocEco, ilMnfDemo, llDemoPop, hlDef, tlDef, ilDefRecLen, llAvgAud, llPopForEst
                        llAvgAud = mRegionAudAdj(hlRaf, llRafCode, llAvgAud)
                        lmAvgAud = llAvgAud
                        lmPopEst = llPopForEst
                        sgAvgAudStamp = slStamp
                        If tlDrf.iRdfCode > 0 Then
                            ilAudFromSource = 0
                            llAudFromCode = tlDrf.iRdfCode
                        Else
                            ilAudFromSource = 1
                            llAudFromCode = tlDrf.lCode
                        End If
                        imAudFromSource = ilAudFromSource
                        lmAudFromCode = llAudFromCode
                        gGetDemoAvgAud = True
                        '10256
                        mSetActOneInfoFromDemoCall tlDrf
'Debug.Print " -gGetDemoExit5"
                        Exit Function
                    End If
                End If
            End If
            ilRet = btrGetNext(hlDrf, tlDrf, ilDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        'MultiDay and Single day test. Placed into same area to reduce disk reads
        ilMDNoSummed = 0
        llMDAvgAud = 0
        ilSDNoSummed = 0
        llSDAvgAud = 0
        '5/13/10:  Add best fit from time
        ilBFTNoSummed = 0
        llBFTAvgAud = 0
        slBFTDays = ""
        For ilDay = LBound(ilSDDay) To UBound(ilSDDay) Step 1
            ilSDDay(ilDay) = False
        Next ilDay
'Debug.Print " LOAD DEMO DATA3: (DT) DnfCode=" & ilDnfCode & " , MnfSocEco=" & ilMnfSocEco & " , RdfCode=0, VefCode=" & ilVefCode
        tlDrfSrchKey.iDnfCode = ilDnfCode
        tlDrfSrchKey.sDemoDataType = "D"    'Demo Data
        tlDrfSrchKey.iMnfSocEco = ilMnfSocEco
        tlDrfSrchKey.iVefCode = ilVefCode
        tlDrfSrchKey.sInfoType = "T"        'Daypart Or Time
        tlDrfSrchKey.iRdfCode = 0
        ilRet = btrGetGreaterOrEqual(hlDrf, tlDrf, ilDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = ilDnfCode) And (tlDrf.iVefCode = ilVefCode) And (tlDrf.iMnfSocEco = ilMnfSocEco) And (tlDrf.sDemoDataType = "D") And (tlDrf.sInfoType = "T")
            'Bypass Cume
            If (tlDrf.sDataType = slDataType) And (tlDrf.sExStdDP <> "Y") And (tlDrf.sExStdDP <> "X") Then
                'Multiday test: Test if days match exactly, then
                'Test if Quarter Hour time is within range of specified times
                ilOk = True
                ilDayIndex = LBound(ilInputDay)
                For ilDay = LBound(tlDrf.sDay) To UBound(tlDrf.sDay) Step 1
                    If (tlDrf.sDay(ilDay) = "Y") And (ilInputDay(ilDayIndex) <> True) Then
                        ilOk = False
                        Exit For
                    ElseIf (tlDrf.sDay(ilDay) <> "Y") And (ilInputDay(ilDayIndex) = True) Then
                        ilOk = False
                        Exit For
                    End If
                    ilDayIndex = ilDayIndex + 1
                Next ilDay
                If ilOk Then
                    gUnpackTimeLong tlDrf.iStartTime(0), tlDrf.iStartTime(1), False, llTime
                    If tlDrf.iQHIndex > 0 Then
                        ilQH1 = tlDrf.iQHIndex
                    Else
                        ilQH1 = (llTime \ 900) + 1
                    End If
                    If (ilRdfCode > 0) And (llOvStartTime = 0) And (llOvEndTime = 0) Then
                        'Check times from rate card
                        For ilIndex = LBound(tgMRdf(ilSvRdf).iStartTime, 2) To UBound(tgMRdf(ilSvRdf).iStartTime, 2) Step 1
                            If (tgMRdf(ilSvRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilSvRdf).iStartTime(1, ilIndex) <> 0) Then
                                gUnpackTimeLong tgMRdf(ilSvRdf).iStartTime(0, ilIndex), tgMRdf(ilSvRdf).iStartTime(1, ilIndex), False, llSTime
                                gUnpackTimeLong tgMRdf(ilSvRdf).iEndTime(0, ilIndex), tgMRdf(ilSvRdf).iEndTime(1, ilIndex), True, llETime
                                ''If (llQH1 >= llSTime) And (llQH1 <= llETime) Then
                                'If (ilQH1 >= (llSTime \ 900) + 1) And (ilQH1 < (llETime \ 900) + 1) Then
                                '    llTAvgAud = llMDAvgAud
                                '    GoSub GetAvgAud
                                '    llMDAvgAud = llTAvgAud
                                '    ilMDNoSummed = ilMDNoSummed + 1
                                'End If
                                If llSTime <= llETime Then
                                    If (ilQH1 >= (llSTime \ 900) + 1) And (ilQH1 < (llETime \ 900) + 1) Then
                                        llTAvgAud = llMDAvgAud
                                        '6/4/16: Replaced GoSub
                                        'GoSub GetAvgAud
                                        mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                                        llMDAvgAud = llTAvgAud
                                        ilMDNoSummed = ilMDNoSummed + 1
                                    End If
                                Else
                                    If ((ilQH1 >= (llSTime \ 900) + 1) And (ilQH1 < (86400 \ 900) + 1)) Or ((ilQH1 >= (0 \ 900) + 1) And (ilQH1 < (llETime \ 900) + 1)) Then
                                        llTAvgAud = llMDAvgAud
                                        '6/4/16: Replaced GoSub
                                        'GoSub GetAvgAud
                                        mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                                        llMDAvgAud = llTAvgAud
                                        ilMDNoSummed = ilMDNoSummed + 1
                                    End If
                                End If
                            End If
                        Next ilIndex
                    Else
                        ''Test override times
                        ''If (llTime >= llOvStartTime) And (llTime <= llOvEndTime) Then
                        ''If ((ilQH1 >= ilOvStartTimeQH) And (ilQH1 < ilOvEndTimeQH)) Then
                        ''If (llTime >= ilOvStartTimeQH) And (llTime <= ilOvEndTimeQH) Then
                        'If (ilQH1 >= ilOvStartTimeQH) And (ilQH1 <= ilOvEndTimeQH) Then
                        '    llTAvgAud = llMDAvgAud
                        '    GoSub GetAvgAud
                        '    llMDAvgAud = llTAvgAud
                        '    ilMDNoSummed = ilMDNoSummed + 1
                        'End If
                        If ilOvStartTimeQH(1) = -1 Then
                            If (ilQH1 >= ilOvStartTimeQH(0)) And (ilQH1 <= ilOvEndTimeQH(0)) Then
                                llTAvgAud = llMDAvgAud
                                '6/4/16: Replaced GoSub
                                'GoSub GetAvgAud
                                mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                                llMDAvgAud = llTAvgAud
                                ilMDNoSummed = ilMDNoSummed + 1
                            End If
                        Else
                            If ((ilQH1 >= ilOvStartTimeQH(0)) And (ilQH1 <= ilOvEndTimeQH(0))) Or ((ilQH1 >= ilOvStartTimeQH(1)) And (ilQH1 <= ilOvEndTimeQH(1))) Then
                                llTAvgAud = llMDAvgAud
                                '6/4/16: Replaced GoSub
                                'GoSub GetAvgAud
                                mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                                llMDAvgAud = llTAvgAud
                                ilMDNoSummed = ilMDNoSummed + 1
                            End If
                        End If
                    End If
                Else    'Best fit from time
                    '5/13/10:  Add Best Fit from Time
                    ilDayMatch = 0
                    slDays = ""
                    ilDayIndex = LBound(ilInputDay)
                    For ilDay = LBound(tlDrf.sDay) To UBound(tlDrf.sDay) Step 1
                        If (tlDrf.sDay(ilDay) = "Y") Then
                            slDays = slDays & "Y"
                        Else
                            slDays = slDays & "N"
                        End If
                        If (tlDrf.sDay(ilDay) = "Y") And (ilInputDay(ilDayIndex) <> True) Then
                            ilDayMatch = ilDayMatch + 1
                        ElseIf (tlDrf.sDay(ilDay) <> "Y") And (ilInputDay(ilDayIndex) = True) Then
                            ilDayMatch = -1
                            Exit For
                        End If
                        ilDayIndex = ilDayIndex + 1
                    Next ilDay
                    If ilDayMatch <> -1 Then
                        gUnpackTimeLong tlDrf.iStartTime(0), tlDrf.iStartTime(1), False, llTime
                        If tlDrf.iQHIndex > 0 Then
                            ilQH1 = tlDrf.iQHIndex
                        Else
                            ilQH1 = (llTime \ 900) + 1
                        End If
                        If (ilRdfCode > 0) And (llOvStartTime = 0) And (llOvEndTime = 0) Then
                            'Check times from rate card
                            For ilIndex = LBound(tgMRdf(ilSvRdf).iStartTime, 2) To UBound(tgMRdf(ilSvRdf).iStartTime, 2) Step 1
                                If (tgMRdf(ilSvRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilSvRdf).iStartTime(1, ilIndex) <> 0) Then
                                    gUnpackTimeLong tgMRdf(ilSvRdf).iStartTime(0, ilIndex), tgMRdf(ilSvRdf).iStartTime(1, ilIndex), False, llSTime
                                    gUnpackTimeLong tgMRdf(ilSvRdf).iEndTime(0, ilIndex), tgMRdf(ilSvRdf).iEndTime(1, ilIndex), True, llETime
                                    If llSTime <= llETime Then
                                        If (ilQH1 >= (llSTime \ 900) + 1) And (ilQH1 < (llETime \ 900) + 1) Then
                                            If (slBFTDays = "") Or (ilDayMatch < ilBestDayCount) Then
                                                slBFTDays = slDays
                                                llBFTAvgAud = 0
                                                ilBFTNoSummed = 0
                                            End If
                                            If slBFTDays = slDays Then
                                                llTAvgAud = llBFTAvgAud
                                                '6/4/16: Replaced GoSub
                                                'GoSub GetAvgAud
                                                mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                                                llBFTAvgAud = llTAvgAud
                                                ilBFTNoSummed = ilBFTNoSummed + 1
                                                ilBestDayCount = ilDayMatch
                                            End If
                                        End If
                                    Else
                                        If ((ilQH1 >= (llSTime \ 900) + 1) And (ilQH1 < (86400 \ 900) + 1)) Or ((ilQH1 >= (0 \ 900) + 1) And (ilQH1 < (llETime \ 900) + 1)) Then
                                            If (slBFTDays = "") Or (ilDayMatch < ilBestDayCount) Then
                                                slBFTDays = slDays
                                                llBFTAvgAud = 0
                                                ilBFTNoSummed = 0
                                            End If
                                            If slBFTDays = slDays Then
                                                llTAvgAud = llBFTAvgAud
                                                '6/4/16: Replaced GoSub
                                                'GoSub GetAvgAud
                                                mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                                                llBFTAvgAud = llTAvgAud
                                                ilBFTNoSummed = ilBFTNoSummed + 1
                                                ilBestDayCount = ilDayMatch
                                            End If
                                        End If
                                    End If
                                End If
                            Next ilIndex
                        Else
                            If ilOvStartTimeQH(1) = -1 Then
                                If (ilQH1 >= ilOvStartTimeQH(0)) And (ilQH1 <= ilOvEndTimeQH(0)) Then
                                    If (slBFTDays = "") Or (ilDayMatch < ilBestDayCount) Then
                                        slBFTDays = slDays
                                        llBFTAvgAud = 0
                                        ilBFTNoSummed = 0
                                    End If
                                    If slBFTDays = slDays Then
                                        llTAvgAud = llBFTAvgAud
                                        '6/4/16: Replaced GoSub
                                        'GoSub GetAvgAud
                                        mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                                        llBFTAvgAud = llTAvgAud
                                        ilBFTNoSummed = ilBFTNoSummed + 1
                                        ilBestDayCount = ilDayMatch
                                    End If
                                End If
                            Else
                                If ((ilQH1 >= ilOvStartTimeQH(0)) And (ilQH1 <= ilOvEndTimeQH(0))) Or ((ilQH1 >= ilOvStartTimeQH(1)) And (ilQH1 <= ilOvEndTimeQH(1))) Then
                                    If (slBFTDays = "") Or (ilDayMatch < ilBestDayCount) Then
                                        slBFTDays = slDays
                                        llBFTAvgAud = 0
                                        ilBFTNoSummed = 0
                                    End If
                                    If slBFTDays = slDays Then
                                        llTAvgAud = llBFTAvgAud
                                        '6/4/16: Replaced GoSub
                                        'GoSub GetAvgAud
                                        mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                                        llBFTAvgAud = llTAvgAud
                                        ilBFTNoSummed = ilBFTNoSummed + 1
                                        ilBestDayCount = ilDayMatch
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                'Single day test: Test that day defined is for one day only and is specified day, then
                'Test if Quarter Hour time is within range of specified times
                ilOk = True
                ilDayIndex = -1
                For ilDay = LBound(tlDrf.sDay) To UBound(tlDrf.sDay) Step 1
                    If (tlDrf.sDay(ilDay) = "Y") Then
                        If ilDayIndex = -1 Then
                            ilDayIndex = ilDay
                        Else
                            ilOk = False
                            Exit For
                        End If
                    End If
                Next ilDay
                If (ilOk) And (ilDayIndex <> -1) Then
                    If (ilInputDay(LBound(ilInputDay) + ilDayIndex) = True) Then
                        ilSDDay(ilDayIndex) = True
                        'llQH1 = (tlDrf.iQHIndex - 1) * CLng(900) + 2
                        gUnpackTimeLong tlDrf.iStartTime(0), tlDrf.iStartTime(1), False, llTime
                        If tlDrf.iQHIndex > 0 Then
                            ilQH1 = tlDrf.iQHIndex
                        Else
                            ilQH1 = (llTime \ 900) + 1
                        End If
                        If (ilRdfCode > 0) And (llOvStartTime = 0) And (llOvEndTime = 0) Then
                            'Check times from rate card
                            For ilIndex = LBound(tgMRdf(ilSvRdf).iStartTime, 2) To UBound(tgMRdf(ilSvRdf).iStartTime, 2) Step 1
                                If (tgMRdf(ilSvRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilSvRdf).iStartTime(1, ilIndex) <> 0) Then
                                    gUnpackTimeLong tgMRdf(ilSvRdf).iStartTime(0, ilIndex), tgMRdf(ilSvRdf).iStartTime(1, ilIndex), False, llSTime
                                    gUnpackTimeLong tgMRdf(ilSvRdf).iEndTime(0, ilIndex), tgMRdf(ilSvRdf).iEndTime(1, ilIndex), True, llETime
                                    ''If (llQH1 >= llSTime) And (llQH1 <= llETime) Then
                                    ''If (llTime >= llSTime) And (llTime <= llETime) Then
                                    'If (ilQH1 >= (llSTime \ 900) + 1) And (ilQH1 < (llETime \ 900) + 1) Then
                                    '    llTAvgAud = llSDAvgAud
                                    '    GoSub GetAvgAud
                                    '    llSDAvgAud = llTAvgAud
                                    '    ilSDNoSummed = ilSDNoSummed + 1
                                    'End If
                                    If llSTime <= llETime Then
                                        If (ilQH1 >= (llSTime \ 900) + 1) And (ilQH1 < (llETime \ 900) + 1) Then
                                            llTAvgAud = llSDAvgAud
                                            '6/4/16: Replaced GoSub
                                            'GoSub GetAvgAud
                                            mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                                            llSDAvgAud = llTAvgAud
                                            ilSDNoSummed = ilSDNoSummed + 1
                                        End If
                                    Else
                                        If ((ilQH1 >= (llSTime \ 900) + 1) And (ilQH1 < (86400 \ 900) + 1)) Or ((ilQH1 >= (0 \ 900) + 1) And (ilQH1 < (llETime \ 900) + 1)) Then
                                            llTAvgAud = llSDAvgAud
                                            '6/4/16: Replaced GoSub
                                            'GoSub GetAvgAud
                                            mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                                            llSDAvgAud = llTAvgAud
                                            ilSDNoSummed = ilSDNoSummed + 1
                                        End If
                                    End If
                                End If
                            Next ilIndex
                        Else
                            ''Test override times
                            ''If (llQH1 >= llOvStartTime) And (llQH1 <= llOvEndTime) Then
                            ''If (llTime >= llOvStartTime) And (llTime <= llOvEndTime) Then
                            ''If (ilQH1 >= ilOvStartTimeQH) And (ilQH1 < ilOvEndTimeQH) Then
                            'If (ilQH1 >= ilOvStartTimeQH) And (ilQH1 <= ilOvEndTimeQH) Then
                            '    llTAvgAud = llSDAvgAud
                            '    GoSub GetAvgAud
                            '    llSDAvgAud = llTAvgAud
                            '    ilSDNoSummed = ilSDNoSummed + 1
                            'End If
                            If ilOvStartTimeQH(1) = -1 Then
                                If (ilQH1 >= ilOvStartTimeQH(0)) And (ilQH1 <= ilOvEndTimeQH(0)) Then
                                    llTAvgAud = llSDAvgAud
                                    '6/4/16: Replaced GoSub
                                    'GoSub GetAvgAud
                                    mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                                    llSDAvgAud = llTAvgAud
                                    ilSDNoSummed = ilSDNoSummed + 1
                                End If
                            Else
                                If ((ilQH1 >= ilOvStartTimeQH(0)) And (ilQH1 <= ilOvEndTimeQH(0))) Or ((ilQH1 >= ilOvStartTimeQH(1)) And (ilQH1 <= ilOvEndTimeQH(1))) Then
                                    llTAvgAud = llSDAvgAud
                                    '6/4/16: Replaced GoSub
                                    'GoSub GetAvgAud
                                    mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                                    llSDAvgAud = llTAvgAud
                                    ilSDNoSummed = ilSDNoSummed + 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            ilRet = btrGetNext(hlDrf, tlDrf, ilDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        'Test if all single days specified
        If ilSDNoSummed > 0 Then
            ilDayIndex = LBound(ilInputDay)
            For ilDay = LBound(ilSDDay) To UBound(ilSDDay) Step 1
                If ilInputDay(ilDayIndex) <> ilSDDay(ilDay) Then
                    ilSDNoSummed = 0
                    Exit For
                End If
                ilDayIndex = ilDayIndex + 1
            Next ilDay
        End If
'        If (ilMDNoSummed > 0) Or (ilSDNoSummed > 0) Then
'            Exit Do
'        End If
'        If ilPass >= 3 Then
'            Exit Do
'        End If
'        If ilPass = 1 Then
'            If (ilAnyDayOff) And (ilRdfCode > 0) Then
'                For ilLoop = 0 To 6 Step 1
'                    ilSvInputDay(ilLoop) = ilInputDay(ilLoop)
'                    ilInputDay(ilLoop) = False
'                Next ilLoop
'                ilRet = gObtainRcfRifRdf()
'                'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) Step 1
'                '    If ilRdfCode = tgMRdf(ilRdf).iCode Then
'                    ilRdf = gBinarySearchRdf(ilRdfCode)
'                    If ilRdf <> -1 Then
'                        For ilIndex = 1 To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
'                            If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
'                                For ilDay = 1 To 7 Step 1
'                                    If (tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) = "Y") Then
'                                        ilInputDay(ilDay - 1) = True
'                                    End If
'                                Next ilDay
'                            End If
'                        Next ilIndex
'                '        Exit For
'                    End If
'                'Next ilRdf
'                ilAnyDayOff = False
'                For ilLoop = 0 To 6 Step 1
'                    If ilSvInputDay(ilLoop) <> ilInputDay(ilLoop) Then
'                        ilAnyDayOff = True
'                        Exit For
'                    End If
'                Next ilLoop
'                If (ilAnyDayOff) Then
'                    ilPass = 2
'                Else
'                    If llInOvStartTime <> 0 Then
'                        llOvStartTime = 0
'                        llOvEndTime = 0
'                        ilPass = 3
'                    Else
'                        Exit Do
'                    End If
'                End If
'            Else
'                If llInOvStartTime <> 0 Then
'                    llOvStartTime = 0
'                    llOvEndTime = 0
'                    ilPass = 3
'                Else
'                    Exit Do
'                End If
'            End If
'        ElseIf ilPass = 2 Then
'            If llInOvStartTime <> 0 Then
'                llOvStartTime = 0
'                llOvEndTime = 0
'                ilPass = 3
'            Else
'                Exit Do
'            End If
'        End If
'    Loop
    If ilMDNoSummed > 0 Then
        llAvgAud = llMDAvgAud / ilMDNoSummed
        '6/4/16: Replaced GoSub
        'GoSub GetEstAdj
        mGetEstAdj ilDnfCode, slEstLOrU, slEstPop, slEstAud, ilEstNoWks, llStartDateForEst, llEndDateForEst, hlDrf, hlMnf, hlDpf, ilMnfSocEco, ilMnfDemo, llDemoPop, hlDef, tlDef, ilDefRecLen, llAvgAud, llPopForEst
        llAvgAud = mRegionAudAdj(hlRaf, llRafCode, llAvgAud)
        lmAvgAud = llAvgAud
        lmPopEst = llPopForEst
        sgAvgAudStamp = slStamp
        ilAudFromSource = 2
        llAudFromCode = 0
        imAudFromSource = ilAudFromSource
        lmAudFromCode = llAudFromCode
        gGetDemoAvgAud = True
        '10256
        mSetActOneInfoFromDemoCall tlDrf
'Debug.Print " -gGetDemoExit6"
        Exit Function
    End If
    If ilSDNoSummed > 0 Then
        llAvgAud = llSDAvgAud / ilSDNoSummed
        '6/4/16: Replaced GoSub
        'GoSub GetEstAdj
        mGetEstAdj ilDnfCode, slEstLOrU, slEstPop, slEstAud, ilEstNoWks, llStartDateForEst, llEndDateForEst, hlDrf, hlMnf, hlDpf, ilMnfSocEco, ilMnfDemo, llDemoPop, hlDef, tlDef, ilDefRecLen, llAvgAud, llPopForEst
        llAvgAud = mRegionAudAdj(hlRaf, llRafCode, llAvgAud)
        lmAvgAud = llAvgAud
        lmPopEst = llPopForEst
        sgAvgAudStamp = slStamp
        ilAudFromSource = 2
        llAudFromCode = 0
        imAudFromSource = ilAudFromSource
        lmAudFromCode = llAudFromCode
        gGetDemoAvgAud = True
        '10256
        mSetActOneInfoFromDemoCall tlDrf
'Debug.Print " -gGetDemoExit7"
        Exit Function
    End If
    '5/13/10:  Add Best Fit from Time
    If ilBFTNoSummed > 0 Then
        llAvgAud = llBFTAvgAud / ilBFTNoSummed
        '6/4/16: Replaced GoSub
        'GoSub GetEstAdj
        mGetEstAdj ilDnfCode, slEstLOrU, slEstPop, slEstAud, ilEstNoWks, llStartDateForEst, llEndDateForEst, hlDrf, hlMnf, hlDpf, ilMnfSocEco, ilMnfDemo, llDemoPop, hlDef, tlDef, ilDefRecLen, llAvgAud, llPopForEst
        llAvgAud = mRegionAudAdj(hlRaf, llRafCode, llAvgAud)
        lmAvgAud = llAvgAud
        lmPopEst = llPopForEst
        sgAvgAudStamp = slStamp
        gGetDemoAvgAud = True
        '10256
        mSetActOneInfoFromDemoCall tlDrf
'Debug.Print " -gGetDemoExit8"
        Exit Function
    End If
    'Test for best fit daypart that supersede the buy
    ilRet = gObtainRcfRifRdf()
    'ilDayIndex = LBound(ilInputDay)
    'For ilLoop = LBound(ilInDay) To UBound(ilInDay) Step 1
    '    ilInputDay(ilDayIndex) = ilInDay(ilLoop)
    '    ilDayIndex = ilDayIndex + 1
    'Next ilLoop
    llSRdfTime = 99999999
    llERdfTime = 0
    If (ilRdfCode <> 0) And (llOvStartTime = 0) And (llOvEndTime = 0) Then
        ilSvRdf = gBinarySearchRdf(ilRdfCode)
        If ilSvRdf <> -1 Then
            ilCrossMid = 0
            For ilIndex = LBound(tgMRdf(ilSvRdf).iStartTime, 2) To UBound(tgMRdf(ilSvRdf).iStartTime, 2) Step 1
                If (tgMRdf(ilSvRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilSvRdf).iStartTime(1, ilIndex) <> 0) Then
                    gUnpackTimeLong tgMRdf(ilSvRdf).iStartTime(0, ilIndex), tgMRdf(ilSvRdf).iStartTime(1, ilIndex), False, llSMidTime
                    gUnpackTimeLong tgMRdf(ilSvRdf).iEndTime(0, ilIndex), tgMRdf(ilSvRdf).iEndTime(1, ilIndex), False, llEMidTime
                    If (llSMidTime <> 0) Or (llEMidTime <> 0) Then
                        If (llSMidTime <= llEMidTime) Or (llEMidTime = 0) Then
                            If llSMidTime = 0 Then
                                If ilCrossMid = 0 Then
                                    ilCrossMid = 1
                                Else
                                    ilCrossMid = 2
                                End If
                            Else
                                If llEMidTime = 0 Then
                                    If ilCrossMid = 0 Then
                                        ilCrossMid = 1
                                    Else
                                        ilCrossMid = 2
                                    End If
                                End If
                            End If
                        Else
                            ilCrossMid = 3
                            Exit For
                        End If
                    End If
                End If
            Next ilIndex
            If ilCrossMid = 2 Then
                For ilDay = LBound(ilInputDay) To UBound(ilInputDay) Step 1
                    ilInputDay(ilDay) = False
                Next ilDay
            End If
            For ilIndex = LBound(tgMRdf(ilSvRdf).iStartTime, 2) To UBound(tgMRdf(ilSvRdf).iStartTime, 2) Step 1
                If (tgMRdf(ilSvRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilSvRdf).iStartTime(1, ilIndex) <> 0) Then
                    If ilCrossMid = 2 Then
                        gUnpackTimeLong tgMRdf(ilSvRdf).iStartTime(0, ilIndex), tgMRdf(ilSvRdf).iStartTime(1, ilIndex), False, llTime
                        If llTime <> 0 Then
                            If llTime < llSRdfTime Then
                                llSRdfTime = llTime
                            End If
                            gUnpackTimeLong tgMRdf(ilSvRdf).iEndTime(0, ilIndex), tgMRdf(ilSvRdf).iEndTime(1, ilIndex), True, llTime
                            If llTime > llERdfTime Then
                                llERdfTime = llTime
                            End If
                            ilDayIndex = LBound(ilInputDay)
                            For ilDay = LBound(tgMRdf(ilSvRdf).sWkDays, 2) To UBound(tgMRdf(ilSvRdf).sWkDays, 2) Step 1
                                'If tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) = "Y" Then
                                If tgMRdf(ilSvRdf).sWkDays(ilIndex, ilDay) = "Y" Then
                                    ilInputDay(ilDayIndex) = True
                                End If
                                ilDayIndex = ilDayIndex + 1
                            Next ilDay
                        End If
                    ElseIf ilCrossMid = 3 Then
                        gUnpackTimeLong tgMRdf(ilSvRdf).iStartTime(0, ilIndex), tgMRdf(ilSvRdf).iStartTime(1, ilIndex), False, llTime
                        If llTime <> 0 Then
                            If llTime < llSRdfTime Then
                                llSRdfTime = llTime
                            End If
                            llTime = 86399
                            If llTime > llERdfTime Then
                                llERdfTime = llTime
                            End If
                            ilDayIndex = LBound(ilInputDay)
                            For ilDay = LBound(tgMRdf(ilSvRdf).sWkDays, 2) To UBound(tgMRdf(ilSvRdf).sWkDays, 2) Step 1
                                'If tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) = "Y" Then
                                If tgMRdf(ilSvRdf).sWkDays(ilIndex, ilDay) = "Y" Then
                                    ilInputDay(ilDayIndex) = True
                                End If
                                ilDayIndex = ilDayIndex + 1
                            Next ilDay
                            Exit For
                        End If
                    Else
                        gUnpackTimeLong tgMRdf(ilSvRdf).iStartTime(0, ilIndex), tgMRdf(ilSvRdf).iStartTime(1, ilIndex), False, llTime
                        If llTime < llSRdfTime Then
                            llSRdfTime = llTime
                        End If
                        gUnpackTimeLong tgMRdf(ilSvRdf).iEndTime(0, ilIndex), tgMRdf(ilSvRdf).iEndTime(1, ilIndex), True, llTime
                        If llTime > llERdfTime Then
                            llERdfTime = llTime
                        End If
                    End If
                End If
            Next ilIndex
        End If
    Else
        If llOvStartTime <= llOvEndTime Then
            llSRdfTime = llOvStartTime
            llERdfTime = llOvEndTime
        Else
            llSRdfTime = llOvStartTime
            llERdfTime = 86399
        End If
    End If
    
'Debug.Print " LOAD DEMO DATA4: (DD) DnfCode=" & ilDnfCode & " , MnfSocEco=" & ilMnfSocEco & " , RdfCode=0, VefCode=" & ilVefCode
    llBestDrfCode = -1
    ilBestDayCount = -1
    llBestTime = -1
    tlDrfSrchKey.iDnfCode = ilDnfCode
    tlDrfSrchKey.sDemoDataType = "D"    'Demo Data
    tlDrfSrchKey.iMnfSocEco = ilMnfSocEco
    tlDrfSrchKey.iVefCode = ilVefCode
    tlDrfSrchKey.sInfoType = "D"        'Daypart Or Time
    tlDrfSrchKey.iRdfCode = 0
    ilRet = btrGetGreaterOrEqual(hlDrf, tlDrf, ilDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = ilDnfCode) And (tlDrf.iVefCode = ilVefCode) And (tlDrf.iMnfSocEco = ilMnfSocEco) And (tlDrf.sDemoDataType = "D") And (tlDrf.sInfoType = "D")
        'Bypass Cume
'Debug.Print " -- reading DRF:" & tlDrf.lCode
        If (tlDrf.sDataType = slDataType) And (tlDrf.iCount >= 0) Then
            ilNoTimes = 1
            ilRdf = tlDrf.iRdfCode
            '12/20/18: Index 0 is now valid
            If ilRdf > 0 Then
                ilRdf = gBinarySearchRdf(tlDrf.iRdfCode)
            ''Else
            ''    ilRdf = -1
            Else
                '2/25/19: rdfCode = 0 is an extra daypart
                If (tlDrf.iRdfCode = 0) And (tlDrf.sInfoType = "D") Then
                    ilRdf = -2
                Else
                    ilRdf = -1
                End If
            End If
            For ilDay = LBound(ilSDDay) To UBound(ilSDDay) Step 1
                ilSDDay(ilDay) = False
            Next ilDay
            llSTime = 9999999
            llETime = 0
            ''12/20/18: Index 0 is now valid
            '''If ilRdf > 0 Then
            '''If ilRdf <> -1 Then
            ''If ilRdf > 0 Then
            ''If ilRdf <> -1 Then
            'Rdf = 0 is extra daypart
            If ilRdf >= 0 Then
                ilCrossMid = 0
                For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                    If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                        gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), False, llSMidTime
                        gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), False, llEMidTime
                        If (llSMidTime <> 0) Or (llEMidTime <> 0) Then
                            If (llSMidTime <= llEMidTime) Or (llEMidTime = 0) Then
                                If llSMidTime = 0 Then
                                    If ilCrossMid = 0 Then
                                        ilCrossMid = 1
                                    Else
                                        ilCrossMid = 2
                                    End If
                                Else
                                    If llEMidTime = 0 Then
                                        If ilCrossMid = 0 Then
                                            ilCrossMid = 1
                                        Else
                                            ilCrossMid = 2
                                        End If
                                    End If
                                End If
                            Else
                                ilCrossMid = 3
                                Exit For
                            End If
                        End If
                    End If
                Next ilIndex
                For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                    If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                        ilDayIndex = LBound(ilSDDay)
                        If ilCrossMid = 2 Then
                            gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), False, llTime
                            If llTime <> 0 Then
                                If llTime < llSTime Then
                                    llSTime = llTime
                                End If
                                gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), True, llTime
                                If llTime > llETime Then
                                    llETime = llTime
                                End If
                                For ilDay = LBound(tgMRdf(ilRdf).sWkDays, 2) To UBound(tgMRdf(ilRdf).sWkDays, 2) Step 1
                                    If tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) = "Y" Then
                                        ilSDDay(ilDayIndex) = True
                                    End If
                                    ilDayIndex = ilDayIndex + 1
                                Next ilDay
                            End If
                        ElseIf ilCrossMid = 3 Then
                            gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), False, llTime
                            If llTime <> 0 Then
                                '3/19/10
                                'If llTime < llSRdfTime Then
                                If llTime < llSTime Then
                                    '3/19/10
                                    'llSRdfTime = llTime
                                    llSTime = llTime
                                End If
                                llTime = 86399
                                '3/19/10
                                'If llTime > llERdfTime Then
                                If llTime > llETime Then
                                    '3/19/10
                                    'llERdfTime = llTime
                                    llETime = llTime
                                End If
                                ilDayIndex = LBound(ilInputDay)
                                For ilDay = LBound(tgMRdf(ilRdf).sWkDays, 2) To UBound(tgMRdf(ilRdf).sWkDays, 2) Step 1
                                    If tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) = "Y" Then
                                        '3/19/10
                                        'ilInputDay(ilDayIndex) = True
                                        ilSDDay(ilDayIndex) = True
                                    End If
                                    ilDayIndex = ilDayIndex + 1
                                Next ilDay
                                Exit For
                            End If
                        Else
                            gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), False, llTime
                            If llTime < llSTime Then
                                llSTime = llTime
                            End If
                            gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), True, llTime
                            If llTime > llETime Then
                                llETime = llTime
                            End If
                            For ilDay = LBound(tgMRdf(ilRdf).sWkDays, 2) To UBound(tgMRdf(ilRdf).sWkDays, 2) Step 1
                                If tgMRdf(ilRdf).sWkDays(ilIndex, ilDay) = "Y" Then
                                    ilSDDay(ilDayIndex) = True
                                End If
                                ilDayIndex = ilDayIndex + 1
                            Next ilDay
                        End If
                    End If
                Next ilIndex
            'ElseIf ilRdf = 0 Then
            ElseIf ilRdf = -2 Then
                For ilDay = LBound(tlDrf.sDay) To UBound(tlDrf.sDay) Step 1
                    If (tlDrf.sDay(ilDay) = "Y") Then
                        ilSDDay(ilDay) = True
                    End If
                Next ilDay
                gUnpackTimeLong tlDrf.iStartTime(0), tlDrf.iStartTime(1), False, llSTime
                gUnpackTimeLong tlDrf.iEndTime(0), tlDrf.iEndTime(1), True, llETime
                If ((tlDrf.iStartTime2(0) <> 1) Or (tlDrf.iStartTime2(1) <> 0)) Then
                    ilNoTimes = 2
                    gUnpackTimeLong tlDrf.iStartTime2(0), tlDrf.iStartTime2(1), False, llTime
                    If llTime < llSTime Then
                        llSTime = llTime
                    End If
                    gUnpackTimeLong tlDrf.iEndTime2(0), tlDrf.iEndTime2(1), True, llTime
                    If llTime > llETime Then
                        llETime = llTime
                    End If
                End If
            End If
            'Three results: Days match; Extra x days; days don't match
            ilDayMatch = 0
            ilDayIndex = LBound(ilInputDay)
            For ilDay = LBound(ilSDDay) To UBound(ilSDDay) Step 1
                If (ilSDDay(ilDay) = True) And (ilInputDay(ilDayIndex) <> True) Then
                    ilDayMatch = -1
                    Exit For
                ElseIf (ilSDDay(ilDay) <> True) And (ilInputDay(ilDayIndex) = True) Then
                    ilDayMatch = -1
                    Exit For
                End If
                ilDayIndex = ilDayIndex + 1
            Next ilDay
            If ilDayMatch = -1 Then
                ilDayMatch = 0
                ilDayIndex = LBound(ilInputDay)
                For ilDay = LBound(ilSDDay) To UBound(ilSDDay) Step 1
                    If (ilSDDay(ilDay) = True) And (ilInputDay(ilDayIndex) <> True) Then
                        ilDayMatch = ilDayMatch + 1
                    ElseIf (ilSDDay(ilDay) <> True) And (ilInputDay(ilDayIndex) = True) Then
                        ilDayMatch = -1
                        Exit For
                    End If
                    ilDayIndex = ilDayIndex + 1
                Next ilDay
            End If
            
'Debug.Print " ilDayMatch=" & ilDayMatch & " , ilNoTimes=" & ilNoTimes
            If (ilDayMatch <> -1) And (ilNoTimes = 1) Then
                llTimeMatch = -1
                If (llSRdfTime = llSTime) And (llERdfTime = llETime) Then
                    llTimeMatch = 0
                ElseIf (llSTime <= llSRdfTime) And (llETime >= llERdfTime) Then
                    llTimeMatch = (llETime - llSTime) - (llERdfTime - llSRdfTime)
                End If
                If llTimeMatch <> -1 Then
                    'Check if research date exits
                    llTAvgAud = 0
                    '6/4/16: Replaced GoSub
                    'GoSub GetAvgAud
                    mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
                    If llTAvgAud >= 0 Then
                        'Check place to store result
                        If (ilDayMatch = 0) And (llTimeMatch = 0) Then
                            llAvgAud = llTAvgAud
                            '6/4/16: Replaced GoSub
                            'GoSub GetEstAdj
                            mGetEstAdj ilDnfCode, slEstLOrU, slEstPop, slEstAud, ilEstNoWks, llStartDateForEst, llEndDateForEst, hlDrf, hlMnf, hlDpf, ilMnfSocEco, ilMnfDemo, llDemoPop, hlDef, tlDef, ilDefRecLen, llAvgAud, llPopForEst
                            llAvgAud = mRegionAudAdj(hlRaf, llRafCode, llAvgAud)
                            lmAvgAud = llAvgAud
                            lmPopEst = llPopForEst
                            sgAvgAudStamp = slStamp
                            If tlDrf.iRdfCode > 0 Then
                                ilAudFromSource = 4
                                llAudFromCode = tlDrf.iRdfCode
                            Else
                                ilAudFromSource = 5
                                llAudFromCode = tlDrf.lCode
                            End If
                            imAudFromSource = ilAudFromSource
                            lmAudFromCode = llAudFromCode
                            gGetDemoAvgAud = True
                            '10256
                            mSetActOneInfoFromDemoCall tlDrf
'Debug.Print " -gGetDemoExit9"
                            Exit Function
                        Else
                            If (Asc(tgSpf.sOverrideOptions) And BESTFITWEIGHTNONE) <> BESTFITWEIGHTNONE Then
                                If llBestDrfCode = -1 Then
                                    llBestDrfCode = tlDrf.lCode
                                    ilBestDayCount = ilDayMatch
                                    llBestTime = llTimeMatch
                                Else
                                    'If llTimeMatch < llBestTime Then
                                    '    llBestDrfCode = tlDrf.lCode
                                    '    ilBestDayCount = ilDayMatch
                                    '    llBestTime = llTimeMatch
                                    'ElseIf (llTimeMatch = llBestTime) Then
                                    '    If (ilDayMatch < ilBestDayCount) Then
                                    '        llBestDrfCode = tlDrf.lCode
                                    '        ilBestDayCount = ilDayMatch
                                    '        llBestTime = llTimeMatch
                                    '    End If
                                    'End If
                                    If (Asc(tgSpf.sUsingFeatures) And BESTFITWEIGHT) <> BESTFITWEIGHT Then
                                        If llTimeMatch < llBestTime Then
                                            llBestDrfCode = tlDrf.lCode
                                            ilBestDayCount = ilDayMatch
                                            llBestTime = llTimeMatch
                                        ElseIf (llTimeMatch = llBestTime) Then
                                            If (ilDayMatch < ilBestDayCount) Then
                                                llBestDrfCode = tlDrf.lCode
                                                ilBestDayCount = ilDayMatch
                                                llBestTime = llTimeMatch
                                            End If
                                        End If
                                    Else
                                        If ilDayMatch < ilBestDayCount Then
                                            llBestDrfCode = tlDrf.lCode
                                            ilBestDayCount = ilDayMatch
                                            llBestTime = llTimeMatch
                                        ElseIf (ilDayMatch = ilBestDayCount) Then
                                            If (llTimeMatch < llBestTime) Then
                                                llBestDrfCode = tlDrf.lCode
                                                ilBestDayCount = ilDayMatch
                                                llBestTime = llTimeMatch
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        ilRet = btrGetNext(hlDrf, tlDrf, ilDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If llBestDrfCode <> -1 Then
        tlDrfSrchKey2.lCode = llBestDrfCode
'Debug.Print " LOAD DEMO DATA5: (DT) DnfCode=" & ilDnfCode & " , MnfSocEco=" & ilMnfSocEco & " , DRFCode=" & llBestDrfCode

        ilRet = btrGetEqual(hlDrf, tlDrf, ilDrfRecLen, tlDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If (ilRet = BTRV_ERR_NONE) Then
            'If Exist, then see which category it should be saved as
            llTAvgAud = 0
            '6/4/16: Replaced GoSub
            'GoSub GetAvgAud
            mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
            llAvgAud = llTAvgAud
            '6/4/16: Replaced GoSub
            'GoSub GetEstAdj
            mGetEstAdj ilDnfCode, slEstLOrU, slEstPop, slEstAud, ilEstNoWks, llStartDateForEst, llEndDateForEst, hlDrf, hlMnf, hlDpf, ilMnfSocEco, ilMnfDemo, llDemoPop, hlDef, tlDef, ilDefRecLen, llAvgAud, llPopForEst
            llAvgAud = mRegionAudAdj(hlRaf, llRafCode, llAvgAud)
            lmAvgAud = llAvgAud
            lmPopEst = llPopForEst
            sgAvgAudStamp = slStamp
            ilAudFromSource = 4
            llAudFromCode = tlDrf.iRdfCode
            If tlDrf.iRdfCode > 0 Then
                ilAudFromSource = 4
                llAudFromCode = tlDrf.iRdfCode
            Else
                ilAudFromSource = 5
                llAudFromCode = tlDrf.lCode
            End If
            imAudFromSource = ilAudFromSource
            lmAudFromCode = llAudFromCode
            gGetDemoAvgAud = True
            '10256
            mSetActOneInfoFromDemoCall tlDrf
'Debug.Print " -gGetDemoExit10"
            Exit Function
        End If
    End If
    'Test if Drf exist for vehicle specified by vehicle
'Debug.Print " LOAD DEMO DATA6: (DV) DnfCode=" & ilDnfCode & " , MnfSocEco=" & ilMnfSocEco & " , RdfCode=0, Vef=" & ilVefCode
    tlDrfSrchKey.iDnfCode = ilDnfCode
    tlDrfSrchKey.sDemoDataType = "D"    'Demo Data
    tlDrfSrchKey.iMnfSocEco = ilMnfSocEco
    tlDrfSrchKey.iVefCode = ilVefCode
    tlDrfSrchKey.sInfoType = "V"        'Daypart
    tlDrfSrchKey.iRdfCode = 0
    ilRet = btrGetGreaterOrEqual(hlDrf, tlDrf, ilDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    'Bypass Cume
    Do While (tlDrf.sDataType <> slDataType) And (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = ilDnfCode) And (tlDrf.iVefCode = ilVefCode) And (tlDrf.iMnfSocEco = ilMnfSocEco) And (tlDrf.sDemoDataType = "D") And (tlDrf.sInfoType = "V")
        ilRet = btrGetNext(hlDrf, tlDrf, ilDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = ilDnfCode) And (tlDrf.iVefCode = ilVefCode) And (tlDrf.iMnfSocEco = ilMnfSocEco) And (tlDrf.sDemoDataType = "D") And (tlDrf.sInfoType = "V") Then
        'If tlMnf.iGroupNo = 0 Then
        '    If (ilDemoSex = 0) Or (ilDemoSex = 2) Then
        '        For ilLoop = ilStartDemoIndex To ilEndDemoIndex Step 1
        '            llAvgAud = llAvgAud + tlDrf.lDemo(ilLoop)
        '        Next ilLoop
        '    End If
        '    If (ilDemoSex = 1) Or (ilDemoSex = 2) Then
        '        For ilLoop = ilStartDemoIndex + 8 To ilEndDemoIndex + 8 Step 1
        '            llAvgAud = llAvgAud + tlDrf.lDemo(ilLoop)
        '        Next ilLoop
        '    End If
        'Else
        '    llAvgAud = tlDrf.lDemo(tlMnf.iGroupNo)
        'End If
        llTAvgAud = llAvgAud
        '6/4/16: Replaced GoSub
        'GoSub GetAvgAud
        mGetAvgAud tlDrf, ilMnfDemo, hlDpf, tlDpf, ilDpfRecLen, llTAvgAud, tlMnf, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex, ilDemoSex, ilCustomIndex
        llAvgAud = llTAvgAud
        '6/4/16: Replaced GoSub
        'GoSub GetEstAdj
        mGetEstAdj ilDnfCode, slEstLOrU, slEstPop, slEstAud, ilEstNoWks, llStartDateForEst, llEndDateForEst, hlDrf, hlMnf, hlDpf, ilMnfSocEco, ilMnfDemo, llDemoPop, hlDef, tlDef, ilDefRecLen, llAvgAud, llPopForEst
        llAvgAud = mRegionAudAdj(hlRaf, llRafCode, llAvgAud)
        lmAvgAud = llAvgAud
        lmPopEst = llPopForEst
        sgAvgAudStamp = slStamp
        ilAudFromSource = 3
        llAudFromCode = 0
        imAudFromSource = ilAudFromSource
        lmAudFromCode = llAudFromCode
        gGetDemoAvgAud = True
        '10256
        mSetActOneInfoFromDemoCall tlDrf
'Debug.Print " -gGetDemoExit11"
        Exit Function
    End If
    '6/4/16: Replaced GoSub
    'GoSub GetEstAdj
    mGetEstAdj ilDnfCode, slEstLOrU, slEstPop, slEstAud, ilEstNoWks, llStartDateForEst, llEndDateForEst, hlDrf, hlMnf, hlDpf, ilMnfSocEco, ilMnfDemo, llDemoPop, hlDef, tlDef, ilDefRecLen, llAvgAud, llPopForEst
    llAvgAud = mRegionAudAdj(hlRaf, llRafCode, llAvgAud)
    lmAvgAud = llAvgAud
    lmPopEst = llPopForEst
    sgAvgAudStamp = slStamp
    gGetDemoAvgAud = True
    '10256
    gSetAct1ToBlank
'Debug.Print " -gGetDemoExit12"
    Exit Function
'GetAvgAud:
'    If lgDpfNoRecs > 0 Then
'        tlDpfSrchKey1.lDrfCode = tlDrf.lCode
'        tlDpfSrchKey1.iMnfDemo = ilMnfDemo
'        ilRet = btrGetEqual(hlDpf, tlDpf, ilDpfRecLen, tlDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
'            llTAvgAud = llTAvgAud + tlDpf.lDemo
'            Return
'        End If
'    End If
'    If tlMnf.iGroupNo = 0 Then
'        If ilStartDemoIndex < 0 Then
'            mComputeDemoIndex tlDrf.sForm, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex
'        End If
'        If (ilStartDemoIndex <> -1) And (ilEndDemoIndex <> -1) Then
'            If (ilDemoSex = 0) Or (ilDemoSex = 2) Then
'                For ilLoop = ilStartDemoIndex To ilEndDemoIndex Step 1
'                    llTAvgAud = llTAvgAud + tlDrf.lDemo(ilLoop)
'                Next ilLoop
'            End If
'            If (ilDemoSex = 1) Or (ilDemoSex = 2) Then
'                If tlDrf.sForm <> "8" Then
'                    For ilLoop = ilStartDemoIndex + 8 To ilEndDemoIndex + 8 Step 1
'                        llTAvgAud = llTAvgAud + tlDrf.lDemo(ilLoop)
'                    Next ilLoop
'                Else
'                    For ilLoop = ilStartDemoIndex + 9 To ilEndDemoIndex + 9 Step 1
'                        llTAvgAud = llTAvgAud + tlDrf.lDemo(ilLoop)
'                    Next ilLoop
'                End If
'            End If
'        End If
'    Else
'        llTAvgAud = llTAvgAud + tlDrf.lDemo(ilCustomIndex)
'    End If
'    Return
'GetEstAdj:
'    If tgSpf.sDemoEstAllowed <> "Y" Then
'        Return
'    End If
'    slEstLOrU = mGetBookEstByLOrU(ilDnfCode)
'    If slEstLOrU <> "U" Then
'        tgSpf.sDemoEstAllowed = "N"
'    End If
'    ilRet = gGetDemoPop(hlDrf, hlMnf, hlDpf, ilDnfCode, ilMnfSocEco, ilMnfDemo, llDemoPop)
'    tgSpf.sDemoEstAllowed = "Y"
'    slEstPop = "0"
'    slEstAud = "0"
'    ilEstNoWks = 0
'    For llDate = llStartDateForEst To llEndDateForEst Step 7
'        ilEstNoWks = ilEstNoWks + 1
'        tlDefSrchKey1.iDnfCode = ilDnfCode
'        gPackDateLong llDate, tlDefSrchKey1.iStartDate(0), tlDefSrchKey1.iStartDate(1)
'        ilRet = btrGetLessOrEqual(hlDef, tlDef, ilDefRecLen, tlDefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
'        If (ilRet = BTRV_ERR_NONE) And (tlDef.iDnfCode = ilDnfCode) Then
'            'clEstAud = clEstAud + (CCur(llAvgAud) * tlDef.lEstimatePct)
'            'clEstPop = clEstPop + (CCur(llDemoPop) * tlDef.lEstimatePct)
'            slEstAud = gAddStr(slEstAud, gMulStr(str$(llAvgAud), str(tlDef.lEstimatePct)))  '(CCur(llAvgAud) * tlDef.lEstimatePct)
'            If slEstLOrU <> "U" Then
'                slEstPop = gAddStr(slEstPop, gMulStr(str$(llDemoPop), str(tlDef.lEstimatePct)))  '+ (CCur(llDemoPop) * tlDef.lEstimatePct)
'            End If
'        Else
'            'clEstAud = clEstAud + (CCur(llAvgAud) * 10000)
'            'clEstPop = clEstPop + CCur(llDemoPop) * 10000
'            slEstAud = gAddStr(slEstAud, gMulStr(str$(llAvgAud), "10000"))  '(CCur(llAvgAud) * tlDef.lEstimatePct)
'            If slEstLOrU <> "U" Then
'                slEstPop = gAddStr(slEstPop, gMulStr(str$(llDemoPop), "10000"))  '+ (CCur(llDemoPop) * tlDef.lEstimatePct)
'            End If
'        End If
'    Next llDate
'    If ilEstNoWks > 0 Then
'        'llAvgAud = ((clEstAud + 50) / ilEstNoWks) \ 10000
'        'llPopForEst = ((clEstPop + 50) / ilEstNoWks) \ 10000
'        llAvgAud = Val(gDivStr(gDivStr(slEstAud, str$(ilEstNoWks)), "10000"))  '((clEstAud + 50) / ilEstNoWks) \ 10000
'        If slEstLOrU <> "U" Then
'            llPopForEst = Val(gDivStr(gDivStr(slEstPop, str$(ilEstNoWks)), "10000"))  '((clEstPop + 50) / ilEstNoWks) \ 10000
'        Else
'            llPopForEst = llDemoPop
'        End If
'    Else
'        llAvgAud = llAvgAud
'        llPopForEst = llDemoPop
'    End If
'    Return
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGetDemoPop                     *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Population for Week        *
'*                                                     *
'*******************************************************
Function gGetDemoPop(hlDrf As Integer, hlMnf As Integer, hlDpf As Integer, ilInDnfCode As Integer, ilMnfSocEco As Integer, ilMnfDemo As Integer, llPop As Long) As Integer
'
'   ilRet = gGetDemoPop(hlDrf, hlMnf, ilDnfCode, ilMnfSocEco, ilMnfDemo, llPop)
'   Where:
'       hlDrf(I)- Handle open to DRF
'       hlMnf(I)- Handle to open MNF
'       ilDnfCode(I)- Rating book code
'       ilMnfSocEco(I)- Social Ecomony MultiName Code or Zero
'       ilMnfDemo(I)- Demo MultiName Code (M18-34)
'       llPop(O)- Population
'
'
    Dim ilDrfRecLen As Integer     'Record length
    Dim tlDrf As DRF
    Dim tlDrfSrchKey As DRFKEY0
    Dim ilDpfRecLen As Integer     'Record length
    Dim tlDpf As DPF
    Dim tlDpfSrchKey2 As DPFKEY2
    Dim ilMnfRecLen As Integer     'Record length
    Dim tlMnf As MNF
    Dim tlMnfSrchKey As INTKEY0
    Dim ilDemoSex As Integer    'Demo Sex: 0=Man; 1=Women; 2= Both
    Dim ilDemoSAge As Integer   'Demo Start Age
    Dim ilDemoEAge As Integer   'Demo End Age
    Dim ilStartDemoIndex As Integer
    Dim ilEndDemoIndex As Integer
    Dim illoop As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilDnfCode As Integer
    Dim slDataType As String    'A=Broadcast data for standard Demo Data; B=Broadcast data for custom demos(1-18); Chr(127) for 19-36; Chr(128) for 37-54;....
    Dim ilCustomIndex As Integer
    'Dim ilUsingSourcePop As Integer

'   Compute from base estimate
'    'If using Research Estimate, bypass getting Pop.  It will be obtained in gGetDemoAvgAud
'    If tgSpf.sDemoEstAllowed = "Y" Then
'        'Don't set or alter llPop
'        gGetDemoPop = True
'        Exit Function
'    End If

    slStamp = gFileDateTime(sgDBPath & "Drf.btr") & Trim$(str$(ilInDnfCode)) & Trim$(str$(ilMnfSocEco)) & Trim$(str$(ilMnfDemo))
    If slStamp = smPopStamp Then
        llPop = lmPop
        gGetDemoPop = True
        Exit Function
    End If
    'ilUsingSourcePop = False
    ilRet = gObtainPopDnfInfo()
    ilDnfCode = ilInDnfCode
    ilRet = gBinarySearchPopDnfInfo(ilDnfCode)
    'If ilRet > 0 Then
    If ilRet >= 0 Then
        If ilMnfSocEco <= 0 Then
            If tgPopDnfInfo(ilRet).iPopDnfCode > 0 Then
                ilDnfCode = tgPopDnfInfo(ilRet).iPopDnfCode
                'ilUsingSourcePop = True
            End If
        Else
            If tgPopDnfInfo(ilRet).iQualPopDnfCode > 0 Then
                ilDnfCode = tgPopDnfInfo(ilRet).iQualPopDnfCode
                'ilUsingSourcePop = True
            End If
        End If
    End If
    llPop = 0
    ilDrfRecLen = Len(tlDrf) 'btrRecordLength(hlDrf)  'Get and save record length
    ilDpfRecLen = Len(tlDpf) 'btrRecordLength(hlpf)  'Get and save record length
    ilMnfRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
    If lgDpfNoRecs = 0 Then
        lgDpfNoRecs = btrRecords(hlDpf)
        If lgDpfNoRecs = 0 Then
            lgDpfNoRecs = -1
        End If
    End If
    tlMnfSrchKey.iCode = ilMnfDemo
    ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        gGetDemoPop = False
        Exit Function
    End If

    'Determine which demos make-up demos
    'Standard 16(M12-17;M18-24;M25-34;M35-44;M45-49;M50-54;M55-64;M65+
    '            W12-17;W18-24;W25-34;W35-44;W45-49;W50-54;W55-64;W65+)
    'MnfDemo names: Mxx-xx; Pxx-xx; Axx-xx; Wxx-xx
    'Parse fields

    If tlMnf.iGroupNo <= 0 Then
        If Left$(tlMnf.sName, 1) = "M" Then
            ilDemoSex = 0
        ElseIf Left$(tlMnf.sName, 1) = "W" Then
            ilDemoSex = 1
        Else
            ilDemoSex = 2
        End If
        ilDemoSAge = Val(Mid$(tlMnf.sName, 2, 2))
        If Mid$(tlMnf.sName, 4, 1) = "+" Then
            ilDemoEAge = 99
        Else
            ilDemoEAge = Val(Mid$(tlMnf.sName, 5, 2))
        End If
'        If ilDemoSAge = 12 Then
'            ilStartDemoIndex = 1
'        ElseIf ilDemoSAge = 18 Then
'            ilStartDemoIndex = 2
'        ElseIf ilDemoSAge = 25 Then
'            ilStartDemoIndex = 3
'        ElseIf ilDemoSAge = 35 Then
'            ilStartDemoIndex = 4
'        ElseIf ilDemoSAge = 45 Then
'            ilStartDemoIndex = 5
'        ElseIf ilDemoSAge = 50 Then
'            ilStartDemoIndex = 6
'        ElseIf ilDemoSAge = 55 Then
'            ilStartDemoIndex = 7
'        Else
'            ilStartDemoIndex = 8
'        End If
'
'        If ilDemoEAge = 17 Then
'            ilEndDemoIndex = 1
'        ElseIf ilDemoEAge = 24 Then
'            ilEndDemoIndex = 2
'        ElseIf ilDemoEAge = 34 Then
'            ilEndDemoIndex = 3
'        ElseIf ilDemoEAge = 44 Then
'            ilEndDemoIndex = 4
'        ElseIf ilDemoEAge = 49 Then
'            ilEndDemoIndex = 5
'        ElseIf ilDemoEAge = 54 Then
'            ilEndDemoIndex = 6
'        ElseIf ilDemoEAge = 64 Then
'            ilEndDemoIndex = 7
'        Else
'            ilEndDemoIndex = 8
'        End If
    Else
        gConvCustomGroup tlMnf.iGroupNo, slDataType, ilCustomIndex
    End If

    'Get Population
    tlDrfSrchKey.iDnfCode = ilDnfCode
    tlDrfSrchKey.sDemoDataType = "P"
    'Restore using the Soc Eco Population- ABC request 5/16/02
    'tlDrfSrchKey.iMnfSocEco = 0 'ilMnfSocEco
    tlDrfSrchKey.iMnfSocEco = ilMnfSocEco
    tlDrfSrchKey.iVefCode = 0
    tlDrfSrchKey.sInfoType = ""
    tlDrfSrchKey.iRdfCode = 0
    ilRet = btrGetGreaterOrEqual(hlDrf, tlDrf, ilDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = ilDnfCode) And (tlDrf.iVefCode = 0) And (tlDrf.sDemoDataType = "P")
        If ((tlMnf.iGroupNo <= 0) And (tlDrf.sDataType = "A")) Or ((tlMnf.iGroupNo > 0) And (tlDrf.sDataType = slDataType)) Then
            'If (lgDpfNoRecs > 0) And (mGetBookEstByLOrU(ilDnfCode) <> "U") And (Not ilUsingSourcePop) Then
            If (lgDpfNoRecs > 0) And (mGetBookEstByLOrU(ilDnfCode) <> "U") Then
                tlDpfSrchKey2.iDnfCode = ilDnfCode
                tlDpfSrchKey2.iMnfDemo = ilMnfDemo
                ilRet = btrGetEqual(hlDpf, tlDpf, ilDpfRecLen, tlDpfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    llPop = tlDpf.lPop
                    smPopStamp = slStamp
                    lmPop = llPop
                    gGetDemoPop = True
                    Exit Function
                End If
            End If
            If tlMnf.iGroupNo = 0 Then
                llPop = 0
                mComputeDemoIndex tlDrf.sForm, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex
                If (ilStartDemoIndex <> -1) And (ilEndDemoIndex <> -1) Then
                    If LBound(tlDrf.lDemo) = 0 Then
                        ilStartDemoIndex = ilStartDemoIndex - 1
                        ilEndDemoIndex = ilEndDemoIndex - 1
                    End If
                    If (ilDemoSex = 0) Or (ilDemoSex = 2) Then
                        For illoop = ilStartDemoIndex To ilEndDemoIndex Step 1
                            llPop = llPop + tlDrf.lDemo(illoop)
                        Next illoop
                    End If
                    If (ilDemoSex = 1) Or (ilDemoSex = 2) Then
                        If tlDrf.sForm <> "8" Then
                            For illoop = ilStartDemoIndex + 8 To ilEndDemoIndex + 8 Step 1
                                llPop = llPop + tlDrf.lDemo(illoop)
                            Next illoop
                        Else
                            For illoop = ilStartDemoIndex + 9 To ilEndDemoIndex + 9 Step 1
                                llPop = llPop + tlDrf.lDemo(illoop)
                            Next illoop
                        End If
                    End If
                End If
            Else
                llPop = tlDrf.lDemo(ilCustomIndex - 1)
            End If
            smPopStamp = slStamp
            lmPop = llPop
            gGetDemoPop = True
            Exit Function
        End If
        ilRet = btrGetNext(hlDrf, tlDrf, ilDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    gGetDemoPop = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gResearchTotals                 *
'*                                                     *
'*             Created:10/22/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute Rating, Gross Impression*
'*                     Gross Rating points             *
'*                     AvgRating=Sum(LnAvgRating)/#Line*
'*                     CntrGrImp=Sum(LnGrImp)          *
'*                     CntrGRP=Sum(LnGRP)              *
'*                     CPP=Sum(LnCost)/CntrGRP         *
'*                     CPM=1000*Sum(LnCost)/CntrGrImp  *
'*                                                     *
'*******************************************************
'Sub gResearchTotals (ilAllowNeg As Integer, llPop As Long, llLnCost() As Long, ilLnAvgRating() As Integer, llLnGrImp() As Long, llLnGRP() As Long, llTotalCost As Long, ilAvgRating As Integer, llGrImp As Long, llGRP As Long, llCPP As Long, llCPM As Long)
'Sub gResearchTotals(sl1or2PlaceRating As String, ilAllowNeg As Integer, llPop As Long, llLnCost() As Long, llLnGrImp() As Long, llLnGRP() As Long, llLnSpots As Long, llTotalCost As Long, ilAvgRating As Integer, llGrImp As Long, llGRP As Long, llCPP As Long, llCPM As Long, llAvgAud As Long)
Sub gResearchTotals(sl1or2PlaceRating As String, ilAllowNeg As Integer, llPop As Long, llLnCost() As Long, llLnGrImp() As Long, llLnGRP() As Long, llLnSpots As Long, dlTotalCost As Double, ilAvgRating As Integer, llGrImp As Long, llGRP As Long, llCPP As Long, llCPM As Long, llAvgAud As Long) 'TTP 10439 - Rerate 21,000,000
'
''   gResearchTotals ilAllowNeg, llPop, llLnCost(), ilLnAvgRating(), llLnGrImp(), llLnGRP(), llTotalCost, ilAvgRating, llGrImp, llGrp, llCPP, llCPM
'   gResearchTotals sl1or2PlaceRating, ilAllowNeg, llPop, llLnCost(), llLnGrImp(), llLnGRP(), llLnSpots, llTotalCost, ilAvgRating, llGrImp, llGrp, llCPP, llCPM, llAvgAud
'   Where:
'       sl1or2PlaceRating(I)- 1 or 2 place rating. Test for 2.
'       ilAllowNeg(I)- True if neg spots and dollars allowed
'       llPop(I)- Population (can be zero- value computed from formulas shown above)
'       llLnCost()(I)- Line cost (size of array must be for the exact number of lines to compute totals for (i.e. no cancel before start lines)
''       ilLnAvgRating()(I)- Line Average Rating- This vaue is not used- replaced 4/5/99 with GrImp/#Spots
'       llLnGrImp()(I)- Line Gross Impressions
'       llLnGRP()(I)- Line Gross Rating Points
'       llLnSpots(I)- Total number of spots across lines (using total because packages)- added 4/5/99
'       dlTotalCost(O)- Total contract cost.  Two decimal places
'       ilAvgRating(O)- Average rating for contract.  One decimal place
'       llGrImp(O)- Gross Impressions for contract.  Whole # (divide by 1000)
'       llGRP(O)- Gross Rating points for contract.  One decimal place
'       llCPP(O)- CPP for contract. Whole number (no decimal places)
'       llCPM(O)- CPM for contract. Two decimal places
'       llAvgAud(O)- Average audience for contract.  Whole number (no decimal point)
'
    Dim illoop As Integer
    Dim llAvgRating As Long
    Dim llCPMAdj As Long
    Dim llTempGRP As Long
    
    On Error GoTo TotalErr:
    dlTotalCost = 0
    ilAvgRating = 0
    llGrImp = 0
    llGRP = 0
    llCPP = 0
    llCPM = 0
    llAvgAud = 0
'    If tgSpf.sSAudData <> "H" Then
'        llCPMAdj = 1
'    Else
'        llCPMAdj = 10
'    End If
    If tgSpf.sSAudData = "H" Then
        llCPMAdj = 10
    ElseIf tgSpf.sSAudData = "N" Then
        llCPMAdj = 100
    ElseIf tgSpf.sSAudData = "U" Then
        llCPMAdj = 1000
    Else
        llCPMAdj = 1
    End If
    For illoop = LBound(llLnCost) To UBound(llLnCost) Step 1
        dlTotalCost = dlTotalCost + llLnCost(illoop) 'TTP 10439 - Rerate 21,000,000
        'llAvgRating = llAvgRating + ilLnAvgRating(ilLoop)
        llGrImp = llGrImp + llLnGrImp(illoop)
        llGRP = llGRP + llLnGRP(illoop)
    Next illoop
    'ilAvgRating = llAvgRating / (UBound(llLnCost) - LBound(llLnCost) + 1)
    If llLnSpots > 0 Then
        'ilAvgRating = llGRP / llLnSpots
        If llPop > 0 Then
            If tgSpf.sSGRPCPPCal = "R" Then
                ilAvgRating = llGRP / llLnSpots
            ElseIf tgSpf.sSGRPCPPCal = "G" Then
                ilAvgRating = ((1000 * CSng(llGrImp)) / llPop) / llLnSpots
            Else
                If sl1or2PlaceRating <> "2" Then
                    ilAvgRating = ((1000 * CSng(llGrImp)) / llPop) / llLnSpots
                Else
                    ilAvgRating = ((10000 * CSng(llGrImp)) / llPop) / llLnSpots
                End If
            End If
        Else
            '10/27/14: 2 places
            If (sl1or2PlaceRating <> "2") Or (tgSpf.sSGRPCPPCal <> "A") Then
                ilAvgRating = llGRP / llLnSpots
            Else
                ilAvgRating = (10 * llGRP) / llLnSpots
            End If
        End If
        llAvgAud = llGrImp / llLnSpots
    Else
        ilAvgRating = 0
        llAvgAud = 0
    End If
    If llPop > 0 Then
        If tgSpf.sSGRPCPPCal = "R" Then
            'Computed above
        ElseIf tgSpf.sSGRPCPPCal = "G" Then
            llTempGRP = (10000 * CSng(llGrImp)) / llPop
            'llGRP = llTempGRP / 10
            llGRP = (1000 * CSng(llGrImp)) / llPop
        Else
            llGRP = (1000 * CSng(llGrImp)) / llPop
        End If
        If llGrImp > 0 Then
            If tgSpf.sSGRPCPPCal = "R" Then
                If llGRP > 0 Then
                    llCPP = (CSng(dlTotalCost) / 10) / llGRP   'TTP 10439 - Rerate 21,000,000
                Else
                    llCPP = 0
                End If
            ElseIf tgSpf.sSGRPCPPCal = "G" Then
                If llTempGRP > 0 Then  'llGRP > 0 Then
                    'llCPP = (CSng(llTotalCost) / 10) / llGRP   '(llTotalCost / llGRP) / 10
                    llCPP = (CSng(dlTotalCost)) / llTempGRP   'TTP 10439 - Rerate 21,000,000
                Else
                    llCPP = 0
                End If
            Else
                llCPP = ((CSng(dlTotalCost) / 10000) * llPop) / llGrImp 'TTP 10439 - Rerate 21,000,000
            End If
            llCPM = (llCPMAdj * CSng(dlTotalCost)) / llGrImp 'TTP 10439 - Rerate 21,000,000
        ElseIf ilAllowNeg And (llGrImp < 0) Then
            If dlTotalCost > 0 Then 'TTP 10439 - Rerate 21,000,000
                If tgSpf.sSGRPCPPCal = "R" Then
                    If llGRP > 0 Then
                        llCPP = (CSng(dlTotalCost) / 10) / llGRP   'TTP 10439 - Rerate 21,000,000
                    Else
                        llCPP = 0
                    End If
                ElseIf tgSpf.sSGRPCPPCal = "G" Then
                    If llTempGRP > 0 Then  'llGRP > 0 Then
                        'llCPP = (CSng(llTotalCost) / 10) / llGRP   '(CSng(llTotalCost) / llGRP) / 10
                        llCPP = (CSng(dlTotalCost)) / llTempGRP   'TTP 10439 - Rerate 21,000,000
                    Else
                        llCPP = 0
                    End If
                Else
                    llCPP = ((CSng(dlTotalCost) / 10000) * llPop) / llGrImp 'TTP 10439 - Rerate 21,000,000
                End If
                llCPM = (llCPMAdj * CSng(dlTotalCost)) / llGrImp 'TTP 10439 - Rerate 21,000,000
            Else
                If tgSpf.sSGRPCPPCal = "R" Then
                    If llGRP > 0 Then
                        llCPP = (CSng(-dlTotalCost) / 10) / llGRP   'TTP 10439 - Rerate 21,000,000
                    Else
                        llCPP = 0
                    End If
                ElseIf tgSpf.sSGRPCPPCal = "G" Then
                    If llTempGRP > 0 Then   'llGRP > 0 Then
                        'llCPP = (CSng(-llTotalCost) / 10) / llGRP   '(CSng(-llTotalCost) / llGRP) / 10
                        llCPP = (CSng(-dlTotalCost)) / llTempGRP   'TTP 10439 - Rerate 21,000,000
                    Else
                        llCPP = 0
                    End If
                Else
                    llCPP = ((CSng(-dlTotalCost) / 10000) * llPop) / llGrImp 'TTP 10439 - Rerate 21,000,000
                End If
                llCPM = (-llCPMAdj * CSng(dlTotalCost)) / llGrImp 'TTP 10439 - Rerate 21,000,000
            End If
        Else
            llCPP = 0
            llCPM = 0
        End If
    Else
        If llGRP > 0 Then
            llCPP = (CSng(dlTotalCost) / 10) / llGRP   'TTP 10439 - Rerate 21,000,000
        ElseIf ilAllowNeg And (llGRP < 0) Then
            If dlTotalCost > 0 Then
                llCPP = (CSng(dlTotalCost) / 10) / llGRP   'TTP 10439 - Rerate 21,000,000
            Else
                llCPP = (CSng(-dlTotalCost) / 10) / llGRP   'TTP 10439 - Rerate 21,000,000
            End If
        Else
            llCPP = 0
        End If
        If llGrImp > 0 Then
            llCPM = (llCPMAdj * CSng(dlTotalCost)) / llGrImp 'TTP 10439 - Rerate 21,000,000
        ElseIf ilAllowNeg And (llGrImp < 0) Then
            If dlTotalCost > 0 Then 'TTP 10439 - Rerate 21,000,000
                llCPM = (llCPMAdj * CSng(dlTotalCost)) / llGrImp 'TTP 10439 - Rerate 21,000,000
            Else
                llCPM = (-llCPMAdj * CSng(dlTotalCost)) / llGrImp 'TTP 10439 - Rerate 21,000,000
            End If
        Else
            llCPM = 0
        End If
    End If
    Exit Sub
TotalErr:
#If programmatic <> 1 Then
    MsgBox "Combination of Price, Spots and Audience Exceeded 21,000,000", vbOKOnly + vbExclamation, "Warning"
#End If
    dlTotalCost = 0
    ilAvgRating = 0
    llAvgAud = 0
    llGrImp = 0
    llGRP = 0
    llCPP = 0
    llCPM = 0
    
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gResearchTotalsFloat            *
'*                                                     *
'*            Created:10/22/93      By:D. LeVine       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute Rating, Gross Impression*
'*                     Gross Rating points             *
'*                     AvgRating=Sum(LnAvgRating)/#Line*
'*                     CntrGrImp=Sum(LnGrImp)          *
'*                     CntrGRP=Sum(LnGRP)              *
'*                     CPP=Sum(LnCost)/CntrGRP         *
'*                     CPM=1000*Sum(LnCost)/CntrGrImp  *
''
''          Return dlTotalCost as negative if over
'*          $21 mill to indicate overflow; otherwise
'*          return the true total Cost
'*                                                     *
'*       8/14/98 - Duplicate gResearchtotals and make  *
'*                 dlTotalCost to a Floating point type*
'*******************************************************
''Sub gResearchTotalsFloat (ilAllowNeg As Integer, llPop As Long, llLnCost() As Long, ilLnAvgRating() As Integer, llLnGrImp() As Long, llLnGRP() As Long, llTotalCost As Long, ilAvgRating As Integer, llGrImp As Long, llGRP As Long, llCPP As Long, llCPM As Long)
'Sub gResearchTotalsFloat(sl1or2PlaceRating As String, ilAllowNeg As Integer, llPop As Long, llLnCost() As Long, llLnGrImp() As Long, llLnGRP() As Long, llLnSpots As Long, llTotalCost As Long, ilAvgRating As Integer, llGrImp As Long, llGRP As Long, llCPP As Long, llCPM As Long, llAvgAud As Long)
'Sub gResearchTotalsFloat(sl1or2PlaceRating As String, ilAllowNeg As Integer, llPop As Long, flLnCost() As Single, llLnGrImp() As Long, llLnGRP() As Long, llLnSpots As Long, llTotalCost As Long, ilAvgRating As Integer, llGrImp As Long, llGRP As Long, llCPP As Long, llCPM As Long, llAvgAud As Long, Optional flTotalCost As Single)
Sub gResearchTotalsFloat(sl1or2PlaceRating As String, ilAllowNeg As Integer, llPop As Long, dlLnCost() As Double, llLnGrImp() As Long, llLnGRP() As Long, llLnSpots As Long, dlTotalCost As Double, ilAvgRating As Integer, llGrImp As Long, llGRP As Long, llCPP As Long, llCPM As Long, llAvgAud As Long, Optional flTotalCost As Single)
'
''   gResearchTotalsFloat ilAllowNeg, llPop, llLnCost(), ilLnAvgRating(), llLnGrImp(), llLnGRP(), llTotalCost, ilAvgRating, llGrImp, llGrp, llCPP, llCPM
'   gResearchTotalsFloat sl1or2PlaceRating, ilAllowNeg, llPop, llLnCost(), llLnGrImp(), llLnGRP(), llLnSpots, llTotalCost, ilAvgRating, llGrImp, llGrp, llCPP, llCPM, llAvgAud
'   Where:
'       sl1or2PlaceRating(I)- 1 or 2 place rating. Test for 2.
'       ilAllowNeg(I)- True if neg spots and dollars allowed
'       llPop(I)- Population (can be zero- value computed from formulas shown above)
'       llLnCost()(I)- Line cost (size of array must be for the exact number of lines to compute totals for (i.e. no cancel before start lines)
''       ilLnAvgRating()(I)- Line Average Rating
'       llLnGrImp()(I)- Line Gross Impressions
'       llLnGRP()(I)- Line Gross Rating Points
'       llLnSpots(I)- Total number of spots across lines (using total because packages)- added 4/5/99
'       dlTotalCost(O)- Total contract cost.  Two decimal places
'       ilAvgRating(O)- Average rating for contract.  One or Two decimal place 'Whole # (divide by 1000)
'       llGrImp(O)- Gross Impressions for contract.  Whole # (divide by 1000)
'       llGRP(O)- Gross Rating points for contract.  One decimal place
'       llCPP(O)- CPP for contract. Whole number (no decimal places)
'       llCPM(O)- CPM for contract. Two decimal places
'       llAvgAud(O)- Average audience for contract.  Whole number (no decimal point)
'
    Dim illoop As Integer
    Dim llAvgRating As Long
    Dim llCPMAdj As Long
    Dim llTempGRP As Long
    'Dim flTotalCost As Single       'TotalCost calc field to handle total over 21 mil
    Dim flGrimp As Single
    
    On Error GoTo TotalErr:
    dlTotalCost = 0 'TTP 10439 - Rerate 21,000,000
    'flTotalCost = 0'TTP 10439 - Rerate 21,000,000
    flGrimp = 0
    ilAvgRating = 0
    llGrImp = 0
    llGRP = 0
    llCPP = 0
    llCPM = 0
    llAvgAud = 0
'    If tgSpf.sSAudData <> "H" Then
'        llCPMAdj = 1
'    Else
'        llCPMAdj = 10
'    End If
    If tgSpf.sSAudData = "H" Then
        llCPMAdj = 10
    ElseIf tgSpf.sSAudData = "N" Then
        llCPMAdj = 100
    ElseIf tgSpf.sSAudData = "U" Then
        llCPMAdj = 1000
    Else
        llCPMAdj = 1
    End If
    'For ilLoop = LBound(llLnCost) To UBound(llLnCost) Step 1
    'For illoop = LBound(flLnCost) To UBound(flLnCost) Step 1
    For illoop = LBound(dlLnCost) To UBound(dlLnCost) Step 1 'TTP 10439 - Rerate 21,000,000
        'flTotalCost = flTotalCost + llLnCost(ilLoop)
        'flTotalCost = flTotalCost + flLnCost(illoop)
        dlTotalCost = dlTotalCost + dlLnCost(illoop) 'TTP 10439 - Rerate 21,000,000
        'llAvgRating = llAvgRating + ilLnAvgRating(ilLoop)
        flGrimp = flGrimp + llLnGrImp(illoop)
        llGRP = llGRP + llLnGRP(illoop)
    Next illoop
    'ilAvgRating = llAvgRating / (UBound(llLnCost) - LBound(llLnCost) + 1)
    If llLnSpots > 0 Then
        '10/27/14: 2 places
        If (sl1or2PlaceRating <> "2") Or (tgSpf.sSGRPCPPCal <> "A") Then
            ilAvgRating = llGRP / llLnSpots
        Else
            ilAvgRating = (10 * llGRP) / llLnSpots
        End If
        llAvgAud = flGrimp / llLnSpots
    Else
        ilAvgRating = 0
        llAvgAud = 0
    End If
    If llPop > 0 Then
        If tgSpf.sSGRPCPPCal = "R" Then
            'Computed above
        ElseIf tgSpf.sSGRPCPPCal = "G" Then
            llTempGRP = (10000 * CSng(flGrimp)) / llPop
            'llGRP = llTempGRP / 10
            llGRP = (1000 * CSng(flGrimp)) / llPop
        Else
            llGRP = (1000 * CSng(flGrimp)) / llPop
        End If
        If flGrimp > 0 Then
            If tgSpf.sSGRPCPPCal = "R" Then
                If llGRP > 0 Then
                    'llCPP = (flTotalCost) / 10 / llGRP   '(llTotalCost / llGRP) / 10
                    llCPP = (dlTotalCost) / 10 / llGRP   'TTP 10439 - Rerate 21,000,000
                Else
                    llCPP = 0
                End If
            ElseIf tgSpf.sSGRPCPPCal = "G" Then
                If llTempGRP > 0 Then  'llGRP > 0 Then
                    'llCPP = (CSng(llTotalCost) / 10) / llGRP   '(llTotalCost / llGRP) / 10
                    'llCPP = flTotalCost / llTempGRP  '(llTotalCost / llGRP) / 10
                    llCPP = dlTotalCost / llTempGRP 'TTP 10439 - Rerate 21,000,000
                Else
                    llCPP = 0
                End If
            Else
                'llCPP = ((flTotalCost) / 10000 * llPop) / flGrimp
                llCPP = ((dlTotalCost) / 10000 * llPop) / flGrimp 'TTP 10439 - Rerate 21,000,000
            End If
            'llCPM = (llCPMAdj * flTotalCost) / flGrimp
            llCPM = (llCPMAdj * dlTotalCost) / flGrimp 'TTP 10439 - Rerate 21,000,000
        ElseIf ilAllowNeg And (flGrimp < 0) Then
            'If flTotalCost > 0 Then
            If dlTotalCost > 0 Then 'TTP 10439 - Rerate 21,000,000
                If tgSpf.sSGRPCPPCal = "R" Then
                    If llGRP > 0 Then
                        'llCPP = flTotalCost / 10 / llGRP   '(CSng(llTotalCost) / llGRP) / 10
                        llCPP = flTotalCost / 10 / llGRP 'TTP 10439 - Rerate 21,000,000
                    Else
                        llCPP = 0
                    End If
                ElseIf tgSpf.sSGRPCPPCal = "G" Then
                    If llTempGRP > 0 Then  'llGRP > 0 Then
                        'llCPP = (CSng(llTotalCost) / 10) / llGRP   '(CSng(llTotalCost) / llGRP) / 10
                        'llCPP = flTotalCost / llTempGRP   '(llTotalCost / llGRP) / 10
                        llCPP = dlTotalCost / llTempGRP 'TTP 10439 - Rerate 21,000,000
                    Else
                        llCPP = 0
                    End If
                Else
                    'llCPP = (flTotalCost / 10000 * llPop) / flGrimp
                    llCPP = (dlTotalCost / 10000 * llPop) / flGrimp 'TTP 10439 - Rerate 21,000,000
                End If
                'llCPM = (llCPMAdj * flTotalCost) / flGrimp
                llCPM = (llCPMAdj * dlTotalCost) / flGrimp 'TTP 10439 - Rerate 21,000,000
            Else
                If tgSpf.sSGRPCPPCal = "R" Then
                    If llGRP > 0 Then
                        'llCPP = (-flTotalCost / 10) / llGRP   '(CSng(-llTotalCost) / llGRP) / 10
                        llCPP = (-dlTotalCost / 10) / llGRP 'TTP 10439 - Rerate 21,000,000
                    Else
                        llCPP = 0
                    End If
                ElseIf tgSpf.sSGRPCPPCal = "G" Then
                    If llTempGRP > 0 Then   'llGRP > 0 Then
                        'llCPP = (CSng(-llTotalCost) / 10) / llGRP   '(CSng(-llTotalCost) / llGRP) / 10
                        'llCPP = (-flTotalCost) / llTempGRP   '(llTotalCost / llGRP) / 10
                        llCPP = (-dlTotalCost) / llTempGRP 'TTP 10439 - Rerate 21,000,000
                    Else
                        llCPP = 0
                    End If
                Else
                    'llCPP = ((-flTotalCost / 10000) * llPop) / flGrimp
                    llCPP = ((-dlTotalCost / 10000) * llPop) / flGrimp 'TTP 10439 - Rerate 21,000,000
                End If
                'llCPM = (-llCPMAdj * flTotalCost) / flGrimp
                llCPM = (-llCPMAdj * dlTotalCost) / flGrimp 'TTP 10439 - Rerate 21,000,000
            End If
        Else
            llCPP = 0
            llCPM = 0
        End If
    Else
        If llGRP > 0 Then
            'llCPP = (flTotalCost / 10) / llGRP   '(llTotalCost / llGRP) / 10
            llCPP = (dlTotalCost / 10) / llGRP 'TTP 10439 - Rerate 21,000,000
        ElseIf ilAllowNeg And (llGRP < 0) Then
            'If flTotalCost > 0 Then
            If dlTotalCost > 0 Then 'TTP 10439 - Rerate 21,000,000
                'llCPP = (flTotalCost / 10) / llGRP   '(llTotalCost / llGRP) / 10
                llCPP = (dlTotalCost / 10) / llGRP 'TTP 10439 - Rerate 21,000,000
            Else
                'llCPP = (-flTotalCost / 10) / llGRP   '(-llTotalCost / llGRP) / 10
                llCPP = (-dlTotalCost / 10) / llGRP 'TTP 10439 - Rerate 21,000,000
            End If
        Else
            llCPP = 0
        End If
        If flGrimp > 0 Then
            'llCPM = (llCPMAdj * flTotalCost) / flGrimp
            llCPM = (llCPMAdj * dlTotalCost) / flGrimp 'TTP 10439 - Rerate 21,000,000
        ElseIf ilAllowNeg And (flGrimp < 0) Then
            'If flTotalCost > 0 Then
            If dlTotalCost > 0 Then 'TTP 10439 - Rerate 21,000,000
                'llCPM = (llCPMAdj * flTotalCost) / flGrimp
                llCPM = (llCPMAdj * dlTotalCost) / flGrimp 'TTP 10439 - Rerate 21,000,000
            Else
                'llCPM = (-llCPMAdj * flTotalCost) / flGrimp
                llCPM = (-llCPMAdj * dlTotalCost) / flGrimp 'TTP 10439 - Rerate 21,000,000
            End If
        Else
            llCPM = 0
        End If
    End If
    'Fix TTP 10439 per Jason email v81 Contract report overflow error - Wed 11/16/22 9:42 AM
    'If flTotalCost > 2100000000 Then
    'If dlTotalCost > 2100000000 Then 'TTP 10439 - Rerate 21,000,000
    '    dlTotalCost = -1
    'Else
    '    dlTotalCost = flTotalCost 'TTP 10439 - Rerate 21,000,000
    'End If
    flTotalCost = dlTotalCost
    If flGrimp > 2100000000 Then
        llGrImp = -1
    Else
        llGrImp = flGrimp
    End If
    Exit Sub
TotalErr:
#If programmatic <> 1 Then
    MsgBox "Combination of Price, Spots and Audience Exceeded 21,000,000", vbOKOnly + vbExclamation, "Warning"
#End If
    dlTotalCost = 0
    ilAvgRating = 0
    llAvgAud = 0
    llGrImp = 0
    llGRP = 0
    llCPP = 0
    llCPM = 0
End Sub

Private Sub mComputeDemoIndex(slForm As String, ilDemoSAge As Integer, ilDemoEAge As Integer, ilStartDemoIndex As Integer, ilEndDemoIndex As Integer)

    'Determine which demos make-up demos
    'Standard 16(M12-17;M18-24;M25-34;M35-44;M45-49;M50-54;M55-64;M65+
    '            W12-17;W18-24;W25-34;W35-44;W45-49;W50-54;W55-64;W65+)
    'Standard 18(M12-17;M18-20;M21-24;M25-34;M35-44;M45-49;M50-54;M55-64;M65+
    '            W12-17;W18-20;W21-24;W25-34;W35-44;W45-49;W50-54;W55-64;W65+)
    'MnfDemo names: Mxx-xx; Pxx-xx; Axx-xx; Wxx-xx
    'Parse fields

    ilStartDemoIndex = -1
    ilEndDemoIndex = -1
    If slForm <> "8" Then
        If ilDemoSAge = 12 Then
            ilStartDemoIndex = 1
        ElseIf ilDemoSAge = 18 Then
            ilStartDemoIndex = 2
        ElseIf ilDemoSAge = 25 Then
            ilStartDemoIndex = 3
        ElseIf ilDemoSAge = 35 Then
            ilStartDemoIndex = 4
        ElseIf ilDemoSAge = 45 Then
            ilStartDemoIndex = 5
        ElseIf ilDemoSAge = 50 Then
            ilStartDemoIndex = 6
        ElseIf ilDemoSAge = 55 Then
            ilStartDemoIndex = 7
        ElseIf ilDemoSAge = 65 Then
            ilStartDemoIndex = 8
        End If

        If ilDemoEAge = 17 Then
            ilEndDemoIndex = 1
        ElseIf ilDemoEAge = 24 Then
            ilEndDemoIndex = 2
        ElseIf ilDemoEAge = 34 Then
            ilEndDemoIndex = 3
        ElseIf ilDemoEAge = 44 Then
            ilEndDemoIndex = 4
        ElseIf ilDemoEAge = 49 Then
            ilEndDemoIndex = 5
        ElseIf ilDemoEAge = 54 Then
            ilEndDemoIndex = 6
        ElseIf ilDemoEAge = 64 Then
            ilEndDemoIndex = 7
        ElseIf ilDemoEAge = 99 Then
            ilEndDemoIndex = 8
        End If
    Else
        If ilDemoSAge = 12 Then
            ilStartDemoIndex = 1
        ElseIf ilDemoSAge = 18 Then
            ilStartDemoIndex = 2
        ElseIf ilDemoSAge = 21 Then
            ilStartDemoIndex = 3
        ElseIf ilDemoSAge = 25 Then
            ilStartDemoIndex = 4
        ElseIf ilDemoSAge = 35 Then
            ilStartDemoIndex = 5
        ElseIf ilDemoSAge = 45 Then
            ilStartDemoIndex = 6
        ElseIf ilDemoSAge = 50 Then
            ilStartDemoIndex = 7
        ElseIf ilDemoSAge = 55 Then
            ilStartDemoIndex = 8
        ElseIf ilDemoSAge = 65 Then
            ilStartDemoIndex = 9
        End If

        If ilDemoEAge = 17 Then
            ilEndDemoIndex = 1
        ElseIf ilDemoEAge = 20 Then
            ilEndDemoIndex = 2
        ElseIf ilDemoEAge = 24 Then
            ilEndDemoIndex = 3
        ElseIf ilDemoEAge = 34 Then
            ilEndDemoIndex = 4
        ElseIf ilDemoEAge = 44 Then
            ilEndDemoIndex = 5
        ElseIf ilDemoEAge = 49 Then
            ilEndDemoIndex = 6
        ElseIf ilDemoEAge = 54 Then
            ilEndDemoIndex = 7
        ElseIf ilDemoEAge = 64 Then
            ilEndDemoIndex = 8
        ElseIf ilDemoEAge = 99 Then
            ilEndDemoIndex = 9
        End If

    End If

End Sub

Private Function mGetBookEstByLOrU(ilDnfCode As Integer) As String
    Dim illoop As Integer

    mGetBookEstByLOrU = "L"
    If tgSpf.sDemoEstAllowed <> "Y" Then
        Exit Function
    End If
    For illoop = LBound(igEstByUSADnfCode) To UBound(igEstByUSADnfCode) - 1 Step 1
        If igEstByUSADnfCode(illoop) = ilDnfCode Then
            mGetBookEstByLOrU = "U"
            Exit Function
        End If
    Next illoop
'    If ilDnfCode = tmDnf.iCode Then
'        mGetBookEstByLOrU = tmDnf.sEstListenerOrUSA
'        Exit Function
'    End If
'    hmDnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
'    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet = BTRV_ERR_NONE Then
'        imDnfRecLen = Len(tmDnf)
'        tmDnfSrchKey0.iCode = ilDnfCode
'        ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'        If ilRet = BTRV_ERR_NONE Then
'            mGetBookEstByLOrU = tmDnf.sEstListenerOrUSA
'        Else
'            tmDnf.iCode = ilDnfCode
'            tmDnf.sEstListenerOrUSA = "L"
'            mGetBookEstByLOrU = "L"
'        End If
'    Else
'        tmDnf.iCode = ilDnfCode
'        tmDnf.sEstListenerOrUSA = "L"
'        mGetBookEstByLOrU = "L"
'    End If
'    btrDestroy hmDnf
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainBooksForEstByUSA         *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get Contract (Chf), Lines (Clf) *
'*                                                     *
'*******************************************************
Public Sub gObtainBooksForEstByUSA()
'
'   gObtainBooksForEstByUSA
'   Where:
'       igEstByUSADnfCode() (O)- DnfCode for books with sEstListenerOrUSA as U
'       ilRet (O)- True = populated; False = error
'
    Dim llNoRec As Long         'Number of records in Mnf
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record


    ReDim igEstByUSADnfCode(0 To 0) As Integer
    If tgSpf.sDemoEstAllowed <> "Y" Then
        Exit Sub
    End If
    hmDnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDnf)
        btrDestroy hmDnf
        Exit Sub
    End If
    imDnfRecLen = Len(tmDnf)
    ilExtLen = Len(tmDnf)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hmDnf) 'Obtain number of records
    btrExtClear hmDnf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hmDnf)
        btrDestroy hmDnf
        Exit Sub
    Else
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmDnf)
            btrDestroy hmDnf
            Exit Sub
        End If
    End If
    Call btrExtSetBounds(hmDnf, llNoRec, -1, "UC", "DNF", "") 'Set extract limits (all records)
    tlCharTypeBuff.sType = "U"
    ilOffSet = gFieldOffset("Dnf", "DnfEstListenerOrUSA")
    ilRet = btrExtAddLogicConst(hmDnf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
    ilOffSet = 0
    ilRet = btrExtAddField(hmDnf, ilOffSet, imDnfRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDnf)
        btrDestroy hmDnf
        Exit Sub
    End If
    'ilRet = btrExtGetNextExt(hmDnf)    'Extract record
    ilRet = btrExtGetNext(hmDnf, tmDnf, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            ilRet = btrClose(hmDnf)
            btrDestroy hmDnf
            Exit Sub
        End If
        ilExtLen = Len(tmDnf)  'Extract operation record size
        'ilRet = btrExtGetFirst(hmDnf, tgAvailAnf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmDnf, tmDnf, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If tmDnf.sEstListenerOrUSA = "U" Then
                ilUpperBound = UBound(igEstByUSADnfCode)
                igEstByUSADnfCode(ilUpperBound) = tmDnf.iCode
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve igEstByUSADnfCode(0 To ilUpperBound) As Integer
            End If
            ilRet = btrExtGetNext(hmDnf, tmDnf, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmDnf, tmDnf, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hmDnf)
    btrDestroy hmDnf
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetGroupName                   *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Group name                 *
'*                                                     *
'*******************************************************
Public Sub gGetGroupName(ilIndex, slName As String, slGroup As String)
    Select Case ilIndex
        Case 1
            slName = "East Region"
            slGroup = "A1"
        Case 2
            slName = "South Region"
            slGroup = "A2"
        Case 3
            slName = "Midwest Region"
            slGroup = "A3"
        Case 4
            slName = "West Region"
            slGroup = "A4"
        Case 5
            slName = "County Size AA"
            slGroup = "B1"
        Case 6
            slName = "County Size A"
            slGroup = "B2"
        Case 7
            slName = "County Size B"
            slGroup = "B3"
        Case 8
            slName = "County Size C"
            slGroup = "B4"
        Case 9
            slName = "County Size D"
            slGroup = "B5"
        Case 10
            slName = "County Size AA & A"
            slGroup = "B6"
        Case 11
            slName = "County Size AA,A&B"
            slGroup = "B7"
        Case 12
            slName = "County Size C & D"
            slGroup = "B8"
        Case 13
            slName = "Household Size 1-2"
            slGroup = "C1"
        Case 14
            slName = "Household Size 3-4"
            slGroup = "C2"
        Case 15
            slName = "Household Size 5+"
            slGroup = "C3"
        Case 16
            slName = "Income < $25k"
            slGroup = "D1"
        Case 17
            slName = "Income $25k-$50k"
            slGroup = "D2"
        Case 18
            slName = "Income $50k-$75k"
            slGroup = "D3"
        Case 19
            slName = "Income $$75k+"
            slGroup = "D4"
        Case 20
            slName = "Income $50k+"
            slGroup = "D5"
'        Case 21
'            slName = "Income $40k+"
'            slGroup = "D6"
'        Case 22
'            slName = "Income $50k+"
'            slGroup = "D7"
        Case 21
            slName = "No College"
            slGroup = "E1"
        Case 22
            slName = "Some College"
            slGroup = "E2"
        Case 23
            slName = "College Graduate"
            slGroup = "E3"
        Case 24
            slName = "TV DMA Market Top 10"
            slGroup = "F1"
        Case 25
            slName = "TV DMA Market Top 25"
            slGroup = "F2"
        Case 26
            slName = "TV DMA Market Top 50"
            slGroup = "F3"
        Case 27
            slName = "TV DMA Top 100"
            slGroup = "F4"
        Case 28
            slName = "Women Job Full Time"
            slGroup = "G1"
        Case 29
            slName = "Women Full/Part Time"
            slGroup = "G2"
        Case 30
            slName = "Household Grad $50k+"
            slGroup = "H1"
'        Case 33
'            slName = "Job Tech/Sales/Cler"
'            slGroup = "H2"
'        Case 34
'            slName = "Job Other"
'            slGroup = "H3"
'        Case 35
'            slName = "Job None"
'            slGroup = "H4"
        Case 31
            slName = "Hispanic" '"College $50k+"
            slGroup = "I1"
        Case 32
            slName = "Black Non-Hispanic"   '"Prof/Manage $50k+"
            slGroup = "I2"
        Case 33
            slName = "Other"
            slGroup = "I3"
        Case 34
            slName = "Location In Home" '"TV Day-Light"
            slGroup = "J1"
        Case 35
            slName = "Location In Auto" '"TV Day-Medium/Heavy"
            slGroup = "J2"
        Case 36
            slName = "Location At Work" '"TV After/Eve-Light"
            slGroup = "J3"
        Case 37
            slName = "Location Other"   '"TV After/Eve-Medium"
            slGroup = "J4"
'        Case 43
'            slName = "TV After/Eve-Heavy"
'            slGroup = "J5"
'        Case 44
'            slName = "TV Night-Light"
'            slGroup = "J6"
'        Case 45
'            slName = "TV Night-Medium"
'            slGroup = "J7"
'        Case 46
'            slName = "TV Night-Heavy"
'            slGroup = "J8"
        Case 38
            slName = "Childen None < 12"    '"White Race"
            slGroup = "K1"
        Case 39
            slName = "Childen Any < 12" '"Black Race"
            slGroup = "K2"
        Case 40
            slName = "Top 10-HH Inc 75k+"    '"Loc-In home only"
            slGroup = "L1"
        Case 41
            slName = "Top 10-College Grad"    '"Loc-In Auto"
            slGroup = "L2"
        Case 42
            slName = "Top 10-Black Non-H"    '"Loc-Out home ex-auto"
            slGroup = "L3"
        Case 43
            slName = "Top 10-Work FT/PT"    '"Loc-Out home ex-auto"
            slGroup = "L4"
        Case 44
            slName = "Top 10-Childen<12"    '"Loc-Out home ex-auto"
            slGroup = "L5"
        Case 45
            slName = "Top 10-Hispanic"    '"Loc-Out home ex-auto"
            slGroup = "L6"
        Case 46
            slName = "Top 25-HH Inc 75k+"    '"Loc-In home only"
            slGroup = "M1"
        Case 47
            slName = "Top 25-College Grad"    '"Loc-In Auto"
            slGroup = "M2"
        Case 48
            slName = "Top 25-Black Non-H"    '"Loc-Out home ex-auto"
            slGroup = "M3"
        Case 49
            slName = "Top 25-Work FT/PT"    '"Loc-Out home ex-auto"
            slGroup = "M4"
        Case 50
            slName = "Top 25-Childen<12"    '"Loc-Out home ex-auto"
            slGroup = "M5"
        Case 51
            slName = "Top 25-Hispanic"    '"Loc-Out home ex-auto"
            slGroup = "M6"
        Case 52
            slName = "Top 50-HH Inc 75k+"    '"Loc-In home only"
            slGroup = "N1"
        Case 53
            slName = "Top 50-College Grad"    '"Loc-In Auto"
            slGroup = "N2"
        Case 54
            slName = "Top 50-Black Non-H"    '"Loc-Out home ex-auto"
            slGroup = "N3"
        Case 55
            slName = "Top 50-Work FT/PT"    '"Loc-Out home ex-auto"
            slGroup = "N4"
        Case 56
            slName = "Top 50-Childen<12"    '"Loc-Out home ex-auto"
            slGroup = "N5"
        Case 57
            slName = "Top 50-Hispanic"    '"Loc-Out home ex-auto"
            slGroup = "N5"
        Case 58
            slName = "HH 75k+ & Coll Grad"    '"Loc-In home only"
            slGroup = "O1"
        Case 59
            slName = "HH 75k+ & Work FT/PT"    '"Loc-In Auto"
            slGroup = "O2"
        Case 60
            slName = "HH 75k+ & Childen<12"    '"Loc-Out home ex-auto"
            slGroup = "O3"
        Case 60
            slName = "Hispanic DMA-Top 10"    '"Loc-Out home ex-auto"
            slGroup = "P1"
        Case 61
            slName = "Hispanic DMA-Top 25"    '"Loc-In home only"
            slGroup = "P2"
        Case 62
            slName = "Hispanic DMA-Top 50"    '"Loc-In Auto"
            slGroup = "P3"
        Case 63
            slName = "Hispanic DMA-Top 100"    '"Loc-Out home ex-auto"
            slGroup = "P4"
        Case 64
            slName = "Black DMA-Top 10"    '"Loc-Out home ex-auto"
            slGroup = "Q1"
        Case 65
            slName = "Black DMA-Top 25"    '"Loc-In home only"
            slGroup = "Q2"
        Case 66
            slName = "Black DMA-Top 50"    '"Loc-In Auto"
            slGroup = "Q3"
        Case 67
            slName = "Black DMA-Top 100"    '"Loc-Out home ex-auto"
            slGroup = "Q4"
        Case 68
            slName = "Top10 His-HH 75K+"
            slGroup = "R1"
        Case 69
            slName = "Top10 His-Coll Grad"
            slGroup = "R2"
        Case 70
            slName = "Top10 His-Hispanic"
            slGroup = "R3"
        Case 71
            slName = "Top10 His-Work FT/PP"
            slGroup = "R4"
        Case 72
            slName = "Top10 His-Child<12"
            slGroup = "R5"
        Case 73
            slName = "Top25 His-HH 75K+"
            slGroup = "S1"
        Case 74
            slName = "Top25 His-Coll Grad"
            slGroup = "S2"
        Case 75
            slName = "Top25 His-Hispanic"
            slGroup = "S3"
        Case 76
            slName = "Top25 His-Work FT/PP"
            slGroup = "S4"
        Case 77
            slName = "Top25 His-Child<12"
            slGroup = "S5"
        Case 78
            slName = "Top50 His-HH 75K+"
            slGroup = "T1"
        Case 79
            slName = "Top50 His-Coll Grad"
            slGroup = "T2"
        Case 80
            slName = "Top50 His-Hispanic"
            slGroup = "T3"
        Case 81
            slName = "Top50 His-Work FT/PP"
            slGroup = "T4"
        Case 82
            slName = "Top50 His-Child<12"
            slGroup = "T5"
        Case 83
            slName = "Top10 Blk-HH 75K+"
            slGroup = "U1"
        Case 84
            slName = "Top10 Blk-Coll Grad"
            slGroup = "U2"
        Case 85
            slName = "Top10 Blk-Hispanic"
            slGroup = "U3"
        Case 86
            slName = "Top10 Blk-Work FT/PP"
            slGroup = "U4"
        Case 87
            slName = "Top10 Blk-Child<12"
            slGroup = "U5"
        Case 88
            slName = "Top25 Blk-HH 75K+"
            slGroup = "V1"
        Case 89
            slName = "Top25 Blk-Coll Grad"
            slGroup = "V2"
        Case 90
            slName = "Top25 Blk-Hispanic"
            slGroup = "V3"
        Case 91
            slName = "Top25 Blk-Work FT/PP"
            slGroup = "V4"
        Case 92
            slName = "Top25 Blk-Child<12"
            slGroup = "V5"
        Case 93
            slName = "Top50 Blk-HH 75K+"
            slGroup = "W1"
        Case 94
            slName = "Top50 Blk-Coll Grad"
            slGroup = "W2"
        Case 95
            slName = "Top50 Blk-Hispanic"
            slGroup = "W3"
        Case 96
            slName = "Top50 Blk-Work FT/PP"
            slGroup = "W4"
        Case 97
            slName = "Top50 Blk-Child<12"
            slGroup = "W5"
        Case 98
            slName = "Hisp&Lang-All Span"
            slGroup = "X1"
        Case 99
            slName = "Hisp&Lang-Most Span"
            slGroup = "X2"
        Case 100
            slName = "Hisp&Lang-All Engl"
            slGroup = "X3"
        Case 101
            slName = "Hisp&Lang-Most Engl"
            slGroup = "X4"
        Case 102
            slName = "Hisp&Lang-Span Dom"
            slGroup = "X5"
        Case 103
            slName = "Hisp&Lang-Engl Dom"
            slGroup = "X6"
        Case 104
            slName = "Hisp&HDMA-Top 10"    '"Loc-Out home ex-auto"
            slGroup = "Y1"
        Case 105
            slName = "Hisp&HDMA-Top 25"    '"Loc-In home only"
            slGroup = "Y2"
        Case 106
            slName = "Hisp&HDMA-Top 50"    '"Loc-In Auto"
            slGroup = "Y3"
        Case 107
            slName = "Hisp&HDMA-Top 100"    '"Loc-Out home ex-auto"
            slGroup = "Y4"
        Case 108
            slName = "Blk Non Hisp-Top 10"    '"Loc-Out home ex-auto"
            slGroup = "AA1"
        Case 109
            slName = "Blk Non Hisp-Top 25"    '"Loc-In home only"
            slGroup = "AA2"
        Case 110
            slName = "Blk Non Hisp-Top 50"    '"Loc-In Auto"
            slGroup = "AA3"
        Case 111
            slName = "Blk Non Hisp-Top 100"    '"Loc-Out home ex-auto"
            slGroup = "AA4"
        Case 112
            slName = "His&Top10HDMA-HH75K+"
            slGroup = "AB1"
        Case 113
            slName = "His&Top10HDMA-Grad"
            slGroup = "AB2"
        Case 114
            slName = "His&Top10HDMA-FT/PP"
            slGroup = "AB3"
        Case 115
            slName = "His&Top10HDMA-Kid<12"
            slGroup = "AB4"
        Case 116
            slName = "His&Top25HDMA-HH75K+"
            slGroup = "AC1"
        Case 117
            slName = "His&Top25HDMA-Grad"
            slGroup = "AC2"
        Case 118
            slName = "His&Top25HDMA-FT/PP"
            slGroup = "AC3"
        Case 119
            slName = "His&Top25HDMA-Kid<12"
            slGroup = "AC4"
        Case 120
            slName = "His&Top50HDMA-HH75K+"
            slGroup = "AD1"
        Case 121
            slName = "His&Top50HDMA-Grad"
            slGroup = "AD2"
        Case 122
            slName = "His&Top50HDMA-FT/PP"
            slGroup = "AD3"
        Case 123
            slName = "His&Top50HDMA-Kid<12"
            slGroup = "AD4"
        Case 124
            slName = "Blk&Top10BDMA-HH75K+"
            slGroup = "AE1"
        Case 125
            slName = "Blk&Top10BDMA-Grad"
            slGroup = "AE2"
        Case 126
            slName = "Blk&Top10BDMA-FT/PP"
            slGroup = "AE3"
        Case 127
            slName = "Blk&Top10BDMA-Kid<12"
            slGroup = "AE4"
        Case 128
            slName = "Blk&Top25BDMA-HH75K+"
            slGroup = "AF1"
        Case 129
            slName = "Blk&Top25BDMA-Grad"
            slGroup = "AF2"
        Case 130
            slName = "Blk&Top25BDMA-FT/PP"
            slGroup = "AF3"
        Case 131
            slName = "Blk&Top25BDMA-Kid<12"
            slGroup = "AF4"
        Case 132
            slName = "Blk&Top50BDMA-HH75K+"
            slGroup = "AG1"
        Case 133
            slName = "Blk&Top50BDMA-Grad"
            slGroup = "AG2"
        Case 134
            slName = "Blk&Top50BDMA-FT/PP"
            slGroup = "AG3"
        Case 135
            slName = "Blk&Top50BDMA-Kid<12"
            slGroup = "AG4"

    End Select
End Sub

Private Function mRegionAudAdj(hlRaf As Integer, llRafCode As Long, llInAvgAud As Long) As Long
    Dim ilRet As Integer
    Dim llAvgAud As Long

    mRegionAudAdj = llInAvgAud
    If llRafCode <= 0 Then
        Exit Function
    End If
    If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) <> SPLITNETWORKS) Then
        Exit Function
    End If
    imRafRecLen = Len(tmRaf)
    tmRafSrchKey0.lCode = llRafCode
    ilRet = btrGetEqual(hlRaf, tmRaf, imRafRecLen, tmRafSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        If (tmRaf.iAudPct > 0) And (tmRaf.iAudPct < 10000) Then
            'llAvgAud = (tmRaf.iAudPct * llInAvgAud) / 10000
            llAvgAud = gStrDecToLong(gDivStr(gMulStr(gIntToStrDec(tmRaf.iAudPct, 2), gLongToStrDec(llInAvgAud, 0)), "100."), 0)
            mRegionAudAdj = llAvgAud
        End If
    End If
End Function
Public Sub gConvCustomGroup(ilGroupNo As Integer, slDataType As String, ilDemoIndex As Integer)
    If ilGroupNo <= 18 Then
        slDataType = "B"
    Else
        slDataType = Chr$(99 + (ilGroupNo - 1) \ 18)
    End If
    ilDemoIndex = ((ilGroupNo - 1) Mod 18) + 1
End Sub
Public Function gReconvertCustomGroup(slDataType As String) As Integer
    'Dan M 12/23/10 reverse of gConvCustomGroup. "d" = 19  "e" = 37
    Dim ilRet As Integer
    
    If slDataType = "B" Then
        ilRet = 1
    Else
        ilRet = CInt(((Asc(slDataType) - 99) * 18) + 1)
    End If
    If ilRet < 0 Then
        ilRet = 1
    End If
    gReconvertCustomGroup = ilRet
End Function
'10/27/14: 1 or 2 place rating
Public Function gSet1or2PlaceRating(ilAgfCode As Integer) As String
    Dim ilAgf As Integer
    
    gSet1or2PlaceRating = "1"
    If tgSpf.sSGRPCPPCal <> "A" Then
        Exit Function
    End If
    If ilAgfCode <= 0 Then
        Exit Function
    End If
    ilAgf = gBinarySearchAgf(ilAgfCode)
    If ilAgf <> -1 Then
        gSet1or2PlaceRating = tgCommAgf(ilAgf).s1or2DigitRating
    End If
End Function

Private Sub mGetEstAdj(ilDnfCode As Integer, slEstLOrU As String, slEstPop As String, slEstAud As String, ilEstNoWks As Integer, llStartDateForEst As Long, llEndDateForEst As Long, hlDrf As Integer, hlMnf As Integer, hlDpf As Integer, ilMnfSocEco As Integer, ilMnfDemo As Integer, llDemoPop As Long, hlDef As Integer, tlDef As DEF, ilDefRecLen As Integer, llAvgAud As Long, llPopForEst As Long)
'Debug.Print "     - mGetEstAdj: ilDnfCode:" & ilDnfCode
    Dim ilRet As Integer
    Dim llDate As Long
    Dim tlDefSrchKey1 As DEFKEY1
    
    If tgSpf.sDemoEstAllowed <> "Y" Then
        'Return
        Exit Sub
    End If
    slEstLOrU = mGetBookEstByLOrU(ilDnfCode)
    If slEstLOrU <> "U" Then
        tgSpf.sDemoEstAllowed = "N"
    End If
    ilRet = gGetDemoPop(hlDrf, hlMnf, hlDpf, ilDnfCode, ilMnfSocEco, ilMnfDemo, llDemoPop)
    tgSpf.sDemoEstAllowed = "Y"
    slEstPop = "0"
    slEstAud = "0"
    ilEstNoWks = 0
    For llDate = llStartDateForEst To llEndDateForEst Step 7
        ilEstNoWks = ilEstNoWks + 1
        tlDefSrchKey1.iDnfCode = ilDnfCode
        gPackDateLong llDate, tlDefSrchKey1.iStartDate(0), tlDefSrchKey1.iStartDate(1)
        ilRet = btrGetLessOrEqual(hlDef, tlDef, ilDefRecLen, tlDefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If (ilRet = BTRV_ERR_NONE) And (tlDef.iDnfCode = ilDnfCode) Then
            'clEstAud = clEstAud + (CCur(llAvgAud) * tlDef.lEstimatePct)
            'clEstPop = clEstPop + (CCur(llDemoPop) * tlDef.lEstimatePct)
            slEstAud = gAddStr(slEstAud, gMulStr(str$(llAvgAud), str(tlDef.lEstimatePct)))  '(CCur(llAvgAud) * tlDef.lEstimatePct)
            If slEstLOrU <> "U" Then
                slEstPop = gAddStr(slEstPop, gMulStr(str$(llDemoPop), str(tlDef.lEstimatePct)))  '+ (CCur(llDemoPop) * tlDef.lEstimatePct)
            End If
        Else
            'clEstAud = clEstAud + (CCur(llAvgAud) * 10000)
            'clEstPop = clEstPop + CCur(llDemoPop) * 10000
            slEstAud = gAddStr(slEstAud, gMulStr(str$(llAvgAud), "10000"))  '(CCur(llAvgAud) * tlDef.lEstimatePct)
            If slEstLOrU <> "U" Then
                slEstPop = gAddStr(slEstPop, gMulStr(str$(llDemoPop), "10000"))  '+ (CCur(llDemoPop) * tlDef.lEstimatePct)
            End If
        End If
    Next llDate
    If ilEstNoWks > 0 Then
        'llAvgAud = ((clEstAud + 50) / ilEstNoWks) \ 10000
        'llPopForEst = ((clEstPop + 50) / ilEstNoWks) \ 10000
        llAvgAud = Val(gDivStr(gDivStr(slEstAud, str$(ilEstNoWks)), "10000"))  '((clEstAud + 50) / ilEstNoWks) \ 10000
        If slEstLOrU <> "U" Then
            llPopForEst = Val(gDivStr(gDivStr(slEstPop, str$(ilEstNoWks)), "10000"))  '((clEstPop + 50) / ilEstNoWks) \ 10000
        Else
            llPopForEst = llDemoPop
        End If
    Else
        llAvgAud = llAvgAud
        llPopForEst = llDemoPop
    End If
End Sub

Private Sub mGetAvgAud(tlDrf As DRF, ilMnfDemo As Integer, hlDpf As Integer, tlDpf As DPF, ilDpfRecLen As Integer, llTAvgAud As Long, tlMnf As MNF, ilDemoSAge As Integer, ilDemoEAge As Integer, ilStartDemoIndex As Integer, ilEndDemoIndex As Integer, ilDemoSex As Integer, ilCustomIndex As Integer)
'Debug.Print "     - mGetAvgAud: DRFCode=" & tlDrf.lCode & " , MnfGroupNo:" & tlMnf.iGroupNo
    Dim ilRet As Integer
    Dim tlDpfSrchKey1 As DPFKEY1
    Dim illoop As Integer
    
    If lgDpfNoRecs > 0 Then
        tlDpfSrchKey1.lDrfCode = tlDrf.lCode
        tlDpfSrchKey1.iMnfDemo = ilMnfDemo
        ilRet = btrGetEqual(hlDpf, tlDpf, ilDpfRecLen, tlDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            llTAvgAud = llTAvgAud + tlDpf.lDemo
            'Return
            Exit Sub
        End If
    End If
    If tlMnf.iGroupNo = 0 Then
        If ilStartDemoIndex < 0 Then
            mComputeDemoIndex tlDrf.sForm, ilDemoSAge, ilDemoEAge, ilStartDemoIndex, ilEndDemoIndex
            If (ilStartDemoIndex <> -1) And (ilEndDemoIndex <> -1) Then
                If LBound(tlDrf.lDemo) = 0 Then
                    ilStartDemoIndex = ilStartDemoIndex - 1
                    ilEndDemoIndex = ilEndDemoIndex - 1
                End If
            End If
        End If
        If (ilStartDemoIndex <> -1) And (ilEndDemoIndex <> -1) Then
            If (ilDemoSex = 0) Or (ilDemoSex = 2) Then
                For illoop = ilStartDemoIndex To ilEndDemoIndex Step 1
                    llTAvgAud = llTAvgAud + tlDrf.lDemo(illoop)
                Next illoop
            End If
            If (ilDemoSex = 1) Or (ilDemoSex = 2) Then
                If tlDrf.sForm <> "8" Then
                    For illoop = ilStartDemoIndex + 8 To ilEndDemoIndex + 8 Step 1
                        llTAvgAud = llTAvgAud + tlDrf.lDemo(illoop)
                    Next illoop
                Else
                    For illoop = ilStartDemoIndex + 9 To ilEndDemoIndex + 9 Step 1
                        llTAvgAud = llTAvgAud + tlDrf.lDemo(illoop)
                    Next illoop
                End If
            End If
        End If
    Else
        llTAvgAud = llTAvgAud + tlDrf.lDemo(ilCustomIndex - 1)
    End If
End Sub
Private Sub mSetActOneInfoFromDemoCall(tlDrf As DRF)
    
    'infotype = daypart and not demodatatype <> population
    If (tlDrf.sInfoType = "D" Or tlDrf.sInfoType = "V") And tlDrf.sDemoDataType <> "P" Then
        sgActLineup = Trim$(gStripChr0(tlDrf.sACTLineupCode))
        If Trim$(gStripChr0(tlDrf.sACT1StoredTime)) = "T" Then
            igActTime = 0
        Else
            igActTime = -1
        End If
        If Trim$(gStripChr0(tlDrf.sACT1StoredSpots)) = "S" Then
            igActSpots = 0
        Else
            igActSpots = -1
        End If
        If Trim$(gStripChr0(tlDrf.sACT1StoreClearPct)) = "C" Then
            igActPct = 0
        Else
            igActPct = -1
        End If
        If Trim$(gStripChr0(tlDrf.sACT1DaypartFilter)) = "F" Then
            igActDaypart = 0
        Else
            igActDaypart = -1
        End If
    Else
        gSetAct1ToBlank
    End If
End Sub
Public Sub gSetAct1ToBlank()
    '10256
    sgActLineup = ""
    igActTime = -1
    igActSpots = -1
    igActPct = -1
    igActDaypart = -1
End Sub

