Attribute VB_Name = "RPTVFYLG"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptvfylg.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelLg.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
'Public tgAirNameCode() As SORTCODE
'Public sgAirNameCodeTag As String
'Public tgCSVNameCode() As SORTCODE
'Public sgCSVNameCodeTag As String
'Public tgSellNameCode() As SORTCODE
'Public sgSellNameCodeTag As String
'Public tgRptSelLgAgencyCode() As SORTCODE
'Public sgRptSelLgAgencyCodeTag As String
'Public tgRptSelLgSalespersonCode() As SORTCODE
'Public sgRptSelLgSalespersonCodeTag As String
'Public tgRptSelLgAdvertiserCode() As SORTCODE
'Public sgRptSelLgAdvertiserCodeTag As String
'Public tgRptSelLgNameCode() As SORTCODE
'Public sgRptSelLgNameCodeTag As String
'Public tgRptSelLgBudgetCode() As SORTCODE
'Public sgRptSelLgBudgetCodeTag As String
'Public tgMultiCntrCode() As SORTCODE
'Public sgMultiCntrCodeTag As String
'Public tgManyCntCode() As SORTCODE
'Public sgManyCntCodeTag As String
'Public tgRptSelLgDemoCode() As SORTCODE
'Public sgRptSelLgDemoCodeTag As String
'Public tgSOCode() As SORTCODE
'Public sgSOCodeTag As String
'Public sgMNFCodeTag As String       '6-19-01
'Public igUsingCrystal As Integer
'Public sgPhoneImage As String
'Public igJobRptNo As Integer
'Public lgStartingCntrNo As Long
'Public lgOrigCntrNo As Long
'Public tgVof As VOF                     'vehicle options for log/cps
'Public sgRnfRptName As String * 3          'Report name from RNF file (L01, L02,... C01, C02,.. ..etc)
'Public igNoCodes As Integer
'Public igCodes() As Integer
'Public sgLogStartDate As String
'Public sgLogNoDays As String
'Public sgLogUserCode As String
'Public sgLogStartTime As String
'Public sgLogEndTime As String
'Public igZones As Integer                   'time zones  (0=all, 1=est, 2=cst, 3=mst, 4=pst)
'Public igRnfCode As Integer
'Public sgLogType As String * 1              'l=log, c=cp, o = other
'Public igGenRpt As Integer     'True = Generating Report ignore user input
'Public igOutput As Integer
'Public igCopies As Integer
'Public igWhen As Integer
'Public igFile As Integer
'Public igOption As Integer
'Public igReportType As Integer
'Public igOutputTo As Integer        '0 = display , 1 = print
''Global spot types for Spots by Advt & spots by Date & Time
''bit selectivity for charged and different types of no charge spots
''bits defined right to left (0 to 9)
'Public Const SPOT_CHARGE = &H1         'charged
'Public Const SPOT_00 = &H2          '0.00
'Public Const SPOT_ADU = &H4         'ADU
'Public Const SPOT_BONUS = &H8       'bonus
'Public Const SPOT_EXTRA = &H10      'Extra
'Public Const SPOT_FILL = &H20       'Fill
'Public Const SPOT_NC = &H40         'no charge
'Public Const SPOT_MG = &H80         'mg
'Public Const SPOT_RECAP = &H100     'recapturable
'Public Const SPOT_SPINOFF = &H200   'spinoff
'Library calendar file- used to obtain post log date status
Dim hmVof As Integer
Dim tmVofSrchKey As VOFKEY0
'
'**************************************************************
'*                                                             *
'*      Procedure Name:gGenReportLg                            *
'*                                                             *
'*             Created:6/16/93       By:D. LeVine              *
'*            Modified:              By:                       *
'*                                                             *
'*         Comments: Formula setups for Crystal                *
'
'       <input> llRegion = split region network code, else
'                          0 for full network
'*                                                             *
'*          Return : 0 =  either error in input, stay in       *
'*                   -1 = error in Crystal, return to          *
'*                        calling program                      *
''*                       failure of gSetformula or another    *
'*                    1 = Crystal successfully completed       *
'*                    2 = successful Bridge                    *
'***************************************************************
Function gCmcGenLg(llRegionCode As Long, slRegionName As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slDateFrom                    slDateTo                                                *
'******************************************************************************************

    Dim slSelection As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slTime As String
    Dim ilPreview As Integer
    Dim slBaseDate As String
    Dim llDate1 As Long
    Dim llDate2 As Long
    Dim ilWeek As Integer
    Dim slLegend As String              'amfm legend showing contact
'    Dim slBrownBag As String
'    Dim slBrownBag1 As String            'amfm vehicle specifications on L28
'    Dim slBrownBag1A As String
'    Dim slBrownBag2 As String
'    Dim slBrownBag3 As String
    ReDim ilStdStart(0 To 1) As Integer
    gCmcGenLg = 0
    If igOutputTo = 0 Then          'display
        ilPreview = True
    Else                        'print
        ilPreview = False
    End If
    'setup controls as tho they were entered on an input screen
    RptSelLg!edcSelCFrom.Text = sgLogStartDate                'setup Start Date from Log Default screen
    RptSelLg!edcSelCFrom1.Text = sgLogNoDays          'setup # days from Log Default screen
    RptSelLg!edcSelCTo.Text = sgLogStartTime
    RptSelLg!edcSelCTo1.Text = sgLogEndTime
    'Hard-coded for AMFM until user defined field implemented
    'slLegend = "Questions? Contact Eric Garst, Clearance Manager, 972-455-6260 or Rebecca Starr, Associate Director of Operations, 972-455-6296"
    slLegend = ""
    
    'Two bridge reports will process and exit
    'If sgRnfRptName = "L07" Then   'commercial schedule
    '    gLogSevenRptAll ilPreview, "LogSeven.Lst", Val(slLogUserCode), "L07"
    '    gCmcGenLg = 2         'successful return
    '    Exit Function
    'ElseIf sgRnfRptName = "L08" Then   'commercial summary (orig version)
    '    gCmmlSumRpt ilPreview, "CmmlSum.Lst", Val(slLogUserCode)
    '    gCmcGenLg = 2         'successful return
    '    Exit Function
    'ElseIf sgRnfRptName = "L27" Then   'commercial schedule w/ new event IDs (ABC 3/9/99)
    '    gLogSevenRptAll ilPreview, "LogL27.Lst", Val(slLogUserCode), "L27"
    '    gCmcGenLg = 2         'successful return
    '    Exit Function
    'End If
    '9-11-00 If sgRnfRptName = "L09" Or sgRnfRptName = "L17" Or sgRnfRptName = "L41" Or sgRnfRptName = "L43" Then            'copy playlist
    If sgRnfRptName = "L09" Or sgRnfRptName = "L17" Or sgRnfRptName = "L41" Or sgRnfRptName = "L70" Then            '9-11-00 copy playlist
        'Place current date and time generated into ignowdate & ignowtime
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for rept headers or prepass keys

        slSelection = "{CPR_Copy_Report.cprGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And {CPR_Copy_Report.cprGenTime} = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        '9-20-16 C89 copy of L10, excludes Politicals.  clients runs 2 "L" logs, have to make this a "C" log to get 3 logs printed
        'L89 (copy of L10 with columns removed), added 2/22/19
    ElseIf sgRnfRptName = "L10" Or sgRnfRptName = "L28" Or sgRnfRptName = "L34" Or sgRnfRptName = "L40" Or sgRnfRptName = "C82" Or sgRnfRptName = "C89" Or sgRnfRptName = "L89" Then             '7 day commercial schedule (amfm version)
        If sgRnfRptName <> "C82" Then           'bypass following if not the customized Commercial Schedule
'           2-28-19 remove this hard coded text
'            slBrownBag = ""
'            slBrownBag1 = ""
'            slBrownBag1A = ""
'            slBrownBag2 = ""
'            slBrownBag3 = ""
'            If InStr(tgSpf.sGClient, "AMFM") > 0 Then
'                'slLegend = "Questions? Contact Eric Garst, Clearance Manager, 972-455-6260 or Rebecca Starr, Associate Director of Operations, 972-455-6296"
'                If sgRnfRptName = "L28" Then
'                    slBrownBag = "BROWN BAG "
'                ElseIf sgRnfRptName = "L40" Then
'                    slBrownBag = ""     '5-24-00 Show vehicle name not their selling network "WEATHER CHANNEL "
'                End If
'                slBrownBag1 = slBrownBag & "SPOTS MUST BE SCHEDULED ON THE DAY & DATE INDICATED ABOVE "
'                slBrownBag1A = "AND WITHIN THE PLEDGE TIME WINDOW SPECIFIED IN YOUR STATION"
'                slBrownBag2 = "S CONTRACTUAL AGREEMENT. "
'                slBrownBag3 = " ANY AND ALL EXCEPTIONS MUST BE APPROVED BY AMFM RADIO NETWORKS IN ADVANCE OF AIR."
'                If sgRnfRptName = "L28" Or sgRnfRptName = "L40" Then
'                    If Not gSetFormula("BrownBag1", "'" & (slBrownBag1) & "'") Then
'                        gCmcGenLg = -1
'                        Exit Function
'                    End If
'                    If Not gSetFormula("BrownBag1A", "'" & (slBrownBag1A) & "'") Then
'                        gCmcGenLg = -1
'                        Exit Function
'                    End If
'                    If Not gSetFormula("BrownBag2", "'" & (slBrownBag2) & "'") Then
'                        gCmcGenLg = -1
'                        Exit Function
'                    End If
'                    If Not gSetFormula("BrownBag3", "'" & (slBrownBag3) & "'") Then
'                        gCmcGenLg = -1
'                        Exit Function
'                    End If
'                End If
'
'            Else
                slLegend = " "
'            End If
            If Not gSetFormula("Legend1", "'" & slLegend & "'") Then
                gCmcGenLg = -1
                Exit Function
            End If
'           2-28-19 remove hard coded text
'            If InStr(tgSpf.sGClient, "AMFM") > 0 Then
'                slLegend = "***** Please insure this schedule is distributed to all stations affiliated with "
'            Else
                slLegend = " "
'            End If
            If Not gSetFormula("Legend2", "'" & slLegend & "'") Then
                gCmcGenLg = -1
                Exit Function
            End If
        End If

        If mShowzone = -1 Then      '2-3-05 chg to function
            gCmcGenLg = -1
            Exit Function
        End If

        'If igZones <> 0 Then             'all time zones, show time zone in header
        '    If Not gSetFormula("ShowZone", "'N'") Then
        '        gCmcGenLg = -1
        '        Exit Function
        '    End If
        'Else                            'selected single time zone, dont show it in the header
        '    If Not gSetFormula("ShowZone", "'Y'") Then
        '        gCmcGenLg = -1
        '        Exit Function
        '    End If
        'End If
        'Place current date and time generated into ignowdate & ignowtime
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for rept headers or prepass keys

        slSelection = "{SVR_7Day_Report.svrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({SVR_7Day_Report.svrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        mSelectZones slSelection, 2
    ElseIf sgRnfRptName = "L11" Or sgRnfRptName = "L32" Or sgRnfRptName = "L35" Then            'Comml summary (amfm or jones version - 15 Dayparts)
'       2-28-19 remove this hardcoded text
'        If InStr(tgSpf.sGClient, "AMFM") > 0 Then   'amfm uses l11 and l35
'            If sgRnfRptName = "L11" Then
'                slLegend = "5A"             '5 DPs
'            ElseIf sgRnfRptName = "L35" Then
'                'slLegend = "Questions? Contact Eric Garst, Clearance Manager, 972-455-6260 or Rebecca Starr, Associate Director of Operations, 972-455-6296"
'                If Not gSetFormula("Legend1", "'" & slLegend & "'") Then
'                    gCmcGenLg = -1
'                    Exit Function
'                End If
'                slLegend = "***** Please insure this schedule is distributed to all stations affiliated with "
'                If Not gSetFormula("Legend2", "'" & slLegend & "'") Then
'                    gCmcGenLg = -1
'                    Exit Function
'                End If
'
'                'get the end time of the first DP, which becomes the start time
'                'in the headers for 2nd daypart.  The first DP description will be O/N (Overnite)
'
'
'                If Not gSetFormula("Early2", "'5A'") Then      'hard-coded for AMFM
'                    gCmcGenLg = -1
'                    Exit Function
'                End If
'                slLegend = "O/N"            'L35 first DP of 6 (6DP for M-f, Sa, Su)
'            Else
'                slLegend = "Early"
'            End If
'        Else
            If sgRnfRptName = "L35" Then        'show extra DP (18 vs 15), not AMFM using this coml summary
                If Not gSetFormula("Early2", "'5A'") Then      'hard-coded for AMFM
                    gCmcGenLg = -1
                    Exit Function
                End If
            End If
            slLegend = "O/N"
'        End If
        If Not gSetFormula("Early", "'" & slLegend & "'") Then
            gCmcGenLg = -1
            Exit Function
        End If


        If mShowzone = -1 Then      '2-3-05 chg to function
            gCmcGenLg = -1
            Exit Function
        End If

        'If igZones <> 0 Then             'all time zones, show time zone in header
        '    If Not gSetFormula("ShowZone", "'N'") Then
        '        gCmcGenLg = -1
        '        Exit Function
        '    End If
        'Else                            'selected single time zone, dont show it in the header
        '    If Not gSetFormula("ShowZone", "'Y'") Then
        '        gCmcGenLg = -1
        '        Exit Function
        '    End If
        'End If


        'Place current date and time generated into ignowdate & ignowtime
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for rept headers or prepass keys

        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        mSelectZones slSelection, 1
    ElseIf sgRnfRptName = "L36" Or sgRnfRptName = "L38" Or sgRnfRptName = "L08" Or sgRnfRptName = "L88" Then      '11-3-99 abc version of comml summary converted from bridge to crystal; 1-12-12 L08
        'Place current date and time generated into ignowdate & ignowtime
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for rept headers or prepass keys

        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        mSelectZones slSelection, 1
    ElseIf sgRnfRptName = "L37" Or sgRnfRptName = "L39" Or sgRnfRptName = "C80" Then    '5-17-01 (c80), 11-4-99 abc version of coml sched converted from bridge to crystl
        'Place current date and time generated into ignowdate & ignowtime
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for rept headers or prepass keys

        slSelection = "{SVR_7Day_Report.svrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({SVR_7Day_Report.svrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        mSelectZones slSelection, 2
    ElseIf sgRnfRptName = "L31" Then
        gUnpackDate igNowDate(0), igNowDate(1), slDate
        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay

        gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
        slSelection = "{ODF_One_Day_Log.odfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({ODF_One_Day_Log.odfGenTime} ) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        mSelectZones slSelection, 0           'setup Crystal formula for zone filtering
    '9-11-00 ElseIf sgRnfRptName = "C22" Or sgRnfRptName = "C23" Then
    ElseIf sgRnfRptName = "C72" Or sgRnfRptName = "C73" Or sgRnfRptName = "C84" Then    '5-4-04 added c84
        'Place current date and time generated into ignowdate & ignowtime
        gCurrDateTime slDate, slTime, slMonth, slDay, slYear    'get system date and time for rept headers or prepass keys

        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        mSelectZones slSelection, 1
    ElseIf sgRnfRptName = "L87" Then                'podcast log, use the generated date and time filter from log generation
        gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
        gUnpackDate igNowDate(0), igNowDate(1), slDate
        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
        slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        mSelectZones slSelection, 3         'determine zones from CBF
    Else
        '2-3-05  show the zones on l74
        If sgRnfRptName = "L74" Then
            If mShowzone = -1 Then      '2-3-05 chg to function
                gCmcGenLg = -1
                Exit Function
            End If
        End If

        'Arrive at start date and end date based on the # days added to input start date

        '4-17-07 remove filter for air date (and # days) test; gen date & time is sufficient.
        'When a game extends across midnight and into the next day, that day was ignored with this code
        'slDate = RptSelLg!edcSelCFrom.Text   'Start date
        'slDateFrom = slDate
        'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
        'slSelection = "{ODF_One_Day_Log.odfAirDate} >= Date(" & slYear & "," & slMonth & "," & slDay & ") "
        'slDate = RptSelLg!edcSelCFrom1.Text '# of days
        'slDate = Format$(gDateValue(slDateFrom) + Val(slDate) - 1, "m/d/yy")
        'slDateTo = slDate
        'gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
        'slSelection = slSelection & " And {ODF_One_Day_Log.odfAirDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"

        '10-9-01
        'slTime = RptSelLg!edcSelCTo.Text   'Start time
        'slSelection = slSelection & " And Round({ODF_One_Day_Log.odfLocalTime}) >= " & Trim$(Str$(CLng(gTimeToCurrency(slTime, False))))
        'slTime = RptSelLg!edcSelCTo1.Text   'End Time
        'slSelection = slSelection & " And Round({ODF_One_Day_Log.odfLocalTime}) <= " & Trim$(Str$(CLng(gTimeToCurrency(slTime, True) - 1))) 'slTime


        gUnpackDate igNowDate(0), igNowDate(1), slDate
        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay

        gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slTime
        slSelection = slSelection & "  {ODF_One_Day_Log.odfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({ODF_One_Day_Log.odfGenTime} ) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))

        mSelectZones slSelection, 0           'setup Crystal formula for zone filtering
        'Vehicle selection
        slSelection = slSelection & " and {ODF_One_Day_Log.odfVefCode} = " & Trim$(str$(igcodes(0)))

        If sgRnfRptName = "L78" Then        'only L78 has been compiled for this new feature.  Once all reports using ODF has been compiled via crystal
                                            'this test can be removd.
            slSelection = slSelection & " and ({ODF_One_Day_Log.odfRafCode} = " & Trim$(str$(llRegionCode)) & " or {ODF_One_Day_Log.odfRafCode} = " & Trim$(str$(0)) & ")"
            'the formula "RegionCode" must also be put into all the logs that use ODF if the test for L78 is removed
            ilRet = mSendRegion(slRegionName)
        End If
        'slOr = " Or "
        'slSelection = slSelection & ")"

        'Include spots record types :  for short & long form logs include everything but avails
        'All other report include just spots
        'If sgRnfRptName = "L02" Or sgRnfRptName = "L03" Or sgRnfRptName = "L04" Or sgRnfRptName = "L05" Or sgRnfRptName = "L29" Or sgRnfRptName = "L30" Or sgRnfRptName = "C15" Or sgRnfRptName = "L72" Or sgRnfRptName = "L73" Or sgRnfRptName = "L75" Then  '3-30-01 C72 , 6-1-01 L73
        '3-30-01 C72 , 6-1-01 L73, 12-28-09 L79
        If sgRnfRptName = "L02" Or sgRnfRptName = "L03" Or sgRnfRptName = "L04" Or sgRnfRptName = "L05" Or sgRnfRptName = "L29" Or sgRnfRptName = "L30" Or sgRnfRptName = "C15" Or sgRnfRptName = "L72" Or sgRnfRptName = "L73" Or sgRnfRptName = "L75" Or sgRnfRptName = "L79" Then
            slSelection = "(" & slSelection & ")" & " And {ODF_One_Day_Log.odfType} <> 3 And {ODF_One_Day_Log.odfmnfSubFeed} = 0"
        '5-29-13 L83 duplicate of L73 , L85 duplicate of L75:  both have all comments removed for less clutter, and both show game info
        'ElseIf sgRnfRptName = "L83" Or sgRnfRptName = "L85" Then
        ElseIf sgRnfRptName = "L83" Or sgRnfRptName = "L85" Or sgRnfRptName = "C88" Then
            slSelection = "(" & slSelection & ")" & " And {ODF_One_Day_Log.odfType} <> 3 And {ODF_One_Day_Log.odfmnfSubFeed} = 0"
       '1/22/99 l19 changed to use avail name comment instead of "Other" event type comment (exclude Other event types in L19)
        '1/25/99 l19 changed back to use "Other" event types comment (include this type in filter)
        '2/9/99 L24 needs "Other" event types to simulate open avails
        '5-22-01 remove l30 from following test and include program events for l30 above
        '5-26-11 create special sports log
        '11-19-13 L86 is a copy of L81. L86 shows Game info only teams and date
        ElseIf sgRnfRptName = "L81" Or sgRnfRptName = "L86" Then             'special Event (sports) log to ignore all BB whether they are
                                                    'automatically created for a comml or they are in BB avail
            slSelection = "(" & slSelection & ")" & " And {ODF_One_Day_Log.odfType} <> 3 And {ODF_One_Day_Log.odfmnfSubFeed} = 0 and  InStr({ODF_One_Day_Log.odfBBDesc},'BB') = 0 and (({ODF_One_Day_Log.odfanfCode} > 0 and {ANF_Avail_Names.anfName} <> 'BB') or ({ODF_One_Day_Log.odfanfCode} = 0)) "
        ElseIf sgRnfRptName = "C12" Or sgRnfRptName = "L19" Or sgRnfRptName = "L20" Or sgRnfRptName = "L24" Or sgRnfRptName = "C18" Then              '12-20-00 (add c18)include comments and spots (1/8/99 )
            slSelection = "(" & slSelection & ")" & " And ({ODF_One_Day_Log.odfType} = 4 or {ODF_One_Day_Log.odfType} = 2) And {ODF_One_Day_Log.odfmnfSubFeed} = 0"
        Else
            slSelection = "(" & slSelection & ")" & " And {ODF_One_Day_Log.odfType} = 4 And {ODF_One_Day_Log.odfmnfSubFeed} = 0"
        End If


        'if L14 (5 day down- 2across & down) L12 - 5 days across, filter out Sat & sun records.
        If sgRnfRptName = "L14" Or sgRnfRptName = "L12" Then
            slSelection = "(" & slSelection & ")" & " And  (DayOfWeek ({ODF_One_Day_Log.odfAirDate}) >= 2 and DayOfWeek({ODF_One_Day_Log.odfAirDate}) <= 6)"
        ElseIf sgRnfRptName = "L13" Then     '2 days across (sa - su only)
            slSelection = "(" & slSelection & ")" & " And  (DayOfWeek ({ODF_One_Day_Log.odfAirDate}) = 1 or DayOfWeek({ODF_One_Day_Log.odfAirDate}) = 7)"
        End If
    End If
    '9-11-00 If sgRnfRptName = "C20" Or sgRnfRptName = "C21" Or sgRnfRptName = "C22" Or sgRnfRptName = "C23" Or sgRnfRptName = "C24" Or sgRnfRptName = "L43" Or sgRnfRptName = "L44" Then         'these are the logs/cps that use VOF table for options
    '9-12-00 L28 & L40 aded to obtain VOF to retrieve comments only to avoid the hard-coded messages
    'C01 added as customized 1-25-01

    'if log or cp is 70 or greater, its customized, as well as C01.  L28, L10, L40, L34 only has comments customized.
    'L89 (customized) copy of L10 added 2/22/19
    If (Val(Mid(sgRnfRptName, 2, 2)) >= 70) Or sgRnfRptName = "C01" Or sgRnfRptName = "L28" Or sgRnfRptName = "L40" Or sgRnfRptName = "L34" Or sgRnfRptName = "L10" Or sgRnfRptName = "L32" Or sgRnfRptName = "C89" Or sgRnfRptName = "L89" Then    '9-20-16 C89 copy of L10, excludes Politicals.  clients runs 2 "L" logs, have to make this a "C" log to get 3 logs printed
    'If sgRnfRptName = "C01" Or sgRnfRptName = "C70" Or sgRnfRptName = "C71" Or sgRnfRptName = "C72" Or sgRnfRptName = "C73" Or sgRnfRptName = "C74" Or sgRnfRptName = "C75" Or sgRnfRptName = "C76" Or sgRnfRptName = "C77" Or sgRnfRptName = "L70" Or sgRnfRptName = "L71" Or sgRnfRptName = "L28" Or sgRnfRptName = "L40" Or sgRnfRptName = "L34" Or sgRnfRptName = "L10" Or sgRnfRptName = "C78" Or sgRnfRptName = "C79" Or sgRnfRptName = "L72" Or sgRnfRptName = "C80" Then         '9-11-00 these are the logs/cps that use VOF table for options
        ilRet = mObtainVof(sgLogType, igcodes(0))
        If ilRet <> BTRV_ERR_NONE Then
            gCmcGenLg = -1
            Exit Function
        End If
    End If

    '12-15-04 send any additional formulas which crystal cant handle with subreports

    ilRet = mExceptionFormulas
    If ilRet <> 0 Then
        gCmcGenLg = -1
        Exit Function
    End If

    If Not gSetSelection(slSelection) Then
        gCmcGenLg = -1
        Exit Function
    End If
    '******* The following formulas must be in all Crystal reports even
    '        if they are not using them for generality
    '        StdYear, Week, InputDate, NumberDays
    'obtain the standard bcst end date
    'Show Std bdct Week # and year for for the Program show reference
    slDate = RptSelLg!edcSelCFrom.Text            'input date
    slDate = gObtainEndStd(slDate)              'get the std month end date to retrieve current year
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    slBaseDate = "1 / 15 /" & slYear
    slBaseDate = gObtainStartStd(slBaseDate)                'find current year's first week
    slDate = gObtainEndStd(slBaseDate)
    'gPackDate slDate, igNowDate(0), igNowDate(1)
    gPackDate slDate, ilStdStart(0), ilStdStart(1)

    If (ilStdStart(1) > 1999) Then                   'adjust for year 2000
        ilStdStart(1) = ilStdStart(1) - 2000
    Else
        ilStdStart(1) = ilStdStart(1) - 1900
    End If
    If Not gSetFormula("StdYear", ilStdStart(1)) Then
        gCmcGenLg = -1
        Exit Function
    End If
    'llDate1 = jan 1 of current year; llDate2 = Input date
    'gUnpackDateLong igNowDate(0), igNowDate(1), llDate1  'convert to long so math can be done
    llDate1 = gDateValue(slBaseDate)
    slDate = RptSelLg!edcSelCFrom.Text            'input date
    llDate2 = gDateValue(slDate)
    ilWeek = ((llDate2 - llDate1) \ 7) + 1
    If Not gSetFormula("Week", ilWeek) Then
        gCmcGenLg = -1
        Exit Function
    End If


    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    If Not gSetFormula("InputDate", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
        gCmcGenLg = -1
        Exit Function
    End If
    slDate = RptSelLg!edcSelCFrom1.Text       '# of days
    If Not gSetFormula("NumberDays", Val(slDate)) Then
        gCmcGenLg = -1
        Exit Function
    End If


    gCmcGenLg = 1
    'ttp 5260 use jpg as logo...here, if it exists, use bitmap because it's been customized
    '5676 not hardcoded c:\
'    If Dir("C:\csi\RptLogo.bmp") > "" Then
'        gSetFormula "LogoLocation", "'C:\csi\RptLogo.bmp'"
'    End If


'8-18-14  do not use root drive any longer, Logo location is now used.  the customized logo was set in rptsellg.frm
'    If Dir(sgRootDrive & "csi\RptLogo.bmp") > "" Then
'        gSetFormula "LogoLocation", "'" & sgRootDrive & "csi\RptLogo.bmp'"
'    End If

    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGenReportLg                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize reports and do      *
'*                      Validity checking, then call   *
'*                      to Crystal for proper report   *
'*                      Validity checking in gCmcDone  *
'*                      should be moved to this rtn    *
'*******************************************************
Function gGenReportLg() As Integer
    Dim slCrystalName As String
    Dim ilLogExpInx As Integer
    Dim ilRet As Integer
    Dim ilError As Integer
    Dim ilPos As Integer
    Dim slFileName As String
    Dim fs As New FileSystemObject
   
    If Not igUsingCrystal Then
        gGenReportLg = True
        Exit Function
    End If
    
    '6-5-19 altho this module is loaded in Reports, it is not called for any log or report that generates thru this code. Other reports go thru
    'other modules within rptsellg & rptselevfy.  This is called thru Logs.
    
    '2-12-15 use the SaveReportPath if defined in traffic.ini for the LogGenExport files
    'it has been initialized to use sgSaveRptPath, or overridden with sgLogGenSavePath if defined in traffic.ini
    sgSaveFilePathToUse = sgLogGenExportPath
   
    ilLogExpInx = gBinarySearchVefForLogExport(igcodes(0))
    If ilLogExpInx <> -1 Then
        sgSaveFilePathToUse = Trim$(tgLogExportLoc(ilLogExpInx).sExportPath)     'vehicle export folder
    End If
    
'    'it has been initialized to use sgSaveRptPath, or overridden with sgLogGenSavePath if defined in traffic.ini
'    sgSaveFilePathToUse = sgLogGenExportPath
   
    '6-4-19 Create an export folder if none exists
    On Error Resume Next
    If Not fs.FolderExists(Left$(sgSaveFilePathToUse, Len(sgSaveFilePathToUse) - 1)) Then
        slFileName = Left$(sgSaveFilePathToUse, Len(sgSaveFilePathToUse))
        fs.CreateFolder (slFileName)
        If Not fs.FolderExists(slFileName) Then
            gMsgBox ("Counterpoint was unable to create the folder: " & slFileName & " Please have your IT manager add this folder; default Export Path used."), vbOKOnly + vbApplicationModal, "gObtainIniValue"
            sgSaveFilePathToUse = Trim$(sgLogGenExportPath)
            tgLogExportLoc(ilLogExpInx).sExportPath = Trim$(sgLogGenExportPath)   'in case multiple logs (time zones) generated, show mesg only once for the vehicle
        End If
    End If
    
    slCrystalName = Trim$(sgRnfRptName) & Trim$(".rpt")
    If Not gOpenPrtJob(slCrystalName) Then
        gGenReportLg = False
        Exit Function
    End If
    gGenReportLg = True
    Exit Function
    
End Function
'
'
'       mObtainVof - retrieve the Vehicle Log options table
'           which contains options to show/not show on log or CP
'           <return> 0 = OK 1, error
'           6/19/00
Function mObtainVof(slLogType As String, ilVefCode As Integer) As Integer
Dim ilRet As Integer
Dim ilError As Integer
    ilError = 0
    hmVof = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVof, "", sgDBPath & "Vof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        MsgBox "VOF file missing"
        ilRet = btrClose(hmVof)
        btrDestroy hmVof
        mObtainVof = ilRet
        Exit Function
    End If

    tmVofSrchKey.iVefCode = ilVefCode
    tmVofSrchKey.sType = slLogType
    ilRet = btrGetEqual(hmVof, tgVof, Len(tgVof), tmVofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        tgVof.iNoDaysCP = 30
        tgVof.sShowLen = "Y"
        tgVof.sShowProduct = "Y"
        tgVof.sShowCreative = "Y"
        tgVof.sShowISCI = "Y"
        tgVof.sShowDP = "Y"
        tgVof.sShowAirTime = "Y"
        tgVof.sShowAirLine = "Y"
        tgVof.sShowHour = "Y"
        tgVof.sSkipPage = "Y"
        tgVof.iLoadFactor = 1
    End If
    If Not gSetFormula("ReturnDays", tgVof.iNoDaysCP) Then
        ilError = -1
    End If
    If Not gSetFormula("ShowLength", "'" & tgVof.sShowLen & "'") Then
        ilError = -1
    End If
    If Not gSetFormula("ShowProd", "'" & tgVof.sShowProduct & "'") Then
        ilError = -1
    End If
    If Not gSetFormula("ShowTitle", "'" & tgVof.sShowCreative & "'") Then
        ilError = -1
    End If
    If Not gSetFormula("ShowISCI", "'" & tgVof.sShowISCI & "'") Then
        ilError = -1
    End If
    If Not gSetFormula("ShowDPDesc", "'" & tgVof.sShowDP & "'") Then
        ilError = -1
    End If
    If Not gSetFormula("ShowDetTimeLine", "'" & tgVof.sShowAirTime & "'") Then
        ilError = -1
    End If
    If Not gSetFormula("ShowAiredLine", "'" & tgVof.sShowAirLine & "'") Then
        ilError = -1
    End If
    If Not gSetFormula("ShowHour", "'" & tgVof.sShowHour & "'") Then
        ilError = -1
    End If
    If Not gSetFormula("SkipPage", "'" & tgVof.sSkipPage & "'") Then
        ilError = -1
    End If
    If tgVof.iLoadFactor = 0 Then
        tgVof.iLoadFactor = 1
    End If
    If Not gSetFormula("RepeatFactor", tgVof.iLoadFactor) Then
        ilError = -1
    End If
    If ilError <> 0 Then
        mObtainVof = ilError
        MsgBox "Error in Crystal VOF formula, mObtainVof"
        ilRet = btrClose(hmVof)
        btrDestroy hmVof
        Exit Function
    Else
        ilRet = btrClose(hmVof)
        btrDestroy hmVof
    End If
End Function
'
'
'
'                       mSelectZones - send Crystal formula to filter
'                       out the proper time zones to print for the ODF
'                       prepass file (code duplicated for GRF pre-pass file)
'
'
'                       <input & output> slselection - Crystal formula text
'                               ilFile: 0 = setup selection for ODF file
'                                        1 = setup selection for GRF file
'                                        2 = setup selection for SVR file
'                                        3 = setup selection for CBF file
'                       Created :  1/29/98   D.Hosaka
'
'
Sub mSelectZones(slSelection As String, ilFile As Integer)
Dim ilUseZone As Integer
Dim ilVpfIndex As Integer
Dim ilLoop As Integer
Dim slOr As String
        ilUseZone = False
        ilVpfIndex = -1
        'For ilLoop = 0 To UBound(tgVpf) Step 1
        '    If igcodes(0) = tgVpf(ilLoop).iVefKCode Then
            ilLoop = gBinarySearchVpf(igcodes(0))
            If ilLoop <> -1 Then
                ilVpfIndex = ilLoop
                'If tgVpf(ilLoop).sGZone(1) = " " Or tgVpf(ilLoop).sGZone(1) = "" Then
                If tgVpf(ilLoop).sGZone(0) = " " Or tgVpf(ilLoop).sGZone(0) = "" Then
        '            Exit For                        'no zones used
                Else
                    ilUseZone = True
                    ilVpfIndex = ilLoop
        '            Exit For
                End If
            End If
        'Next ilLoop
        If ilVpfIndex >= 0 Then
            If igZones = 0 Then                     'all requested or if there isnt any zones defined in vpf, force to none
                If ilUseZone Then                   'using zones, possibly get all depending which report
                    ' slSelection = "(" & slSelection & ") "
                    'slOr = ""
                    'only certain reports need the use of zones
                    'If Trim$(sgRnfRptName) = "L01" Then
                        'take everything (use no filters)
                    'If Trim$(sgRnfRptName) = "L02" Or Trim$(sgRnfRptName) = "L03" Or Trim$(sgRnfRptName) = "L04" Or Trim$(sgRnfRptName) = "L05" Then
                        'take everything (use no filters)
                    'If sgRnfRptName = "C01" Or sgRnfRptName = "C04" Or sgRnfRptName = "C05" Or sgRnfRptName = "C06" Or sgRnfRptName = "L06" Then
                    If ilFile = 0 Then                  'ODF file selection
                        If (Left$(sgRnfRptName, 1) = "C") Or (sgRnfRptName = "L06") Then
                            'these CP should only print 1 set regardless of how many time zones are defined
                            '7-17-09 CR 2008 needs a trim to test blanks in the zone
                            'slSelection = slSelection & " and Trim({ODF_One_Day_Log.odfZone}[1 to 3]) = " & "'" & tgVpf(ilVpfIndex).sGZone(1) & "'"
                            slSelection = slSelection & " and Trim({ODF_One_Day_Log.odfZone}[1 to 3]) = " & "'" & tgVpf(ilVpfIndex).sGZone(0) & "'"
                            'slSelection = slSelection & " and {ODF_One_Day_Log.odfZone}[1 to 3] = " & "'" & tgVpf(ilVpfIndex).sGZone(1) & "'"
                        End If
                    ElseIf ilFile = 1 Then              'GRF file
                        If (Left$(sgRnfRptName, 1) = "C") Or (sgRnfRptName = "L06") Or (sgRnfRptName = "C72") Or (sgRnfRptName = "C73") Or (sgRnfRptName = "C84") Then
                            'these CP should only print 1 set regardless of how many time zones are defined
                            '7-17-09 CR 2008 needs a trim to test blanks
                            'slSelection = slSelection & " and Trim({GRF_Generic_Report.grfBktType}[1 to 1]) = " & "'" & Mid$(tgVpf(ilVpfIndex).sGZone(1), 1, 1) & "'"
                            slSelection = slSelection & " and Trim({GRF_Generic_Report.grfBktType}[1 to 1]) = " & "'" & Mid$(tgVpf(ilVpfIndex).sGZone(0), 1, 1) & "'"
                        End If
                    ElseIf ilFile = 2 Then                                'SVR file
                        If (Left$(sgRnfRptName, 1) = "C") Or (sgRnfRptName = "L06") Then
                            'these CP should only print 1 set regardless of how many time zones are defined
                            '7-17-09 CR 2008 needs a trim to test blanks in the zone
                            'slSelection = slSelection & " and Trim({SVR_7Day_Report.svrZone}[1 to 3]) = " & "'" & tgVpf(ilVpfIndex).sGZone(1) & "'"
                            slSelection = slSelection & " and Trim({SVR_7Day_Report.svrZone}[1 to 3]) = " & "'" & tgVpf(ilVpfIndex).sGZone(0) & "'"
                        End If
                    Else                            'CBF
                        'no filters for CBF if using zones, take all
                    End If
                Else                                'not using zones, get blank zones
                    If ilFile = 0 Then                  'ODF file
                        slSelection = "(" & slSelection & ") " & " And ("
                        slOr = ""
                        '7-17-09 Crystal 2008 needs trim on the testing of blank field
                        slSelection = slSelection & slOr & "Trim({ODF_One_Day_Log.odfZone}[1 to 3]) = '   '"
                        slOr = " Or "
                    ElseIf ilFile = 1 Then              'GRF selection
                        slSelection = "(" & slSelection & ") " & " And ("
                        slOr = ""
                        '7-17-09 Crystal 2008 needs trim on the testing of blank field
                        slSelection = slSelection & slOr & "Trim({GRF_Generic_Report.grfBktType}) = '   '"
                        slOr = " Or "
                    ElseIf ilFile = 2 Then                             'svr SELECTION
                        slSelection = "(" & slSelection & ") " & " And ("
                        slOr = ""
                        slSelection = slSelection & slOr & "Trim({SVR_7Day_Report.svrZone}[1 to 3]) = '   '"
                        slOr = " Or "
                    Else
                        slSelection = "(" & slSelection & ") " & " And ("
                        slOr = ""
                        slSelection = slSelection & slOr & "Trim({CBF_Contract_BR.cbfResort}[1 to 3]) = '   '"
                        slOr = " Or "
                    End If
                End If
            End If
            If igZones > 0 Then             '"all" already taken care of
                If ilFile = 0 Then          'ODF file selection
                    If igZones = 1 Then
                        slSelection = slSelection & " and {ODF_One_Day_Log.odfZone} = 'EST'"
                    ElseIf igZones = 2 Then
                        slSelection = slSelection & " and {ODF_One_Day_Log.odfZone} = 'CST'"
                    ElseIf igZones = 3 Then
                        slSelection = slSelection & " and {ODF_One_Day_Log.odfZone} = 'MST'"
                    ElseIf igZones = 4 Then
                        slSelection = slSelection & " and {ODF_One_Day_Log.odfZone} = 'PST'"
                    End If
                ElseIf ilFile = 1 Then          'GRF file selection
                    If igZones = 1 Then
                        slSelection = slSelection & " and {GRF_Generic_Report.grfBktType} = 'E'"
                    ElseIf igZones = 2 Then
                        slSelection = slSelection & " and {GRF_Generic_Report.grfBktType} = 'C'"
                    ElseIf igZones = 3 Then
                        slSelection = slSelection & " and {GRF_Generic_Report.grfBktType} = 'M'"
                    ElseIf igZones = 4 Then
                        slSelection = slSelection & " and {GRF_Generic_Report.grfBktType} = 'P'"
                    End If
                ElseIf ilFile = 2 Then                          'SVR file selection
                    If igZones = 1 Then
                        slSelection = slSelection & " and {SVR_7Day_Report.svrZone} = 'EST'"
                    ElseIf igZones = 2 Then
                        slSelection = slSelection & " and {SVR_7Day_Report.svrZone} = 'CST'"
                    ElseIf igZones = 3 Then
                        slSelection = slSelection & " and {SVR_7Day_Report.svrZone} = 'MST'"
                    ElseIf igZones = 4 Then
                        slSelection = slSelection & " and {SVR_7Day_Report.svrZone} = 'PST'"
                    End If
                Else           'CBF file selection
                    If igZones = 1 Then
                        slSelection = slSelection & " and Trim({CBF_Contract_BR.cbfResort}[1 to 3]) = 'EST'"
                    ElseIf igZones = 2 Then
                        slSelection = slSelection & " and Trim({CBF_Contract_BR.cbfResort}[1 to 3])  = 'CST'"
                    ElseIf igZones = 3 Then
                        slSelection = slSelection & " and Trim({CBF_Contract_BR.cbfResort}[1 to 3]) = 'MST'"
                    ElseIf igZones = 4 Then
                        slSelection = slSelection & " and Trim({CBF_Contract_BR.cbfResort}[1 to 3])  = 'PST'"
                    End If
                End If
            End If
            'slSelection = slSelection & ")"
        End If
End Sub
'           12-15-04
'           ilret = mExceptionFormulas() as integer
'           Crystal formulas sent to selective logs  because of the Site Preference
'           parameters that are needed.
'           When there is grouping on a Site Preference field (which was previously
'           put into a subreport), that site preference field cannot be grouped.
'           For example, if UseCartNo is tested to see if sorting should occur on
'           the cart # or not, the UseCartNo field is not recognized.
'           Therefore, the site preference parameter must be sent as a formula to Crystal
'
'           <input>  None
'           <output> None
'           <Return> 0 = OK, 1 error in formula call
Public Function mExceptionFormulas() As Integer
    mExceptionFormulas = 0          'ok

    If sgRnfRptName = "L18" Then
        If Not gSetFormula("SiteUseCart", "'" & tgSpf.sUseCartNo & "'") Then
            mExceptionFormulas = -1
        End If
    End If
End Function
'
'           determine if any zones, or only 1 selected, do not show
'           the zone desription in the header of the log/cp.
'           Pass formula indicating that to Crystal
'           2-3-05
'
Public Function mShowzone() As Integer

    mShowzone = 0
    If igZones <> 0 Then             'all time zones, show time zone in header
        If Not gSetFormula("ShowZone", "'N'") Then
            mShowzone = -1
            Exit Function
        End If
    Else                            'selected single time zone, dont show it in the header
        If Not gSetFormula("ShowZone", "'Y'") Then
            mShowzone = -1
            Exit Function
        End If
    End If
End Function
'
'           mSendRegion - send the region code or 0 if full network to ODF
'           generated logs
'       <input>  split network region code, else 0 for full network
'               used to filter the type of records for log
Public Function mSendRegion(slRegionName As String) As Integer
    mSendRegion = 0
    'obtain region name for log

    If Not gSetFormula("RegionName", "'" & Trim$(slRegionName) & "'") Then
        mSendRegion = -1
        Exit Function
    End If
End Function
