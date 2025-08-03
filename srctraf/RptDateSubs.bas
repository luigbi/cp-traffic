Attribute VB_Name = "RptDateSubs"
'
Option Explicit
Option Compare Text

'
'
'                   gSetupDates - Build array of start and end dates
'                                 to gather history and contract $
'                   <input> ilCorpStd - 1 = corp, 2 = std, 3 = calendar
'                           llPacingDate - effective pacing date (currently for Billed & Booked)
'                           ilMonth - sometimes input is by quarter, need the month # to calculate all the monthly dates to process
'                   <output> llStdStartDates - array of 13 start dates
'                           (i.e. billed and booked has 13 start dates
'                                 Sales comparisons as many start dates,
'                                 that is requested)
'                            llLastbilled - From site pref, last date invoiced
'                            ilLastbilledInx - Index to last complete month billed
'
'           11-21-05 add effective pacing date and all places that reference it.  Pacing added to
'           Billed & Booked
'           7-3-08 When obtaining dates for the corporate calendar, the year entered is the Corporate year,
'           not the actual year when the fiscal year starts.  for example:  ABC fiscal year 2008 starts Oct 2007 thru Sep 2008
'           11-10-08 handle wrap around of year for corporate calendar
Sub gSetupBOBDates(ilCorpStd As Integer, llStdStartDates() As Long, llLastBilled As Long, ilLastBilledInx As Integer, llPacingDate As Long, ilMonth As Integer, Optional blGetCalLastBilled = False)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************
Dim slStr As String
Dim ilLoop As Integer
Dim llDate As Long
Dim ilYearInx As Integer
Dim ilStartMonth As Integer
Dim ilNextDate As Integer
Dim llSaveLastDate As Long
    'Determine calendar month requested, and retrieve all History and Receivables
    'records that fall within the beginning of the cal year and end of calendar month requested
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr       'convert last bdcst billing date to string
    If blGetCalLastBilled Then                                              '1-13-21 get the true last billing date for calendar bill cycle
        gUnpackDate tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), slStr
    End If
    If Trim$(slStr) = "" Then
        slStr = "1/1/1975"
    End If
    llLastBilled = gDateValue(slStr)            'convert last month billed to long

    'ilLoop = (igMonthOrQtr - 1) * 3 + 1
    ilLoop = ilMonth                 'starting month (starting qtrs have been converted to month #)

    If ilCorpStd = 1 Then                                       'build array of corp start months
        'Determine what month the fiscal year starts
        ilYearInx = gGetCorpCalIndex(igYear)
        If ilYearInx < 0 Then           'corp calendar doesnt exist
            ilLastBilledInx = 1
            gMsgBox "Corporate Calendar does not exist", vbOKOnly + vbExclamation, "Corporate Calendar"
            Exit Sub
        End If
        ilNextDate = 1

        '8-13-08 comment this code out, determine the corp dates based on the user entered month
        'relative to start of the corporate year.  I.e. if Corp year starts in Oct, then user
        'user entered corp 2007, month 1 would need to obtain Oct 2006.
'        ilStartMonth = 0                            'month inx into corp table
'        For ilLoop = tgMCof(ilYearInx).iStartMnthNo To tgMCof(ilYearInx).iStartMnthNo + 11
'            If ilLoop > 12 Then        'wrap around to start of next calendar year
'                ilStartMonth = ilStartMonth + 1
'                If ilLoop - 12 = igMonthOrQtr Then
'                    Exit For
'                End If
'             Else
'                ilStartMonth = ilStartMonth + 1
'                If ilLoop = igMonthOrQtr Then
'                    Exit For
'                End If
'            End If
'        Next ilLoop
'        ilLoop = ilStartMonth
        For ilStartMonth = ilLoop To 12             'start from requested qtr of corp cal, may wrap around to next year if not starting at 1st qtr
            gUnpackDateLong tgMCof(ilYearInx).iStartDate(0, ilStartMonth - 1), tgMCof(ilYearInx).iStartDate(1, ilStartMonth - 1), llStdStartDates(ilNextDate)
            gUnpackDateLong tgMCof(ilYearInx).iEndDate(0, ilStartMonth - 1), tgMCof(ilYearInx).iEndDate(1, ilStartMonth - 1), llSaveLastDate
            ilNextDate = ilNextDate + 1
        Next ilStartMonth
        '11-10-08 always test the wraparound
        'If ilNextDate <> 12 Then                'corp year wrap around; started from qtr other than 1
            ilYearInx = gGetCorpCalIndex(igYear + 1)
            If ilYearInx = -1 Then          '8-18-09
                MsgBox "Corporate year " & str$(igYear + 1) & " must be defined in Site"
                ilLastBilledInx = -1
                Exit Sub
            End If
            For ilStartMonth = 1 To (12 - ilNextDate) + 1       'do the next qtr from the following year on wrap-around
                gUnpackDateLong tgMCof(ilYearInx).iStartDate(0, ilStartMonth - 1), tgMCof(ilYearInx).iStartDate(1, ilStartMonth - 1), llStdStartDates(ilNextDate)
                gUnpackDateLong tgMCof(ilYearInx).iEndDate(0, ilStartMonth - 1), tgMCof(ilYearInx).iEndDate(1, ilStartMonth - 1), llSaveLastDate
                ilNextDate = ilNextDate + 1
            Next ilStartMonth
        'End If
        'get the 13th month start date.  Increment the last saved end date by 1 day
        llStdStartDates(13) = llSaveLastDate + 1

        'For ilLoop = 1 To 13 Step 1
        '    slStr = gObtainStartCorp(slStr, True)
        '    llStdStartDates(ilLoop) = gDateValue(slStr)
        '    slStr = gObtainEndCorp(slStr, True)
        '    llDate = gDateValue(slStr) + 1                      'increment for next month
        '    slStr = Format$(llDate, "m/d/yy")
        'Next ilLoop
    ElseIf ilCorpStd = 2 Then                                     'build array of std start months
        slStr = Trim$(str$(ilLoop)) & "/15/" & Trim$(str$(igYear))      'format xx/xx/xxxx
        For ilLoop = 1 To 13 Step 1
            slStr = gObtainStartStd(slStr)
            llStdStartDates(ilLoop) = gDateValue(slStr)
            slStr = gObtainEndStd(slStr)
            llDate = gDateValue(slStr) + 1                      'increment for next month
            slStr = Format$(llDate, "m/d/yy")
        Next ilLoop
    Else                'calendar
        slStr = Trim$(str$(ilLoop)) & "/15/" & Trim$(str$(igYear))      'format xx/xx/xxxx
        For ilLoop = 1 To 13 Step 1
            slStr = gObtainStartCal(slStr)
            llStdStartDates(ilLoop) = gDateValue(slStr)
            slStr = gObtainEndCal(slStr)
            llDate = gDateValue(slStr) + 1                      'increment for next month
            slStr = Format$(llDate, "m/d/yy")
        Next ilLoop

    End If
    'determine what month index the actual is (versus the future dates)
    'assume everything in the past if by std; if by corp it wont use
    'history files since billing is done all by std
    If ilCorpStd = 1 Then                           'corp
        ilLastBilledInx = 1                         'retrieve everything from contracts
    Else
        If llPacingDate = 0 Then                    'no pacing option
            'std or cal
            ilLastBilledInx = 12                    'assume everything in the past
            If llLastBilled < llStdStartDates(1) Then
                ilLastBilledInx = 1                     'everything in the future
            End If
            'determine what month, if any, is in the past.  Set ilLastBilledInx to the last month billed
            For ilLoop = 1 To 12 Step 1
                If llLastBilled > llStdStartDates(ilLoop) And llLastBilled < llStdStartDates(ilLoop + 1) Then
                    ilLastBilledInx = ilLoop
                    Exit For
                End If
            Next ilLoop
        Else                            '11-21-05 pacing Billed &Booked
            ilLastBilledInx = 1          'retrieve everything from contracts
            '8-4-08 if pacing, last billed date isnt applicable.  Need ANs and all contracts
            'llLastBilled = llStdStartDates(1) - 1       'force last date billed to first period minus 1
        End If
    End If
End Sub

'
'
'
'             gFindMaxDates - Determine the earliest and latest dates of the contract
'                   header.  If Differences option, see if previous revision has an earlier
'                   start date and/or later end date
'
'           <input> slStartDate - Contract header start date
'                   slEnd Date - contract header end date
'                   ilShowStdQtr - 12-19-20 0 = std, 1 = cal, 2 = corp: was true if showing quarters by std months
'                   blDefaultToQtr - true to start on qtrs, false to start on whatever month the start date is (i.e. order audit) 12-19-20
'           <output> llchfStart - Start date of quarter (based on contract header start date)
'                   llchfEnd - End date of quarter (based on contract header end date)
'                   ilCurrTotalMonths - Total months of order
'                   llStdStartDates - start dates of each std month (to summarize $ and spots)
'                   ilYear - corp or std Year that this contract belongs in
'                   ilStartMonth - Start month of corp or std year that this contracts starts in
'
'
Sub gFindMaxDates(slStartDate As String, slEndDate As String, llChfStart As Long, llChfEnd As Long, ilCurrStartQtr() As Integer, ilCurrTotalMonths As Integer, llStdStartDates() As Long, ilShowStdQtr As Integer, ilYear As Integer, ilStartMonth As Integer, blDefaultToQtr As Boolean)
Dim llStartQtr As Long
Dim llEndQtr As Long
Dim slDay As String
Dim ilLoop3 As Integer
Dim llFltStart As Long
Dim slStr As String
Dim ilLoop As Integer
'    If ilShowStdQtr Then
    If ilShowStdQtr = 0 Then                    'std dates
        'determine current qtr start date
        llStartQtr = gDateValue(gObtainStartStd(slStartDate))    'get the std start date of the contract
        llEndQtr = gDateValue(gObtainEndStd(slEndDate))       'get the std end date of the contract
        'Save the earliest/latest dates from contr header
        llChfStart = gDateValue(slStartDate)
        llChfEnd = gDateValue(slEndDate)
        'Calculate the total number of std airng months for this order
        'ilCurrStartQtr(0) = 0                       'init to place start date of starting qtr of order
        'ilCurrStartQtr(1) = 0
        'ilCurrTotalMonths = 0
        slDay = gObtainEndStd(Format$(llStartQtr, "m/d/yy"))
        
        If blDefaultToQtr = True Then       '12-19-20 earliest date must fall on a qtr start date
            ilLoop3 = Month(Format$(gDateValue(slDay), "m/d/yy"))
            'Calculate the first qtr, backup months until start of a qtr month
            Do While ilLoop3 <> 1 And ilLoop3 <> 4 And ilLoop3 <> 7 And ilLoop3 <> 10
                slStr = gObtainStartStd(slDay)      'start date of month to backup to previous month
                llFltStart = gDateValue(slStr) - 1
                slDay = gObtainEndStd(Format$(llFltStart, "m/d/yy"))
                ilLoop3 = Month(Format$(gDateValue(slDay), "m/d/yy"))
            Loop
        End If
        'slDay contains start of contract's qtr
        gPackDate slDay, ilCurrStartQtr(0), ilCurrStartQtr(1)   'save this contracts start month to store in cbf
        llStartQtr = gDateValue(slDay)
        Do While llStartQtr <= llEndQtr
            slStr = gObtainEndStd(Format$(llStartQtr, "m/d/yy"))
            'Determine what the starting qtr and month is for this order
            ilLoop3 = Month(Format$(gDateValue(slStr), "m/d/yy"))
            llStartQtr = gDateValue(slStr) + 1
            ilCurrTotalMonths = ilCurrTotalMonths + 1         'accum total # of airing std months   (to be stored in cbf)
        Loop
        'Calc # std month start and date dates to total the week into
        'build array of 13 start standard dates
        For ilLoop = 1 To 37 Step 1
            slDay = gObtainStartStd(slDay)
            llStdStartDates(ilLoop) = gDateValue(slDay)
            slDay = gObtainEndStd(slDay)
            llStartQtr = gDateValue(slDay) + 1                      'increment for next month
            slDay = Format$(llStartQtr, "m/d/yy")
        Next ilLoop
    ElseIf ilShowStdQtr = 1 Then                                    'calendar dates
     'determine current qtr start date
        llStartQtr = gDateValue(gObtainStartCal(slStartDate))    'get the std start date of the contract
        llEndQtr = gDateValue(gObtainEndCal(slEndDate))       'get the std end date of the contract
        'Save the earliest/latest dates from contr header
        llChfStart = gDateValue(slStartDate)
        llChfEnd = gDateValue(slEndDate)
        'Calculate the total number of std airng months for this order
        'ilCurrStartQtr(0) = 0                       'init to place start date of starting qtr of order
        'ilCurrStartQtr(1) = 0
        'ilCurrTotalMonths = 0
        slDay = gObtainEndCal(Format$(llStartQtr, "m/d/yy"))
        
        If blDefaultToQtr = True Then       '12-19-20 earliest date must fall on a qtr start date
            ilLoop3 = Month(Format$(gDateValue(slDay), "m/d/yy"))
            'Calculate the first qtr, backup months until start of a qtr month
            Do While ilLoop3 <> 1 And ilLoop3 <> 4 And ilLoop3 <> 7 And ilLoop3 <> 10
                slStr = gObtainStartCal(slDay)      'start date of month to backup to previous month
                llFltStart = gDateValue(slStr) - 1
                slDay = gObtainEndCal(Format$(llFltStart, "m/d/yy"))
                ilLoop3 = Month(Format$(gDateValue(slDay), "m/d/yy"))
            Loop
        End If
        'slDay contains start of contract's qtr
        gPackDate slDay, ilCurrStartQtr(0), ilCurrStartQtr(1)   'save this contracts start month to store in cbf
        llStartQtr = gDateValue(slDay)
        Do While llStartQtr <= llEndQtr
            slStr = gObtainEndCal(Format$(llStartQtr, "m/d/yy"))
            'Determine what the starting qtr and month is for this order
            ilLoop3 = Month(Format$(gDateValue(slStr), "m/d/yy"))
            llStartQtr = gDateValue(slStr) + 1
            ilCurrTotalMonths = ilCurrTotalMonths + 1         'accum total # of airing cal months   (to be stored in cbf)
        Loop
        'Calc # std month start and date dates to total the week into
        'build array of 13 start calendar dates
        For ilLoop = 1 To 37 Step 1
            slDay = gObtainStartCal(slDay)
            llStdStartDates(ilLoop) = gDateValue(slDay)
            slDay = gObtainEndCal(slDay)
            llStartQtr = gDateValue(slDay) + 1                      'increment for next month
            slDay = Format$(llStartQtr, "m/d/yy")
        Next ilLoop
    Else                                                            'corporate dates (unused for now)
        'determine current qtr start date
        llStartQtr = gDateValue(gObtainStartCorp(slStartDate, False))   'get the std start date of the contract
        llEndQtr = gDateValue(gObtainEndCorp(slEndDate, False))      'get the std end date of the contract
        'Save the earliest/latest dates from contr header
        llChfStart = gDateValue(slStartDate)
        llChfEnd = gDateValue(slEndDate)

        'slDay = gObtainEndCorp(Format$(llStartQtr + 14, "m/d/yy"), False)      'get to middle of the month to find its true month  #
        slDay = Format$(llStartQtr + 14, "m/d/yy")
        ilLoop3 = Month(Format$(gDateValue(slDay), "m/d/yy"))
        'Calculate the first qtr, backup months until start of a qtr month
        Do While ilLoop3 <> 1 And ilLoop3 <> 4 And ilLoop3 <> 7 And ilLoop3 <> 10
            slStr = gObtainStartCorp(slDay, False)     'start date of month to backup to previous month
            'llFltStart = gDateValue(slStr) - 14
            slDay = Format$((gDateValue(slStr) - 14), "m/d/yy")
            'slDay = gObtainEndCorp(Format$(llFltStart, "m/d/yy"), False)
            ilLoop3 = Month(Format$(gDateValue(slDay), "m/d/yy"))
        Loop
        'slDay contains start of contract's qtr

        gPackDate slDay, ilCurrStartQtr(0), ilCurrStartQtr(1)   'save this contracts start month to store in cbf
        slDay = gObtainStartCorp(slDay, False)
        llStartQtr = gDateValue(slDay)
        Do While llStartQtr <= llEndQtr
            slStr = gObtainEndCorp(Format$(llStartQtr, "m/d/yy"), False)
            'Determine what the starting qtr and month is for this order
            ilLoop3 = Month(Format$(gDateValue(slStr), "m/d/yy"))
            llStartQtr = gDateValue(slStr) + 1
            ilCurrTotalMonths = ilCurrTotalMonths + 1         'accum total # of airing std months   (to be stored in cbf)
        Loop

        'Calc # std month start and date dates to total the week into
        'build array of 13 start corp dates
        For ilLoop = 1 To 37 Step 1         '12-29-06
            slDay = gObtainStartCorp(slDay, False)
            llStdStartDates(ilLoop) = gDateValue(slDay)
            slDay = gObtainEndCorp(slDay, False)
            llStartQtr = gDateValue(slDay) + 1                      'increment for next month
            slDay = Format$(llStartQtr, "m/d/yy")
        Next ilLoop
    End If
    slStr = Format$(llStdStartDates(1), "m/d/yy")
    gGetYearStartMo ilShowStdQtr, slStr, ilYear, ilStartMonth
End Sub
'
'
'           gGetYearStartMo - get the Year and Starting Month of the Corp
'               calendar or standard bdcst year based on any date
'
'           <input> slInpDate - Date string to determine year and start month
'                   ilShowStdQtr - 12-19-20 0 = std, 1 = cal, 2 = corp: was true if using std qtrs, else false
'           <output> ilYear -  year of corp calendar or std bdcst year
'                   ilStartMonth - start month of corp cal or std bdcst
'
Sub gGetYearStartMo(ilShowStdQtr As Integer, slInpDate As String, ilYear As Integer, ilStartMonth As Integer)
Dim ilLoop As Integer
Dim llStartDate As Long
Dim llEndDate As Long
Dim llInpDate As Long
Dim slTempDate As String
Dim slYear As String
Dim slMonth As String
Dim slDay As String
'    If ilShowStdQtr Then                'using the std bdcst month for output
    If ilShowStdQtr = 0 Then                '12-19-20 using the std bdcst month for output
        ilStartMonth = 1                'standard always starts with Jan
        slTempDate = gObtainEndStd(slInpDate)        'get std bdcst end date
        gObtainYearMonthDayStr slTempDate, True, slYear, slMonth, slDay
        ilYear = Val(slYear)
        Exit Sub
    ElseIf ilShowStdQtr = 1 Then           'using calendar months
        ilStartMonth = 1
        slTempDate = gObtainEndCal(slInpDate)        'get std bdcst end date
        gObtainYearMonthDayStr slTempDate, True, slYear, slMonth, slDay
        ilYear = Val(slYear)
        Exit Sub
    Else
        llInpDate = gDateValue(slInpDate)
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            gUnpackDateLong tgMCof(ilLoop).iStartDate(0, 0), tgMCof(ilLoop).iStartDate(1, 0), llStartDate
            gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iStartDate(1, 11), llEndDate
            If llInpDate >= llStartDate And llInpDate <= llEndDate Then
                ilYear = tgMCof(ilLoop).iYear
                ilStartMonth = tgMCof(ilLoop).iStartMnthNo
                Exit Sub
            End If
        Next ilLoop
    End If
End Sub
'*********************************************************************
'
'               gCurrDateTime
'               <Output> slDate - current date (xx/xx/xx)
'                        slTime - current time (xx:xx:xxa/p)
'                        Some routines may not use these return values
'                        slMonth - xx  (1-12)
'                        slDay - XX  (1-31)
'                        slYear - xxxx (19xx-20xx)
'               obtain system current date and time and return it
'               in string format
'
'               Created:  7/3/96
'*********************************************************************
Sub gCurrDateTime(slDate As String, slTime As String, slMonth As String, slDay As String, slYear As String)
    slDate = Format$(gNow(), "m/d/yy")
    gPackDate slDate, igNowDate(0), igNowDate(1)
    slTime = Format$(gNow(), "h:mm:ssAM/PM")
    gPackTime slTime, igNowTime(0), igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
End Sub
'*********************************************************************
'
'               gRandomDateTime
'               <Output> slDate - current date (xx/xx/xx)
'                        slTime - current time (xx:xx:xxa/p)
'                        Some routines may not use these return values
'                        slMonth - xx  (1-12)
'                        slDay - XX  (1-31)
'                        slYear - xxxx (19xx-20xx)
'               obtain random date and current time and return it
'               Random date is based on the low date of 1/1/1970 (value of 25569)
'                   Hi date of 12/31/2069 (value of 62093)
'                   Dates will be adjusted to lo value of 26000 and hi value of 62000)
'               in string format
'
'               Created:  10/26/20
'*********************************************************************
Sub gRandomDateTime(slDate As String, slTime As String, slMonth As String, slDay As String, slYear As String)
    Dim llRandomDate As Long
        Randomize
        llRandomDate = CLng(((62000 - 26000) + 1) * Rnd + 26000)
        slDate = Format$(llRandomDate, "ddddd")
        gPackDate slDate, igNowDate(0), igNowDate(1)
        slTime = Format$(gNow(), "h:mm:ssAM/PM")
        gPackTime slTime, igNowTime(0), igNowTime(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
        
    Exit Sub
End Sub
'***************************************************************
'
'           gUnpackCurDateTime - from igNowDate & igNowTime
'           convert to string
'           <output> slCurrDate xx/xx/xxxx
'                    slCurrTime xx:xx:xxa/p
'                    slMonth as string
'                    slDay as string
'                    slYear as string (xxxx)
'***************************************************************
Sub gUnpackCurrDateTime(slCurrDate As String, slCurrTime As String, slMonth As String, slDay As String, slYear As String)
    gUnpackDate igNowDate(0), igNowDate(1), slCurrDate
    gUnpackTime igNowTime(0), igNowTime(1), "A", "1", slCurrTime
    gObtainYearMonthDayStr slCurrDate, True, slYear, slMonth, slDay
End Sub
'
'           Obtain the periods (calendar, standard or corporate start and end dates)
'           <input>  ilWhichPeriod : 0 = corporate, 1 = standard, 2 = calendar
'                    llStartToConvert - Date to obtain the requested period start date
'                    llEndToConvert - date to obtain the requested period end date
'           <output>  llPeriodStart = calendars start date
'                     llPeriodEnd = calendars enddate
'
Public Sub gObtainPeriodDates(ilWhichPeriod As Integer, llStartToConvert As Long, llEndToConvert As Long, llPeriodStart As Long, llPeriodEnd As Long)

        llPeriodStart = 0
        llPeriodEnd = 0
        If ilWhichPeriod = 0 Then          'corp
            'get the corp start/end date of the ad server entry
            llPeriodStart = gDateValue(gObtainStartCorp(Format$(llStartToConvert, "ddddd"), False))
            llPeriodEnd = gDateValue(gObtainEndCorp(Format$(llEndToConvert, "ddddd"), False))
        ElseIf ilWhichPeriod = 1 Then      'std
            'get the std start/end date of the ad server entry
            llPeriodStart = gDateValue(gObtainStartStd(Format$(llStartToConvert, "ddddd")))
            llPeriodEnd = gDateValue(gObtainEndStd(Format$(llEndToConvert, "ddddd")))
        ElseIf ilWhichPeriod = 2 Then     'cal contract
            'get the cal start/end date of the ad server entry
            llPeriodStart = gDateValue(gObtainStartCal(Format$(llStartToConvert, "ddddd")))
            llPeriodEnd = gDateValue(gObtainEndCal(Format$(llEndToConvert, "ddddd")))
        End If
       
        Exit Sub
End Sub
