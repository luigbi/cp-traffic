Attribute VB_Name = "RPTEXTRA"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptextra.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmSdfSrchKey4                 tmRsr                         tmRsrSrchKey              *
'*  hmRsr                         imRsrRecLen                   hmCrf                     *
'*  hmLlf                                                                                 *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  gCRRSRClear                   gObtainVaf                                              *
'******************************************************************************************

Option Explicit
Option Compare Text
'Public lgNowTime As Long
'8297
'Public ogPdf As PDFSplitMerge.CPDFSplitMergeObj
Public tgMMnf() As MNF
Public sgMnfVehGrpTag As String
Public smVehGp5CodeTag As String
Public tgVehicleSets1() As POPICODENAME
Public tgVehicleSets2() As POPICODENAME

Public lgSTime1 As Long
Public lgSTime2 As Long
Public lgSTime3 As Long
Public lgSTime4 As Long
Public lgSTime5 As Long
Public lgSTime6 As Long

Public lgETime1 As Long
Public lgETime2 As Long
Public lgETime3 As Long
Public lgETime4 As Long
Public lgETime5 As Long
Public lgETime6 As Long

Public lgTtlTime1 As Long
Public lgTtlTime2 As Long
Public lgTtlTime3 As Long
Public lgTtlTime4 As Long
Public lgTtlTime5 As Long
Public lgTtlTime6 As Long

Dim hmMsg As Integer    'd.s. 11/6/01

Dim tmAfr As AFR                  'Generic AFfiliate prepass file
Dim hmAfr As Integer
Dim tmAfrSrchKey As AFRKEY0       'Gen date and time
Dim imAfrRecLen As Integer        'Generic record length

'Copy Report
Dim hmCpr As Integer            'Copy Report file handle
Dim tmCpr() As CPR                'CPR record image
Dim tmCprSrchKey As CPRKEY0            'CPR record image
Dim imCprRecLen As Integer        'CPR record length

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length
'Log Calendar
Dim hmLcf As Integer            'Log Calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim imLcfRecLen As Integer        'LCF record length
Dim tmPjf As PJF                  'Slsp Projections
Dim imPjfRecLen As Integer        'PJF record length

Dim tmSbf As SBF                  '8-3-092 Special Billing
Dim imSbfRecLen As Integer        'SBF record length

Dim tmCbf As CBF                  'BR prepass file
Dim hmCbf As Integer
Dim tmCbfSrchKey As CBFKEY0       'Gen date and time
Dim tmCbfSrchKey1 As CBFKEY1      '2-17-13 gen date & time & user ID
Dim imCbfRecLen As Integer        'BR record length
Dim tmGrf As GRF                  'Generic  prepass file
Dim hmGrf As Integer
Dim tmGrfSrchKey As GRFKEY0       'Gen date and time
Dim imGrfRecLen As Integer        'Generic record length
Dim tmRvf As RVF                  'Receivables/History
Dim hmRvf As Integer
Dim imRvfRecLen As Integer        'PJF record length
Dim tmSlf As SLF                  'Slsp
Dim imSlfRecLen As Integer        'PJF record length
Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey2 As SDFKEY2            'SDF record image (key 2)
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSdf As SDF
Dim tmSdfSrchKey3 As LONGKEY0     'SDF record image (SDF code as keyfield)
Dim tmSdfSrchKey1 As SDFKEY1        'sdfkey:  vefcode, date, time, schstatus
Dim tmSdfSrchKey4 As SDFKEY4        'sdf key:  date, chfcode
Dim tmSdfSrchKey0 As SDFKEY0        'vefcode, chfcode, date

Dim hmIbf As Integer             'Ibf detail file handle
Dim imIbfRecLen As Integer       'Ibf record length
Dim tmIbf As IBF
Dim tmIbfSrchKey0 As IBFKEY0     'Ibf record image (ibfCode as keyfield)
Dim tmIbfSrchKey1 As IBFKEY1     'Ibf record image (contrNo, podCPMID as keyfield)
Dim tmIbfSrchKey2 As IBFKEY2     'Ibf record image (contrNo, vefCode as keyfield)
Dim tmIbfSrchKey3 As IBFKEY3     'Ibf record image (billYear, billMonth as keyfield)
'-----------------

Dim hmSmf As Integer            'mg/outside  file handle
Dim imSmfRecLen As Integer
Dim tmSmf As SMF

Dim hmSsf As Integer
Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim tmChf As CHF
Dim imCHFRecLen As Integer
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim tmRcf As RCF

Dim hmClf As Integer
Dim tmClf As CLF
Dim imClfRecLen As Integer
Dim tmClfSrchKey As CLFKEY0             'Sch Line record image

Dim hmCff As Integer
Dim tmCff As CFF
Dim imCffRecLen As Integer
 
'Quarterly Avails
Dim hmAvr As Integer            'Quarterly Avails file handle
Dim tmAvr() As AVR                'AVR record image
Dim tmInvValAmtSold() As INVVALAMTSOLD      '$ accumulated for spots sold to use to calc avg price per 30"unit

Dim tmAvrSrchKey As AVRKEY0            'AVR record image
Dim imAvrRecLen As Integer        'AVR record length
'Invoices
Dim hmIvr As Integer            'Invoices file handle
Dim tmIvrSrchKey As IVRKEY0       'IVR record image
Dim imIvrRecLen As Integer        'IVR record length
Dim hmImr As Integer            'Invoices main prepass for combined air time & ntr file handle
Dim tmImrSrchKey As AVRKEY0            'IMR record image
Dim imImrRecLen As Integer        'IMR record length
'Dim lmSAvailsDates(1 To 14) As Long   '6-30-00 Start Dates of avail week
Dim lmSAvailsDates(0 To 14) As Long   '6-30-00 Start Dates of avail week. Index zero ignored
'Dim lmEAvailsDates(1 To 14) As Long   '6-30-00End dates of avail week
Dim lmEAvailsDates(0 To 14) As Long   '6-30-00End dates of avail week. Index zero ignored
Dim tmAnr As ANR                    'prepass analysis file
Dim hmAnr As Integer
Dim tmAnrSrchKey As GRFKEY0       'Gen date and time
Dim imAnrRecLen As Integer        'ANR record length

Dim hmScr As Integer        'Script dump file handle
Dim imScrRecLen As Integer  'sCR record length
Dim tmScr As SCR            'SCR record image
Dim tmScrSrchKey As SCRKEY1     'Gen date and time

Dim hmTxr As Integer        'Text dump file handle
Dim imTxrRecLen As Integer  'TXR record length
Dim tmTxr As TXR            'TXR record image
Dim tmTxrSrchKey As TXRKEY0     'Gen date and time

Dim hmUor As Integer        'User Options file handle
Dim imUorRecLen As Integer  'UOR record length
Dim tmUor As UOR            'UOR record image
Dim tmUorSrchKey As UORKEY0     'Gen date and time

Dim imFsfRecLen As Integer  'FSF record length
Dim tmFsf As FSF            'FSF record image

Dim imFpfRecLen As Integer  'FPF record length
Dim tmFpf As FPF            'FPF record image

Dim imFdfRecLen As Integer  'FDF record length
Dim tmFdf As FDF            'FDF record image

Dim tmOdf As ODF                  'One Day log file file
Dim hmOdf As Integer
Dim tmOdfSrchKey2 As ODFKEY2 'ODF key record image
Dim imOdfRecLen As Integer

Dim tmUaf As UAF                  'User Activity
Dim hmUaf As Integer
Dim imUafRecLen As Integer        'User Activity record length

'10-4-05
Dim tmCrf As CRF                  'Copy Rotation Header
Dim imCrfRecLen As Integer

'12-14-05
Dim tmLlf As LLF
Dim imLlfRecLen As Integer

'2-10-06                Demo search key
Dim tmDrfSrchKey As DRFKEY0
Dim imDrfRecLen As Integer
Dim tmDrf As DRF

'9-19-06
Dim tmRaf As RAF            'split regions
Dim imRafRecLen As Integer

'5-6-10
Dim tmPff As PFF            'Prefeed
Dim imPffRecLen As Integer

Dim tmRvr As RVR                  'Receivables prepass tale
Dim hmRvr As Integer              'handle
Dim tmRvrSrchKey As RVRKEY0       'Gen date and time
Dim imRvrRecLen As Integer        'Generic record length


'3-25-14 move sdfsortbyline & spottypes to rptrec.bas
'Type SDFSORTBYLINE
'    sKey As String * 20         'line ID
'    tSdf As SDF
'End Type
'
'Type SPOTTYPES
'    iSched As Integer
'    iMissed As Integer
'    iMG As Integer
'    iOutside As Integer
'    iHidden As Integer
'    iCancel As Integer
'    iFill As Integer
'    iOpen As Integer        '1-18-11
'    iClose As Integer       '1-18-11
'End Type

Type RVFCNTSORT
    sKey As String * 60         'contract #|vehicle name | cash/trade/merch/promo flag
    tRVF As RVF
End Type

'AVail information for Avail Inventory VAluation calc
Type VALUATIONINFO
    sBaseReptFlag As String * 1          'B = test BaseDP Flag, R = test Show on Report Flag
    iRCvsAvgPrice As Integer            '0 = use r/c, 1 = use avg 30" spot price
    iUnsoldPctAdj As Integer            '+/- pct adjustment factor to adjust r/c or avg spot price value (0 indicates 100)
    iEstPctSellout As Integer           'estimated percent sellout of inv valuation (0 = 100)
End Type

Type INVVALAMTSOLD
    'lRate(1 To 14) As Long              'total $ from spot costs to calc 30" avg spot price after finding spots sold
    lRate(0 To 14) As Long              'total $ from spot costs to calc 30" avg spot price after finding spots sold. Index zero ignored
End Type

Type SPOTLENRATIO
    iLen(0 To 9) As Integer             'spot lengths (will be used as spans):  i.e. index 0 = 10, index 1 = 15, index 2 = 30 ,etc
                                        'means any spot length from 0 to 10" will be "x" radio, then 11 - 15 will be "y" ratio, then 16-30 will be "z" ratio
    iRatio(0 To 9) As Integer              'spot length ratios in hundreds I.e.  15" = .5, entered as 50, 30" = 1.00, entered as 100
End Type

Type AVRCOUNTS
    'Index zero ignored in all arrays below
    l30Count(0 To 14) As Long   'Count of Inventory or Avails or Sold for 30sec (for qtrly detail this contains sold values)
    l60Count(0 To 14) As Long   'Count of Inventory or Avails or Sold for 60sec (for qtrly detail this contains sold values)
    l30InvCount(0 To 14) As Long   'Count of Inventory for 30sec
    l60InvCount(0 To 14) As Long   'Count of Inventory for 60sec
    'the remaining variables are for the qtrly detail report
    l30Hold(0 To 14) As Long     'count of hold units sold for 30sec
    l60Hold(0 To 14) As Long     'count of hold units sold for 60sec
    l30Reserve(0 To 14) As Long     'count of reserve units sold for 30sec
    l60Reserve(0 To 14) As Long     'count of reserve units sold for 60sec
    l30Avail(0 To 14) As Long     'count of available units sold for 30sec
    l60Avail(0 To 14) As Long     'count of available units sold for 60sec
    l30Prop(0 To 14) As Long  'count of 30 proposals
    l60Prop(0 To 14) As Long  'count of 60 proposals
End Type

'this is for the Billed and Booked when requesting by Billing Method (cycle)
'When any of the date options are requested (other than bill method), the report uses one set of dates.
'When bill method is selected, 2 sets of dates are required:  one for standard bdcst billing; other other for calendar billing (currently for Podcast CPM only)
Type BILLCYCLE                      '1-13-21
    blBillCycle As Boolean          'true if use billing cycle of contract
    lBillCycleStartDates(0 To 13) As Long   'start dates of calendar (std dates also retained that are in orig code)
    lBillCycleLastBilled As Long
    iBillCycleLastBilledInx As Integer
End Type

Type BILLCYCLERAB                   '1-27-21
    ilBillCycle As Integer          '0=Std, 1=Cal, 2=billing cycle of contract
    lStdBillCycleStartDates(0 To 25) As Long   'start dates by std Broadcast Calendar
    lStdBillCycleLastBilled As Long
    iStdBillCycleLastBilledInx As Integer
    lCalBillCycleStartDates(0 To 25) As Long   'start dates by Monthly Calendar
    lCalBillCycleLastBilled As Long
    iCalBillCycleLastBilledInx As Integer
End Type

'TTP 10666 - daily avg Digital lines in RAB & B&B Cal Spots
Type DIGITALLINEAVERAGE
    lPcfCode As Long                    'What PCF code is this Averaging for
    iLineNo As Integer                  'What Line # is this on the Contract (podCPMID) - TTP 10743
    iStartYear As Integer               'Starting Period the data (the period in index 1 of the following arrays) - the 1st period in the Line (StartDate)
    iStartMonth As Integer              'Starting Period the data
    iTotalDays As Integer               'How many days does this line Span
    iFirstInx As Integer                'What is the 1st used index in this type (usually 1, but sometimes 0)
    iProjectedInx As Integer            'Which Index is Projected (Future)
    iBilledInx As Integer               'Which Index is the last Billed period (-1 = not yet billed)
    iLastInx As Integer                 'Which Index has the End of the Line
    dLineCost As Double                 'How Much does this Line cost
    iRemainingDays As Integer           'How Many Days remain from Last Billed
    
    iYears(0 To 25) As Integer          'what Year is this index
    iMonths(0 To 25) As Integer         'what Month is this index - 0 through 25 (26 periods)
    
    dUnbilledAmt As Double     'Amounts Remaining by Std Period
    iDaysInMonthCal(0 To 25) As Integer  'How many Cal days are in this Month
    iDaysInMonthStd(0 To 25) As Integer  'How many Std days are in this Month
    iDaysAdjPrior(0 To 25) As Integer   'How many days are moved into Previous month to make this into a Cal Month
    iDaysAdjNext(0 To 25) As Integer    'How many days are taken from Next month to make this into a Cal Month
    iDaysExt(0 To 25) As Integer        'How many days are in This Month + Adj Prior + Adj Next
    
    dDailyAmt(0 To 25) As Double        'Average Daily Amount
    
    dBilledAmt(0 To 25) As Double       'Amounts Billed by Std Period
    dBilledRemainder(0 To 25) As Double 'When Billed $ is applied across 2 Cal Months, this is the remaining Amount otherwise is the Full Billed $ for the Month
    dExtAmountGross(0 To 25) As Double  'The Number for the Report for this month
    
    sPCFLineStartDate As String         'Start Date
    sPCFLineEndDate As String           'End Date
    dExtAmountTotal As Double           'A total of the Ext Amounts Gross (which should be the same as the Line Cost)
    
    sComment(0 To 25) As String         'Auto Generated Comment to help Audit this mess
End Type

'-----------------------------------------------------------------------
'TTP 10666 - daily avg Digital lines in RAB & B&B Cal Spots
'Inputs:
'   sBillCycle, "S" for Standard Bill Cycle, "C" for Calendar Bill Cycle
'   lLastBilled, the Date we Last Billed - everything beyond this date is "projected" future.
'   tlPCF, the PCF record for a Digital Line
'   tlRVF, the Table of RVF records for a Contract
'   tlPCF2, a Table of ALL PCF records for all versions of the the contract (TTP 10955)
'Returns:
'   DIGITALLINEAVERAGE, a Type with Monthly Actual Amounts & projected Gross Avg Estimates (dExtAmountGross)
'                       based on Daily Averaging for a Digital Line in a Contract.
'                       Resulting Amounts are based in Calendar Months (despite Contract Bill Cycle being Standard or Calendar)
'v81 TTP 10666 - new issue Wed 5/3/23 4:01 PM - Added lLastBilledCal parameter

'Function gBuildDigitalLineAverage(sBillCycle As String, lLastBilledStd As Long, lLastBilledCal As Long, tlPcf As PCF, tlRvf() As RVF) As DIGITALLINEAVERAGE
Function gBuildDigitalLineAverage(sBillCycle As String, lLastBilledStd As Long, lLastBilledCal As Long, tlPcf As PCF, tlRvf() As RVF, tlPcf2() As PCF) As DIGITALLINEAVERAGE
    Dim tlDigitalLineAvg As DIGITALLINEAVERAGE
    Dim llCPMStartDate As Long
    Dim llCPMEndDate As Long
    Dim ilLoop As Integer
    Dim ilLoop2 As Integer
    Dim llAmount As Long
    Dim dlAmount As Double
    Dim dlTotalAmt As Double
    Dim slTemp As String
    Dim slTemp2 As String
    Dim llTemp As Long
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim strComment As String
    Dim ilDaysSoFar As Integer
    Dim lLastBilled As Long
    Dim ilDoe As Integer
    Dim ilPcf2Loop As Integer
    
    tlDigitalLineAvg.iLastInx = -1
    tlDigitalLineAvg.iFirstInx = -1
    On Error GoTo ExitBuildDigitalLine

    '-------------------------------------
    'Get Line Start & End Date
    gUnpackDateLong tlPcf.iStartDate(0), tlPcf.iStartDate(1), llCPMStartDate
    gUnpackDateLong tlPcf.iEndDate(0), tlPcf.iEndDate(1), llCPMEndDate
    tlDigitalLineAvg.lPcfCode = tlPcf.lCode
    tlDigitalLineAvg.sPCFLineStartDate = Format$(llCPMStartDate, "m/d/yy")
    tlDigitalLineAvg.sPCFLineEndDate = Format$(llCPMEndDate, "m/d/yy")

    Debug.Print "gBuildDigitalLineAverage - Line:" & tlPcf.iPodCPMID & ", PCFCode:" & tlPcf.lCode & ", CPMStart:" & tlDigitalLineAvg.sPCFLineStartDate & ", CPMEnd:" & tlDigitalLineAvg.sPCFLineEndDate

    '-------------------------------------
    'CBS (Canceled before start)
    If llCPMStartDate > llCPMEndDate Then
        Debug.Print " - CBS"
        gBuildDigitalLineAverage = tlDigitalLineAvg
        Exit Function
    End If
    
    '-------------------------------------
    'Determine Std Period this Line starts in
    slTemp = Format$(llCPMStartDate, "m/d/yy")
    slTemp = gObtainEndStd(slTemp)
    gObtainYearMonthDayStr slTemp, True, slYear, slMonth, slDay
    
    '-------------------------------------
    'Store the Line's Std Start Month/Year
    tlDigitalLineAvg.iStartMonth = Val(slMonth)
    tlDigitalLineAvg.iStartYear = Val(slYear)
    tlDigitalLineAvg.iLineNo = tlPcf.iPodCPMID 'TTP 10743
    
    '-------------------------------------
    'Line Cost
    tlDigitalLineAvg.dLineCost = tlPcf.lTotalCost / 100
    
    '-------------------------------------
    'Fill in the Month and Year Indexes (PCF StartDate [month/year] through 2 years in the future)
    'Also Set the iProjectedInx
    tlDigitalLineAvg.iProjectedInx = -1 'No Projected (All Past)
    tlDigitalLineAvg.iBilledInx = -1 'Not Yet billed
    For ilLoop = 0 To 25
        llAmount = slMonth + (ilLoop - 1)
        If llAmount = 0 Then
            tlDigitalLineAvg.iMonths(ilLoop) = 12
            tlDigitalLineAvg.iYears(ilLoop) = Val(slYear) - 1
        ElseIf llAmount < 13 Then
            tlDigitalLineAvg.iMonths(ilLoop) = llAmount
            tlDigitalLineAvg.iYears(ilLoop) = Val(slYear)
        ElseIf llAmount >= 13 And llAmount <= 24 Then
            tlDigitalLineAvg.iMonths(ilLoop) = llAmount - 12
            tlDigitalLineAvg.iYears(ilLoop) = Val(slYear) + 1
        Else
            tlDigitalLineAvg.iMonths(ilLoop) = llAmount - 24
            tlDigitalLineAvg.iYears(ilLoop) = Val(slYear) + 2
        End If
        
        '----------------------------------
        'Set iProjectedInx
        'Get End of Current Month
        slTemp = tlDigitalLineAvg.iMonths(ilLoop) & "/15/" & tlDigitalLineAvg.iYears(ilLoop)
        If sBillCycle = "C" Then
            slTemp = gObtainEndCal(slTemp)
        Else
            slTemp = gObtainEndStd(slTemp)
        End If
        'Get Last Billed
        slTemp2 = Format(lLastBilledStd, "ddddd")
        lLastBilled = lLastBilledStd
        If sBillCycle = "C" Then
            'v81 TTP 10666 - new issue Wed 5/3/23 4:01 PM
            'slTemp2 = gObtainEndCal(slTemp2)
            slTemp2 = Format(lLastBilledCal, "ddddd")
            lLastBilled = lLastBilledCal
        End If
        'Determine if Current Month > Last Billed
        If DateDiff("d", slTemp, slTemp2) >= 0 Then
            tlDigitalLineAvg.iProjectedInx = ilLoop + 1
            'Debug.Print slTemp & " has been billed.  Last Billed=" & slTemp2 & ", ProjectedIndex=" & tlDigitalLineAvg.iProjectedInx
        End If
    Next ilLoop
    
    '-------------------------------------
    'Fill in the Day Counts
    tlDigitalLineAvg.iTotalDays = DateDiff("d", Format(llCPMStartDate, "ddddd"), Format(llCPMEndDate, "ddddd")) + 1
    For ilLoop = 1 To 24
        '----------------------------------
        'How many Days in Std Cal period
        slTemp = tlDigitalLineAvg.iMonths(ilLoop) & "/15/" & tlDigitalLineAvg.iYears(ilLoop)
        slTemp = gObtainStartStd(slTemp)
        slTemp2 = tlDigitalLineAvg.iMonths(ilLoop) & "/15/" & tlDigitalLineAvg.iYears(ilLoop)
        slTemp2 = gObtainEndStd(slTemp)
        'Limit days in Month to CPM Start/End Date
        If DateDiff("d", Format(llCPMStartDate, "ddddd"), slTemp) < 0 Then
            slTemp = Format(llCPMStartDate, "ddddd")
        End If
        If DateDiff("d", Format(llCPMEndDate, "ddddd"), slTemp2) > 0 Then
            slTemp2 = Format(llCPMEndDate, "ddddd")
        End If
        If DateDiff("d", slTemp, slTemp2) + 1 > 0 Then tlDigitalLineAvg.iDaysInMonthStd(ilLoop) = DateDiff("d", slTemp, slTemp2) + 1
        '----------------------------------
        'How many Days in Monthly Cal period
        slTemp = tlDigitalLineAvg.iMonths(ilLoop) & "/15/" & tlDigitalLineAvg.iYears(ilLoop)
        slTemp = gObtainStartCal(slTemp)
        slTemp2 = tlDigitalLineAvg.iMonths(ilLoop) & "/15/" & tlDigitalLineAvg.iYears(ilLoop)
        slTemp2 = gObtainEndCal(slTemp)
        'Limit days in Month to CPM Start/End Date
        If DateDiff("d", Format(llCPMStartDate, "ddddd"), slTemp) < 0 Then
            slTemp = Format(llCPMStartDate, "ddddd")
        End If
        If DateDiff("d", Format(llCPMEndDate, "ddddd"), slTemp2) > 0 Then
            slTemp2 = Format(llCPMEndDate, "ddddd")
        End If
        If DateDiff("d", slTemp, slTemp2) + 1 > 0 Then tlDigitalLineAvg.iDaysInMonthCal(ilLoop) = DateDiff("d", slTemp, slTemp2) + 1
    Next ilLoop
    
    '-------------------------------------
    'TTP 10822 - determine which index is the FIRST index (the index which is at the StartDate of the CPM)
    For ilLoop = 0 To 25
        slTemp = tlDigitalLineAvg.iMonths(ilLoop) & "/15/" & tlDigitalLineAvg.iYears(ilLoop)
        slTemp = gObtainStartStd(slTemp)
        If tlDigitalLineAvg.iFirstInx = -1 Then
            If DateDiff("d", DateValue(tlDigitalLineAvg.sPCFLineStartDate), DateValue(slTemp)) >= 0 Then
                tlDigitalLineAvg.iFirstInx = ilLoop
                Exit For
            End If
        End If
    Next ilLoop

    '-------------------------------------
    'determine which index is the LAST index (the index which is at the EndDate of the CPM)
    For ilLoop = 1 To 25
        slTemp = tlDigitalLineAvg.iMonths(ilLoop) & "/15/" & tlDigitalLineAvg.iYears(ilLoop)
        slTemp = gObtainEndStd(slTemp)
        If tlDigitalLineAvg.iLastInx = -1 Then
            If DateDiff("d", tlDigitalLineAvg.sPCFLineEndDate, slTemp) >= 0 Then
                tlDigitalLineAvg.iLastInx = ilLoop
                Exit For
            End If
        End If
    Next ilLoop
    
    '-------------------------------------
    'Fill in the Adjustment Days (Adjustments needed to convert Standard to Calendar)
    If sBillCycle = "S" Then
        ilDaysSoFar = 0
        For ilLoop = 1 To tlDigitalLineAvg.iLastInx + 1
            slTemp = tlDigitalLineAvg.iMonths(ilLoop) & "/15/" & tlDigitalLineAvg.iYears(ilLoop)
            slTemp2 = gObtainStartStd(slTemp)
            slTemp = gObtainStartCal(slTemp)
            If DateDiff("d", Format(llCPMStartDate, "ddddd"), slTemp) < 0 Then
                slTemp = Format(llCPMStartDate, "ddddd")
            End If
            If DateDiff("d", Format(llCPMStartDate, "ddddd"), slTemp2) < 0 Then
                slTemp2 = Format(llCPMStartDate, "ddddd")
            End If
            ilDaysSoFar = ilDaysSoFar + tlDigitalLineAvg.iDaysInMonthStd(ilLoop)
            If Day(slTemp) <> Day(slTemp2) Then
                tlDigitalLineAvg.iDaysAdjPrior(ilLoop) = DateDiff("d", slTemp, slTemp2)
                tlDigitalLineAvg.iDaysAdjNext(ilLoop - 1) = DateDiff("d", slTemp2, slTemp)
            End If
            If ilDaysSoFar >= tlDigitalLineAvg.iTotalDays Then Exit For
        Next ilLoop
    End If
    
    '-------------------------------------
    'Fill in the Ext Days (Std Days + Adj Prior + Adj Next)
    If sBillCycle = "S" Then
        For ilLoop = 0 To 25
            tlDigitalLineAvg.iDaysExt(ilLoop) = tlDigitalLineAvg.iDaysInMonthStd(ilLoop) + tlDigitalLineAvg.iDaysAdjNext(ilLoop) + tlDigitalLineAvg.iDaysAdjPrior(ilLoop)
            If ilLoop > 0 And tlDigitalLineAvg.iDaysExt(ilLoop) = 0 Then Exit For
        Next ilLoop
    Else        'Fill in ExtDays for Monthly Cal
        For ilLoop = 0 To 25
            tlDigitalLineAvg.iDaysExt(ilLoop) = tlDigitalLineAvg.iDaysInMonthCal(ilLoop)
            If ilLoop > 0 And tlDigitalLineAvg.iDaysExt(ilLoop) = 0 Then Exit For
        Next ilLoop
    End If

    '-------------------------------------
    'get the adserver Received amts (RVF)
    'Debug.Print " - Rcvd Amounts:";
    dlTotalAmt = 0
    For ilLoop = 0 To 25
        tlDigitalLineAvg.dBilledAmt(ilLoop) = 0
    Next ilLoop
    For ilLoop = LBound(tlRvf) To UBound(tlRvf) - 1
        'TTP 10725 - Billed and Booked Cal Spots and RAB Cal Spots: not including digital line contract that should be included
        'If tlRvf(ilLoop).lPcfCode = tlPcf.lCode Then
        
        'TTP 10955 - Billed and Booked Cal Spots: slowness reported, possibly due to including digital lines
        'If gObtainPcfCPMID(tlRvf(illoop).lPcfCode) = tlPcf.iPodCPMID Then
        For ilPcf2Loop = 0 To UBound(tlPcf2) - 1
            If tlPcf2(ilPcf2Loop).lCode = tlRvf(ilLoop).lPcfCode And tlPcf2(ilPcf2Loop).iPodCPMID = tlPcf.iPodCPMID Then
                ilDoe = ilDoe + 1
                If ilDoe > 100 Then
                    ilDoe = 0
                    DoEvents
                End If                'Accumulate the billing (for this digital line)
                'TTP 10725 - Billed and Booked Cal Spots and RAB Cal Spots: not including digital line contract that should be included
                'gUnpackDateLong tlRvf(ilLoop).iTranDate(0), tlRvf(ilLoop).iTranDate(1), llTemp
                'slTemp = Format$(llTemp, "m/d/yy")
                'If sBillCycle = "S" Then
                '    slTemp = gObtainEndStd(slTemp)
                'End If
                If tgSpf.sSEnterAgeDate = "E" Then      'use E=Entered Date or ageing date (Sales tab)
                    'JW 5/18/23 - Fixed RAB and B&B for V81 TTP 10725 – new issue 5-17-23.zip
                    gUnpackDate tlRvf(ilLoop).iTranDate(0), tlRvf(ilLoop).iTranDate(1), slTemp
                    If sBillCycle = "S" Then
                        slTemp = gObtainEndStd(slTemp)
                    End If
                Else
                    'JW 5/18/23 - Fixed RAB and B&B for V81 TTP 10725 – new issue 5-17-23.zip
                    slTemp = Trim$(str$(tlRvf(ilLoop).iAgePeriod) & "/15/" & Trim$(str$(tlRvf(ilLoop).iAgingYear)))
                    slTemp = gObtainEndStd(slTemp)
                End If
                
                gObtainYearMonthDayStr slTemp, True, slYear, slMonth, slDay
                gPDNToLong tlRvf(ilLoop).sGross, llAmount
                'Apply Amount to correct Period
                For ilLoop2 = 0 To 25
                    If tlDigitalLineAvg.iMonths(ilLoop2) = Val(slMonth) And tlDigitalLineAvg.iYears(ilLoop2) = Val(slYear) Then
                        tlDigitalLineAvg.dBilledAmt(ilLoop2) = tlDigitalLineAvg.dBilledAmt(ilLoop2) + (llAmount / 100)
                        'Debug.Print "(" & tlDigitalLineAvg.iMonths(ilLoop2) & "/" & tlDigitalLineAvg.iYears(ilLoop2) & ")" & Format(tlDigitalLineAvg.dBilledAmt(ilLoop2), "$#.00") & ", ";
    
                        'ReCalc Daily Amount for [Index 0] / [Index 1]
                        If sBillCycle = "S" Then
                            If tlDigitalLineAvg.iDaysInMonthStd(ilLoop2) <> 0 Then 'WWO B&B by CAL Spots TTP 10760
                                tlDigitalLineAvg.dDailyAmt(ilLoop2) = tlDigitalLineAvg.dBilledAmt(ilLoop2) / tlDigitalLineAvg.iDaysInMonthStd(ilLoop2)
                            End If
                        Else
                            If tlDigitalLineAvg.iDaysInMonthCal(ilLoop2) <> 0 Then 'WWO B&B by CAL Spots TTP 10760
                                tlDigitalLineAvg.dDailyAmt(ilLoop2) = tlDigitalLineAvg.dBilledAmt(ilLoop2) / tlDigitalLineAvg.iDaysInMonthCal(ilLoop2)
                            End If
                        End If
                        dlTotalAmt = dlTotalAmt + llAmount / 100
                        tlDigitalLineAvg.iBilledInx = ilLoop2
                        Exit For
                    End If
                Next ilLoop2
                Exit For
            End If
        Next ilPcf2Loop
    Next ilLoop
    
    'Debug.Print ""
    'Debug.Print " - LastBilled=" & Format(lLastBilled, "ddddd") & ", FirstInx=" & tlDigitalLineAvg.iFirstInx & ", ProjectedInx=" & tlDigitalLineAvg.iProjectedInx & ", BilledInx=" & tlDigitalLineAvg.iBilledInx & ", LastInx=" & tlDigitalLineAvg.iLastInx
    
    '-------------------------------------
    'Calc Unbilled $ Amount & Remaining Days
    tlDigitalLineAvg.dUnbilledAmt = tlDigitalLineAvg.dLineCost - dlTotalAmt
    If lLastBilled < llCPMStartDate Then
        tlDigitalLineAvg.iRemainingDays = DateDiff("d", Format(llCPMStartDate, "ddddd"), Format(llCPMEndDate, "ddddd")) + 1
    Else
        If sBillCycle = "S" Then
            tlDigitalLineAvg.iRemainingDays = DateDiff("d", Format(lLastBilled + 1, "ddddd"), Format(llCPMEndDate, "ddddd")) + 1
        Else
            'Note: I dont like that I had to remove the +1 on the lastBilled
            tlDigitalLineAvg.iRemainingDays = DateDiff("d", gObtainEndCal(Format(lLastBilled, "ddddd")), Format(llCPMEndDate, "ddddd"))
        End If
    End If
    If tlDigitalLineAvg.iRemainingDays < 0 Then tlDigitalLineAvg.iRemainingDays = 0
    
    '-------------------------------------
    'Calc DailyAmt & Projected Ext Amounts
    If tlDigitalLineAvg.iRemainingDays <> 0 Then
        dlAmount = (tlDigitalLineAvg.dUnbilledAmt / tlDigitalLineAvg.iRemainingDays)
    Else
        dlAmount = 0
    End If
    tlDigitalLineAvg.dDailyAmt(0) = dlAmount
    For ilLoop = IIF(tlDigitalLineAvg.iProjectedInx = -1, 0, IIF(tlDigitalLineAvg.iBilledInx = -1, 0, tlDigitalLineAvg.iProjectedInx)) To tlDigitalLineAvg.iLastInx + 1
        tlDigitalLineAvg.dDailyAmt(ilLoop) = dlAmount
        tlDigitalLineAvg.dExtAmountGross(ilLoop) = (tlDigitalLineAvg.iDaysExt(ilLoop) * dlAmount)
        If ilLoop < 25 Then
            tlDigitalLineAvg.dExtAmountGross(ilLoop) = tlDigitalLineAvg.dExtAmountGross(ilLoop) + (tlDigitalLineAvg.iDaysAdjNext(ilLoop) * tlDigitalLineAvg.dDailyAmt(ilLoop + 1)) 'TTP 10741
        End If
        If tlDigitalLineAvg.iProjectedInx = 1 And ilLoop < 2 Then
            If tlDigitalLineAvg.iDaysExt(0) <> 0 Then
                tlDigitalLineAvg.dExtAmountGross(0) = (tlDigitalLineAvg.iDaysExt(0) * dlAmount)
            End If
        End If
    Next ilLoop

    '-------------------------------------
    'Calc Billed Ext Amounts [Index 0] and [Index 1] (split First Payment to prior Cal Month if needed)
    For ilLoop = 0 To tlDigitalLineAvg.iBilledInx
        'Copy the Billed amounts.  They [Index 0] and [Index 1] might be adjusted below
        tlDigitalLineAvg.dExtAmountGross(ilLoop) = tlDigitalLineAvg.dBilledAmt(ilLoop)
        tlDigitalLineAvg.dExtAmountGross(ilLoop) = tlDigitalLineAvg.dExtAmountGross(ilLoop) + (tlDigitalLineAvg.iDaysAdjNext(ilLoop) * tlDigitalLineAvg.dDailyAmt(ilLoop + 1)) 'TTP 10741
    Next ilLoop
    If sBillCycle = "S" Then
        If tlDigitalLineAvg.iDaysExt(0) <> 0 And tlDigitalLineAvg.dBilledAmt(1) <> 0 Then
            'Ext Amount [Period 0]
            tlDigitalLineAvg.dBilledRemainder(0) = tlDigitalLineAvg.dBilledAmt(1)
            tlDigitalLineAvg.dDailyAmt(0) = tlDigitalLineAvg.dDailyAmt(1)
            tlDigitalLineAvg.dExtAmountGross(0) = tlDigitalLineAvg.iDaysExt(0) * tlDigitalLineAvg.dDailyAmt(1)
            tlDigitalLineAvg.dBilledRemainder(1) = tlDigitalLineAvg.dBilledAmt(1) - tlDigitalLineAvg.dExtAmountGross(0)
            
            'Ext Amount [Period 1]
            tlDigitalLineAvg.dExtAmountGross(1) = tlDigitalLineAvg.dBilledRemainder(1)
            tlDigitalLineAvg.dExtAmountGross(1) = tlDigitalLineAvg.dExtAmountGross(1) + (tlDigitalLineAvg.iDaysAdjNext(1) * tlDigitalLineAvg.dDailyAmt(2))
        End If
    End If
    
    '-------------------------------------
    'Calc Billed Ext Amounts for remaining billed periods.
    If sBillCycle = "S" Then
        'Std Cal, [Index 0] and [Index 1] computed above
        For ilLoop = 2 To tlDigitalLineAvg.iProjectedInx - 1
             tlDigitalLineAvg.dExtAmountGross(ilLoop) = tlDigitalLineAvg.dBilledAmt(ilLoop) _
             + (tlDigitalLineAvg.iDaysAdjNext(ilLoop) * tlDigitalLineAvg.dDailyAmt(ilLoop + 1)) _
             - (tlDigitalLineAvg.iDaysAdjNext(ilLoop - 1) * tlDigitalLineAvg.dDailyAmt(ilLoop))
        Next ilLoop
    Else
        'Monthly Calendar
        For ilLoop = 1 To tlDigitalLineAvg.iProjectedInx - 1
            tlDigitalLineAvg.dExtAmountGross(ilLoop) = tlDigitalLineAvg.dBilledAmt(ilLoop)
        Next ilLoop
    End If
    
    '-------------------------------------
    'Move Adjustment from Last+1 Period into Last Period (dont include adjustment for last index + 1)
    If sBillCycle = "S" And tlDigitalLineAvg.iLastInx < 25 Then
        tlDigitalLineAvg.dExtAmountGross(tlDigitalLineAvg.iLastInx) = tlDigitalLineAvg.dExtAmountGross(tlDigitalLineAvg.iLastInx) + tlDigitalLineAvg.dExtAmountGross(tlDigitalLineAvg.iLastInx + 1)
        tlDigitalLineAvg.dExtAmountGross(tlDigitalLineAvg.iLastInx + 1) = 0
        tlDigitalLineAvg.dDailyAmt(tlDigitalLineAvg.iLastInx + 1) = 0
    End If

    '-------------------------------------
    'Move Adjustment from Last Period into Last-1 Period if Negative
    If sBillCycle = "S" And tlDigitalLineAvg.dExtAmountGross(tlDigitalLineAvg.iLastInx) < 0 Then
        tlDigitalLineAvg.dExtAmountGross(tlDigitalLineAvg.iLastInx - 1) = tlDigitalLineAvg.dExtAmountGross(tlDigitalLineAvg.iLastInx - 1) + tlDigitalLineAvg.dExtAmountGross(tlDigitalLineAvg.iLastInx)
        tlDigitalLineAvg.dExtAmountGross(tlDigitalLineAvg.iLastInx) = 0
        tlDigitalLineAvg.iDaysInMonthStd(tlDigitalLineAvg.iLastInx - 1) = tlDigitalLineAvg.iDaysInMonthStd(tlDigitalLineAvg.iLastInx - 1) + tlDigitalLineAvg.iDaysInMonthStd(tlDigitalLineAvg.iLastInx)
        tlDigitalLineAvg.dExtAmountGross(tlDigitalLineAvg.iLastInx) = 0
        tlDigitalLineAvg.dDailyAmt(tlDigitalLineAvg.iLastInx) = 0
    End If

    '-------------------------------------
    'TTP 10822 - Adjust for one day contracts
    If sBillCycle = "S" Then
        If tlDigitalLineAvg.iDaysExt(0) > 0 And tlDigitalLineAvg.iDaysExt(1) < 0 And tlDigitalLineAvg.iFirstInx = 1 And tlDigitalLineAvg.sPCFLineStartDate = tlDigitalLineAvg.sPCFLineEndDate Then
            tlDigitalLineAvg.iDaysExt(0) = tlDigitalLineAvg.iDaysExt(0) + tlDigitalLineAvg.iDaysExt(1)
            tlDigitalLineAvg.iDaysExt(1) = 0

            tlDigitalLineAvg.iDaysAdjPrior(0) = 0
            tlDigitalLineAvg.iDaysAdjPrior(1) = 0

            tlDigitalLineAvg.iDaysAdjNext(0) = 0
            tlDigitalLineAvg.iDaysAdjNext(1) = 0
            
            tlDigitalLineAvg.iFirstInx = 0
            tlDigitalLineAvg.iStartMonth = tlDigitalLineAvg.iStartMonth - 1
            If tlDigitalLineAvg.iStartMonth < 1 Then
                tlDigitalLineAvg.iStartMonth = 12
                tlDigitalLineAvg.iStartYear = tlDigitalLineAvg.iStartYear - 1
            End If
        End If
    End If
    
    '-------------------------------------
    'Calc Total Ext Amount (Add each months Ext Amounts)
    'Debug.Print " - Monthly Amounts:";
    tlDigitalLineAvg.dExtAmountTotal = 0
    For ilLoop = 0 To tlDigitalLineAvg.iLastInx
        If tlDigitalLineAvg.dExtAmountGross(ilLoop) <> 0 Then
            'Debug.Print "(" & tlDigitalLineAvg.iMonths(ilLoop) & "/" & tlDigitalLineAvg.iYears(ilLoop) & ")" & Format(tlDigitalLineAvg.dExtAmountGross(ilLoop), "$#.00") & ", ";
            tlDigitalLineAvg.dExtAmountTotal = tlDigitalLineAvg.dExtAmountTotal + tlDigitalLineAvg.dExtAmountGross(ilLoop)
        End If
    Next ilLoop
    
    '-------------------------------------
    'Create Comments for each period on how the period was calculated
    For ilLoop = 0 To tlDigitalLineAvg.iLastInx
        tlDigitalLineAvg.sComment(ilLoop) = ""
        strComment = ""
        If tlDigitalLineAvg.dExtAmountGross(ilLoop) > 0 Then
            If ilLoop < tlDigitalLineAvg.iProjectedInx And tlDigitalLineAvg.iBilledInx > -1 Then
                If tlDigitalLineAvg.dBilledAmt(ilLoop) <> 0 Or tlDigitalLineAvg.dBilledRemainder(ilLoop) <> 0 Then
                    If tlDigitalLineAvg.dBilledRemainder(0) > 0 And tlDigitalLineAvg.dExtAmountGross(0) > 0 And ilLoop = 1 Then
                        strComment = "Billed Remainder " & mGetMonthName(tlDigitalLineAvg.iMonths(ilLoop)) & ":" & tlDigitalLineAvg.dBilledRemainder(ilLoop)
                    Else
                        If tlDigitalLineAvg.dBilledRemainder(ilLoop) <> 0 Then
                            If ilLoop = 0 And tlDigitalLineAvg.dBilledRemainder(0) <> 0 Then
                                strComment = "Billed " & mGetMonthName(tlDigitalLineAvg.iMonths(ilLoop) + 1) & ":" & tlDigitalLineAvg.dBilledRemainder(ilLoop)
                            Else
                                strComment = "Billed " & mGetMonthName(tlDigitalLineAvg.iMonths(ilLoop)) & ":" & tlDigitalLineAvg.dBilledRemainder(ilLoop)
                            End If
                            If ilLoop = 0 Then
                                strComment = strComment & " (" & tlDigitalLineAvg.iDaysExt(ilLoop) & " days "
                                strComment = strComment & "@" & tlDigitalLineAvg.dDailyAmt(ilLoop) & ")"
                            End If
                        Else
                            strComment = "Billed " & mGetMonthName(tlDigitalLineAvg.iMonths(ilLoop)) & ":" & tlDigitalLineAvg.dBilledAmt(ilLoop)
                            If ilLoop = 0 Then
                                strComment = strComment & " (" & tlDigitalLineAvg.iDaysExt(ilLoop) & " days "
                                strComment = strComment & "@" & tlDigitalLineAvg.dDailyAmt(ilLoop) & ")"
                            End If
                        End If
                    End If
                    If ilLoop > 0 Then
                        If tlDigitalLineAvg.iDaysAdjNext(ilLoop) <> 0 Then
                            'Show adjusted Days from Prior Month
                            If InStr(1, strComment, "Billed Remainder") = 0 And tlDigitalLineAvg.iDaysAdjPrior(ilLoop) <> 0 Then 'TTP 10741
                                strComment = strComment & " + "
                                strComment = strComment & "'" & tlDigitalLineAvg.iDaysAdjPrior(ilLoop) & " days " & mGetMonthName(tlDigitalLineAvg.iMonths(ilLoop) - 1)
                                strComment = strComment & "@" & tlDigitalLineAvg.dDailyAmt(ilLoop)
                                strComment = strComment & "'"
                            End If
                            'Show adjusted Days from Next Month
                            If tlDigitalLineAvg.dDailyAmt(ilLoop + 1) <> 0 Then 'TTP 10741
                                strComment = strComment & " + "
                                strComment = strComment & "'" & tlDigitalLineAvg.iDaysAdjNext(ilLoop) & " days " & mGetMonthName(tlDigitalLineAvg.iMonths(ilLoop) + 1)
                                strComment = strComment & "@" & tlDigitalLineAvg.dDailyAmt(ilLoop + 1)
                                strComment = strComment & "'"
                            End If
                        End If
                    End If
                End If
            Else
                strComment = "Projected:"
                'Current Month
                If tlDigitalLineAvg.iDaysInMonthStd(ilLoop) > 0 And tlDigitalLineAvg.dDailyAmt(ilLoop) <> 0 Then
                    If sBillCycle = "S" Then
                        strComment = strComment & "'" & tlDigitalLineAvg.iDaysInMonthStd(ilLoop) & " days " & mGetMonthName(tlDigitalLineAvg.iMonths(ilLoop))
                    Else
                        strComment = strComment & "'" & tlDigitalLineAvg.iDaysInMonthCal(ilLoop) & " days " & mGetMonthName(tlDigitalLineAvg.iMonths(ilLoop))
                    End If
                    strComment = strComment & "@" & tlDigitalLineAvg.dDailyAmt(ilLoop)
                    strComment = strComment & "'"
                End If
                If ilLoop > 0 Then
                    'Prior Month Adjustments
                    If tlDigitalLineAvg.iDaysAdjPrior(ilLoop) <> 0 And tlDigitalLineAvg.dDailyAmt(ilLoop - 1) <> 0 Then
                        If right(strComment, 1) = "'" Then strComment = strComment & " + "
                        If ilLoop = tlDigitalLineAvg.iProjectedInx Then
                            strComment = strComment & "'" & tlDigitalLineAvg.iDaysAdjPrior(ilLoop) & " days " & mGetMonthName(tlDigitalLineAvg.iMonths(ilLoop - 1))
                            strComment = strComment & "@" & tlDigitalLineAvg.dDailyAmt(ilLoop)
                        Else
                            strComment = strComment & "'" & tlDigitalLineAvg.iDaysAdjPrior(ilLoop) & " days " & mGetMonthName(tlDigitalLineAvg.iMonths(ilLoop - 1))
                            strComment = strComment & "@" & tlDigitalLineAvg.dDailyAmt(ilLoop - 1)
                        End If
                        strComment = strComment & "'"
                    End If
                End If
                If tlDigitalLineAvg.iDaysAdjNext(ilLoop) <> 0 And tlDigitalLineAvg.dDailyAmt(ilLoop + 1) <> 0 Then
                    'Next Month Adjustments
                    If ilLoop <> tlDigitalLineAvg.iLastInx Then 'dont include adjustment for last index + 1
                        If ilLoop = tlDigitalLineAvg.iLastInx - 1 And tlDigitalLineAvg.iDaysExt(tlDigitalLineAvg.iLastInx) < 0 Then
                            If right(strComment, 1) = "'" Then strComment = strComment & " + "
                            strComment = strComment & "'" & (tlDigitalLineAvg.iDaysAdjNext(ilLoop) + tlDigitalLineAvg.iDaysExt(tlDigitalLineAvg.iLastInx)) & " days " & mGetMonthName(tlDigitalLineAvg.iMonths(ilLoop) + 1)
                            strComment = strComment & "@" & tlDigitalLineAvg.dDailyAmt(ilLoop + 1)
                            strComment = strComment & "'"
                        Else
                            If right(strComment, 1) = "'" Then strComment = strComment & " + "
                            strComment = strComment & "'" & tlDigitalLineAvg.iDaysAdjNext(ilLoop) & " days " & mGetMonthName(tlDigitalLineAvg.iMonths(ilLoop) + 1)
                            strComment = strComment & "@" & tlDigitalLineAvg.dDailyAmt(ilLoop + 1)
                            strComment = strComment & "'"
                        End If
                    End If
                End If
            End If
            If strComment = "Projected:" Then
                strComment = ""
            End If
        End If
        tlDigitalLineAvg.sComment(ilLoop) = strComment
    Next ilLoop
    
    'Debug.Print ""
    'Debug.Print " - Comments:";
    For ilLoop = 0 To tlDigitalLineAvg.iLastInx
        If tlDigitalLineAvg.sComment(ilLoop) <> "" Then
            'Debug.Print "(" & tlDigitalLineAvg.iMonths(ilLoop) & "/" & tlDigitalLineAvg.iYears(ilLoop) & ")" & tlDigitalLineAvg.sComment(ilLoop) & ", ";
        End If
    Next ilLoop
    'Debug.Print ""
    gBuildDigitalLineAverage = tlDigitalLineAvg
    Exit Function
    
ExitBuildDigitalLine:
    Debug.Print ""
    Debug.Print " ! gBuildDigitalLineAverage ERROR: " & err & " - " & Error(err)
    On Error GoTo 0
    gBuildDigitalLineAverage = tlDigitalLineAvg
End Function

'TTP 10666 - this function returns the Abbreviated Month Name.
'If the Month is > 12, it assumes to wrap to next year.  For example: 13=Jan
'If the Month is < 1, it assumes to wrap to prior year.  For example: 0=Dec, -1=Nov
'only designed for -11 to 24 months
Function mGetMonthName(ilMonthNo) As String
    If ilMonthNo < -11 Then Exit Function
    If ilMonthNo > 24 Then Exit Function
    If ilMonthNo > 12 Then ilMonthNo = ilMonthNo - 12
    If ilMonthNo < 1 Then ilMonthNo = ilMonthNo + 12
    mGetMonthName = MonthName(ilMonthNo, True)
End Function

'    gObtainUAFbyDate (hlUAF, slActiveDate,  tlUAF())
'           <input>  RptForm - form name source
'                    hlUAF - User Activity handle
'                    slActiveDate - Active date to gather activity
'           <output> tlUAF() array of UAF records
'           <return> true if valid reads
Public Function gObtainUafByDate(RptForm As Form, hlUAF, tlUAF() As UAF, slActiveDate As String) As Integer
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim slDate As String
    
    btrExtClear hlUAF   'Clear any previous extend operation
    ilExtLen = Len(tlUAF(0))  'Extract operation record size
    imUafRecLen = Len(tlUAF(0))

    ilRet = btrGetFirst(hlUAF, tmUaf, imUafRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hlUAF, llNoRec, -1, "UC", "UAF", "") '"EG") 'Set extract limits (all records)

        gPackDate slActiveDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("UAF", "UAFStartDate")
        ilRet = btrExtAddLogicConst(hlUAF, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        ilRet = btrExtAddField(hlUAF, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainUAFErr
        gBtrvErrorMsg ilRet, "gObtainUAF (btrExtAddField):" & "UAF.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlUAF, tmUaf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainUAFErr
            gBtrvErrorMsg ilRet, "gObtainUAF (btrExtGetNextExt):" & "UAF.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmUaf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlUAF, tmUaf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tlUAF(UBound(tlUAF)) = tmUaf           'save entire record
                ReDim Preserve tlUAF(0 To UBound(tlUAF) + 1) As UAF
                ilRet = btrExtGetNext(hlUAF, tmUaf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlUAF, tmUaf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainUafByDate = True
    Exit Function
mObtainUAFErr:
    On Error GoTo 0
    MsgBox "RptExtra: gObtainUAF error", vbCritical + vbOkOnly, "UAF I/O Error"
    gObtainUafByDate = False
    Exit Function
End Function

'               Obtain all Affiliate Users in memory
'               Build all Affiliate users into tgUST array
Public Function gObtainUst() As Integer
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim hlUst As Integer
    ReDim tgUst(0 To 0) As UST

    gObtainUst = BTRV_ERR_NONE
    hlUst = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlUst, "", sgDBPath & "Ust.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainUst = ilRet
        ilRet = btrClose(hlUst)
        btrDestroy hlUst
        Exit Function
    End If

    ilRecLen = Len(tgUst(0))
    ilRet = btrGetFirst(hlUst, tgUst(UBound(tgUst)), ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tgUst(0 To UBound(tgUst) + 1) As UST
        ilRet = btrGetNext(hlUst, tgUst(UBound(tgUst)), ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilRet = btrClose(hlUst)
    btrDestroy hlUst
    Exit Function
End Function

'           gPopANF - build Named AVails into array
'
Public Sub gPopAnf(hlAnf As Integer, tlAnf() As ANF)
    Dim ilRet As Integer
    Dim ilRecLen As Integer

    ReDim tlAnf(0 To 0) As ANF
    ilRecLen = Len(tlAnf(0))
    ilRet = btrGetFirst(hlAnf, tlAnf(UBound(tlAnf)), ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tlAnf(0 To UBound(tlAnf) + 1) As ANF
        ilRet = btrGetNext(hlAnf, tlAnf(UBound(tlAnf)), ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Exit Sub
End Sub

'           gBinarySearchVaf - find the matching vehicle code in array that contains
'           index references to Great Plains options by vehicle
'           <input> ilVefcode = vehicle code to match
'                   tlPIFKEY - vehicle table to search
'           return - index to matching vehicle entry
'                    -1 if not found
Public Function gBinarySearchVaf(ilVefCode As Integer, tlVaf() As VAF) As Integer
    Dim ilMiddle As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    ilMin = LBound(tlVaf)
    ilMax = UBound(tlVaf)
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilVefCode = tlVaf(ilMiddle).iVefCode Then
            'found the match
            gBinarySearchVaf = ilMiddle
            Exit Function
        ElseIf ilVefCode < tlVaf(ilMiddle).iVefCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchVaf = -1
End Function

Public Sub gPopVaf(hlVaf As Integer, tlVaf() As VAF)
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    
    ReDim tlVaf(0 To 0) As VAF
    ilRecLen = Len(tlVaf(0))
    ilRet = btrGetFirst(hlVaf, tlVaf(UBound(tlVaf)), ilRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tlVaf(0 To UBound(tlVaf) + 1) As VAF
        ilRet = btrGetNext(hlVaf, tlVaf(UBound(tlVaf)), ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Exit Sub
End Sub

Function gMnfTermsPop(Form As Form, cbcList As Control) As Integer
    Dim ilRet As Integer
    Dim slName As String

    gMnfTermsPop = False           'assume OK
    ilRet = gPopMnfPlusFieldsBox(Form, cbcList, tgMNFCodeRpt(), sgMNFCodeTagRpt, "J")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo gMnfTermsPopErr
        gCPErrorMsg ilRet, "gMnfTermsPop (gPopMnfPlusFieldsBox)", Form
        On Error GoTo 0
        cbcList.AddItem "[Use Assigned]", 0

        gFindMatch slName, 1, cbcList
        If gLastFound(cbcList) > 0 Then
            cbcList.ListIndex = gLastFound(cbcList)
        Else
            cbcList.ListIndex = -1
        End If
    End If
    cbcList.ListIndex = 0
    Exit Function
gMnfTermsPopErr:
    On Error GoTo 0
    gMnfTermsPop = True           'error
End Function

Sub gCbfClearBylgNowTime()
'*******************************************************
'*                                                     *
'*      Procedure Name:gCRCbf    Clear                 *
'*                                                     *
'*             Created:05/29/96      By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear Contract/Proposal  data   *
'*                     for Crystal report              *
'*          9-14-09 Clear the prepass table (CBF)      *
'*          using lgGenTime (not igGenTime(0 to 1) as
'*          the time used was converted to milliseconds
'           in various reports (i.e. Contract/Proposals
'           Insertion Order
'*******************************************************
    Dim ilRet As Integer
    Dim llNowTime As Long       '10-10-01 gen time
    hmCbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCbf)
        btrDestroy hmCbf
        Exit Sub
    End If
    '10-10-01
    'gUnpackTimeLong igNowTime(0), igNowTime(1), False, llNowTime       '9-14-09
    imCbfRecLen = Len(tmCbf)
    tmCbfSrchKey.iGenDate(0) = igNowDate(0)
    tmCbfSrchKey.iGenDate(1) = igNowDate(1)
    '10-10-01
    'tmCbfSrchKey.lGenTime = llNowTime
     '9-14-09
     tmCbfSrchKey.lGenTime = lgNowTime
     
     ilRet = btrGetGreaterOrEqual(hmCbf, tmCbf, imCbfRecLen, tmCbfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    '10-10-01
    'Do While (ilRet = BTRV_ERR_NONE) And (tmCbf.iGenDate(0) = igNowDate(0)) And (tmCbf.iGenDate(1) = igNowDate(1)) And (tmCbf.lGenTime = llNowTime)
     '9-14-09
    Do While (ilRet = BTRV_ERR_NONE) And (tmCbf.iGenDate(0) = igNowDate(0)) And (tmCbf.iGenDate(1) = igNowDate(1)) And (tmCbf.lGenTime = lgNowTime)
        ilRet = btrDelete(hmCbf)
        ilRet = btrGetNext(hmCbf, tmCbf, imCbfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmCbf)
    btrDestroy hmCbf
End Sub

Public Function gObtainCrfByRegionDateSpan(RptForm As Form, hlCrf, tlCrf() As CRF, slActiveStart As String, slActiveEnd As String) As Integer
'
'    gObtainCRF (hlCRF, slActiveDate,  tlCRF())
'           <input>  RptForm - form name source
'                    hlCrf - Copy Rotation Header handle
'                    slActiveStart - Active start date to gather rotation headers
'                    slActiveEnd - active end date to gather rotation headers
'           <output> tlCRF() array of CRF records
'           <return> true if valid reads
'
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim slDate As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record

    btrExtClear hlCrf   'Clear any previous extend operation
    ilExtLen = Len(tlCrf(0))  'Extract operation record size
    imCrfRecLen = Len(tlCrf(0))

    ilRet = btrGetFirst(hlCrf, tmCrf, imCrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hlCrf, llNoRec, -1, "UC", "CRF", "") '"EG") 'Set extract limits (all records)

        tlIntTypeBuff.iType = 0
        ilOffSet = gFieldOffset("Crf", "CrfRafCode")
        ilRet = btrExtAddLogicConst(hlCrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_GT, BTRV_EXT_AND, tlIntTypeBuff, 2)

         ' chfEndDate >= InputStartDate And chfStartDate <= InputEndDate
        If slActiveStart = "" Then
            slDate = "1/1/1970"
        Else
            slDate = slActiveStart
        End If
        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Crf", "CrfEndDate")
        ilRet = btrExtAddLogicConst(hlCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        If slActiveEnd = "" Then
            slDate = "12/31/2069"
        Else
            slDate = slActiveEnd
        End If

        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Crf", "CrfStartDate")
        ilRet = btrExtAddLogicConst(hlCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        'gPackDate slActiveDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        'ilOffset = gFieldOffset("Crf", "CrfEndDate")
        'ilRet = btrExtAddLogicConst(hlCrf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        ilRet = btrExtAddField(hlCrf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainCrfErr
        gBtrvErrorMsg ilRet, "gObtainCrf (btrExtAddField):" & "Crf.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlCrf, tmCrf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainCrfErr
            gBtrvErrorMsg ilRet, "gObtainCrf (btrExtGetNextExt):" & "Crf.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmCrf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlCrf, tmCrf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tlCrf(UBound(tlCrf)) = tmCrf           'save entire record
                ReDim Preserve tlCrf(0 To UBound(tlCrf) + 1) As CRF
                ilRet = btrExtGetNext(hlCrf, tmCrf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlCrf, tmCrf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainCrfByRegionDateSpan = True
    Exit Function
mObtainCrfErr:
    On Error GoTo 0
    MsgBox "RptExtra: gObtainCrfByRegionDateSpan error", vbCritical + vbOkOnly, "Crf I/O Error"
    gObtainCrfByRegionDateSpan = False
    Exit Function
End Function

'                   gSbfAdjustForNTR -  determine where item billing $ goes
'                   <input> tlSbf() - array of SBF records to process
'                           llstdstartDates() - array of 13 start dates of the 12 months to gather
'                           ilFirstProjInx - index of first month to start projection (earlier is from receivables)
'                           llStartAdjust -  Earliest date to start searching for missed, etc.
'                           llEndAdjust - latest date to stop searchng for missed, etc.
'                           ilHowManyPer - # of periods to gather
'
'
Sub gSbfAdjustForNTR(tlSbf() As SBF, tlSBFAdjust() As ADJUSTLIST, llStdStartDates() As Long, ilFirstProjInx As Integer, llStartAdjust As Long, llEndAdjust As Long, ilHowManyPer As Integer, tlMMnf() As MNF)
    Dim slDate As String
    Dim llDate As Long
    Dim ilMonthInx As Integer
    Dim ilFoundMonth As Integer
    Dim ilFoundVef As Integer
    Dim ilTemp As Integer
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSBFLoop As Integer
    Dim llSBFLoop As Long
    Dim ilIsItHardCost As Integer
    Dim ilFoundOption As Integer

    For llSBFLoop = LBound(tlSbf) To UBound(tlSbf) - 1
        tmSbf = tlSbf(llSBFLoop)
        gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), slDate
        llDate = gDateValue(slDate)
        If llDate > llEndAdjust Then
            Exit Sub
        End If
        'SBF is OK with dates, adjust the $

        'determine if hard cost to be included
        ilFoundOption = True                'default to include the NTR if no options exist and its not B & B or Recap report
        ilIsItHardCost = gIsItHardCost(tlSbf(llSBFLoop).iMnfItem, tlMMnf())

        ilFoundMonth = False
        For ilMonthInx = ilFirstProjInx To ilHowManyPer Step 1         'loop thru months to find the match
            If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                ilFoundMonth = True
                Exit For
            End If
        Next ilMonthInx

        ilFoundVef = False
        'setup vehicle that spot was moved to
        For ilTemp = LBound(tlSBFAdjust) To UBound(tlSBFAdjust) - 1 Step 1
            If tlSBFAdjust(ilTemp).iVefCode = tlSbf(llSBFLoop).iBillVefCode And tlSBFAdjust(ilTemp).iSlsCommPct = tlSbf(llSBFLoop).iCommPct And tlSBFAdjust(ilTemp).sAgyComm = tlSbf(llSBFLoop).sAgyComm And tlSBFAdjust(ilTemp).iMnfItem = tlSbf(llSBFLoop).iMnfItem Then      '8-10-06
                ilFoundVef = True
                Exit For
            End If
        Next ilTemp
        If Not (ilFoundVef) Then
            ilTemp = UBound(tlSBFAdjust)
            tlSBFAdjust(ilTemp).iVefCode = tlSbf(llSBFLoop).iBillVefCode
            tlSBFAdjust(ilTemp).iSlsCommPct = tlSbf(llSBFLoop).iCommPct
            tlSBFAdjust(ilTemp).sAgyComm = tlSbf(llSBFLoop).sAgyComm
            tlSBFAdjust(ilTemp).iIsItHardCost = ilIsItHardCost              '8-10-06
            tlSBFAdjust(ilTemp).iMnfItem = tlSbf(llSBFLoop).iMnfItem        '8-10-06

            If ilFoundMonth Then
                tlSBFAdjust(ilTemp).lProject(ilMonthInx) = tlSBFAdjust(ilTemp).lProject(ilMonthInx) + (tlSbf(llSBFLoop).lGross * tlSbf(llSBFLoop).iNoItems)
                tlSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) = tlSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) + (tlSbf(llSBFLoop).lAcquisitionCost * tlSbf(llSBFLoop).iNoItems)
            End If
            ReDim Preserve tlSBFAdjust(0 To UBound(tlSBFAdjust) + 1) As ADJUSTLIST
        Else
            If ilFoundMonth Then
                tlSBFAdjust(ilTemp).lProject(ilMonthInx) = tlSBFAdjust(ilTemp).lProject(ilMonthInx) + (tlSbf(llSBFLoop).lGross * tlSbf(llSBFLoop).iNoItems)
                tlSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) = tlSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) + (tlSbf(llSBFLoop).lAcquisitionCost * tlSbf(llSBFLoop).iNoItems)
            End If
        End If


    Next llSBFLoop
    Exit Sub
End Sub

'               gSbfAdjustForInstall - get the Installment billing from SBF file
'               for projecting into future
'
Sub gSbfAdjustForInstall(tlSbf() As SBF, tlSBFAdjust() As ADJUSTLIST, llStdStartDates() As Long, ilFirstProjInx As Integer, llStartAdjust As Long, llEndAdjust As Long, ilHowManyPer As Integer, ilAgfCode As Integer)
    Dim slDate As String
    Dim llDate As Long
    Dim ilMonthInx As Integer
    Dim ilFoundMonth As Integer
    Dim ilFoundVef As Integer
    Dim ilTemp As Integer
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSBFLoop As Integer
    Dim llSBFLoop As Long
    Dim ilFoundOption As Integer

    For llSBFLoop = LBound(tlSbf) To UBound(tlSbf) - 1
        tmSbf = tlSbf(llSBFLoop)
        gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), slDate
        llDate = gDateValue(slDate)
        If llDate > llEndAdjust Then
            Exit Sub
        End If
        'SBF is OK with dates, adjust the $

        ilFoundOption = True
        If tlSbf(llSBFLoop).sTranType <> "F" Then            'installment
            ilFoundOption = False
        End If

        If ilFoundOption Then
            ilFoundMonth = False
            For ilMonthInx = ilFirstProjInx To ilHowManyPer Step 1         'loop thru months to find the match
                If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                    ilFoundMonth = True
                    Exit For
                End If
            Next ilMonthInx

            ilFoundVef = False
            'setup vehicle that spot was moved to
            For ilTemp = LBound(tlSBFAdjust) To UBound(tlSBFAdjust) - 1 Step 1
                If tlSBFAdjust(ilTemp).iVefCode = tlSbf(llSBFLoop).iBillVefCode Then
                    ilFoundVef = True
                    Exit For
                End If
            Next ilTemp
            If Not (ilFoundVef) Then
                ilTemp = UBound(tlSBFAdjust)
                tlSBFAdjust(ilTemp).iVefCode = tlSbf(llSBFLoop).iBillVefCode
                'Installments from SBF need to set the agy Commission flag since its not coming from the SBF record,
                'but from the contract header
                tlSBFAdjust(ilTemp).iSlsCommPct = 0                         '4-25-08
                If ilAgfCode > 0 Then
                    tlSBFAdjust(ilTemp).sAgyComm = "Y"   '4-25-08
                Else
                    tlSBFAdjust(ilTemp).sAgyComm = "N"
                End If
                tlSBFAdjust(ilTemp).iIsItHardCost = False                   '4-25-08
                tlSBFAdjust(ilTemp).iMnfItem = 0                           '4-25-08

                If ilFoundMonth Then
                    tlSBFAdjust(ilTemp).lProject(ilMonthInx) = tlSBFAdjust(ilTemp).lProject(ilMonthInx) + tlSbf(llSBFLoop).lGross
                    tlSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) = tlSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) + (tlSbf(llSBFLoop).lAcquisitionCost)
                End If
                ReDim Preserve tlSBFAdjust(0 To UBound(tlSBFAdjust) + 1) As ADJUSTLIST
            Else
                If ilFoundMonth Then
                    tlSBFAdjust(ilTemp).lProject(ilMonthInx) = tlSBFAdjust(ilTemp).lProject(ilMonthInx) + tlSbf(llSBFLoop).lGross
                    tlSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) = tlSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) + (tlSbf(llSBFLoop).lAcquisitionCost)
                End If
            End If
        End If
    Next llSBFLoop
    Exit Sub
End Sub

Function gObtainPhfRvfforSort(RptForm As Form, slEarliestDate As String, slLatestDate As String, tlTranType As TRANTYPES, tlRvfSort() As RVFCNTSORT, ilWhichDate As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                                                                                 *
'******************************************************************************************

'****************************************************************
'*
'*      Obtain all History and Receivables transactions whose
'*      transaction date falls within the earliest and latest
'*      dates requested.  Test for transaction types "I", "A"
'*      or "W".
'
'*      <input>  RptForm - Form calling this populate rtn
'*               slEarliestDate - get all trans starting with
'*                   this date
'*               slLatestDAte - get all trans equal or prior to
'*                  this pacing date (effective date)
'*              ilWhichDate, 0 = trandate, 1 = entry date
'*
'*      <I/O>    tlRvfSort() - array of matching Phf/Rvf recds
'*               funtion return - true if receivables populated
'*                       false if no receivables, error
'*
'*             Created:4/17/98       By:D. Hosaka
'*            Modified:              By:
'*
'*            Comments: make 2 passes, read Phf, the Rvf, build
'*               all transactions types "A", "I", "W" or "P"
'*               based on parameters in tlTranType.
'*            3-14-01 Include "HI" with the "I" transactions
'            9-17-02 Add NTR flag to list of filters (test for presence of item bill mnfcode
'*          5-26-04 Exclude NTR "AN" transactions when NTR to be excluded
'*          2-11-05 Change array from single to double integer to prevent Overflow
'                   (records in excess of 32000)
'****************************************************************
'
'    ilRet = gObtainPhfRvf (RptForm,  slEarliestDate, slLatestDate, tlTranType, tlRvf())
'
    Dim ilRet As Integer    'Return status
    ReDim ilEarliestDate(0 To 1) As Integer
    ReDim ilLatestDate(0 To 1) As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilLoop As Integer
    'Dim ilRVFUpper As Integer
    Dim llRVFUpper As Long          '2-11-05 chg to long
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilVefInx As Integer
    Dim slVehicleName As String
    Dim slContract As String
    Dim ilLowLimit As Integer

    'On Error GoTo gObtainPhfRvfForSortErr
    'ilRet = 0
    'ilLowLimit = LBound(tlRvfSort)
    'If ilRet <> 0 Then
    '    ilLowLimit = 1
    'End If
    'On Error GoTo 0
    If PeekArray(tlRvfSort).Ptr <> 0 Then
        ilLowLimit = LBound(tlRvfSort)
    Else
        ilLowLimit = 0
    End If

    ReDim tlRvfSort(ilLowLimit To ilLowLimit) As RVFCNTSORT
    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        gObtainPhfRvfforSort = False
        Exit Function
    End If
    imRvfRecLen = Len(tmRvf)
    gPackDate slEarliestDate, ilEarliestDate(0), ilEarliestDate(1)
    gPackDate slLatestDate, ilLatestDate(0), ilLatestDate(1)
    btrExtClear hmRvf   'Clear any previous extend operation
    ilExtLen = Len(tmRvf)  'Extract operation record size
    llRVFUpper = UBound(tlRvfSort)
    For ilLoop = 1 To 2         'pass 1- get PHF, pass 2 get RVF

        ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
            Call btrExtSetBounds(hmRvf, llNoRec, -1, "UC", "RVF", "") '"EG") 'Set extract limits (all records)

            tlDateTypeBuff.iDate0 = ilEarliestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilEarliestDate(1)
            If ilWhichDate = 0 Then                         '12-14-06
                ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
            Else
                ilOffSet = gFieldOffset("Rvf", "RvfDateEntrd")
            End If
            ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainPhfRvfforSort (btrExtAddLogicConst):" & "Rvf.Btr", RptForm
            On Error GoTo 0

            tlDateTypeBuff.iDate0 = ilLatestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilLatestDate(1)
            If ilWhichDate = 0 Then                         '12-14-06
                ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
            Else
                ilOffSet = gFieldOffset("Rvf", "RvfDateEntrd")
            End If

            ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainPhfRvfforSort (btrExtAddLogicConst):" & "Rvf.Btr", RptForm
            On Error GoTo 0

            ilRet = btrExtAddField(hmRvf, 0, ilExtLen)  'Extract the whole record
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainRVF (btrExtAddField):" & "RVF.Btr", RptForm
            On Error GoTo 0
            ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo mRvfErr
                gBtrvErrorMsg ilRet, "gObtainRVF (btrExtGetNextExt):" & "RVF.Btr", RptForm
                On Error GoTo 0
                ilExtLen = Len(tmRvf)  'Extract operation record size
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    'first test for valid trans types (Invoices, adjustments, write-off & payments
                    If ((Left$(tmRvf.sTranType, 1) = "I" Or tmRvf.sTranType = "HI") And tlTranType.iInv) Or (Left$(tmRvf.sTranType, 1) = "A" And tlTranType.iAdj) Or (Left$(tmRvf.sTranType, 1) = "W" And tlTranType.iWriteOff) Or (Left$(tmRvf.sTranType, 1) = "P" And tlTranType.iPymt) Then
                        'obtain the info to build the key for sortin
                        ilVefInx = gBinarySearchVef(tmRvf.iAirVefCode)
                        If ilVefInx = -1 Then
                            slVehicleName = "Unknown Vehicle"
                        Else
                            slVehicleName = Trim$(tgMVef(ilVefInx).sName)
                        End If
                        slContract = Trim$(str$(tmRvf.lCntrNo))
                        Do While Len(slContract) < 9
                            slContract = "0" & slContract
                        Loop

                        If (tlTranType.iNTR) Then       'NTR option, tested separately because it shouldnt be tested with Cash transactions
                            If tmRvf.iMnfItem > 0 Then
                                tlRvfSort(UBound(tlRvfSort)).tRVF = tmRvf           'save entire record
                                tlRvfSort(UBound(tlRvfSort)).sKey = Trim$(slContract) & "|" & Trim$(slVehicleName) & "|" & tmRvf.sCashTrade
                                ReDim Preserve tlRvfSort(ilLowLimit To UBound(tlRvfSort) + 1) As RVFCNTSORT
                            Else            'its not an NTR
                                'got valid trans type - test for Cash, Trade, Merchandising or Promotions
                                If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                    tlRvfSort(UBound(tlRvfSort)).tRVF = tmRvf           'save entire record
                                    tlRvfSort(UBound(tlRvfSort)).sKey = Trim$(slContract) & "|" & Trim$(slVehicleName) & "|" & tmRvf.sCashTrade
                                    ReDim Preserve tlRvfSort(ilLowLimit To UBound(tlRvfSort) + 1) As RVFCNTSORT

                                End If
                            End If
                        Else
                            'got valid trans type - test for Cash, Trade, Merchandising or Promotions
                            '05-26-04 dont include NTR, exclude this if it is
                            '2-27-08 chg test to test for NTR using mnitem only; Installment records will also have an sbfcode
                            If tmRvf.iMnfItem = 0 Then      'and tmRvf.isbfcode = 0
                                If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                    tlRvfSort(UBound(tlRvfSort)).tRVF = tmRvf           'save entire record
                                    tlRvfSort(UBound(tlRvfSort)).sKey = Trim$(slContract) & "|" & Trim$(slVehicleName) & "|" & tmRvf.sCashTrade
                                    ReDim Preserve tlRvfSort(ilLowLimit To UBound(tlRvfSort) + 1) As RVFCNTSORT
                                End If
                            End If
                        End If
                    End If
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
        If ilLoop = 1 Then                          'if 1, then just finished history, go do Receivables
            btrExtClear hmRvf   'Clear any previous extend operation
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmRvf)
                btrDestroy hmRvf
                gObtainPhfRvfforSort = False
                Exit Function
            End If
            imRvfRecLen = Len(tmRvf)
            llRVFUpper = UBound(tlRvfSort)
        End If
    Next ilLoop
    ilRet = btrClose(hmRvf)
    btrDestroy hmRvf
    gObtainPhfRvfforSort = True
    Exit Function
gObtainPhfRvfForSortErr:
    ilRet = 0
    Resume Next
mRvfErr:
    On Error GoTo 0
    gObtainPhfRvfforSort = False
    Exit Function
End Function

'           gTestIncludeExclude - test array of codes (stored in ilUseCodes) to determine if the value
'           should be included or excluded (determined by flag ilIncludeCodes)
'           <input>  ilValue = value to test for inclusion/exclusion
'                    ilIncludeCodes - true = include code, false = exclude code
'                    ilUseCodes() - array of codes to include or exclude
'           <return> true to include, false to exclude
'
Public Function gTestIncludeExclude(ilValue As Integer, ilIncludeCodes As Integer, ilUseCodes() As Integer) As Integer
    Dim ilTemp As Integer
    Dim ilFoundOption As Integer
    If ilIncludeCodes Then          'include the any of the codes in array?
        ilFoundOption = False
        For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
            If ilUseCodes(ilTemp) = ilValue Then
                ilFoundOption = True                    'include the matching vehicle
                Exit For
            End If

        Next ilTemp
    Else                            'exclude any of the codes in array?
        ilFoundOption = True
        For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
            If ilUseCodes(ilTemp) = ilValue Then
                ilFoundOption = False                  'exclude the matching vehicle
                Exit For
            End If
        Next ilTemp
    End If
    gTestIncludeExclude = ilFoundOption
    Exit Function
End Function

'           gObtainDRFByCode -obtain all matching demo records by book name code
'
Public Function gObtainDRFByCode(RptForm As Form, hlDrf As Integer, tlDrf() As DRF, ilDnfCode As Integer) As Integer
'
'    gObtainDRF (hlDRF, tlDRF(), llDrfCode)
'           <input>  RptForm - form name source
'                    hlDRF - DRF handle
'                    ilDnfCode - book name code
'           <output> tlDrF() array of DrF demo records
'           <return> true if valid reads
'
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim tlIntTypeBuff As POPICODE

    ReDim tlDrf(0 To 0) As DRF
    btrExtClear hlDrf   'Clear any previous extend operation
    ilExtLen = Len(tlDrf(0))  'Extract operation record size
    imDrfRecLen = Len(tlDrf(0))

    tmDrfSrchKey.iDnfCode = ilDnfCode
    tmDrfSrchKey.iMnfSocEco = 0
    tmDrfSrchKey.iRdfCode = 0
    tmDrfSrchKey.iVefCode = 0
    tmDrfSrchKey.sDemoDataType = ""
    tmDrfSrchKey.sInfoType = ""

    ilRet = btrGetGreaterOrEqual(hlDrf, tmDrf, imDrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hlDrf, llNoRec, -1, "UC", "DRF", "") '"EG") 'Set extract limits (all records)
        tlIntTypeBuff.iCode = ilDnfCode
        ilOffSet = gFieldOffset("DRF", "DRFDnfCode")
        ilRet = btrExtAddLogicConst(hlDrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)

        ilRet = btrExtAddField(hlDrf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainDRFErr
        gBtrvErrorMsg ilRet, "gObtainDRF (btrExtAddField):" & "DRF.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlDrf, tmDrf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainDRFErr
            gBtrvErrorMsg ilRet, "gObtainDRF (btrExtGetNextExt):" & "DRF.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmDrf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlDrf, tmDrf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE And tmDrf.iDnfCode = ilDnfCode
                tlDrf(UBound(tlDrf)) = tmDrf           'save entire record
                ReDim Preserve tlDrf(0 To UBound(tlDrf) + 1) As DRF
                ilRet = btrExtGetNext(hlDrf, tmDrf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlDrf, tmDrf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainDRFByCode = True
    Exit Function
mObtainDRFErr:
    On Error GoTo 0
    MsgBox "RptExtra: gObtainDRF error", vbCritical + vbOkOnly, "DRF I/O Error"
    gObtainDRFByCode = False
    Exit Function
End Function

'***************************************************************************************
'*
'*      Procedure Name:gGetSpotsByVefDate - obtain spots from SDF by Key1:
'*                                          Vehicle, Date, Time, SchStatus
'*      <input> hlSdf - SDF handle
'*              ilVefCode - vehicle code
'*              slStartDate - earliest date to retrieve spots
'*              slEndDate - latest date to retrieve spots
'*              tlSdf() - array of spots for the requested span
'*      <return>
'*
'*     Created:8/19/05       By:D. Hosaka
'*
'***************************************************************************************
Public Function gGetSpotsbyVefDate(hlSdf As Integer, ilVefCode As Integer, slStartDate As String, slEndDate As String, tlSdf() As SDF) As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim llChfCode As Long
    Dim tlSdfSrchKey1 As SDFKEY1
    Dim ilSdfRecLen As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record

    ReDim tlSdf(0 To 0) As SDF
    Dim ilUpper As Integer

    gGetSpotsbyVefDate = True
    btrExtClear hlSdf   'Clear any previous extend operation
    ilExtLen = Len(tlSdf(0))  'Extract operation record size
    tlSdfSrchKey1.iVefCode = ilVefCode
    gPackDate slStartDate, tlSdfSrchKey1.iDate(0), tlSdfSrchKey1.iDate(1)
    tlSdfSrchKey1.iTime(0) = 0
    tlSdfSrchKey1.iTime(1) = 0
    tlSdfSrchKey1.sSchStatus = ""   'slType
    ilSdfRecLen = Len(tlSdf(0))
    ilUpper = 0
    ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf(0), ilSdfRecLen, tlSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
    If (tlSdf(ilUpper).iVefCode = ilVefCode) And (ilRet <> BTRV_ERR_END_OF_FILE) Then

        ' Prepare to execute an extended operation.
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC", "SDF", "") '"EG") 'Set extract limits (all records)

        ' We only the records for the passed in vehicle code.
        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, ilVefCode, 2)


        ' And only records where the ChfCode = 0
        llChfCode = 0
        ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, llChfCode, 4)

        gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)


        ' And on the records where the date is between the passed  date
        gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        ilRet = btrExtAddField(hlSdf, 0, ilExtLen) 'Extract the whole record

        ilRet = btrExtGetNext(hlSdf, tlSdf(ilUpper), ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            ilExtLen = Len(tlSdf(0))  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSdf, tlSdf(ilUpper), ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE

                'tlSDF(ilUpper) = tlSDF
                ReDim Preserve tlSdf(0 To UBound(tlSdf) + 1) As SDF
                ilUpper = ilUpper + 1
                ilExtLen = Len(tlSdf(0))
                ilRet = btrExtGetNext(hlSdf, tlSdf(ilUpper), ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSdf, tlSdf(ilUpper), ilExtLen, llRecPos)
                Loop
                DoEvents
            Loop
        End If
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gSocEcoPop                      *
'*                                                     *
'       <input> Form - form name
'               cbcList - combo box to set social economic entries
'       <return> False = OK, else true = error
'
'*             Created:10/28/03     D.Hosaka           *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Soc Eco Code          *
'*                      box if required                *
'*                                                     *
'*******************************************************
Function gSocEcoPop(Form As Form, cbcList As Control) As Integer
'
'   gSocEcoPopPop
'   Where:
'
    Dim ilRet As Integer
    Dim slSocEco As String      'Demo name, saved to determine if changed

    gSocEcoPop = False           'assume OK
    ilRet = gPopMnfPlusFieldsBox(Form, cbcList, tgSocEcoCode(), sgSocEcoCodeTag, "FNG")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo gSocEcoPopErr
        gCPErrorMsg ilRet, "gSocEcoPop (gPopMnfPlusFieldsBox)", Form
        On Error GoTo 0
        cbcList.AddItem "[None]", 0

        gFindMatch slSocEco, 1, cbcList
        If gLastFound(cbcList) > 0 Then
            cbcList.ListIndex = gLastFound(cbcList)
        Else
            cbcList.ListIndex = -1
        End If
    End If
    cbcList.ListIndex = 0
    Exit Function
gSocEcoPopErr:
    On Error GoTo 0
    gSocEcoPop = True           'error
End Function

Function gObtainPhfRvfbyCntr(RptForm As Form, llCntrNo As Long, slEarliestDate As String, slLatestDate As String, tlTranType As TRANTYPES, tlRvf() As RVF) As Integer
Debug.Print "gObtainPhfRvfbyCntr:" & llCntrNo & ", " & slEarliestDate & " - " & slLatestDate
'****************************************************************
'*
'*      Obtain all History and Receivables transactions whose
'*      transaction date falls within the earliest and latest
'*      dates requested.  Test for transaction types "I", "A"
'*      or "W".  Access PHF & RVF by key 4: contr # & Tran Date.
'*      Retrieve all or a single contract for a span of tran dates
'*
'*      <input>  RptForm - Form calling this populate rtn
'*               llCntrNo - single contract #, zero if all
'*                         -1 indicates test for cnt#0
'*               slEarliestDate - get all trans starting with
'*                   this date
'*               slLatestDAte - get all trans equal or prior to
'*                  this pacing date (effective date)
'*
'*      <I/O>    tlRvf() - array of matching Phf/Rvf recds
'*               funtion return - true if receivables populated
'*                       false if no receivables, error
'*
'*             Created:4/17/98       By:D. Hosaka
'*            Modified:              By:
'*
'*            Comments: make 2 passes, read Phf, the Rvf, build
'*               all transactions types "A", "I", "W" or "P"
'*               based on parameters in tlTranType.
'*          4-21-03 Add subroutine to retrieve by Key 4 (contr # & tran date), along
'*              with all or single contract
'*          5-26-04 Exclude NTR "AN" transactions when NTR to be excluded
'****************************************************************
'
'    ilRet = gObtainPhfRvfbyCntr (RptForm,  llCntrNo as long, slEarliestDate, slLatestDate, tlTranType, tlRvf())
'
    Dim ilRet As Integer    'Return status
    ReDim ilEarliestDate(0 To 1) As Integer
    ReDim ilLatestDate(0 To 1) As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilLoop As Integer
    Dim ilRVFUpper As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlCntrTypeBuff As POPLCODE
    Dim ilLowLimit As Integer
    Dim ilDoe As Integer
    
    'On Error GoTo gObtainPhfRvfbyCntrErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlRvf)
    'If ilRet <> 0 Then
    '    ilLowLimit = 1
    'End If
    'On Error GoTo 0
    If PeekArray(tlRvf).Ptr <> 0 Then
        ilLowLimit = LBound(tlRvf)
    Else
        ilLowLimit = 0
    End If

    ReDim tlRvf(ilLowLimit To ilLowLimit) As RVF
    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        gObtainPhfRvfbyCntr = False
        Exit Function
    End If
    imRvfRecLen = Len(tlRvf(ilLowLimit))
    gPackDate slEarliestDate, ilEarliestDate(0), ilEarliestDate(1)
    gPackDate slLatestDate, ilLatestDate(0), ilLatestDate(1)

    btrExtClear hmRvf   'Clear any previous extend operation
    ilExtLen = Len(tlRvf(ilLowLimit))  'Extract operation record size
    ilRVFUpper = UBound(tlRvf)
    For ilLoop = 1 To 2         'pass 1- get PHF, pass 2 get RVF

        ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
            Call btrExtSetBounds(hmRvf, llNoRec, -1, "UC", "RVF", "") '"EG") 'Set extract limits (all records)

            If llCntrNo > 0 Then                        'single contract
                tlCntrTypeBuff.lCode = llCntrNo
                ilOffSet = gFieldOffset("Rvf", "RvfCntrNo")
                ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCntrTypeBuff, 4)
                On Error GoTo mRvfErr
                gBtrvErrorMsg ilRet, "gObtainPhfRvfbyCntr (btrExtAddLogicConst):" & "Rvf.Btr", RptForm
                On Error GoTo 0
                '3-20-18 trouble sending contract # 0 because the routine that gets rvf/phf by contract # assumes 0 means get ALL, not any selective contract
                'Change the routine gObtainPhfRvfbyCntr to test for -1 and use 0 as a matching contract test, vs retrieve ALL
            ElseIf llCntrNo = -1 Then
                tlCntrTypeBuff.lCode = 0
                ilOffSet = gFieldOffset("Rvf", "RvfCntrNo")
                ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCntrTypeBuff, 4)
                On Error GoTo mRvfErr
                gBtrvErrorMsg ilRet, "gObtainPhfRvfbyCntr (btrExtAddLogicConst):" & "Rvf.Btr", RptForm
                On Error GoTo 0
            End If

            tlDateTypeBuff.iDate0 = ilEarliestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilEarliestDate(1)
            ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
            ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainPhfRvfbyCntr (btrExtAddLogicConst):" & "Rvf.Btr", RptForm
            On Error GoTo 0

            tlDateTypeBuff.iDate0 = ilLatestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilLatestDate(1)
            ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
            ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainPhfRvfbyCntr (btrExtAddLogicConst):" & "Rvf.Btr", RptForm
            On Error GoTo 0

            ilRet = btrExtAddField(hmRvf, 0, ilExtLen)  'Extract the whole record
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainRVF (btrExtAddField):" & "RVF.Btr", RptForm
            On Error GoTo 0
            ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
            DoEvents
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo mRvfErr
                gBtrvErrorMsg ilRet, "gObtainRVF (btrExtGetNextExt):" & "RVF.Btr", RptForm
                On Error GoTo 0
                ilExtLen = Len(tmRvf)  'Extract operation record size
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    'first test for valid trans types (Invoices, adjustments, write-off & payments
                    If ((Left$(tmRvf.sTranType, 1) = "I" Or tmRvf.sTranType = "HI") And tlTranType.iInv) Or (Left$(tmRvf.sTranType, 1) = "A" And tlTranType.iAdj) Or (Left$(tmRvf.sTranType, 1) = "W" And tlTranType.iWriteOff) Or (Left$(tmRvf.sTranType, 1) = "P" And tlTranType.iPymt) Then
                        ilDoe = ilDoe + 1
                        If ilDoe > 60 Then
                            ilDoe = 0
                            DoEvents
                        End If
                        If (tlTranType.iNTR) Then       'NTR option, tested separately because it shouldnt be tested with Cash transactions
                            If tmRvf.iMnfItem > 0 Then
                                tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                            Else            'its not an NTR
                                'got valid trans type - test for Cash, Trade, Merchandising or Promotions
                                If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                    tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                    ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                                End If
                            End If
                        Else
                            'got valid trans type - test for Cash, Trade, Merchandising or Promotions
                            '05-26-04 dont include NTR, exclude this if it is
                            If tmRvf.iMnfItem = 0 And tmRvf.lSbfCode = 0 Then       'not NTR
                                If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                    tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                    ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                                End If
                            Else
                                'must be an installment record, so it should be included
                                'If trans is an NTR, it must have the item type and pointer to SBF; otherwise its assumed to be an installment
                                'if it has an SBF pointer only
                                If tmRvf.iMnfItem = 0 And tmRvf.lSbfCode > 0 Then
                                    If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                        tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                        ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                                    End If
                                End If
                            End If
                        End If
                    End If
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
        If ilLoop = 1 Then                          'if 1, then just finished history, go do Receivables
            btrExtClear hmRvf   'Clear any previous extend operation
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmRvf)
                btrDestroy hmRvf
                gObtainPhfRvfbyCntr = False
                Exit Function
            End If
            imRvfRecLen = Len(tmRvf)
            ilRVFUpper = UBound(tlRvf)
        End If
    Next ilLoop
    ilRet = btrClose(hmRvf)
    btrDestroy hmRvf
    gObtainPhfRvfbyCntr = True
    Exit Function
gObtainPhfRvfbyCntrErr2:
    ilRet = 1
    Resume Next
mRvfErr:
    On Error GoTo 0
    gObtainPhfRvfbyCntr = False
    Exit Function
End Function

'****************************************************************
'*                                                              *
'*      Procedure Name:gMarketPop                               *
'*      <input>  List box                                       *
'*               ilTestCluster: true to test for cluster market *
'*                              false if show regardless of cluster
'*            Modified:              By:                        *
'*                                                              *
'*            Comments: Populate Market combobox                *
'*                                                              *
'*                                                              *
'****************************************************************
Sub gMarketPop(lbcSelection As Control, ilTestCluster As Integer, imTerminate As Integer)
'
'
    Dim ilRet As Integer
    Dim slName As String
    Dim slNameCode As String
    Dim ilSortCode As Integer
    Dim ilLoop As Integer
    Dim llLen As Long
    Dim slStr As String
    lbcSelection.Clear
    ilSortCode = 0
    ReDim tgMktCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    ilRet = gObtainMnfForType("H3", slStr, tgMktMnf())
    For ilLoop = LBound(tgMktMnf) To UBound(tgMktMnf) - 1 Step 1
        If (ilTestCluster And Trim$(tgMktMnf(ilLoop).sRPU) = "Y") Or (Not ilTestCluster) Then
            slName = Trim$(tgMktMnf(ilLoop).sName)
            slName = slName & "\" & Trim$(str$(tgMktMnf(ilLoop).iCode))
            tgMktCode(ilSortCode).sKey = slName
            If ilSortCode >= UBound(tgMktCode) Then
                ReDim Preserve tgMktCode(0 To UBound(tgMktCode) + 100) As SORTCODE
            End If
            ilSortCode = ilSortCode + 1
        End If
    Next ilLoop
    ReDim Preserve tgMktCode(0 To ilSortCode) As SORTCODE
    If UBound(tgMktCode) - 1 > 0 Then
        ArraySortTyp fnAV(tgMktCode(), 0), UBound(tgMktCode), 0, LenB(tgMktCode(0)), 0, LenB(tgMktCode(0).sKey), 0
    End If
    llLen = 0
    For ilLoop = 0 To UBound(tgMktCode) - 1 Step 1
        slNameCode = tgMktCode(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet = CP_MSG_NONE Then
            slName = Trim$(slName)
            If Not gOkAddStrToListBox(slName, llLen, True) Then
                Exit For
            End If
            lbcSelection.AddItem slName  'Add ID to list box
        End If
    Next ilLoop
    Exit Sub

    On Error GoTo 0
    imTerminate = True
End Sub

Sub gClearTxr()
'*******************************************************
'*                                                     *
'*      Procedure Name:Clear Prepass file for Text
'*                  Dump
'*                                                     *
'*             Created:05/22/01      By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*                                                     *
'*******************************************************
    Dim ilRet As Integer
    hmTxr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmTxr, "", sgDBPath & "Txr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmTxr)
        btrDestroy hmTxr
        Exit Sub
    End If
    imTxrRecLen = Len(tmTxr)
    tmTxrSrchKey.iGenDate(0) = igNowDate(0)
    tmTxrSrchKey.iGenDate(1) = igNowDate(1)
    'tmTxrSrchKey.iGenTime(0) = igNowTime(0)
    'tmTxrSrchKey.iGenTime(1) = igNowTime(1)
    tmTxrSrchKey.lGenTime = lgNowTime       '10-20-01
    ilRet = btrGetGreaterOrEqual(hmTxr, tmTxr, imTxrRecLen, tmTxrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    'Do While (ilRet = BTRV_ERR_NONE) And (tmTxr.iGenDate(0) = igNowDate(0)) And (tmTxr.iGenDate(1) = igNowDate(1)) And (tmTxr.iGenTime(0) = igNowTime(0)) And (tmTxr.iGenTime(1) = igNowTime(1))
    Do While (ilRet = BTRV_ERR_NONE) And (tmTxr.iGenDate(0) = igNowDate(0)) And (tmTxr.iGenDate(1) = igNowDate(1)) And (tmTxr.lGenTime = lgNowTime)
        ilRet = btrDelete(hmTxr)
        ilRet = btrGetNext(hmTxr, tmTxr, imTxrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmTxr)
    btrDestroy hmTxr
End Sub

Sub gCRGrfClear()
'*******************************************************
'*                                                     *
'*      Procedure Name:gCRBgt    Clear                 *
'*                                                     *
'*             Created:04/18/96      By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear Comparisons Budget data   *
'*                     for Crystal report              *
'*                                                     *
'*******************************************************
    Dim ilRet As Integer
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    tmGrfSrchKey.iGenDate(0) = igNowDate(0)
    tmGrfSrchKey.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrfSrchKey.lGenTime = lgNowTime
    ilRet = btrGetGreaterOrEqual(hmGrf, tmGrf, imGrfRecLen, tmGrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmGrf.iGenDate(0) = igNowDate(0)) And (tmGrf.iGenDate(1) = igNowDate(1)) And (tmGrf.lGenTime = lgNowTime)
        ilRet = btrDelete(hmGrf)
        tmGrfSrchKey.iGenDate(0) = igNowDate(0)
        tmGrfSrchKey.iGenDate(1) = igNowDate(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrfSrchKey.lGenTime = lgNowTime
        ilRet = btrGetGreaterOrEqual(hmGrf, tmGrf, imGrfRecLen, tmGrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        '8-20-13 chg way in which records are removed to avoid microkernel losing position with multi-users deleting from same table
        'ilRet = btrGetNext(hmGrf, tmGrf, imGrfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmGrf)
    btrDestroy hmGrf
End Sub

Sub gCrAnrClear()
'*******************************************************
'*                                                     *
'*      Procedure Name:gCrAnrClear    Clear            *
'*                                                     *
'*             Created:07/13/97      By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear Pre-pass Analysis file    *
'*                     for Crystal report              *
'*                                                     *
'*******************************************************
    Dim ilRet As Integer
    hmAnr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnr, "", sgDBPath & "Anr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAnr)
        btrDestroy hmAnr
        Exit Sub
    End If
    imAnrRecLen = Len(tmAnr)
    tmAnrSrchKey.iGenDate(0) = igNowDate(0)
    tmAnrSrchKey.iGenDate(1) = igNowDate(1)
    'tmAnrSrchKey.iGenTime(0) = igNowTime(0)
    'tmAnrSrchKey.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmAnrSrchKey.lGenTime = lgNowTime
    ilRet = btrGetGreaterOrEqual(hmAnr, tmAnr, imAnrRecLen, tmAnrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmAnr.iGenDate(0) = igNowDate(0)) And (tmAnr.lGenTime = lgNowTime)
        ilRet = btrDelete(hmAnr)
        ilRet = btrGetNext(hmAnr, tmAnr, imAnrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmAnr)
    btrDestroy hmAnr
End Sub

'       Populate Named Avails
'       Created 1-22-01      D Hosaka
'
Function gAvailsPop(frm As Form, lbcLocal As Control, tlSortCode() As SORTCODE) As Integer
'
'   ilRet = gAvailsPop (MainForm, lbcLocal )
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       lbcLocal (I)- List box to be populated from the master list box
'       ilRet (O)- Error code (0 if no error)
'

    Dim ilLoop As Integer
    Dim slName As String
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim ilSortCode As Integer
    Dim llLen As Long

    ilRet = gObtainAvail()
    If ilRet = False Then
        gAvailsPop = False
        Exit Function
    End If

    ReDim tlSortCode(0 To UBound(tgAvailAnf) - LBound(tgAvailAnf)) As SORTCODE
    For ilLoop = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1 Step 1
        slName = tgAvailAnf(ilLoop).sName & "\" & Trim$(str$(tgAvailAnf(ilLoop).iCode))
        tlSortCode(ilSortCode).sKey = slName
        ilSortCode = ilSortCode + 1
    Next ilLoop
    If UBound(tlSortCode) - 1 > 0 Then
        ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
    End If
    lbcLocal.Clear
    llLen = 0
    'For ilLoop = 1 To UBound(tgAvailAnf) - 1 Step 1
    For ilLoop = 0 To UBound(tgAvailAnf) - 1 Step 1
        'slNameCode = tlSortCode(ilLoop - 1).sKey
        slNameCode = tlSortCode(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gAvailsPop = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next ilLoop
    gAvailsPop = True
    Exit Function
End Function

Sub gCRAvrClear()
'*******************************************************
'*                                                     *
'*      Procedure Name:gCRAvrClear                     *
'*                                                     *
'*             Created:02/26/98      By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear Qtrly Avails Prepass      *
'*                     for Crystal report              *
'*                                                     *
'*******************************************************
    Dim ilRet As Integer
    Dim tlAvr As AVR
    Dim llNowTime As Long       '10-10-01 generation time
    hmAvr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAvr, "", sgDBPath & "Avr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAvr)
        btrDestroy hmAvr
        Exit Sub
    End If
    '10-10-01
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, llNowTime
    imAvrRecLen = Len(tlAvr)
    tmAvrSrchKey.iGenDate(0) = igNowDate(0)
    tmAvrSrchKey.iGenDate(1) = igNowDate(1)
    '10-10-01
    tmAvrSrchKey.lGenTime = llNowTime
    'tmAvrSrchKey.iGenTime(0) = igNowTime(0)
    'tmAvrSrchKey.iGenTime(1) = igNowTime(1)
    ilRet = btrGetGreaterOrEqual(hmAvr, tlAvr, imAvrRecLen, tmAvrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    '10-10-01
    Do While (ilRet = BTRV_ERR_NONE) And (tlAvr.iGenDate(0) = igNowDate(0)) And (tlAvr.iGenDate(1) = igNowDate(1)) And (tlAvr.lGenTime = llNowTime)
    'Do While (ilRet = BTRV_ERR_NONE) And (tlAvr.iGenDate(0) = igNowDate(0)) And (tlAvr.iGenDate(1) = igNowDate(1)) And (tlAvr.iGenTime(0) = igNowTime(0)) And (tlAvr.iGenTime(1) = igNowTime(1))
        ilRet = btrDelete(hmAvr)
        ilRet = btrGetNext(hmAvr, tlAvr, imAvrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmAvr)
    btrDestroy hmAvr
End Sub

Sub gCrCbfClear()
'*******************************************************
'*                                                     *
'*      Procedure Name:gCRCbf    Clear                 *
'*                                                     *
'*             Created:05/29/96      By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear Contract/Proposal  data   *
'*                     for Crystal report              *
'*                                                     *
'*******************************************************
    Dim ilRet As Integer
    Dim llNowTime As Long       '10-10-01 gen time
    hmCbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCbf)
        btrDestroy hmCbf
        Exit Sub
    End If
    '10-10-01
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, llNowTime
    imCbfRecLen = Len(tmCbf)
    tmCbfSrchKey.iGenDate(0) = igNowDate(0)
    tmCbfSrchKey.iGenDate(1) = igNowDate(1)
    '10-10-01
    tmCbfSrchKey.lGenTime = llNowTime
    'tmCbfSrchKey.iGenTime(0) = igNowTime(0)
    'tmCbfSrchKey.iGenTime(1) = igNowTime(1)
    ilRet = btrGetGreaterOrEqual(hmCbf, tmCbf, imCbfRecLen, tmCbfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    '10-10-01
    Do While (ilRet = BTRV_ERR_NONE) And (tmCbf.iGenDate(0) = igNowDate(0)) And (tmCbf.iGenDate(1) = igNowDate(1)) And (tmCbf.lGenTime = llNowTime)
    'Do While (ilRet = BTRV_ERR_NONE) And (tmCbf.iGenDate(0) = igNowDate(0)) And (tmCbf.iGenDate(1) = igNowDate(1)) And (tmCbf.iGenTime(0) = igNowTime(0)) And (tmCbf.iGenTime(1) = igNowTime(1))
        ilRet = btrDelete(hmCbf)
       ' ilRet = btrGetNext(hmCbf, tmCbf, imCbfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
       '8-26-13 change way in which deletions are performed to avoid losing positioning
        tmCbfSrchKey.iGenDate(0) = igNowDate(0)
        tmCbfSrchKey.iGenDate(1) = igNowDate(1)
        tmCbfSrchKey.lGenTime = llNowTime
        ilRet = btrGetGreaterOrEqual(hmCbf, tmCbf, imCbfRecLen, tmCbfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Loop
    ilRet = btrClose(hmCbf)
    btrDestroy hmCbf
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gCreateAvails                   *
'*                                                     *
'*             Created:10/24/97      By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Generate avails Data for       *
'*                      any report requiring Inventory,*
'*                      avails, sold.                  *
'*            Find availabilty by vehicle for a given  *
'*            rate card.                               *
'*                                                     *
'*   3/11/98 Add new 3rd parameter as string:
'            B = Test Base DP flag
'            R = Test Show on Report flag
'    7/1/98 Test valid RC reference in RIF
'*                                                     *
'*******************************************************
Sub gCRQAvails(hlChf As Integer, tlValuationInfo As VALUATIONINFO, slEffDate As String, slQtrEnd As String, tlCntTypes As CNTTYPES, lbcVehicle As Control, lbcRC As Control, tlTAvr() As AVR, ilCorpStd As Integer, tlSpotLenRatio As SPOTLENRATIO, ilAnfCodes() As Integer)
'
'   <input> hlChf - Contract file handle
'           ilNoQtrs - # quarters requested (1-4)  (6-30-00 force to 1 qtr only)
'           tlValuationInfo - has information with how to calc the inventory valuation
'           3-11-13 replaced slBAseReptFlag with tlValuationINfo:  slBaseReptFlag - B = test BaseDP Flag, R = test Show on Report Flag
'           slEffDate - Start Date of quarter of start of week to gather avails
'                       any week between the start of the quarter and slEffDate is zero
'           slEndDAte as string (end of qtr)
'           tlCntTypes - contract or spot types to include (cash/trade/remnants/missed/nc etc)
'           lbcVehicle - Vehicle list box
'           lbcRC- Rate Card list box
'           ilCorpStd - 1 = corp, 2 = std (corp assumes the Cof file in in tgCof array)
'           tlSpotLenRatio - table of spot lengths and their associated index for a 30"unit spot (ie. 30" spot = 1 unit, 15 = .50 unit)
'           ilAnfCodes() - array of anf codes selected for a sports vehicle only.  if not sports vehicle, follow normal DP rules
'  <output> tlTAvr - array of DP avails/inventory/sold by DP per vehicle
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilVehicle As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ll1WkDate As Long
    Dim ilFirstQ As Integer
    Dim ilRec As Integer
    Dim ilVpfIndex As Integer
    Dim ilUpper As Integer
    Dim ilDateOk As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilRcf As Integer
    Dim ilFound As Integer
    Dim llStart As Long
    Dim llEnd As Long
    Dim slDPFlag As String
    Dim ilSaveSort As Integer
    Dim ilNoWks As Integer          'weeks in qtr (may be 12-14)
    Dim llNewEndQtr As Long
    Dim slBAseReptFlag As String * 1
    Dim slMsgFile As String
    
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
    'Use AVR as a buffer, no writing to disk to AVR file
    ReDim tmAvr(0 To 0) As AVR

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
    
    If tlValuationInfo.iRCvsAvgPrice = 1 Then       'use avg price
        hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmClf)
            btrDestroy hmClf
            Exit Sub
        End If
        imClfRecLen = Len(tmClf)
    
        hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            btrDestroy hmCff
            btrDestroy hmClf
            Exit Sub
        End If
        imCffRecLen = Len(tmCff)
        
        hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmSmf)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            
            btrDestroy hmSmf
            btrDestroy hmCff
            btrDestroy hmClf
            Exit Sub
        End If
        imSmfRecLen = Len(tmSmf)

    End If
    
        ilRet = mOpenMsgFile(slMsgFile)
        slName = " by Avg 30 Rate Valuation"
        If tlValuationInfo.iRCvsAvgPrice = 0 Then       'use R/C
            slName = " by Rate Card Valuation"
        End If
        Print #hmMsg, "Sales vs Plan on "; Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM") & slName
        
        slName = ""
        For ilLoop = 0 To 9
            If tlSpotLenRatio.iLen(ilLoop) = 0 Then         'done
                Exit For
            Else
                slNameCode = gIntToStrDec(tlSpotLenRatio.iRatio(ilLoop), 2)
                If Trim$(slName) <> "" Then
                    slName = slName & ","
                End If
                slName = slName & str$(tlSpotLenRatio.iLen(ilLoop)) & " @" & Trim$(slNameCode)
            End If
        Next ilLoop
        Print #hmMsg, slName
    
        ReDim tlTAvr(0 To 0) As AVR
        ilRet = gObtainRcfRifRdf()          'get the rate cards and assoc dayparts
        
        slBAseReptFlag = tlValuationInfo.sBaseReptFlag
        
        ll1WkDate = gDateValue(slEffDate)   '6-30-00
        llNewEndQtr = gDateValue(slQtrEnd)  '6-30-00
        ilNoWks = (llNewEndQtr - ll1WkDate) / 7  '6-30-00
        tmVef.iCode = 0
        For ilLoop = 1 To ilNoWks Step 1    '6-30-00
            lmSAvailsDates(ilLoop) = ll1WkDate
            lmEAvailsDates(ilLoop) = ll1WkDate + 6
            ll1WkDate = ll1WkDate + 7
        Next ilLoop

        
        ilFirstQ = 1            '6-30-00 first valid week will always be 1 for quarterly mgmt reports
        For ilVehicle = 0 To lbcVehicle.ListCount - 1 Step 1
            If (lbcVehicle.Selected(ilVehicle)) Then
                slNameCode = tgCSVNameCode(ilVehicle).sKey 'RptSelSP!lbcCSVNameCode.List(ilVehicle)
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slName, 3, "|", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                ilVpfIndex = -1
                ilLoop = gBinarySearchVpf(ilVefCode)
                If ilLoop <> -1 Then
                    ilVpfIndex = ilLoop
                End If
                If ilVpfIndex >= 0 Then
                    For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
                        tmRcf = tgMRcf(ilRcf)
                        ilDateOk = False
                        For ilLoop = 0 To lbcRC.ListCount - 1 Step 1
                            slNameCode = tgRateCardCode(ilLoop).sKey
                            ilRet = gParseItem(slNameCode, 3, "\", slCode)
                            If Val(slCode) = tgMRcf(ilRcf).iCode Then
                                If (lbcRC.Selected(ilLoop)) Then
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
                                If tgMRif(llRif).iRcfCode = tgMRcf(ilRcf).iCode And tgMRif(llRif).iVefCode = ilVefCode Then
                                   ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
                                    If ilRdf <> -1 Then
                                        'Determine if using base DP or show on report field from Items record (RIF)
                                        If Trim$(slBAseReptFlag) = "B" Then             'test Base DP flag
                                            If tgMRif(llRif).sBase <> "Y" And tgMRif(llRif).sBase <> "N" Then
                                                slDPFlag = Trim$(tgMRdf(ilRdf).sBase)
                                            Else
                                                slDPFlag = Trim$(tgMRif(llRif).sBase)
                                            End If
                                        Else                                  'test show on report flag
                                            If tgMRif(llRif).sRpt <> "Y" And tgMRif(llRif).sRpt <> "N" Then
                                                slDPFlag = Trim$(tgMRdf(ilRdf).sReport)
                                            Else
                                                slDPFlag = Trim$(tgMRif(llRif).sRpt)
                                            End If
                                        End If
                                        If tgMRif(llRif).iSort = 0 Then
                                            ilSaveSort = tgMRdf(ilRdf).iSortCode
                                        Else
                                            ilSaveSort = tgMRif(llRif).iSort
                                        End If

                                        If tgMRdf(ilRdf).iCode = tgMRif(llRif).iRdfCode And slDPFlag = "Y" And tgMRdf(ilRdf).sState <> "D" And tgMRif(llRif).iVefCode = ilVefCode Then
                                        'If tgMRdf(ilRdf).iCode = tgMRif(llRif).iRdfCode And tgMRdf(ilRdf).sState <> "D" And tgMRif(llRif).ivefcode = ilVefCode Then
                                            ilFound = False
                                            For ilLoop = LBound(tmAvRdf) To ilUpper - 1 Step 1
                                                If tmAvRdf(ilLoop).iCode = tgMRdf(ilRdf).iCode Then
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
                                        End If
                                    End If
                                End If
                            Next llRif
                            llStart = lmSAvailsDates(ilFirstQ)
                            llEnd = lmEAvailsDates(ilNoWks)     '6-30-00
                            gGetAvailCounts hlChf, ilVefCode, ilVpfIndex, ilFirstQ, llStart, llEnd, lmSAvailsDates(), lmEAvailsDates(), tmAvRdf(), tmRifRate(), tmAvr(), tlCntTypes, ilCorpStd, tlValuationInfo, tlSpotLenRatio, ilAnfCodes()
                            'These avail records are not written to disk.  Send back all avails for all vehicles in tmAVR array
                            'for the calling subroutine to handle
                            'Output records
                            'For ilRec = 0 To UBound(tmAvr) - 1 Step 1
                            '    ilRet = btrInsert(hmAvr, tmAvr(ilRec), imAvrRecLen, INDEXKEY0)
                            'Next ilRec
                        End If
                    Next ilRcf
                End If
                
                For ilRec = 0 To UBound(tmAvr) - 1 Step 1
                    tlTAvr(UBound(tlTAvr)) = tmAvr(ilRec)
                    ReDim Preserve tlTAvr(0 To UBound(tlTAvr) + 1)
                Next ilRec
            End If
        Next ilVehicle
        Erase tmAvRdf, tmAvr, tmInvValAmtSold
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmAvr)
        ilRet = btrClose(hmLcf)
        
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmSsf
        btrDestroy hmLcf
        
        If tlValuationInfo.iRCvsAvgPrice = 1 Then       'use avg price
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCff)
            btrDestroy hmClf
            btrDestroy hmCff
        End If
        Print #hmMsg, "Sales vs Plan Completed "; Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")

        Close #hmMsg
    Exit Sub
End Sub

'           Find the Vehicle Group Index with the list box
'           from the user selected index
'
'           ilRet = gFindVehGroupInx(ilSelectedInx, tlVehicleGroup())
'           <input> ilselectedinx = index to selected item (vehicle group) in list box
'                   tlVehicleGroups - array of vehicle groups
'           <output> - altered
Function gFindVehGroupInx(ilSelectedInx As Integer, tlVehicleSets() As POPICODENAME) As Integer
    Dim slStr As String
    gFindVehGroupInx = 0
    slStr = Trim$(tlVehicleSets(ilSelectedInx).sChar)

        If slStr = "Participants" Then
            gFindVehGroupInx = 1
            Exit Function
        ElseIf slStr = "Sub-Totals" Then
            gFindVehGroupInx = 2
            Exit Function
        ElseIf slStr = "Market" Then
            gFindVehGroupInx = 3
            Exit Function
        ElseIf slStr = "Format" Then
            gFindVehGroupInx = 4
            Exit Function
        ElseIf slStr = "Research" Then
            gFindVehGroupInx = 5
            Exit Function
        ElseIf slStr = "Sub-Company" Then
            gFindVehGroupInx = 6
            Exit Function
        End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetAvailCounts                 *
'*                                                     *
'*             Created:10/20/97      By:D. Hosaka      *
'*             Copy of mGetAvails Counts made into     *
'*             generalized subroutine                  *
''*                                                    *
'*                                                     *
'*            Comments:Obtain the Avail counts         *
'*      4/27/98 Change to use DP sort code from RIF
'*       not RDF (Use RDF only if RIF not defined)
'*                                                     *
'*******************************************************
Sub gGetAvailCounts(hlChf As Integer, ilVefCode As Integer, ilVpfIndex As Integer, ilFirstQ As Integer, llSDate As Long, llEDate As Long, llSAvails() As Long, llEAvails() As Long, tlAvRdf() As RDF, tlRif() As RIF, tlAvr() As AVR, tlCntTypes As CNTTYPES, ilCorpStd As Integer, tlValuationInfo As VALUATIONINFO, tlSpotLenRatio As SPOTLENRATIO, ilAnfCodes() As Integer)
'
'   Where:
'

'   hlChf (I) - handle to Chf file
'   ilVefCode (I) - vehicle code to process
'   ilVpfIndex (I) - vehicle options pointer
'   ilFirstQ (I)
'   llSDate (I) - start date to begin searching Avails
'   llEDate (I) - end date to stop searching avails
'   llSAvails(I)- Array of bucket start dates
'   llEAvails(I)- Array of bucket end dates
'   tlAvRdf() (I) - array of Dayparts
'   tlAvr() (O) - array of AVR records built for avails
'   tlCntTypes (I) - contract and spot types to include in search
'   tlCntTypes.iHold(I)- True = include hold contracts
'   tlCntTypes.iOrder(I)- True= include complete order contracts
'   tlCntTypes.iMissed(I)- True=Include missed
'   tlCntTypes.iXtra(I)- True=Include Xtra bonus spots
'   tlCntTypes.iTrade(I)- True = include trade contracts
'   tlCntTypes.iNC(I)- True = include NC spots
'   tlCntTypes.iReserv(I) - True = include Reservations spots
'   tlCntTypes.iRemnant(I)- True=Include Remnant
'   tlCntTypes.iStandard(I)- true = include std contracts
'   tlCntTypes.iDR(I)- True=Include Direct Response
'   tlCntTypes.iPI(I)- True=Include per Inquiry
'   tlCntTypes.iPSA(I)- True=Include PSA
'   tlCntTypes.iPromo(I)- True=Include Promo
'   ilCorpSTd (1 = corp, 2 = std) 1 assumes corp calendar is in tgCof array
'   tlValuationInfo - flags on how to process valuation:  r/c vs avg rate, use base dp, % for adjustment factors
'   tlSpotLenRatio - array of spot lengths and association index for a 30" spot
'   ilAnfCodes(_) -array of avail name codes selected for sports vehicles only; non-sports vehicles will follow DP rules
'   Note: Remnants; Direct Response; per Inquiry; PSA and Promos are not
'         saved with a miss status
'         For scheduled spots the rank is used to determine if it is one
'         of the above (Direct reponse=1010; Remnant=1020; per Inquiry= 1030;
'         PSA=1060; Promo=1050.
'
'   slBucketType(I): A=Avail; S=Sold; I=Inventory  , P = Percent sellout    'forced to "A" for avail
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
    Dim slBucketType As String
    Dim ilLo As Integer
    Dim ilHi As Integer
    Dim ilWkNo As Integer               'week index to rate card
    Dim ilNoWks As Integer          '6-30-00
    ReDim ilSAvailsDates(0 To 1) As Integer
    ReDim ilEvtType(0 To 14) As Integer
    ReDim ilRdfCodes(0 To 1) As Integer
    ReDim tlAvr(0 To 0) As AVR
    ReDim tlavrcounts(0 To 0) As AVRCOUNTS
    ReDim tmInvValAmtSold(0 To 0) As INVVALAMTSOLD
    Dim ilVefIndex As Integer
    Dim llPrice As Long
    Dim slPrice As String
    Dim slAvgRate As String
    Dim ilRdfNameIndex As Integer
    Dim slChfType As String * 1     '4-12-18
    Dim ilPctTrade As Integer       '4-12-18

    slBucketType = "S"                          'force to Sellout, which calculations inventory, avails & sold
    slDate = Format$(llSAvails(1), "m/d/yy")
    gPackDate slDate, ilSAvailsDates(0), ilSAvailsDates(1)

    'ReDim ilWksInMonth(1 To 3) As Integer
    ReDim ilWksInMonth(0 To 3) As Integer   'Index zero ignored
    slStr = slDate
    If ilCorpStd = 1 Then               'std
        ilIndex = -1
        For ilLoop = LBound(tgMCof) To UBound(tgMCof)
            'loop thru each corporate calendar definition and determine what index the requested quarter is in
            gUnpackDateLong tgMCof(ilLoop).iStartDate(0, 0), tgMCof(ilLoop).iStartDate(1, 0), llDate
            gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iEndDate(1, 11), llLatestDate
            If llSDate >= llDate And llSDate <= llLatestDate Then
                ilIndex = ilLoop
                Exit For
            End If
        Next ilLoop
        If ilIndex < 0 Then
            MsgBox "Please define Corporate Calendar "
            Exit Sub
        End If
        'For ilLoop = 1 To 12            'find the matching quarter based on the start date given
        For ilLoop = 0 To 11            'find the matching quarter based on the start date given
            gUnpackDateLong tgMCof(ilIndex).iStartDate(0, ilLoop), tgMCof(ilIndex).iStartDate(1, ilLoop), llDate
            If llDate = llSDate Then
                'found matching start date of quarter, determine # weeks of each month
                For ilWkNo = 1 To 3
                    ilWksInMonth(ilWkNo) = tgMCof(ilIndex).iNoWks(ilWkNo - 1)
                Next ilWkNo
                Exit For
            End If
        Next ilLoop
    Else
        For ilLoop = 1 To 3 Step 1
            llDate = gDateValue(gObtainStartStd(slStr))
            llLoopDate = gDateValue(gObtainEndStd(slStr)) + 1
            ilWksInMonth(ilLoop) = ((llLoopDate - llDate) / 7)
            slStr = Format(llLoopDate, "m/d/yy")
        Next ilLoop
    End If

    'slType = "O"
    ilType = 0
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
    imSdfRecLen = Len(tmSdf)
    imCHFRecLen = Len(tmChf)
    ilNoWks = (llEDate - llSDate) / 7
    ilVefIndex = gBinarySearchVef(ilVefCode)
    If ilVefIndex = -1 Then
        MsgBox "Vehicle Missing, Internal Code: " & ilVefCode
        Exit Sub
    End If
    For llLoopDate = llSDate To llEDate Step 1
        slDate = Format$(llLoopDate, "m/d/yy")
        gPackDate slDate, ilDate0, ilDate1
        gObtainWkNo 0, slDate, ilWkNo, ilLo        'obtain the week bucket number
        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
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

        Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
            gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llDate
            ilBucketIndex = -1
            For ilLoop = 1 To ilNoWks Step 1    '6-30-00
                If (llDate >= llSAvails(ilLoop)) And (llDate <= llEAvails(ilLoop)) Then
                    ilBucketIndex = ilLoop
                    If ilBucketIndex = 14 Then      '6-30-00 dump 14th week info into 13th week
                        ilBucketIndex = 13
                    End If

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
                        For ilRdf = LBound(tlAvRdf) To UBound(tlAvRdf) - 1 Step 1
                            ilAvailOk = False
                            If (tlAvRdf(ilRdf).iLtfCode(0) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(1) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(2) <> 0) Then
                                If (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(0)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Then
                                    ilAvailOk = False    'True- code later
                                End If
                            Else

                                For ilLoop = LBound(tlAvRdf(ilRdf).iStartTime, 2) To UBound(tlAvRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                    If (tlAvRdf(ilRdf).iStartTime(0, ilLoop) <> 1) Or (tlAvRdf(ilRdf).iStartTime(1, ilLoop) <> 0) Then
                                        gUnpackTimeLong tlAvRdf(ilRdf).iStartTime(0, ilLoop), tlAvRdf(ilRdf).iStartTime(1, ilLoop), False, llStartTime
                                        gUnpackTimeLong tlAvRdf(ilRdf).iEndTime(0, ilLoop), tlAvRdf(ilRdf).iEndTime(1, ilLoop), True, llEndTime
                                        'If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDay + 1) = "Y") Then
                                        If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDay) = "Y") Then
                                            ilAvailOk = True
                                            ilLoopIndex = ilLoop
                                            slDays = ""
                                            For ilDayIndex = 1 To 7 Step 1
                                                If (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "Y") Or (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "N") Then
                                                    slDays = slDays & tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1)
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
                                If tgMVef(ilVefIndex).sType = "G" Then          'sports vehicles uses the user input selection for avail names
                                    ilAvailOk = False
                                    For ilLoop = LBound(ilAnfCodes) To UBound(ilAnfCodes) - 1
                                        If tmAvail.ianfCode = ilAnfCodes(ilLoop) Then
                                            ilAvailOk = True
                                            Exit For
                                        End If
                                    Next ilLoop
                                Else                                            'non sports vehicles follow DP rules for avail names
                                    If tlAvRdf(ilRdf).sInOut = "I" Then   'Book into
                                        If tmAvail.ianfCode <> tlAvRdf(ilRdf).ianfCode Then
                                            ilAvailOk = False
                                        End If
                                    ElseIf tlAvRdf(ilRdf).sInOut = "O" Then   'Exclude
                                        If tmAvail.ianfCode = tlAvRdf(ilRdf).ianfCode Then
                                            ilAvailOk = False
                                        End If
                                    End If
                                End If
                            End If
                            If ilAvailOk Then
                                'Determine if Avr created
                                ilFound = False
                                ilSaveDay = ilDay
                                'If RptSelCt!rbcSelCInclude(0).Value Then       'daypart option, place all values in same record
                                                                                'to get better availability
                                ilDay = 0                                       'force all data in same day of week
                                'End If
                                For ilRec = 0 To UBound(tlAvr) - 1 Step 1
                                    'If (tlAvr(ilRec).iRdfCode = tlAvRdf(ilRdf).iCode) And (tlAvr(ilRec).iFirstBucket = ilFirstQ) And (tlAvr(ilRec).iDay = ilDay) Then
                                    If (ilRdfCodes(ilRec) = tlAvRdf(ilRdf).iCode) And (tlAvr(ilRec).iFirstBucket = ilFirstQ) And (tlAvr(ilRec).iDay = ilDay) Then
                                        ilFound = True
                                        ilRecIndex = ilRec
                                        Exit For
                                    End If
                                Next ilRec
                                If Not ilFound Then
                                    ilRecIndex = UBound(tlAvr)
                                    tlAvr(ilRecIndex).iGenDate(0) = igNowDate(0)
                                    tlAvr(ilRecIndex).iGenDate(1) = igNowDate(1)
                                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                    tlAvr(ilRecIndex).lGenTime = lgNowTime
                                    tlAvr(ilRecIndex).iVefCode = ilVefCode
                                    tlAvr(ilRecIndex).iDay = ilDay
                                    tlAvr(ilRecIndex).iQStartDate(0) = ilSAvailsDates(0)
                                    tlAvr(ilRecIndex).iQStartDate(1) = ilSAvailsDates(1)
                                    tlAvr(ilRecIndex).iFirstBucket = ilFirstQ
                                    tlAvr(ilRecIndex).sBucketType = slBucketType
                                    'tlAvr(ilRecIndex).iRdfCode = tlAvRdf(ilRdf).iCode
                                    ilRdfCodes(ilRecIndex) = tlAvRdf(ilRdf).iCode
                                    'tlAvr(ilRecIndex).iRdfCode = tlAvRdf(ilRdf).iSortCode
                                    tlAvr(ilRecIndex).iRdfCode = tlRif(ilRdf).iSort
                                    tlAvr(ilRecIndex).sInOut = tlAvRdf(ilRdf).sInOut
                                    tlAvr(ilRecIndex).ianfCode = tlAvRdf(ilRdf).ianfCode
                                    tlAvr(ilRecIndex).iDPStartTime(0) = tlAvRdf(ilRdf).iStartTime(0, ilLoopIndex)
                                    tlAvr(ilRecIndex).iDPStartTime(1) = tlAvRdf(ilRdf).iStartTime(1, ilLoopIndex)
                                    tlAvr(ilRecIndex).iDPEndTime(0) = tlAvRdf(ilRdf).iEndTime(0, ilLoopIndex)
                                    tlAvr(ilRecIndex).iDPEndTime(1) = tlAvRdf(ilRdf).iEndTime(1, ilLoopIndex)
                                    tlAvr(ilRecIndex).sDPDays = slDays
                                    tlAvr(ilRecIndex).sNot30Or60 = "N"
                                    ReDim Preserve tlAvr(0 To ilRecIndex + 1) As AVR
                                    ReDim Preserve tlavrcounts(0 To ilRecIndex + 1) As AVRCOUNTS
                                    ReDim Preserve tmInvValAmtSold(0 To ilRecIndex + 1) As INVVALAMTSOLD
                                    ReDim Preserve ilRdfCodes(0 To ilRecIndex + 1)
                                End If
                                tlAvr(ilRecIndex).lRate(ilBucketIndexMinusOne) = tlRif(ilRdf).lRate(ilWkNo)
                                ilDay = ilSaveDay
                                'Always gather inventory
                                ilLen = tmAvail.iLen
                                ilUnits = tmAvail.iAvInfo And &H1F
                                ilNo30 = 0
                                ilNo60 = 0
                                If tgVpf(ilVpfIndex).sSSellOut = "B" Then
'                                    'Convert inventory to number of 30's and 60's
'                                    Do While ilLen >= 60
'                                        ilNo60 = ilNo60 + 1
'                                        ilLen = ilLen - 60
'                                    Loop

                                    ilRet = gDetermineSpotLenRatio(30, tlSpotLenRatio)     'determine 30" avail inventory by the user defined table
                                    If ilRet < 0 Then
                                        ilRet = -ilRet          'get positive number back and use it
                                        Print #hmMsg, tgMVef(ilVefIndex).sName & " spot length 30 not found. 1 unit calculated"
                                    End If
                                    'get # of 30s and and their equivalent index for Inventory counts
                                    Do While ilLen >= 30
                                        ilNo30 = ilNo30 + 1
                                        ilLen = ilLen - 30
                                    Loop
                                    ilNo30 = ilNo30 * ilRet            '# of 30s calculated times the index factor (in hundreds for decimals); i.e. 3 30" * 1.00 unit each = 300

                                    If ilLen < 30 And ilLen > 0 Then    '7-6-00 assume anything under 30" is 1-30" unit availability
                                        ilRet = gDetermineSpotLenRatio(ilLen, tlSpotLenRatio)     'determine 30" avail inventory by the user defined table
                                        If ilRet < 0 Then
                                            ilRet = -ilRet
                                            Print #hmMsg, tgMVef(ilVefIndex).sName & " spot length " & str$(ilLen) & " not found. 1 unit calculated"
                                        End If
                                        'ilNo30 = ilNo30 + 1
                                        ilNo30 = ilNo30 + ilRet
                                        ilLen = 0
                                    End If

                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                    'Count 30 or 60 and set flag if neither
                                    If ilLen = 60 Then
                                        ilNo60 = 1
                                    ElseIf ilLen = 30 Then
                                        ilNo30 = 1
                                    Else
                                        tlAvr(ilRecIndex).sNot30Or60 = "Y"
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
                                        tlAvr(ilRecIndex).sNot30Or60 = "Y"
                                    End If
                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                End If
                                If slBucketType <> "S" And slBucketType <> "P" Then    'sellout by min or pcts, don't update these yet
                                    tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) + ilNo30
                                    tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) + ilNo60
                                End If
                                'always put total inventory into record and avail bucket (avail bucket for qtrly detail)
                                tlavrcounts(ilRecIndex).l30InvCount(ilBucketIndex) = tlavrcounts(ilRecIndex).l30InvCount(ilBucketIndex) + ilNo30
                                tlavrcounts(ilRecIndex).l60InvCount(ilBucketIndex) = tlavrcounts(ilRecIndex).l60InvCount(ilBucketIndex) + ilNo60
                                tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) + ilNo30
                                tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) + ilNo60
                                'Always calculate Avails
                                For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                   LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                                    ilSpotOK = True                             'assume spot is OK to include

                                    If ((tmSpot.iRank And RANKMASK) = REMNANTRANK) And (Not tlCntTypes.iRemnant) Then
                                        ilSpotOK = False
                                    End If
                                    If ((tmSpot.iRank And RANKMASK) = PERINQUIRYRANK) And (Not tlCntTypes.iPI) Then
                                        ilSpotOK = False
                                    End If
                                    '4-12-18 previously DR not tested
                                    If ((tmSpot.iRank And RANKMASK) = DIRECTRESPONSERANK) And (Not tlCntTypes.iDR) Then
                                        ilSpotOK = False
                                    End If

'                                    'Added 4/1/18
'                                    If ((tmSpot.iRank And RANKMASK) = 1010) Then      'DR
'                                        If (Asc(tgSaf(0).sFeatures4) And AVAILINCLDEDIRECTRESPONSES) <> AVAILINCLDEDIRECTRESPONSES Then
'                                            ilSpotOK = False
'                                        End If
'                                    End If
'                                    'End add
                                    
                                    If ((tmSpot.iRank And RANKMASK) = TRADERANK) And (Not tlCntTypes.iTrade) Then
                                        ilSpotOK = False
                                    End If
                                    If ((tmSpot.iRank And RANKMASK) = EXTRARANK) And (Not tlCntTypes.iXtra) Then
                                        ilSpotOK = False
                                    End If
                                    If ((tmSpot.iRank And RANKMASK) = PROMORANK) And (Not tlCntTypes.iPromo) Then
                                        ilSpotOK = False
                                    End If
                                    If ((tmSpot.iRank And RANKMASK) = PSARANK) And (Not tlCntTypes.iPSA) Then
                                        ilSpotOK = False
                                    End If
                                    '4-12-18 reservation previously not tested
                                    If ((tmSpot.iRank And RANKMASK) = RESERVATIONRANK) And (Not tlCntTypes.iReserv) Then
                                        ilSpotOK = False
                                    End If

'                                    'Added 4/1/18
'                                    If (tmSpot.iRank And RANKMASK) = RESERVATION Then     'Reservation
'                                        If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDERESERVATION) <> AVAILINCLUDERESERVATION Then
'                                            ilSpotOK = False
'                                        End If
'                                    End If
'                                    'End add
                                    
                                    If (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                        ilSpotOK = False
                                    End If
                                    ilLen = tmSpot.iPosLen And &HFFF
                                    If ilSpotOK Then                            'continue testing other filters
                                        tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If tmSpot.lSdfCode = tmSdf.lCode And ilRet = BTRV_ERR_NONE Then
                                            If tmSdf.lChfCode <> tmChf.lCode Then               'if already in mem, don't reread
                                                tmChfSrchKey.lCode = tmSdf.lChfCode
                                                ilRet = btrGetEqual(hlChf, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                            Else
                                                ilRet = BTRV_ERR_NONE
                                            End If
                                        End If
                                        If ilRet <> BTRV_ERR_NONE Then
                                            ilSpotOK = False
                                        Else
                                            ilLen = tmSdf.iLen
                                            If tmChf.sStatus = "H" Then
                                                If Not tlCntTypes.iHold Then
                                                    ilSpotOK = False
                                                End If
                                            ElseIf tmChf.sStatus = "O" Then
                                                If Not tlCntTypes.iOrder Then
                                                    ilSpotOK = False
                                                End If
                                            Else
                                                ilSpotOK = False
                                            End If
'                                            '3-16-10 wrong code was tested for standard (tested S not C)
'                                            If tmChf.sType = "C" And Not tlCntTypes.iStandard Then      'include Standard types?
'                                                ilSpotOK = False
'                                            ElseIf tmChf.sType = "V" And Not tlCntTypes.iReserv Then      'include reservations ?
'                                                ilSpotOK = False
'
'                                            ElseIf tmChf.sType = "R" And Not tlCntTypes.iDR Then      'include DR?
'                                                ilSpotOK = False
'                                            End If
                                            '4-16-18
                                            If (tmChf.sType = "C") And (Not tlCntTypes.iStandard) Then
                                                ilSpotOK = False
                                            End If
                                            If (tmChf.sType = "V") And (Not tlCntTypes.iReserv) Then
                                                ilSpotOK = False
                                            End If
                                            If (tmChf.sType = "R") And (Not tlCntTypes.iDR) Then
                                                ilSpotOK = False
                                            End If
                                            If (tmChf.sType = "T") And (Not tlCntTypes.iRemnant) Then
                                                ilSpotOK = False
                                            End If
                                            If (tmChf.sType = "Q") And (Not tlCntTypes.iPI) Then
                                                ilSpotOK = False
                                            End If
                                                          
                                            If (tmChf.sType = "M") And (Not tlCntTypes.iPromo) Then
                                                ilSpotOK = False
                                            End If
                                            If (tmChf.sType = "S") And (Not tlCntTypes.iPSA) Then
                                                ilSpotOK = False
                                            End If
                                            If (tmChf.iPctTrade = 100) And (Not tlCntTypes.iTrade) Then
                                                ilSpotOK = False
                                            End If

                                        End If
                                        If ilSpotOK Then
                                            '3-11-13 valid spot, if processing for avg 30" rate vs R/C, get the spot rate
                                            If tlValuationInfo.iRCvsAvgPrice = 1 Then                   'use avg 30" rate vs rate card
                                                'setup schedule line info so clf doesnt have to be read, and the flight routine can get flight rate info
                                                'tmClf.iLine = tmSdf.iLineNo 'TTP 10743
                                                tmClf.iLine = 0 '5/26/23 - Per Jason suppress the line number on the RAB export for the air time and NTR records
                                                tmClf.iPropVer = tmChf.iPropVer
                                                tmClf.iCntRevNo = tmChf.iCntRevNo
                                                ilRet = gGetFlightPrice(tmSdf, tmClf, hmCff, hmSmf, slPrice)
                                                llPrice = 0     'init incase a decimal number isnt in price field (its adu, nc, fill,etc.)
                                                If (InStr(slPrice, ".") <> 0) Then        'found spot cost
                                                    llPrice = gStrDecToLong(slPrice, 2)
                                                End If
                                                tmInvValAmtSold(ilRecIndex).lRate(ilBucketIndex) = tmInvValAmtSold(ilRecIndex).lRate(ilBucketIndex) + llPrice
                                            End If
                                            
                                            ilNo30 = 0
                                            ilNo60 = 0
                                            If tgVpf(ilVpfIndex).sSSellOut = "B" Then                   'both units and seconds
                                            
                                                '3-14-13 return the ratio of the spot length (i.e. 30 = 100, 15 = 50, etc ; based on the spot length table and associated ratios
                                                ilNo30 = gDetermineSpotLenRatio(ilLen, tlSpotLenRatio)
                                                If ilNo30 < 0 Then
                                                    ilNo30 = -ilNo30
                                                    Print #hmMsg, tgMVef(ilVefIndex).sName & " spot length " & str$(ilLen) & " not found. " & str$(ilNo30 / 100) & " units calculated"
                                                End If
'
'                                               'Convert inventory to number of 30's and 60's
'                                                Do While ilLen >= 60
'                                                    ilNo60 = ilNo60 + 1
'                                                    ilLen = ilLen - 60
'                                                Loop
'                                                Do While ilLen >= 30
'                                                    ilNo30 = ilNo30 + 1
'                                                    ilLen = ilLen - 30
'                                                Loop
'                                                If ilLen < 30 And ilLen > 0 Then    '7-6-00 assume anything under 30" is 1-30" unit availability
'                                                    ilNo30 = ilNo30 + 1
'                                                    ilLen = 0
'                                                End If

                                                If (slBucketType = "S") Or (slBucketType = "P") Then    'sellout or %sellout, accum sold
                                                    'If RptSelCt!rbcSelC4(1).Value Then                'qtrly detail
                                                    If tmChf.sType = "V" Then                       'Type reserve
                                                        'Show on separate line or buy in sold?
                                                        'If RptSelCt!rbcSelC7(0).Value Then                       'hide the reserves
                                                        '    tlAvr(ilRecIndex).i30Count(ilBucketIndex) = tlAvr(ilRecIndex).i30Count(ilBucketIndex) + ilNo30
                                                        '    tlAvr(ilRecIndex).i60Count(ilBucketIndex) = tlAvr(ilRecIndex).i60Count(ilBucketIndex) + ilNo60
                                                        'Else                                            'show the reserves- (If Excluding reserves, it has already been
                                                                                                        'excluded by the time it gets here)
                                                            tlavrcounts(ilRecIndex).l30Reserve(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Reserve(ilBucketIndex) + ilNo30
                                                            tlavrcounts(ilRecIndex).l60Reserve(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Reserve(ilBucketIndex) + ilNo60
                                                        'End If
                                                    ElseIf tmChf.sStatus = "H" Then                         'staus "Hold" , always show on separate line
                                                        tlavrcounts(ilRecIndex).l30Hold(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Hold(ilBucketIndex) + ilNo30
                                                        tlavrcounts(ilRecIndex).l60Hold(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Hold(ilBucketIndex) + ilNo60
                                                    Else
                                                        tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) + ilNo30
                                                        tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) + ilNo60
                                                    End If
                                                    'Else
                                                    '    tlAvr(ilRecIndex).i30Count(ilBucketIndex) = tlAvr(ilRecIndex).i30Count(ilBucketIndex) + ilNo30
                                                    '    tlAvr(ilRecIndex).i60Count(ilBucketIndex) = tlAvr(ilRecIndex).i60Count(ilBucketIndex) + ilNo60
                                                    'End If
                                                'Else                                                    'avails: accum available
                                                '    tlAvr(ilRecIndex).i60Count(ilBucketIndex) = tlAvr(ilRecIndex).i60Count(ilBucketIndex) - ilNo60
                                                '    tlAvr(ilRecIndex).i30Count(ilBucketIndex) = tlAvr(ilRecIndex).i30Count(ilBucketIndex) - ilNo30
                                                End If
                                                'adjust the available buckets (used for qtrly detail  report only)
                                                tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) - ilNo60
                                                tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) - ilNo30
                                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then               'units sold
                                                'Count 30 or 60 and set flag if neither
                                                If ilLen = 60 Then
                                                    ilNo60 = 1
                                                ElseIf ilLen = 30 Then
                                                    ilNo30 = 1
                                                Else
                                                    tlAvr(ilRecIndex).sNot30Or60 = "Y"
                                                    If ilLen <= 30 Then
                                                        ilNo30 = 1
                                                    Else
                                                        ilNo60 = 1
                                                    End If
                                                End If
                                                If (ilNo60 <> 0) Or (ilNo30 <> 0) Then
                                                    If (slBucketType = "S") Or (slBucketType = "P") Then
                                                        'If RptSelCt!rbcSelC4(1).Value Then                'qtrly detail
                                                        If tmChf.sType = "V" Then                       'Type reserve
                                                            'Show on separate line or buy in sold?
                                                            'If RptSelCt!rbcSelC7(0).Value Then                       'hide the reserves
                                                            '    tlAvr(ilRecIndex).i30Count(ilBucketIndex) = tlAvr(ilRecIndex).i30Count(ilBucketIndex) + ilNo30
                                                            '    tlAvr(ilRecIndex).i60Count(ilBucketIndex) = tlAvr(ilRecIndex).i60Count(ilBucketIndex) + ilNo60
                                                            'Else                                            'show the reserves- (if excluding reserves, it has
                                                                                                            'already been excluded by the time it gets here)
                                                                tlavrcounts(ilRecIndex).l30Reserve(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Reserve(ilBucketIndex) + ilNo30
                                                                tlavrcounts(ilRecIndex).l60Reserve(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Reserve(ilBucketIndex) + ilNo60
                                                            'End If
                                                        ElseIf tmChf.sStatus = "H" Then                         'staus "Hold", always show on separate line
                                                            tlavrcounts(ilRecIndex).l30Hold(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Hold(ilBucketIndex) + ilNo30
                                                            tlavrcounts(ilRecIndex).l60Hold(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Hold(ilBucketIndex) + ilNo60
                                                        Else
                                                            tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) + ilNo30
                                                            tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) + ilNo60
                                                        End If
                                                        'Else
                                                        '    tlAvr(ilRecIndex).i30Count(ilBucketIndex) = tlAvr(ilRecIndex).i30Count(ilBucketIndex) + ilNo30
                                                        '    tlAvr(ilRecIndex).i60Count(ilBucketIndex) = tlAvr(ilRecIndex).i60Count(ilBucketIndex) + ilNo60
                                                        'End If
                                                    'Else                                                   'staus hold or reserve n/a for other qtrly summary options
                                                        'Gather Inventory
                                                        'If ilNo60 > 0 Then                     'spot found a 60?
                                                        '    tlAvr(ilRecIndex).i60Count(ilBucketIndex) = tlAvr(ilRecIndex).i60Count(ilBucketIndex) - ilNo60
                                                        'Else
                                                        '    If tlAvr(ilRecIndex).i30Count(ilBucketIndex) > 0 Then
                                                        '        tlAvr(ilRecIndex).i30Count(ilBucketIndex) = tlAvr(ilRecIndex).i30Count(ilBucketIndex) - ilNo30
                                                        '    Else
                                                        '        If tlAvr(ilRecIndex).i60Count(ilBucketIndex) > 0 Then
                                                        '            tlAvr(ilRecIndex).i60Count(ilBucketIndex) = tlAvr(ilRecIndex).i60Count(ilBucketIndex) - ilNo30
                                                        '        Else                        'oversold units
                                                        '            tlAvr(ilRecIndex).i30Count(ilBucketIndex) = tlAvr(ilRecIndex).i30Count(ilBucketIndex) - ilNo30
                                                        '        End If
                                                        '    End If
                                                        'End If
                                                    End If
                                                End If
                                                'adjust the available buckets (used for qtrly detail report only)
                                                tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) - ilNo60
                                                tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) - ilNo30
                                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then               'matching units
                                                'Count 30 or 60 and set flag if neither
                                                If ilLen = 60 Then
                                                    ilNo60 = 1
                                                ElseIf ilLen = 30 Then
                                                    ilNo30 = 1
                                                Else
                                                    tlAvr(ilRecIndex).sNot30Or60 = "Y"
                                                End If
                                                If (slBucketType = "S") Or (slBucketType = "P") Then        'if Sellout or % sellout, accum the seconds sold
                                                    'Qtrly detail has been forced to "Sellout" for internal testing
                                                    'If RptSelCt!rbcSelC4(1).Value Then            'qtrly detail
                                                    If tmChf.sType = "V" Then                       'Type reserve
                                                        'Show on separate line or bury in sold?
                                                        'If RptSelCt!rbcSelC7(0).Value Then                       'hide the reserves
                                                        '    tlAvr(ilRecIndex).i30Count(ilBucketIndex) = tlAvr(ilRecIndex).i30Count(ilBucketIndex) + ilNo30
                                                        '    tlAvr(ilRecIndex).i60Count(ilBucketIndex) = tlAvr(ilRecIndex).i60Count(ilBucketIndex) + ilNo60
                                                        'Else                                            'show the reserves - (if excluding reserves, it has
                                                                                                        'already been excluded by the time it gets here)
                                                            tlavrcounts(ilRecIndex).l30Reserve(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Reserve(ilBucketIndex) + ilNo30
                                                            tlavrcounts(ilRecIndex).l60Reserve(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Reserve(ilBucketIndex) + ilNo60
                                                        'End If
                                                    ElseIf tmChf.sStatus = "H" Then                         'staus "Hold", always show on separate line
                                                        tlavrcounts(ilRecIndex).l30Hold(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Hold(ilBucketIndex) + ilNo30
                                                        tlavrcounts(ilRecIndex).l60Hold(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Hold(ilBucketIndex) + ilNo60
                                                    Else            'not held or reserved, put in sold
                                                        tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) + ilNo30
                                                        tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) + ilNo60
                                                    End If
                                                    'Else
                                                    '    tlAvr(ilRecIndex).i30Count(ilBucketIndex) = tlAvr(ilRecIndex).i30Count(ilBucketIndex) + ilNo30
                                                    '    tlAvr(ilRecIndex).i60Count(ilBucketIndex) = tlAvr(ilRecIndex).i60Count(ilBucketIndex) + ilNo60
                                                    'End If
                                                'Else                                                    'holds & reserve n/a for othr qtrly summary options
                                                '    tlAvr(ilRecIndex).i60Count(ilBucketIndex) = tlAvr(ilRecIndex).i60Count(ilBucketIndex) - ilNo60
                                                '    tlAvr(ilRecIndex).i30Count(ilBucketIndex) = tlAvr(ilRecIndex).i30Count(ilBucketIndex) - ilNo30
                                                End If
                                                'adjust the available bucket (used for qrtrly detail report only)
                                                tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) - ilNo60
                                                tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) - ilNo30
                                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                            End If
                                        End If                              'ilspotOK
                                    End If
                                Next ilSpot                             'loop from ssf file for # spots in avail
                            End If                                          'Avail OK
                        Next ilRdf                                          'ilRdf = lBound(tlAvRdf)
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
    'If (tlCntTypes.iMissed) And (slBucketType <> "I") Then
    If (tlCntTypes.iMissed) Then
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
                
                '4-12-18 implement testing of contract types
                'gGetContractParameters tmSdf.lChfCode, slChfType, ilPctTrade
                If tmSdf.lChfCode <> tmChf.lCode Then               'if already in mem, don't reread
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hlChf, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                    slChfType = tmChf.sType
                    ilPctTrade = tmChf.iPctTrade
                Else
                    ilRet = BTRV_ERR_NONE
                End If
                If ilRet <> BTRV_ERR_NONE Then
                    ilSpotOK = False
                Else
                    If (slChfType = "C") And (Not tlCntTypes.iStandard) Then
                        ilSpotOK = False
                    End If
                    If (slChfType = "V") And (Not tlCntTypes.iReserv) Then
                        ilSpotOK = False
                    End If
                    If (slChfType = "R") And (Not tlCntTypes.iDR) Then
                        ilSpotOK = False
                    End If
                    If (slChfType = "T") And (Not tlCntTypes.iRemnant) Then
                        ilSpotOK = False
                    End If
                    If (slChfType = "Q") And (Not tlCntTypes.iPI) Then
                        ilSpotOK = False
                    End If
                                  
                    If (slChfType = "M") And (Not tlCntTypes.iPromo) Then
                        ilSpotOK = False
                    End If
                    If (slChfType = "S") And (Not tlCntTypes.iPSA) Then
                        ilSpotOK = False
                    End If
                    If (ilPctTrade = 100) And (Not tlCntTypes.iTrade) Then
                        ilSpotOK = False
                    End If
                End If
                If ilSpotOK Then
                'End add
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                    If (llDate >= llSAvails(ilFirstQ)) And (llDate <= llEAvails(ilNoWks)) Then      '6-30-00
                        ilBucketIndex = -1
                        For ilLoop = 1 To ilNoWks Step 1    '6-30-00
                            If (llDate >= llSAvails(ilLoop)) And (llDate <= llEAvails(ilLoop)) Then
                                ilBucketIndex = ilLoop
                                If ilBucketIndex = 14 Then  '6-30-00 dump 14th week info into 13th week
                                    ilBucketIndex = 13
                                End If
                                Exit For
                            End If
                        Next ilLoop
                        If ilBucketIndex > 0 Then
                            ilBucketIndexMinusOne = ilBucketIndex - 1
                            ilDay = gWeekDayLong(llDate)
                            slDate = Format$(llDate, "m/d/yy")
                            gPackDate slDate, ilDate0, ilDate1
                            gObtainWkNo 0, slDate, ilWkNo, ilLo        'obtain the week bucket number
                            gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                            For ilRdf = LBound(tlAvRdf) To UBound(tlAvRdf) - 1 Step 1
                                ilAvailOk = False
                                If (tlAvRdf(ilRdf).iLtfCode(0) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(1) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(2) <> 0) Then
                                    If (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(0)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Then
                                        ilAvailOk = False    'True- code later
                                    End If
                                Else
                                    For ilLoop = LBound(tlAvRdf(ilRdf).iStartTime, 2) To UBound(tlAvRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                        If (tlAvRdf(ilRdf).iStartTime(0, ilLoop) <> 1) Or (tlAvRdf(ilRdf).iStartTime(1, ilLoop) <> 0) Then
                                            gUnpackTimeLong tlAvRdf(ilRdf).iStartTime(0, ilLoop), tlAvRdf(ilRdf).iStartTime(1, ilLoop), False, llStartTime
                                            gUnpackTimeLong tlAvRdf(ilRdf).iEndTime(0, ilLoop), tlAvRdf(ilRdf).iEndTime(1, ilLoop), True, llEndTime
                                            If UBound(tlAvRdf) - 1 = LBound(tlAvRdf) Then   'could be a conv bumped spot sched in
                                                                                        'in conven veh.  The VV has DP times different than the
                                                                                        'conven veh.
                                                llStartTime = llTime
                                                llEndTime = llTime + 1              'actual time of spot
                                            End If
                                            'Don't include the end time i.e. 10a-3p is 10a thru 2:59:59p
                                            'If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDay + 1) = "Y") Then
                                            If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDay) = "Y") Then
                                                ilAvailOk = True
                                                ilLoopIndex = ilLoop
                                                slDays = ""
                                                For ilDayIndex = 1 To 7 Step 1
                                                    If (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "Y") Or (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "N") Then
                                                        slDays = slDays & tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1)
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
                                    'If RptSelCt!rbcSelCInclude(0).Value Then              'daypart option, place all values in same record
                                                                                        'to get better availability
                                        ilDay = 0                                       'force all data in same day of week
                                    'End If
                                    For ilRec = 0 To UBound(tlAvr) - 1 Step 1
                                        'If (tlAvr(ilRec).iRdfCode = tlAvRdf(ilRdf).iCode) And (tlAvr(ilRec).iFirstBucket = ilFirstQ) And (tlAvr(ilRec).iDay = ilDay) Then
                                        If (ilRdfCodes(ilRec) = tlAvRdf(ilRdf).iCode) And (tlAvr(ilRec).iFirstBucket = ilFirstQ) And (tlAvr(ilRec).iDay = ilDay) Then
                                            ilFound = True
                                            ilRecIndex = ilRec
                                            Exit For
                                        End If
                                    Next ilRec
                                    If Not ilFound Then
                                        ilRecIndex = UBound(tlAvr)
                                        tlAvr(ilRecIndex).iGenDate(0) = igNowDate(0)
                                        tlAvr(ilRecIndex).iGenDate(1) = igNowDate(1)
                                        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                        tlAvr(ilRecIndex).lGenTime = lgNowTime
                                        tlAvr(ilRecIndex).iDay = ilDay
                                        tlAvr(ilRecIndex).iQStartDate(0) = ilSAvailsDates(0)
                                        tlAvr(ilRecIndex).iQStartDate(1) = ilSAvailsDates(1)
                                        tlAvr(ilRecIndex).iFirstBucket = ilFirstQ
                                        tlAvr(ilRecIndex).sBucketType = slBucketType
                                        tlAvr(ilRecIndex).iDPStartTime(0) = tlAvRdf(ilRdf).iStartTime(0, ilLoopIndex)
                                        tlAvr(ilRecIndex).iDPStartTime(1) = tlAvRdf(ilRdf).iStartTime(1, ilLoopIndex)
                                        tlAvr(ilRecIndex).iDPEndTime(0) = tlAvRdf(ilRdf).iEndTime(0, ilLoopIndex)
                                        tlAvr(ilRecIndex).iDPEndTime(1) = tlAvRdf(ilRdf).iEndTime(1, ilLoopIndex)
                                        tlAvr(ilRecIndex).sDPDays = slDays
                                        tlAvr(ilRecIndex).sNot30Or60 = "N"
                                        tlAvr(ilRecIndex).iVefCode = ilVefCode
                                        'tlAvr(ilRecIndex).iRdfCode = tlAvRdf(ilRdf).iCode
                                        tlAvr(ilRecIndex).iRdfCode = tlAvRdf(ilRdf).iSortCode
                                        ilRdfCodes(ilRecIndex) = tlAvRdf(ilRdf).iCode
                                        tlAvr(ilRecIndex).sInOut = tlAvRdf(ilRdf).sInOut
                                        tlAvr(ilRecIndex).ianfCode = tlAvRdf(ilRdf).ianfCode
                                        ReDim Preserve tlAvr(0 To ilRecIndex + 1) As AVR
                                        ReDim Preserve tlavrcounts(0 To ilRecIndex + 1) As AVRCOUNTS
                                        ReDim Preserve tmInvValAmtSold(0 To ilRecIndex + 1) As INVVALAMTSOLD
                                        ReDim Preserve ilRdfCodes(0 To ilRecIndex + 1)
                                    End If
                                    tlAvr(ilRecIndex).lRate(ilBucketIndexMinusOne) = tlRif(ilRdf).lRate(ilWkNo)
                                    ilDay = ilSaveDay
                                    ilNo30 = 0
                                    ilNo60 = 0
                                    ilLen = tmSdf.iLen
                                    
                                    '3-11-13 valid spot, if processing for avg 30" rate vs R/C, get the spot rate
                                    If tlValuationInfo.iRCvsAvgPrice = 1 Then           'use avg 30" rate vs rate card
                                        'setup schedule line info so clf doesnt have to be read, and the flight routine can get flight rate info
                                        'tmClf.iLine = tmSdf.iLineNo 'TTP 10743
                                        tmClf.iLine = 0 '5/26/23 - Per Jason suppress the line number on the RAB export for the air time and NTR records
                                        tmClf.iPropVer = tmChf.iPropVer
                                        tmClf.iCntRevNo = tmChf.iCntRevNo
                                        
                                        ilRet = gGetFlightPrice(tmSdf, tmClf, hmCff, hmSmf, slPrice)
                                        llPrice = 0     'init incase a decimal number isnt in price field (its adu, nc, fill,etc.)
                                        If (InStr(slPrice, ".") <> 0) Then        'found spot cost
                                            llPrice = gStrDecToLong(slPrice, 2)
                                        End If
                                        tmInvValAmtSold(ilRecIndex).lRate(ilBucketIndex) = tmInvValAmtSold(ilRecIndex).lRate(ilBucketIndex) + llPrice
                                    End If
                                    
                                    If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                        ilNo30 = gDetermineSpotLenRatio(ilLen, tlSpotLenRatio)
                                        If ilNo30 < 0 Then
                                            ilNo30 = -ilNo30
                                            Print #hmMsg, tgMVef(ilVefIndex).sName & " spot length " & str$(ilLen) & " not found. " & str$(ilNo30 / 100) & " units calculated"
                                        End If
    '                                    'Convert inventory to number of 30's and 60's
    '                                    Do While ilLen >= 60
    '                                        ilNo60 = ilNo60 + 1
    '                                        ilLen = ilLen - 60
    '                                    Loop
    '                                    Do While ilLen >= 30
    '                                        ilNo30 = ilNo30 + 1
    '                                        ilLen = ilLen - 30
    '                                    Loop
    '                                    If ilLen < 30 And ilLen > 0 Then    '7-6-00 assume anything under 30" is 1-30" unit availability
    '                                        ilNo30 = ilNo30 + 1
    '                                        ilLen = 0
    '                                    End If
    
                                        If (slBucketType = "S") Or (slBucketType = "P") Then
                                            tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) + ilNo30
                                            tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) + ilNo60
                                        End If
                                        'adjust the available bucket (used for qrtrly detail report only)
                                        tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) - ilNo60
                                        tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) - ilNo30
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                        'Count 30 or 60 and set flag if neither
                                        If ilLen = 60 Then
                                            ilNo60 = 1
                                        ElseIf ilLen = 30 Then
                                            ilNo30 = 1
                                        Else
                                            tlAvr(ilRecIndex).sNot30Or60 = "Y"
                                            If ilLen <= 30 Then
                                                ilNo30 = 1
                                            Else
                                                ilNo60 = 1
                                            End If
                                        End If
                                        If (slBucketType = "S") Or (slBucketType = "P") Then
                                            tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) + ilNo30
                                            tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) + ilNo60
                                        End If
                                        'adjust the available bucket (used for qrtrly detail report only)
                                        If ilNo60 > 0 Then                      'spot found a 60?
                                            tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) - ilNo60
                                        Else
                                            If tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) > 0 Then
                                                tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) - ilNo30
                                            Else
                                                If tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) > 0 Then
                                                    tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) - ilNo30
                                                Else                        'oversold units
                                                    tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) - ilNo30
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
                                            tlAvr(ilRecIndex).sNot30Or60 = "Y"
                                        End If
                                        If (slBucketType = "S") Or (slBucketType = "P") Then
                                            tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Count(ilBucketIndex) + ilNo30
                                            tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Count(ilBucketIndex) + ilNo60
                                        End If
                                        'adjust the available bucket (used for qrtrly detail report only)
                                        tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l60Avail(ilBucketIndex) - ilNo60
                                        tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) = tlavrcounts(ilRecIndex).l30Avail(ilBucketIndex) - ilNo30
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

    'Adjust counts for qtrly detail availbilty
    ilNo30 = gDetermineSpotLenRatio(30, tlSpotLenRatio)
    If ilNo30 < 0 Then
        ilNo30 = -ilNo30
        Print #hmMsg, tgMVef(ilVefIndex).sName & " spot length 30 not found. 1 unit calculated"
    End If
    ilNo60 = gDetermineSpotLenRatio(60, tlSpotLenRatio)
    If ilNo60 < 0 Then
        ilNo60 = -ilNo60
        Print #hmMsg, tgMVef(ilVefIndex).sName & " spot length 60 not found. 2 units calculated"
    End If
    'there shouldn't be any 60s avail since its all based on 30" avails
    If (tgVpf(ilVpfIndex).sSSellOut = "B") Then                   'qtrly detail?
        For ilRec = 0 To UBound(tlAvr) - 1 Step 1
            ilLoopIndex = ilNoWks
            If ilNoWks = 14 Then
                ilLoopIndex = ilNoWks - 1
            End If
            For ilLoop = 1 To ilLoopIndex Step 1               '6-30-00
                If tlavrcounts(ilRec).l30Avail(ilLoop) < 0 Then
                    Do While (tlavrcounts(ilRec).l60Avail(ilLoop) > 0) And (tlavrcounts(ilRec).l30Avail(ilLoop) < 0)
'                        tlAvr(ilRec).i60Avail(ilLoop) = tlAvr(ilRec).i60Avail(ilLoop) - 1
'                        tlAvr(ilRec).i30Avail(ilLoop) = tlAvr(ilRec).i30Avail(ilLoop) + 2
                        tlavrcounts(ilRec).l60Avail(ilLoop) = tlavrcounts(ilRec).l60Avail(ilLoop) - ilNo60
                        tlavrcounts(ilRec).l30Avail(ilLoop) = tlavrcounts(ilRec).l30Avail(ilLoop) + ilNo30
                    Loop
                ElseIf (tlavrcounts(ilRec).l60Avail(ilLoop) < 0) Then
                End If
            Next ilLoop
        Next ilRec
    End If
    
    'Calculate the avg 30" rate vs by R/C prices
    For ilRec = 0 To UBound(tmInvValAmtSold) - 1
        ilRdfNameIndex = gBinarySearchRdf(ilRdfCodes(ilRec))
        slAvgRate = tgMVef(ilVefIndex).sName & " Avg Rates for DP " & Trim(tgMRdf(ilRdfNameIndex).sName) & ":"
        'i30count contains the sold to calculate the avg rate
        For ilLoop = 1 To 14
             If tlValuationInfo.iRCvsAvgPrice = 1 Then       'do by avg 30 vs r/c, need to alter the avg rate; other the r/c is already in that tlavr.lrate field
                If (tlavrcounts(ilRec).l30Count(ilLoop) + tlavrcounts(ilRec).l30Reserve(ilLoop) + tlavrcounts(ilRec).l30Hold(ilLoop)) > 0 Then
                    'get the avg rate and truncate pennies
                    tlAvr(ilRec).lRate(ilLoop - 1) = tmInvValAmtSold(ilRec).lRate(ilLoop) / (tlavrcounts(ilRec).l30Count(ilLoop) + tlavrcounts(ilRec).l30Reserve(ilLoop) + tlavrcounts(ilRec).l30Hold(ilLoop))
                Else
                    tlAvr(ilRec).lRate(ilLoop - 1) = 0
                End If
            End If
            tlAvr(ilRec).lRate(ilLoop - 1) = (tlAvr(ilRec).lRate(ilLoop - 1) * tlValuationInfo.iUnsoldPctAdj) / 100
            'the available counts are using the table, adjust to remove the hundreds
            tlavrcounts(ilRec).l30Avail(ilLoop) = tlavrcounts(ilRec).l30Avail(ilLoop) / 100
            tlavrcounts(ilRec).l30Count(ilLoop) = tlavrcounts(ilRec).l30Count(ilLoop) / 100
            tlavrcounts(ilRec).l30InvCount(ilLoop) = tlavrcounts(ilRec).l30InvCount(ilLoop) / 100
            tlavrcounts(ilRec).l30Hold(ilLoop) = tlavrcounts(ilRec).l30Hold(ilLoop) / 100
            tlavrcounts(ilRec).l30Prop(ilLoop) = tlavrcounts(ilRec).l30Prop(ilLoop) / 100
            tlavrcounts(ilRec).l30Reserve(ilLoop) = tlavrcounts(ilRec).l30Reserve(ilLoop) / 100
            slAvgRate = slAvgRate & str$(tlAvr(ilRec).lRate(ilLoop - 1)) & ", "
        Next ilLoop
        ilRet = ilRet
        Print #hmMsg, slAvgRate
    Next ilRec
    
    'the % of sellout can be adjusted, but wont be adjusted until all the numbers have been gathered.
    
    'Combines weeks into the proper months for monthly figures
    For ilRec = 0 To UBound(tlAvr) - 1 Step 1  'next daypart
        For ilLoop = 1 To 3 Step 1
            If ilLoop = 1 Then
                ilLo = 1
                ilHi = ilWksInMonth(1)
            Else
                ilLo = ilHi + 1
                ilHi = ilHi + ilWksInMonth(ilLoop)
            End If
            If ilHi = 14 Then           'only go 13 buckets, the avails have been placed in 13 bucket for 14 week quarter
                ilHi = 13
            End If
            For ilLoopIndex = ilLo To ilHi Step 1       '6-30-00
                'avails remaining * the avg rate = inventory valuation
                ilIndex = ilLoopIndex           '6-30-00
                If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                    tlAvr(ilRec).lMonth(ilLoop - 1) = tlAvr(ilRec).lMonth(ilLoop - 1) + (((tlavrcounts(ilRec).l60Avail(ilIndex) * 2) + tlavrcounts(ilRec).l30Avail(ilIndex)) * tlAvr(ilRec).lRate(ilIndex - 1))
                ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Or tgVpf(ilVpfIndex).sSSellOut = "U" Then
                    tlAvr(ilRec).lMonth(ilLoop - 1) = tlAvr(ilRec).lMonth(ilLoop - 1) + ((tlavrcounts(ilRec).l60Avail(ilIndex) + tlavrcounts(ilRec).l30Avail(ilIndex)) * tlAvr(ilRec).lRate(ilIndex - 1))
                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                End If
            Next ilLoopIndex
        Next ilLoop
    Next ilRec
        
    'All the values have been reduced back down without pennies and decimals positions
    'put back all the data into original buffers
    For ilRec = 0 To UBound(tlAvr) - 1
        For ilLoop = 1 To 14
            tlAvr(ilRec).i30Avail(ilLoop - 1) = tlavrcounts(ilRec).l30Avail(ilLoop)
            tlAvr(ilRec).i60Avail(ilLoop - 1) = tlavrcounts(ilRec).l60Avail(ilLoop)
            tlAvr(ilRec).i30Count(ilLoop - 1) = tlavrcounts(ilRec).l30Count(ilLoop)
            tlAvr(ilRec).i60Avail(ilLoop - 1) = tlavrcounts(ilRec).l60Count(ilLoop)
            tlAvr(ilRec).i30InvCount(ilLoop - 1) = tlavrcounts(ilRec).l30InvCount(ilLoop)
            tlAvr(ilRec).i60InvCount(ilLoop - 1) = tlavrcounts(ilRec).l60InvCount(ilLoop)
            tlAvr(ilRec).i30Hold(ilLoop - 1) = tlavrcounts(ilRec).l30Hold(ilLoop)
            tlAvr(ilRec).i60Hold(ilLoop - 1) = tlavrcounts(ilRec).l60Hold(ilLoop)
            tlAvr(ilRec).i30Prop(ilLoop - 1) = tlavrcounts(ilRec).l30Prop(ilLoop)
            tlAvr(ilRec).i60Prop(ilLoop - 1) = tlavrcounts(ilRec).l60Prop(ilLoop)
            tlAvr(ilRec).i30Reserve(ilLoop - 1) = tlavrcounts(ilRec).l30Reserve(ilLoop)
            tlAvr(ilRec).i60Reserve(ilLoop - 1) = tlavrcounts(ilRec).l60Reserve(ilLoop)
            
        Next ilLoop
    Next ilRec
    
    Erase ilSAvailsDates, tlavrcounts
    Erase ilEvtType
    Erase ilRdfCodes
    Erase tlLLC
End Sub

Function gGetRate(tlSdf As SDF, tlClf As CLF, hlCff As Integer, tlSmf As SMF, tlCff As CFF, slPrice As String) As Integer
'
'   ilRet = gGetRate(tlSdf, tlClf, hlCff, tlSmf, tlCff)
'   Where:
'       tlSdf(I)- Spot image to get flight for
'       tlClf(I)- Line image to get iCntRevNo and iPropVer
'       hlCff(I)- Open Handle to CFF
'       tlSmf(I)- SMF image
'       tlCff(O)- Spot flight
'   slPrice(0) - Price xxxxx.xx
'       ilRet = True if found; False if error
'
'   1-19-04 change manner in which fill/extra spots are shown on invoice
'           if spot price type is other than "-" or "+", then use advt to dtermine default

    Dim llSpotDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim tlCffSrchKey As CFFKEY0 'CFF key record image
    Dim ilCffRecLen As Integer     'CFF record length
    Dim ilRet As Integer
    Dim slShowOnInv As String * 1       '1-19-04
    gGetRate = False
    If tlSdf.sSpotType = "X" Then
        '1-19-04
        slShowOnInv = gTestShowFill(tlSdf.sPriceType, tlSdf.iAdfCode)
        If slShowOnInv = "Y" Then
        'If tlSdf.sPriceType <> "N" Then
            slPrice = "Extra"
        Else
            slPrice = "Fill"
        End If
        gGetRate = True
        Exit Function
    End If
    If (tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O") Then
        gUnpackDateLong tlSmf.iMissedDate(0), tlSmf.iMissedDate(1), llSpotDate
    Else
        gUnpackDateLong tlSdf.iDate(0), tlSdf.iDate(1), llSpotDate
    End If
    tlCffSrchKey.lChfCode = tlSdf.lChfCode  'llChfCode
    tlCffSrchKey.iClfLine = tlClf.iLine 'tlSdf.iLineNo using line so avg price can be obtained for package line which bill by airing
    tlCffSrchKey.iCntRevNo = tlClf.iCntRevNo
    tlCffSrchKey.iPropVer = tlClf.iPropVer
    tlCffSrchKey.iStartDate(0) = 0
    tlCffSrchKey.iStartDate(1) = 0
    ilCffRecLen = Len(tlCff)
    ilRet = btrGetGreaterOrEqual(hlCff, tlCff, ilCffRecLen, tlCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tlCff.lChfCode = tlSdf.lChfCode) And (tlCff.iClfLine = tlClf.iLine)
        If (tlCff.iCntRevNo = tlClf.iCntRevNo) And (tlCff.iPropVer = tlClf.iPropVer) Then 'And (tmCff(2).sDelete <> "Y") Then
            gUnpackDateLong tlCff.iStartDate(0), tlCff.iStartDate(1), llStartDate    'Week Start date
            gUnpackDateLong tlCff.iEndDate(0), tlCff.iEndDate(1), llEndDate    'Week Start date
            If (llSpotDate >= llStartDate) And (llSpotDate <= llEndDate) Then
                gGetRate = True
                Select Case tlCff.sPriceType   'tlCff.sPriceType
                    Case "T"    'True
                        slPrice = gLongToStrDec(tlCff.lActPrice, 2)    'tlCff.lActPrice, 2)
                    Case "N"    'No Charge
                        slPrice = "N/C"
                    Case "M"    'MG Line
                        slPrice = "MG"
                    Case "B"    'Bonus
                        slPrice = "Bonus"
                    Case "S"    'Spinoff
                        slPrice = "Spinoff"
                    Case "P"    'Package
                        slPrice = ".00"
                    Case "R"    'Recapturable
                        slPrice = "Recapturable"
                    Case "A"    'ADU
                        slPrice = "ADU"
                End Select
                Exit Function
            End If
        End If
        ilRet = btrGetNext(hlCff, tlCff, ilCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop

    If Not ilRet Then
        slPrice = ".00"
        gGetRate = False
    End If
End Function

'           gGetVehGrpSets - Given the vehicle code, find the vehicle
'               record in tgMVef, and return the major and minor vehicle
'               group set.
'           <input> ilvefcode = Vehicle code
'                   ilMinorSet = set # that determines the minor sort
'                   ilMajorSet = set # that determines the major sort
'           <output> ilmnfMinorCode - mnf code for the minor sort field
'                   ilmnfMajorCode - mnf code for the major sort field
'
'           Created 6/10/98 D. Hosaka
'
Sub gGetVehGrpSets(ilVefCode As Integer, ilMinorSet As Integer, ilMajorSet As Integer, ilmnfMinorCode As Integer, ilMnfMajorCode As Integer)
Dim ilLoop As Integer
    ilmnfMinorCode = 0
    ilMnfMajorCode = 0
    'For ilLoop = LBound(tgMVef) To UBound(tgMVef) Step 1
    '    If tgMVef(ilLoop).iCode = ilVefCode Then
    ilLoop = gBinarySearchVef(ilVefCode)
    If ilLoop <> -1 Then
        If ilMinorSet = 1 Then          'using first vehicle group
            'ilmnfMinorCode = tgMVef(ilLoop).iMnfGroup(1)            'participants
            ilmnfMinorCode = tgMVef(ilLoop).iOwnerMnfCode
        ElseIf ilMinorSet = 2 Then
            ilmnfMinorCode = tgMVef(ilLoop).iMnfVehGp2              'sub-totals
        ElseIf ilMinorSet = 3 Then
            ilmnfMinorCode = tgMVef(ilLoop).iMnfVehGp3Mkt           'markets
        ElseIf ilMinorSet = 4 Then
            ilmnfMinorCode = tgMVef(ilLoop).iMnfVehGp4Fmt           'formats
        ElseIf ilMinorSet = 5 Then
            ilmnfMinorCode = tgMVef(ilLoop).iMnfVehGp5Rsch           'reserach
        ElseIf ilMinorSet = 6 Then
            ilmnfMinorCode = tgMVef(ilLoop).iMnfVehGp6Sub           'Sub-Company
        End If
        If ilMajorSet = 1 Then
            'ilMnfMajorCode = tgMVef(ilLoop).iMnfGroup(1)            'participants
            ilMnfMajorCode = tgMVef(ilLoop).iOwnerMnfCode
        ElseIf ilMajorSet = 2 Then
            ilMnfMajorCode = tgMVef(ilLoop).iMnfVehGp2              'sub-totals
        ElseIf ilMajorSet = 3 Then
            ilMnfMajorCode = tgMVef(ilLoop).iMnfVehGp3Mkt           'markets
        ElseIf ilMajorSet = 4 Then
            ilMnfMajorCode = tgMVef(ilLoop).iMnfVehGp4Fmt           'formats
        ElseIf ilMajorSet = 5 Then
            ilMnfMajorCode = tgMVef(ilLoop).iMnfVehGp5Rsch           'reserach
        ElseIf ilMajorSet = 6 Then
            ilMnfMajorCode = tgMVef(ilLoop).iMnfVehGp6Sub           'Sub-Company
        End If
    End If
    'Next ilLoop
End Sub

'           Clear the Invoice Prepass file
'
'           Created:  10/19/98
'
'
Sub gIvrClear()
    Dim ilRet As Integer
    Dim tlIvr As IVR
    hmIvr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmIvr, "", sgDBPath & "Ivr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmIvr)
        btrDestroy hmIvr
        Exit Sub
    End If
    imIvrRecLen = Len(tlIvr)
    tmIvrSrchKey.iGenDate(0) = igNowDate(0)
    tmIvrSrchKey.iGenDate(1) = igNowDate(1)
    'tmIvrSrchKey.iGenTime(0) = igNowTime(0)
    'tmIvrSrchKey.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmIvrSrchKey.lGenTime = lgNowTime
    tmIvrSrchKey.lInvNo = 0
    ilRet = btrGetGreaterOrEqual(hmIvr, tlIvr, imIvrRecLen, tmIvrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    'Do While (ilRet = BTRV_ERR_NONE) And (tlIvr.iGenDate(0) = igNowDate(0)) And (tlIvr.iGenDate(1) = igNowDate(1)) And (tlIvr.iGenTime(0) = igNowTime(0)) And (tlIvr.iGenTime(1) = igNowTime(1))
    Do While (ilRet = BTRV_ERR_NONE) And (tlIvr.iGenDate(0) = igNowDate(0)) And (tlIvr.iGenDate(1) = igNowDate(1)) And (tlIvr.lGenTime = lgNowTime)
        ilRet = btrDelete(hmIvr)
        ilRet = btrGetNext(hmIvr, tlIvr, imIvrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmIvr)
    btrDestroy hmIvr
End Sub

Function gObtainPhfRvf(RptForm As Form, slEarliestDate As String, slLatestDate As String, tlTranType As TRANTYPES, tlRvf() As RVF, ilWhichDate As Integer) As Integer
'****************************************************************
'*
'*      Obtain all History and Receivables transactions whose
'*      transaction date falls within the earliest and latest
'*      dates requested.  Test for transaction types "I", "A"
'*      or "W".
'
'*      <input>  RptForm - Form calling this populate rtn
'*               slEarliestDate - get all trans starting with
'*                   this date
'*               slLatestDAte - get all trans equal or prior to
'*                  this pacing date (effective date)
'*              ilWhichDate, 0 = trandate, 1 = entry date
'*
'*      <I/O>    tlRvf() - array of matching Phf/Rvf recds
'*               funtion return - true if receivables populated
'*                       false if no receivables, error
'*
'*             Created:4/17/98       By:D. Hosaka
'*            Modified:              By:
'*
'*            Comments: make 2 passes, read Phf, the Rvf, build
'*               all transactions types "A", "I", "W" or "P"
'*               based on parameters in tlTranType.
'*            3-14-01 Include "HI" with the "I" transactions
'            9-17-02 Add NTR flag to list of filters (test for presence of item bill mnfcode
'*          5-26-04 Exclude NTR "AN" transactions when NTR to be excluded
'*          2-11-05 Change array from single to double integer to prevent Overflow
'                   (records in excess of 32000)
'****************************************************************
'
'    ilRet = gObtainPhfRvf (RptForm,  slEarliestDate, slLatestDate, tlTranType, tlRvf())
'
    Dim ilRet As Integer    'Return status
    ReDim ilEarliestDate(0 To 1) As Integer
    ReDim ilLatestDate(0 To 1) As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilLoop As Integer
    'Dim ilRVFUpper As Integer
    Dim llRVFUpper As Long          '2-11-05 chg to long
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilLowLimit As Integer
    Dim ilDoe As Integer
    
    'On Error GoTo gObtainPhfRvfErr
    'ilRet = 0
    'ilLowLimit = LBound(tlRvf)
    'If ilRet <> 0 Then
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlRvf).Ptr <> 0 Then
        ilLowLimit = LBound(tlRvf)
    Else
        ilLowLimit = 0
    End If
    
    ReDim tlRvf(ilLowLimit To ilLowLimit) As RVF
    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        gObtainPhfRvf = False
        Exit Function
    End If
    imRvfRecLen = Len(tlRvf(ilLowLimit))
    gPackDate slEarliestDate, ilEarliestDate(0), ilEarliestDate(1)
    gPackDate slLatestDate, ilLatestDate(0), ilLatestDate(1)
    btrExtClear hmRvf   'Clear any previous extend operation
    ilExtLen = Len(tlRvf(ilLowLimit))  'Extract operation record size
    llRVFUpper = UBound(tlRvf)
    For ilLoop = 1 To 2         'pass 1- get PHF, pass 2 get RVF
        '7-1-14 use key3 (rvftrandate) instead of key 0 (rvfagfcode)
        ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
            Call btrExtSetBounds(hmRvf, llNoRec, -1, "UC", "RVF", "") '"EG") 'Set extract limits (all records)

            tlDateTypeBuff.iDate0 = ilEarliestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilEarliestDate(1)
            If ilWhichDate = 0 Then                         '12-14-06
                ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
            Else
                ilOffSet = gFieldOffset("Rvf", "RvfDateEntrd")
            End If
            ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainPhfRvf (btrExtAddLogicConst):" & "Rvf.Btr", RptForm
            On Error GoTo 0


            tlDateTypeBuff.iDate0 = ilLatestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilLatestDate(1)
            If ilWhichDate = 0 Then                         '12-14-06
                ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
            Else
                ilOffSet = gFieldOffset("Rvf", "RvfDateEntrd")
            End If

            ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainPhfRvf (btrExtAddLogicConst):" & "Rvf.Btr", RptForm
            On Error GoTo 0

            ilRet = btrExtAddField(hmRvf, 0, ilExtLen)  'Extract the whole record
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainRVF (btrExtAddField):" & "RVF.Btr", RptForm
            On Error GoTo 0
            ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
            ilDoe = ilDoe + 1
            If ilDoe > 3000 Then
                ilDoe = 0
                DoEvents
            End If
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo mRvfErr
                gBtrvErrorMsg ilRet, "gObtainRVF (btrExtGetNextExt):" & "RVF.Btr", RptForm
                On Error GoTo 0
                ilExtLen = Len(tmRvf)  'Extract operation record size
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    ilDoe = ilDoe + 1
                    If ilDoe > 25000 Then
                        ilDoe = 0
                        DoEvents
                    End If
                    'first test for valid trans types (Invoices, adjustments, write-off & payments
                    If ((Left$(tmRvf.sTranType, 1) = "I" Or tmRvf.sTranType = "HI") And tlTranType.iInv) Or (Left$(tmRvf.sTranType, 1) = "A" And tlTranType.iAdj) Or (Left$(tmRvf.sTranType, 1) = "W" And tlTranType.iWriteOff) Or (Left$(tmRvf.sTranType, 1) = "P" And tlTranType.iPymt) Then
                        If (tlTranType.iNTR) Then       'NTR option, tested separately because it shouldnt be tested with Cash transactions
                            If tmRvf.iMnfItem > 0 Then
                                tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                            Else            'its not an NTR
                                'got valid trans type - test for Cash, Trade, Merchandising or Promotions
                                If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                    tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                    ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                                End If
                            End If
                        Else
                            'got valid trans type - test for Cash, Trade, Merchandising or Promotions
                            '05-26-04 dont include NTR, exclude this if it is
                            '2-27-08 chg test to test for NTR using mnitem only; Installment records will also have an sbfcode
                            If tmRvf.iMnfItem = 0 Then      'and tmRvf.isbfcode = 0
                                If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                    tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                    ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                                End If
                            End If
                        End If
                    End If
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
        If ilLoop = 1 Then                          'if 1, then just finished history, go do Receivables
            btrExtClear hmRvf   'Clear any previous extend operation
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmRvf)
                btrDestroy hmRvf
                gObtainPhfRvf = False
                Exit Function
            End If
            imRvfRecLen = Len(tmRvf)
            llRVFUpper = UBound(tlRvf)
        End If
    Next ilLoop
    ilRet = btrClose(hmRvf)
    btrDestroy hmRvf
    gObtainPhfRvf = True
    Exit Function
gObtainPhfRvfErr:
    ilRet = 1
    Resume Next
mRvfErr:
    On Error GoTo 0
    gObtainPhfRvf = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainPjf                      *
'*      Extended read to get all matching projection   *
'*      records with passed rollover date              *
'*      <input>  hlPjf - Projection handle (file must  *
'*                       open                          *
'*               ilRODate(0 to 1) - Rollover date to   *
'*                       match                         *
'*      <I/O>    tlPjf() - array of matching PJF recds *
'*                                                     *
'*             Created:7/24/97       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read all of pjf by rollover    *
'*                      date                           *
'*******************************************************
Function gObtainPjf(RptForm As Form, hlPjf As Integer, ilRODate() As Integer, tlPjf() As PJF) As Integer
'
'    gObtainPjf (hlPjf, ilRODate(), tlPjf())
'
    Dim ilRet As Integer    'Return status
    Dim slStr As String
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilPjfUpper As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record

    ReDim tlPjf(0 To 0) As PJF
    btrExtClear hlPjf   'Clear any previous extend operation
    ilExtLen = Len(tlPjf(0))  'Extract operation record size
    imPjfRecLen = Len(tmPjf)
    ilPjfUpper = UBound(tlPjf)
    ilRet = btrGetFirst(hlPjf, tmPjf, imPjfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hlPjf, llNoRec, -1, "UC", "PJF", "") '"EG") 'Set extract limits (all records)
        tlDateTypeBuff.iDate0 = ilRODate(0)                       'retrieve past  projection records
        tlDateTypeBuff.iDate1 = ilRODate(1)
        ilOffSet = gFieldOffset("Pjf", "PjfRolloverDate")
        ilRet = btrExtAddLogicConst(hlPjf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        On Error GoTo mRolloverErr
        gBtrvErrorMsg ilRet, "gObtainPjf (btrExtAddLogicConst):" & "Pjf.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtAddField(hlPjf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mRolloverErr
        gBtrvErrorMsg ilRet, "gObtainPjf (btrExtAddField):" & "Pjf.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlPjf, tmPjf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mRolloverErr
            gBtrvErrorMsg ilRet, "gObtainPjf (btrExtGetNextExt):" & "Pjf.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmPjf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlPjf, tmPjf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slStr = ""
                tlPjf(UBound(tlPjf)) = tmPjf           'save entire record
                ReDim Preserve tlPjf(0 To UBound(tlPjf) + 1) As PJF
                ilRet = btrExtGetNext(hlPjf, tmPjf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlPjf, tmPjf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainPjf = True
    Exit Function
mRolloverErr:
    On Error GoTo 0
    gObtainPjf = False
    Exit Function
End Function

Function gObtainSlf(RptForm As Form, hlSlf As Integer, tlSlf() As SLF) As Integer
'*******************************************************
'*      <input>  hlSlf - Salesperson handle (file must *
'*                       open                          *
'*      <I/O>    tlSlf() - array of matching SLF recds *
'*                                                     *
'*             Created:8/14/97       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read all of slf records        *
'*                                              *
'*******************************************************
'
'    gObtainSlf (hlSlf,  tlSlf())
'
    Dim ilRet As Integer    'Return status
    Dim slStr As String
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilSlfUpper As Integer
    Dim ilLowLimit As Integer

    'On Error GoTo gObtainSlfErr
    'ilRet = 0
    'ilLowLimit = LBound(tlSlf)
    'If ilRet <> 0 Then
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSlf).Ptr <> 0 Then
        ilLowLimit = LBound(tlSlf)
    Else
        ilLowLimit = 0
    End If

    ReDim tlSlf(ilLowLimit To ilLowLimit) As SLF
    btrExtClear hlSlf   'Clear any previous extend operation
    ilExtLen = Len(tlSlf(ilLowLimit))  'Extract operation record size
    imSlfRecLen = Len(tmSlf)
    ilSlfUpper = UBound(tlSlf)
    'ilRet = btrGetFirst(hlSlf, tlSlf(0), imSlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ilRet = btrGetFirst(hlSlf, tmSlf, imSlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hlSlf, llNoRec, -1, "UC", "SLF", "") '"EG") 'Set extract limits (all records)
        ilRet = btrExtAddField(hlSlf, 0, ilExtLen)  'Extract the whole record
        On Error GoTo mSlspErr
        gBtrvErrorMsg ilRet, "gObtainSlf (btrExtAddField):" & "Slf.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlSlf, tmSlf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mSlspErr
            gBtrvErrorMsg ilRet, "gObtainSlf (btrExtGetNextExt):" & "Slf.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmSlf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSlf, tmSlf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slStr = ""
                tlSlf(UBound(tlSlf)) = tmSlf           'save entire record
                ReDim Preserve tlSlf(ilLowLimit To UBound(tlSlf) + 1) As SLF
                ilRet = btrExtGetNext(hlSlf, tmSlf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSlf, tmSlf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainSlf = True
    Exit Function
gObtainSlfErr:
    ilRet = 1
    Resume Next
mSlspErr:
    On Error GoTo 0
    gObtainSlf = False
    Exit Function
End Function

'
'                   Populate the unique groups for selection
'                   <input>  cbcSet1 - control name to populate
'                            ilUseNone  - add Item to List Box with "None" (true/false)
'
'                   D Hosaka 10/1/98
Sub gPopVehicleGroups(cbcSet As Control, tlVehicleSets() As POPICODENAME, ilUseNone As Integer)
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilLoop2 As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    ilRet = gObtainMnfForType("H", sgMnfVehGrpTag, tgMMnf())
    'cbcSet1.AddItem "None"
    'ReDim ilVehGroup1(1 To 1) As Integer
    ReDim ilVehGroup1(0 To 0) As Integer
    'ReDim slVehGroup1(1 To 1) As String * 1
    ReDim slVehGroup1(0 To 0) As String * 1
    'For ilLoop = 1 To UBound(tgMMnf) - 1
    For ilLoop = LBound(tgMMnf) To UBound(tgMMnf) - 1
        ilFound = False
        'For ilIndex = 1 To UBound(slVehGroup1) - 1 Step 1
        For ilIndex = LBound(slVehGroup1) To UBound(slVehGroup1) - 1 Step 1
            If Trim$(tgMMnf(ilLoop).sUnitType) = slVehGroup1(ilIndex) Then          'look for the vehicle set # built in array
                ilFound = True
                Exit For
            End If
        Next ilIndex
        If Not ilFound Then
            slVehGroup1(ilIndex) = tgMMnf(ilLoop).sUnitType          'vehicle set #
            ilVehGroup1(ilIndex) = Val(slVehGroup1(ilIndex))
            'ReDim Preserve slVehGroup1(1 To UBound(slVehGroup1) + 1)
            ReDim Preserve slVehGroup1(0 To UBound(slVehGroup1) + 1)
            'ReDim Preserve ilVehGroup1(1 To UBound(ilVehGroup1) + 1)
            ReDim Preserve ilVehGroup1(0 To UBound(ilVehGroup1) + 1)
        End If
    Next ilLoop
    'sort the unique set #s
    'For ilLoop = 1 To UBound(ilVehGroup1) - 1
    For ilLoop = LBound(ilVehGroup1) To UBound(ilVehGroup1) - 1
        For ilLoop2 = ilLoop + 1 To UBound(ilVehGroup1) - 1
            If ilVehGroup1(ilLoop) > ilVehGroup1(ilLoop2) Then
                'swap the two
                ilFound = ilVehGroup1(ilLoop)
                ilVehGroup1(ilLoop) = ilVehGroup1(ilLoop2)
                ilVehGroup1(ilLoop2) = ilFound
            End If
        Next ilLoop2
    Next ilLoop
    'If sgVehicleSetsTag = "" Then
    cbcSet.Clear
    'sgVehicleSetsTag = Format$(Now, "h/m/yy h:mm")
    ReDim tlVehicleSets(0 To 0) As POPICODENAME
    If ilUseNone Then
        tlVehicleSets(0).iCode = 0
        tlVehicleSets(0).sChar = "None"
        ReDim Preserve tlVehicleSets(0 To UBound(tlVehicleSets) + 1) As POPICODENAME
    End If
    'For ilLoop = 1 To UBound(ilVehGroup1) - 1 Step 1
    For ilLoop = LBound(ilVehGroup1) To UBound(ilVehGroup1) - 1 Step 1
        'ReDim Preserve tgVehicleSets(0 To UBound(tgVehicleSets) + 1) As POPICODENAME
        tlVehicleSets(UBound(tlVehicleSets)).iCode = ilVehGroup1(ilLoop)
        If ilVehGroup1(ilLoop) = 1 Then
            tlVehicleSets(UBound(tlVehicleSets)).sChar = "Participants"
        ElseIf ilVehGroup1(ilLoop) = 2 Then
            tlVehicleSets(UBound(tlVehicleSets)).sChar = "Sub-Totals"
        ElseIf ilVehGroup1(ilLoop) = 3 Then
            tlVehicleSets(UBound(tlVehicleSets)).sChar = "Market"
        ElseIf ilVehGroup1(ilLoop) = 4 Then
            tlVehicleSets(UBound(tlVehicleSets)).sChar = "Format"
        ElseIf ilVehGroup1(ilLoop) = 5 Then
            tlVehicleSets(UBound(tlVehicleSets)).sChar = "Research"
        ElseIf ilVehGroup1(ilLoop) = 6 Then
            tlVehicleSets(UBound(tlVehicleSets)).sChar = "Sub-Company"
        End If
        ReDim Preserve tlVehicleSets(0 To UBound(tlVehicleSets) + 1) As POPICODENAME
    Next ilLoop

    'Fill the List boxes with the unique set #s
    For ilLoop = 0 To UBound(tlVehicleSets) Step 1
        cbcSet.AddItem Trim$(tlVehicleSets(ilLoop).sChar)
        'cbcSet1.AddItem Str$(ilVehGroup1(ilLoop))
    Next ilLoop
    cbcSet.ListIndex = 0
    'End If
End Sub

Public Function gObtainPhfOrRvf(RptForm As Form, slEarliestDate As String, slLatestDate As String, tlTranType As TRANTYPES, tlRvf() As RVF, ilWhichFile As Integer, ilWhichDate) As Integer
'****************************************************************
'*
'*      Obtain all History OR Receivables transactions whose
'*      transaction date falls within the earliest and latest
'*      dates requested.  Test for transaction types "I", "A"
'*      or "W".
'
'*      <input>  RptForm - Form calling this populate rtn
'*               slEarliestDate - get all trans starting with
'*                   this date
'*               slLatestDAte - get all trans equal or prior to
'*                  this pacing date (effective date)
'*              ilWhichFile - 1 = PHF, 2 = RVF, 3= both
'*
'*      <I/O>    tlRvf() - array of matching Phf/Rvf recds
'*               funtion return - true if receivables populated
'*                       false if no receivables, error
'*              ilWhichDate, 0 = trandate, 1 = entry date  added 6-04-08 Dan M
'*             Created:4/17/98       By:D. Hosaka
'*            Modified:              By:
'*
'*            Comments: make up to 2 passes, read Phf, the Rvf, build
'*               all transactions types "A", "I", "W" or "P"
'*               based on parameters in tlTranType.
'*          9-11-03 duplicated from gobtainPHForRVf to be able to search
'*          RVF only, PHF only, or both
'*
'*          5-26-04 Exclude NTR "AN" transactions when NTR to be excluded
'****************************************************************
'
'    ilRet = gObtainPhfORRvf (RptForm,  slEarliestDate, slLatestDate, tlTranType, tlRvf(),ilWhichFile)
'
    Dim ilRet As Integer    'Return status
    ReDim ilEarliestDate(0 To 1) As Integer
    ReDim ilLatestDate(0 To 1) As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilLoop As Integer
    Dim ilRVFUpper As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilStartFile As Integer
    Dim ilEndFile As Integer
    Dim ilLowLimit As Integer
    
    'On Error GoTo gObtainPhfOrRvfErr
    'ilRet = 0
    'ilLowLimit = LBound(tlRvf)
    'If ilRet <> 0 Then
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlRvf).Ptr <> 0 Then
        ilLowLimit = LBound(tlRvf)
    Else
        ilLowLimit = 0
    End If
    
    ReDim tlRvf(ilLowLimit To ilLowLimit) As RVF

    If ilWhichFile = 1 Then
        ilStartFile = 1
        ilEndFile = 1
        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            gObtainPhfOrRvf = False
            Exit Function
        End If
    ElseIf ilWhichFile = 2 Then
        ilStartFile = 2
        ilEndFile = 2
        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            gObtainPhfOrRvf = False
            Exit Function
        End If
    Else
        ilStartFile = 1
        ilEndFile = 2
        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            gObtainPhfOrRvf = False
            Exit Function
        End If
    End If


    imRvfRecLen = Len(tlRvf(ilLowLimit))
    gPackDate slEarliestDate, ilEarliestDate(0), ilEarliestDate(1)
    gPackDate slLatestDate, ilLatestDate(0), ilLatestDate(1)
    btrExtClear hmRvf   'Clear any previous extend operation
    ilExtLen = Len(tlRvf(ilLowLimit))  'Extract operation record size
    'ilRVFUpper = UBound(tlRvf)             1-11-18 removed, unused
    For ilLoop = ilStartFile To ilEndFile         'pass 1- get PHF, pass 2 get RVF
        '7-1-14 use key3 (rvftrandate) instead of key 0 (rvfagfcode)
        ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
            Call btrExtSetBounds(hmRvf, llNoRec, -1, "UC", "RVF", "") '"EG") 'Set extract limits (all records)

            tlDateTypeBuff.iDate0 = ilEarliestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilEarliestDate(1)
            If ilWhichDate = 0 Then                         '6-04-08 Dan M copied from gObtainPhfRvf
                ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
            Else
                ilOffSet = gFieldOffset("Rvf", "RvfDateEntrd")
            End If

            ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainPhfOrRvf (btrExtAddLogicConst):" & "Rvf.Btr", RptForm
            On Error GoTo 0


            tlDateTypeBuff.iDate0 = ilLatestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilLatestDate(1)

            If ilWhichDate = 0 Then                         '6-04-08 Dan M
                ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
            Else
                ilOffSet = gFieldOffset("Rvf", "RvfDateEntrd")
            End If


            ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainPhfOrRvf (btrExtAddLogicConst):" & "Rvf.Btr", RptForm
            On Error GoTo 0

            ilRet = btrExtAddField(hmRvf, 0, ilExtLen)  'Extract the whole record
            On Error GoTo mRvfErr
            gBtrvErrorMsg ilRet, "gObtainRVF (btrExtAddField):" & "RVF.Btr", RptForm
            On Error GoTo 0
            ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo mRvfErr
                gBtrvErrorMsg ilRet, "gObtainRVF (btrExtGetNextExt):" & "RVF.Btr", RptForm
                On Error GoTo 0
                ilExtLen = Len(tmRvf)  'Extract operation record size
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    'first test for valid trans types (Invoices, adjustments, write-off & payments

                    If ((Left$(tmRvf.sTranType, 1) = "I" Or tmRvf.sTranType = "HI") And tlTranType.iInv) Or (Left$(tmRvf.sTranType, 1) = "A" And tlTranType.iAdj) Or (Left$(tmRvf.sTranType, 1) = "W" And tlTranType.iWriteOff) Or (Left$(tmRvf.sTranType, 1) = "P" And tlTranType.iPymt) Then
                        If (tlTranType.iNTR) Then       'NTR option, tested separately because it shouldnt be tested with Cash transactions
                            If tmRvf.iMnfItem > 0 Then
                                tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                            Else            'its not an NTR
                                'got valid trans type - test for Cash, Trade, Merchandising or Promotions
                                '05-26-04 dont include NTR, exclude this if it is
                                If tmRvf.iMnfItem = 0 And tmRvf.lSbfCode = 0 Then       'not NTR
                                    If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                        tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                        ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                                    End If
                                Else
                                    'must be an installment record, so it should be included
                                    'If trans is an NTR, it must have the item type and pointer to SBF; otherwise its assumed to be an installment
                                    'if it has an SBF pointer only
                                    If tmRvf.iMnfItem = 0 And tmRvf.lSbfCode > 0 Then
                                        If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                            tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                            ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                                        End If
                                    End If

                                End If
                            End If
                        Else
                            'got valid trans type - test for Cash, Trade, Merchandising or Promotions
                            If (tmRvf.sCashTrade = "C" And tlTranType.iCash) Or (tmRvf.sCashTrade = "T" And tlTranType.iTrade) Or (tmRvf.sCashTrade = "M" And tlTranType.iMerch) Or (tmRvf.sCashTrade = "P" And tlTranType.iPromo) Then
                                tlRvf(UBound(tlRvf)) = tmRvf           'save entire record
                                ReDim Preserve tlRvf(ilLowLimit To UBound(tlRvf) + 1) As RVF
                            End If
                        End If
                    End If
                    ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmRvf, tmRvf, ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
        If ilWhichFile = 3 And ilLoop = 1 Then                           'if 1, then just finished history, go do Receivables
            btrExtClear hmRvf   'Clear any previous extend operation
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
            ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmRvf)
                btrDestroy hmRvf
                gObtainPhfOrRvf = False
                Exit Function
            End If
            imRvfRecLen = Len(tmRvf)
            'ilRVFUpper = UBound(tlRvf)     1-11-18 removed, unused .  causing overflow error

        End If
    Next ilLoop
    ilRet = btrClose(hmRvf)
    btrDestroy hmRvf
    gObtainPhfOrRvf = True
    Exit Function
gObtainPhfOrRvfErr:
    ilRet = 1
    Resume Next
mRvfErr:
    On Error GoTo 0
    gObtainPhfOrRvf = False
    Exit Function
End Function

'           gObtainFSF - obtain Feed spots (FSF) for selected dates/times
'           and build into array
'
Public Function gObtainFSF(RptForm As Form, hlFsf, tlFsf() As FSF, slStartDate As String, slEndDate As String, slStartTime As String, slEndTime As String) As Integer
'
'    gObtainFSF (hlFSF, slStartDate, slEndDate, tlFSF())
'           <input>  RptForm - form name source
'                    hlFsf - FSF handle
'                    slStartDate - earliest date to gather spots
'                    slEndDate - latest date to gather spots
'                    slStartTime - earliest time to gather spots
'                    slEndTime - late time to gather spots
'           <output> tlFSF() array of FSF records
'           <return> true if valid reads
'
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim slAlteredEndTime As String      'altered to 11:59:59P if 12M end of day
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record

    slAlteredEndTime = slEndTime
    If Trim$(slEndTime) = "12M" Then
        slAlteredEndTime = "11:59:59PM"
    End If
    ReDim tlFsf(0 To 0) As FSF
    btrExtClear hlFsf   'Clear any previous extend operation
    ilExtLen = Len(tlFsf(0))  'Extract operation record size
    imFsfRecLen = Len(tlFsf(0))

    ilRet = btrGetFirst(hlFsf, tmFsf, imFsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hlFsf, llNoRec, -1, "UC", "FSF", "") '"EG") 'Set extract limits (all records)

        gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Fsf", "FsfEndDate")
        ilRet = btrExtAddLogicConst(hlFsf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

        gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Fsf", "FsfStartDate")
        ilRet = btrExtAddLogicConst(hlFsf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)


        gPackTime slStartTime, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Fsf", "FsfEndTime")
        ilRet = btrExtAddLogicConst(hlFsf, BTRV_KT_TIME, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

        gPackTime slAlteredEndTime, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Fsf", "FsfStartTime")
        ilRet = btrExtAddLogicConst(hlFsf, BTRV_KT_TIME, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        ilRet = btrExtAddField(hlFsf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainFSFErr
        gBtrvErrorMsg ilRet, "gObtainFSF (btrExtAddField):" & "FSF.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlFsf, tmFsf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainFSFErr
            gBtrvErrorMsg ilRet, "gObtainFSF (btrExtGetNextExt):" & "FSF.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmFsf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlFsf, tmFsf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tlFsf(UBound(tlFsf)) = tmFsf           'save entire record
                ReDim Preserve tlFsf(0 To UBound(tlFsf) + 1) As FSF
                ilRet = btrExtGetNext(hlFsf, tmFsf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlFsf, tmFsf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainFSF = True
    Exit Function
mObtainFSFErr:
    On Error GoTo 0
    MsgBox "RptExtra: gObtainFSF error", vbCritical + vbOkOnly, "FSF I/O Error"
    gObtainFSF = False
    Exit Function
End Function

Public Function gObtainFPF(RptForm As Form, hlFpf, tlFPF() As FPF, slStartDate As String, slEndDate As String, ilFnfCode As Integer, ilVefCode As Integer) As Integer
'
'    gObtainFPF (hlFPF, slStartDate, slEndDate, tlFPF(), ilFnfCode, ilVefCode)
'           <input>  RptForm - form name source
'                    hlFPF - FPF handle
'                    slStartDate - earliest date to gather spots
'                    slEndDate - latest date to gather spots
'                    ilFnfCode - selective feed name or minus 1 (-1) if all
'                    ilVefCode - selective vehicle or minus 1 (-1) if all
'           <output> tlFPF() array of FPF records
'           <return> true if valid reads
'
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlIntTypeBuff As POPICODE

    ReDim tlFPF(0 To 0) As FPF
    btrExtClear hlFpf   'Clear any previous extend operation
    ilExtLen = Len(tlFPF(0))  'Extract operation record size
    imFpfRecLen = Len(tlFPF(0))

    ilRet = btrGetFirst(hlFpf, tmFpf, imFpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hlFpf, llNoRec, -1, "UC", "FPF", "") '"EG") 'Set extract limits (all records)


        If ilFnfCode > 0 Then           'check for selective feed names (negative indicates all feed names)
            tlIntTypeBuff.iCode = ilFnfCode
            ilOffSet = gFieldOffset("FPF", "FPFFnfCode")
            ilRet = btrExtAddLogicConst(hlFpf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        End If

        If ilVefCode > 0 Then           'check for selective vehicles (negative indicates all vehicle feed names)
            tlIntTypeBuff.iCode = ilVefCode
            ilOffSet = gFieldOffset("FPF", "FPFVefCode")
            ilRet = btrExtAddLogicConst(hlFpf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        End If

        gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("FPF", "FPFEffEndDate")
        ilRet = btrExtAddLogicConst(hlFpf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

        gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("FPF", "FPFEffStartDate")
        ilRet = btrExtAddLogicConst(hlFpf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        ilRet = btrExtAddField(hlFpf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainFPFErr
        gBtrvErrorMsg ilRet, "gObtainFPF (btrExtAddField):" & "FPF.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlFpf, tmFpf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainFPFErr
            gBtrvErrorMsg ilRet, "gObtainFPF (btrExtGetNextExt):" & "FPF.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmFpf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlFpf, tmFpf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tlFPF(UBound(tlFPF)) = tmFpf           'save entire record
                ReDim Preserve tlFPF(0 To UBound(tlFPF) + 1) As FPF
                ilRet = btrExtGetNext(hlFpf, tmFpf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlFpf, tmFpf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainFPF = True
    Exit Function
mObtainFPFErr:
    On Error GoTo 0
    MsgBox "RptExtra: gObtainFPF error", vbCritical + vbOkOnly, "FPF I/O Error"
    gObtainFPF = False
    Exit Function
End Function

'           gObtainFdfByCode -obtain all matching detail pledge records for
'           a given Feed Pledge Header
'
Public Function gObtainFDFByCode(RptForm As Form, hlFDF As Integer, tlFdf() As FDF, ilFpfCode As Integer) As Integer
'
'    gObtainFDF (hlFDF, tlFDF(), ilFpfCode)
'           <input>  RptForm - form name source
'                    hlFDF - FDF handle
'                    ilFpfCode - pledge header to retrieve all detail data
'           <output> tlFDF() array of FDF records
'           <return> true if valid reads
'
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim tlIntTypeBuff As POPICODE

    ReDim tlFdf(0 To 0) As FDF
    btrExtClear hlFDF   'Clear any previous extend operation
    ilExtLen = Len(tlFdf(0))  'Extract operation record size
    imFdfRecLen = Len(tlFdf(0))

    ilRet = btrGetFirst(hlFDF, tmFdf, imFdfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hlFDF, llNoRec, -1, "UC", "FDF", "") '"EG") 'Set extract limits (all records)
        tlIntTypeBuff.iCode = ilFpfCode
        ilOffSet = gFieldOffset("FDF", "FDFFpfCode")
        ilRet = btrExtAddLogicConst(hlFDF, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)

        ilRet = btrExtAddField(hlFDF, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainFDFErr
        gBtrvErrorMsg ilRet, "gObtainFDF (btrExtAddField):" & "FDF.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlFDF, tmFdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainFDFErr
            gBtrvErrorMsg ilRet, "gObtainFDF (btrExtGetNextExt):" & "FDF.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmFdf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlFDF, tmFdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tlFdf(UBound(tlFdf)) = tmFdf           'save entire record
                ReDim Preserve tlFdf(0 To UBound(tlFdf) + 1) As FDF
                ilRet = btrExtGetNext(hlFDF, tmFdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlFDF, tmFdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainFDFByCode = True
    Exit Function
mObtainFDFErr:
    On Error GoTo 0
    MsgBox "RptExtra: gObtainFDF error", vbCritical + vbOkOnly, "FDF I/O Error"
    gObtainFDFByCode = False
    Exit Function
End Function
Public Sub gClearRsr()
'*******************************************************
'*                                                     *
'*      Procedure Name:Clear Prepass file  RSR
'*
'*                                                     *
'*             Created:01/24/11      By:D.Michaelson   *
'*            Modified:              By:               *
'*                                                     *
'*                                                     *
'*******************************************************
    Dim ilRet As Integer
    Dim hlRsr As Integer
    Dim ilRsrRecLen As Integer
    Dim tlRsr As RSR
    Dim tlRsrSrchKey As RSRKEY0
    
    hlRsr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlRsr, "", sgDBPath & "rsr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlRsr)
        btrDestroy hlRsr
        Exit Sub
    End If
    ilRsrRecLen = Len(tlRsr)
    tlRsrSrchKey.iGenDate(0) = igNowDate(0)
    tlRsrSrchKey.iGenDate(1) = igNowDate(1)

    tlRsrSrchKey.lGenTime = lgNowTime       '10-20-01
    ilRet = btrGetGreaterOrEqual(hlRsr, tlRsr, ilRsrRecLen, tlRsrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlRsr.iGenDate(0) = igNowDate(0)) And (tlRsr.iGenDate(1) = igNowDate(1)) And (tlRsr.lGenTime = lgNowTime)
        ilRet = btrDelete(hlRsr)
        ilRet = btrGetNext(hlRsr, tlRsr, ilRsrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hlRsr)
    btrDestroy hlRsr
End Sub

'Public Sub gClearScr()
'*******************************************************
'*                                                     *
'*      Procedure Name:Clear Prepass file for SCR      *
'*                  Dump                               *
'*                                                     *
'*             Created:08/29/05      By:D. Hosaka      *
'*            Modified:10/17/2018    By:FYM            *
'*      Duplicate procedure called gClearScr           *
'*                                                     *
'*                                                     *
'*******************************************************
'    Dim ilRet As Integer
'    hmScr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmScr, "", sgDBPath & "Scr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmScr)
'        btrDestroy hmScr
'        Exit Sub
'    End If
'    imScrRecLen = Len(tmScr)
'    tmScrSrchKey.iGenDate(0) = igNowDate(0)
'    tmScrSrchKey.iGenDate(1) = igNowDate(1)
'
'    tmScrSrchKey.lGenTime = lgNowTime       '10-20-01
'    ilRet = btrGetGreaterOrEqual(hmScr, tmScr, imScrRecLen, tmScrSrchKey, INDEXKEY1, BTRV_LOCK_NONE)
'    Do While (ilRet = BTRV_ERR_NONE) And (tmScr.iGenDate(0) = igNowDate(0)) And (tmScr.iGenDate(1) = igNowDate(1)) And (tmScr.lGenTime = lgNowTime)
'        ilRet = btrDelete(hmScr)
'        ilRet = btrGetNext(hmScr, tmScr, imScrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
'    Loop
'    ilRet = btrClose(hmScr)
'    btrDestroy hmScr
'End Sub

Public Sub gClearODF()
'*******************************************************
'*                                                     *
'*      Procedure Name:gClearODF                       *
'*                                                     *
'*         Created:09/29/05      By:D. Hosaka          *
'*         Modified:              By:                  *
'*                                                     *
'*         Comments:Clear OneDay Log file by gen date  *
'*                     and time                        *
'*                                                     *
'*******************************************************
    Dim ilRet As Integer
    hmOdf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmOdf, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmOdf)
        btrDestroy hmOdf
        Exit Sub
    End If
    imOdfRecLen = Len(tmOdf)
    tmOdfSrchKey2.iGenDate(0) = igNowDate(0)
    tmOdfSrchKey2.iGenDate(1) = igNowDate(1)
    '10-10-01
    tmOdfSrchKey2.lGenTime = lgNowTime
    ilRet = btrGetGreaterOrEqual(hmOdf, tmOdf, imOdfRecLen, tmOdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmOdf.iGenDate(0) = igNowDate(0)) And (tmOdf.iGenDate(1) = igNowDate(1)) And (tmOdf.lGenTime = lgNowTime)
        ilRet = btrDelete(hmOdf)
        'change way to remove records to avoid losing positioning
        'ilRet = btrGetNext(hmOdf, tmOdf, imOdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        tmOdfSrchKey2.iGenDate(0) = igNowDate(0)
        tmOdfSrchKey2.iGenDate(1) = igNowDate(1)
        tmOdfSrchKey2.lGenTime = lgNowTime
        ilRet = btrGetGreaterOrEqual(hmOdf, tmOdf, imOdfRecLen, tmOdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
    Loop
    ilRet = btrClose(hmOdf)
    btrDestroy hmOdf
End Sub

Public Function gObtainCrfByDate(RptForm As Form, hlCrf, tlCrf() As CRF, slActiveDate As String, slActiveEndDate As String) As Integer
'
'    gObtainCRF (hlCRF, slActiveDate,  tlCRF())
'           <input>  RptForm - form name source
'                    hlCrf - Copy Rotation Header handle
'                    slActiveDate - Active date to gather rotation
'                    slActiveEndDate - active end date span to gather rotation
'           <output> tlCRF() array of CRF records
'           <return> true if valid reads
'
'       8-3-10 change from effective date to active start/end date span
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim slDate As String

    btrExtClear hlCrf   'Clear any previous extend operation
    ilExtLen = Len(tlCrf(0))  'Extract operation record size
    imCrfRecLen = Len(tlCrf(0))

    ilRet = btrGetFirst(hlCrf, tmCrf, imCrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hlCrf, llNoRec, -1, "UC", "CRF", "") '"EG") 'Set extract limits (all records)

'        gPackDate slActiveDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
'        ilOffset = gFieldOffset("Crf", "CrfEndDate")
'        ilRet = btrExtAddLogicConst(hlCrf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
         
         ' crfEndDate >= InputStartDate And crfStartDate <= InputEndDate
        '8-3-10 implement span of dates vs just an effective date
        If slActiveDate = "" Then
            slDate = "1/1/1970"
        Else
            slDate = slActiveDate
        End If
        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Crf", "CrfEndDate")
        ilRet = btrExtAddLogicConst(hlCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        If slActiveEndDate = "" Then
            slDate = "12/31/2069"
        Else
            slDate = slActiveEndDate
        End If

        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Crf", "CrfStartDate")
        ilRet = btrExtAddLogicConst(hlCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        ilRet = btrExtAddField(hlCrf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainCrfErr
        gBtrvErrorMsg ilRet, "gObtainCrf (btrExtAddField):" & "Crf.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlCrf, tmCrf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainCrfErr
            gBtrvErrorMsg ilRet, "gObtainCrf (btrExtGetNextExt):" & "Crf.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmCrf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlCrf, tmCrf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tlCrf(UBound(tlCrf)) = tmCrf           'save entire record
                ReDim Preserve tlCrf(0 To UBound(tlCrf) + 1) As CRF
                ilRet = btrExtGetNext(hlCrf, tmCrf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlCrf, tmCrf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainCrfByDate = True
    Exit Function
mObtainCrfErr:
    On Error GoTo 0
    MsgBox "RptExtra: gObtainCrf error", vbCritical + vbOkOnly, "Crf I/O Error"
    gObtainCrfByDate = False
    Exit Function
End Function

Public Function gObtainLLFbyDate(RptForm As Form, hlLlf, tlLlf() As LLF, slStartDate As String, slEndDate As String) As Integer
'
'    gObtainLlf (hlLlf, slStartDate, slEndDate, tlLlf())
'           <input>  RptForm - form name source
'                    hlLlf - Llf handle
'                    slStartDate - earliest date to gather spots
'                    slEndDate - latest date to gather spots
'           <output> tlLlf() array of Llf records
'           <return> true if valid reads
'
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record

    ReDim tlLlf(0 To 0) As LLF
    btrExtClear hlLlf   'Clear any previous extend operation
    ilExtLen = Len(tlLlf(0))  'Extract operation record size
    imLlfRecLen = Len(tlLlf(0))

    ilRet = btrGetFirst(hlLlf, tmLlf, imLlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hlLlf, llNoRec, -1, "UC", "Llf", "") '"EG") 'Set extract limits (all records)

        gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Llf", "LlfAirDate")
        ilRet = btrExtAddLogicConst(hlLlf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

        gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Llf", "LlfAirDate")
        ilRet = btrExtAddLogicConst(hlLlf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        ilRet = btrExtAddField(hlLlf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainLlfErr
        gBtrvErrorMsg ilRet, "gObtainLlfByDate (btrExtAddField):" & "Llf.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlLlf, tmLlf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainLlfErr
            gBtrvErrorMsg ilRet, "gObtainLlfByDate (btrExtGetNextExt):" & "Llf.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmLlf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlLlf, tmLlf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tlLlf(UBound(tlLlf)) = tmLlf           'save entire record
                ReDim Preserve tlLlf(0 To UBound(tlLlf) + 1) As LLF
                ilRet = btrExtGetNext(hlLlf, tmLlf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlLlf, tmLlf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainLLFbyDate = True
    Exit Function
mObtainLlfErr:
    On Error GoTo 0
    MsgBox "RptExtra: gObtainLlfByDate error", vbCritical + vbOkOnly, "Llf I/O Error"
    gObtainLLFbyDate = False
    Exit Function
End Function

'           gobtainRAFByType - obtain all applicable Region records by option
'           and build in returned array
'       <input> ilInclRegCopy - true/false to return regional copy
'               ilInclSplitNet - true/false to return split network regions
'               ilInclSplitCopy - true/false to return split copy regions
'       <return> array of RAF records that meet user option
'
Public Function gObtainRAFByType(RptForm As Form, hlRaf As Integer, tlRaf() As RAF, ilInclRegCopy As Integer, ilInclSplitNet As Integer, ilInclSplitCopy As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilOffset                      tlDateTypeBuff                tlIntTypeBuff             *
'*                                                                                        *
'******************************************************************************************
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long

    ReDim tlRaf(0 To 0) As RAF
    btrExtClear hlRaf   'Clear any previous extend operation
    ilExtLen = Len(tlRaf(0))  'Extract operation record size
    imRafRecLen = Len(tlRaf(0))

    ilRet = btrGetFirst(hlRaf, tmRaf, imRafRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hlRaf, llNoRec, -1, "UC", "RAF", "") '"EG") 'Set extract limits (all records)

        ilRet = btrExtAddField(hlRaf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainRAFErr
        gBtrvErrorMsg ilRet, "gObtainRAF (btrExtAddField):" & "RAF.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlRaf, tmRaf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainRAFErr
            gBtrvErrorMsg ilRet, "gObtainRAF (btrExtGetNextExt):" & "RAF.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmRaf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlRaf, tmRaf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If (tmRaf.sType = "R" And ilInclRegCopy) Or (tmRaf.sType = "N" And ilInclSplitNet) Or (tmRaf.sType = "C" And ilInclSplitCopy) Then
                    tlRaf(UBound(tlRaf)) = tmRaf           'save entire record
                    ReDim Preserve tlRaf(0 To UBound(tlRaf) + 1) As RAF
                End If
                ilRet = btrExtGetNext(hlRaf, tmRaf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlRaf, tmRaf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainRAFByType = True
    Exit Function
mObtainRAFErr:
    On Error GoTo 0
    MsgBox "RptExtra: gObtainRAF error", vbCritical + vbOkOnly, "RAF I/O Error"
    gObtainRAFByType = False
    Exit Function
End Function

'       gBuildStationsByIntCategory - build array of station codes that belong to a regions
'       by category name
'       <input> tlSort() array to search, stored stations by market or owner category
'               ilCode - SEF category
'       <output> array of station integers
'       <return> true = found entry
'
Public Function gBuildStationsByIntCategory(tlSort() As REGIONINTSORT, ilCode As Integer, tlStationsClf() As INTKEY0) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim ilresult As Integer
    Dim ilLoop As Integer

    ilresult = -1
    ilMin = LBound(tlSort)
    ilMax = UBound(tlSort) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tlSort(ilMiddle).iIntCode Then
            'found the match
            ilresult = ilMiddle
            Exit Do
        ElseIf ilCode < tlSort(ilMiddle).iIntCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    If ilresult <> -1 Then
        For ilLoop = ilresult To UBound(tlSort) - 1 Step 1
            If tlSort(ilLoop).iIntCode = ilCode Then
                tlStationsClf(UBound(tlStationsClf)).iCode = tlSort(ilLoop).iShttCode
                ReDim Preserve tlStationsClf(LBound(tlStationsClf) To UBound(tlStationsClf) + 1) As INTKEY0
            Else
                Exit For
            End If
        Next ilLoop
        For ilLoop = ilresult - 1 To LBound(tlSort) Step -1
            If tlSort(ilLoop).iIntCode = ilCode Then
                tlStationsClf(UBound(tlStationsClf)).iCode = tlSort(ilLoop).iShttCode
                ReDim Preserve tlStationsClf(LBound(tlStationsClf) To UBound(tlStationsClf) + 1) As INTKEY0
            Else
                Exit For
            End If
        Next ilLoop
        gBuildStationsByIntCategory = True
    Else
        gBuildStationsByIntCategory = False
    End If
End Function

'
'       gBuildStationsByStrCategory - build array of station codes that belong to a regions
'       by category name
'       <input> tlSort() array to search, stored stations by zipcode or state category
'               ilCode - SEF category
'       <output> array of station integers
'       <return> true = found entry
'
Public Function gBuildStationsByStrCategory(tlSort() As REGIONSTRSORT, slInCode As String, tlStationsClf() As INTKEY0) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim slCode As String
    Dim ilresult As Integer
    Dim ilLoop As Integer

    slCode = Trim$(slInCode)
    ilresult = -1
    ilMin = LBound(tlSort)
    ilMax = UBound(tlSort) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If StrComp(slCode, Trim$(tlSort(ilMiddle).sStr), vbTextCompare) = 0 Then
            'found the match
            ilresult = ilMiddle
            Exit Do
        ElseIf StrComp(slCode, Trim$(tlSort(ilMiddle).sStr), vbTextCompare) < 0 Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    If ilresult <> -1 Then
        For ilLoop = ilresult To UBound(tlSort) - 1 Step 1
            If StrComp(slCode, Trim$(tlSort(ilLoop).sStr), vbTextCompare) = 0 Then
                tlStationsClf(UBound(tlStationsClf)).iCode = tlSort(ilLoop).iShttCode
                ReDim Preserve tlStationsClf(LBound(tlStationsClf) To UBound(tlStationsClf) + 1) As INTKEY0
            Else
                Exit For
            End If
        Next ilLoop
        For ilLoop = ilresult - 1 To LBound(tlSort) Step -1
            If StrComp(slCode, Trim$(tlSort(ilLoop).sStr), vbTextCompare) = 0 Then
                tlStationsClf(UBound(tlStationsClf)).iCode = tlSort(ilLoop).iShttCode
                ReDim Preserve tlStationsClf(LBound(tlStationsClf) To UBound(tlStationsClf) + 1) As INTKEY0
            Else
                Exit For
            End If
        Next ilLoop
        gBuildStationsByStrCategory = True
    Else
        gBuildStationsByStrCategory = False
    End If
End Function

'       Format the selection by Generation date and time from GRF
'       for Crystal report filter
'       <input> none
'       <output> none
'       <return>  selection string for GRF prepass
'
Public Function gGRFSelectionForCrystal() As String
Dim slDate As String
Dim slTime As String
Dim slMonth As String
Dim slDay As String
Dim slYear As String
Dim slSelection As String

        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        gGRFSelectionForCrystal = slSelection
End Function

'       Format the selection by Generation date and time from GRF
'       for Crystal report filter using a Random Date and Time
'       <input> none
'       <output> none
'       <return>  selection string for GRF prepass
'                 sets igNowDate() , igNowTime() , lgNowTime
'       <comment> For TTP 10077 -Spots by Advertiser Report Speed-up for muti-users running report
Public Function gGRFSelectionForCrystalRandom() As String
    Dim slDate As String
    Dim slTime As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim slSelection As String

    gRandomDateTime slDate, slTime, slMonth, slDay, slYear
    slSelection = "{GRF_Generic_Report.grfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
    slSelection = slSelection & " And Round({GRF_Generic_Report.grfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
    gGRFSelectionForCrystalRandom = slSelection
End Function

'       Clear combined Air Time and NTR prepass file
'       01-18-07
Public Sub gIMRClear()
    Dim ilRet As Integer
    Dim tlImr As IMR
    hmImr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmImr, "", sgDBPath & "Imr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmImr)
        btrDestroy hmImr
        Exit Sub
    End If
    imImrRecLen = Len(tlImr)
    tmImrSrchKey.iGenDate(0) = igNowDate(0)
    tmImrSrchKey.iGenDate(1) = igNowDate(1)
    'tmImrSrchKey.iGenTime(0) = igNowTime(0)
    'tmImrSrchKey.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmImrSrchKey.lGenTime = lgNowTime
    ilRet = btrGetGreaterOrEqual(hmImr, tlImr, imImrRecLen, tmImrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    'Do While (ilRet = BTRV_ERR_NONE) And (tlImr.iGenDate(0) = igNowDate(0)) And (tlImr.iGenDate(1) = igNowDate(1)) And (tlImr.iGenTime(0) = igNowTime(0)) And (tlImr.iGenTime(1) = igNowTime(1))
    Do While (ilRet = BTRV_ERR_NONE) And (tlImr.iGenDate(0) = igNowDate(0)) And (tlImr.iGenDate(1) = igNowDate(1)) And (tlImr.lGenTime = lgNowTime)
        ilRet = btrDelete(hmImr)
        ilRet = btrGetNext(hmImr, tlImr, imImrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmImr)
    btrDestroy hmImr
End Sub

'           gObtainSDFByKey - retrieve Spots (SDF) by Vehicle, Date (key1) or
'           Date, chfCode (Kye4), or Single Contract (keyy5)
'           <input> hlSdf - SDF file handle
'                   ilVefCode - if 0, all vehicles
'                   slStartDate as string
'                   slEndDate as string
'                   llChfCode as long - selective contract code if applicable, else 0
'                   ilWhichKey = INDEXKEY1, INDEXKEY4, INDEXKEY5
'                               Key1:  vehicle, date, time, sch status
'                               Key5: ChfCode
'                   tlSpottypes - sched types to include/exclude
'           <output>  llSdfCodes() - array of spot codes gathered from vehicle or
'                     tlSdfInfo() - array of spot records along with a sort key to sort by line ID for single contract
Public Sub gObtainSDFByKey(hlSdf As Integer, ilVefCode As Integer, slStartDate As String, slEndDate As String, llChfCode As Long, ilWhichKey As Integer, llSdfCodes() As Long, tlSdfInfo() As SDFSORTBYLINE, tlSpotTypes As SPOTTYPES)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilOk                          ilFound                       ilLoop                    *
'*  ilIndex                       slPrice                       slSpotAmount              *
'*  llSpotAmount                  slSpotOrNTR                                             *
'******************************************************************************************
    Dim llDate As Long
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE      '7-19-04
    ReDim llSdfCodes(0 To 0) As Long
    ReDim tlSdfInfo(0 To 0) As SDFSORTBYLINE       'setup sort key and maintain entire recd
    Dim llUpperSDF As Long
    Dim slStr As String
    Dim ilInclude As Integer

    llUpperSDF = 0              'init the # records retreived from SDF
    btrExtClear hlSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hlSdf   'Clear any previous extend operation

    If ilWhichKey = INDEXKEY1 Then       'by vehicle
        tmSdfSrchKey1.iVefCode = ilVefCode
        gPackDate slStartDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
        tmSdfSrchKey1.iTime(0) = 0
        tmSdfSrchKey1.iTime(1) = 0
        tmSdfSrchKey1.sSchStatus = ""
        ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, Len(tmSdf), tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    ElseIf ilWhichKey = INDEXKEY5 Then      'by selective contract
        tmSdfSrchKey3.lCode = llChfCode          'selective contracts within date
        ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, Len(tmSdf), tmSdfSrchKey3, INDEXKEY5, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    ElseIf ilWhichKey = INDEXKEY4 Then      'by date, chfcode
        gPackDate slStartDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
        tmSdfSrchKey4.lChfCode = llChfCode
        ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, Len(tmSdf), tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Else                    'invalid, exit
        Exit Sub
    End If

    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)

        If ilWhichKey = INDEXKEY1 Then
            tlIntTypeBuff.iType = ilVefCode

            '7-19-04 exclude network spots in the extended read, indicated by the SDFCode having a value
            ilOffSet = gFieldOffset("Sdf", "sdfCode")
            tlLongTypeBuff.lCode = 0
            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GT, BTRV_EXT_AND, tlLongTypeBuff, 4)

            ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        ElseIf ilWhichKey = 5 Then                           'selective contract
            ilOffSet = gFieldOffset("Sdf", "sdfChfCode")
            tlLongTypeBuff.lCode = llChfCode                    'match on selective contract code
            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)
        Else
            ilOffSet = gFieldOffset("Sdf", "sdfChfCode")
            tlLongTypeBuff.lCode = llChfCode                    'match on selective contract code
            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GT, BTRV_EXT_AND, tlLongTypeBuff, 4)
        End If

        'all options require date test
        If slStartDate <> "" Then
            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Sdf", "SdfDate")
            If slEndDate <> "" Then
                ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            End If
        End If
        If slEndDate <> "" Then
            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Sdf", "SdfDate")
            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        End If
        ilRet = btrExtAddField(hlSdf, 0, ilExtLen)  'Extract Name
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE

                'see if this is a spot type to include
                ilInclude = True
                If tmSdf.sSpotType = "X" And tlSpotTypes.iFill = False Then
                    ilInclude = False
                End If
                
                '1-18-11 open/close tests
                If tmSdf.sSpotType = "O" And tlSpotTypes.iOpen = False Then
                    ilInclude = False
                End If
                If tmSdf.sSpotType = "C" And tlSpotTypes.iClose = False Then
                    ilInclude = False
                End If
                
                If tmSdf.sSchStatus = "S" And tlSpotTypes.iSched = False Then
                    ilInclude = False
                End If
                If tmSdf.sSchStatus = "M" And tlSpotTypes.iMissed = False Then
                    ilInclude = False
                End If

                If tmSdf.sSchStatus = "G" And tlSpotTypes.iMG = False Then
                    ilInclude = False
                End If
                If tmSdf.sSchStatus = "O" And tlSpotTypes.iOutside = False Then
                    ilInclude = False
                End If
                If tmSdf.sSchStatus = "C" And tlSpotTypes.iCancel = False Then
                    ilInclude = False
                End If
                If tmSdf.sSchStatus = "H" And tlSpotTypes.iHidden = False Then
                    ilInclude = False
                End If

                If ilInclude Then
                    If ilWhichKey = INDEXKEY1 Then      'driven by vehicle, retain just the sdf code rather then entire spot record
                        llSdfCodes(llUpperSDF) = tmSdf.lCode
                        llUpperSDF = llUpperSDF + 1
                        ReDim Preserve llSdfCodes(LBound(llSdfCodes) To llUpperSDF) As Long
                    Else
                        'create key for sorting:  Line # only
                        slStr = Trim$(str$(tmSdf.iLineNo))
                        Do While Len(slStr) < 5
                            slStr = "0" & slStr
                        Loop
                        tlSdfInfo(llUpperSDF).sKey = slStr & "|"

                        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                        slStr = Trim$(str$(llDate))
                        Do While Len(slStr) < 6
                            slStr = "0" & slStr
                        Loop

                        tlSdfInfo(llUpperSDF).sKey = Trim$(tlSdfInfo(llUpperSDF).sKey) & slStr
                        tlSdfInfo(llUpperSDF).tSdf = tmSdf
                        llUpperSDF = llUpperSDF + 1
                        ReDim Preserve tlSdfInfo(LBound(tlSdfInfo) To llUpperSDF) As SDFSORTBYLINE
                    End If
                End If

                ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If              'ilRet <> BTRV_ERR_END_OF_FILE

    Exit Sub
End Sub

'         gObtainIBFByKey - retrieve Impressions records using IBF dB indexes
'           <input> hlSdf - SDF file handle
'                   ilWhichKey = INDEXKEY0, INDEXKEY1, INDEXKEY2, INDEXKEY3
'                               Key0: ivfcode
'                               Key1: cntrNo, podCPMID
'                               Key2: cntrNo, VefCode
'                               Key3: billYear, BillMonth
'
'                   ilBillYear -
'                   ilBillMonth -
'                   ilVefCode -
'                   llContrNo -
'                   ilIbfCode -
'                   ilIbfPodCPMID -
'
'         <output>  llIbfCodes() - array of [Index,IBF record(s]) gathered
Public Sub gObtainIBFByKey(hlIbf As Integer, ilWhichKey As Integer, ilBillYear As Integer, ilBillMonth As Integer, ilVefCode As Integer, llContrNo As Long, ilIbfCode As Integer, ilIbfPodCPMID As Integer, tlIbfInfo() As IBFSORTBYLINE)
    Dim llDate As Long
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE      '7-19-04
    ReDim tlIbfInfo(0 To 0) As IBFSORTBYLINE
    Dim llUpperIBF As Long
    Dim slStr As String
    Dim ilInclude As Integer

    llUpperIBF = 0              'init the # records retreived from Ibf
    btrExtClear hlIbf   'Clear any previous extend operation
    ilExtLen = Len(tmIbf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlIbf) 'Obtain number of records
    btrExtClear hlIbf   'Clear any previous extend operation

    '-----------------------------
    'Get first record as starting point of extend operation
    If ilWhichKey = 0 Then
        'key0: not used or tested yet...
        'tmIbfSrchKey0 = ibfCode
        tmIbfSrchKey0.lCode = ilIbfCode
        ilRet = btrGetGreaterOrEqual(hlIbf, tmIbf, Len(tmIbf), tmIbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
    ElseIf ilWhichKey = 1 Then
        'key1: not used or tested yet...
        'tmIbfSrchKey1 = contrNo, podCPMID
        tmIbfSrchKey1.lCntrNo = llContrNo
        tmIbfSrchKey1.iPodCPMID = ilIbfPodCPMID
        ilRet = btrGetGreaterOrEqual(hlIbf, tmIbf, Len(tmIbf), tmIbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    ElseIf ilWhichKey = 2 Then
        'key2: not used or tested yet...
        'tmIbfSrchKey2 = contrNo, vefCode
        tmIbfSrchKey2.lCntrNo = llContrNo
        tmIbfSrchKey2.iVefCode = ilVefCode
        ilRet = btrGetGreaterOrEqual(hlIbf, tmIbf, Len(tmIbf), tmIbfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
    ElseIf ilWhichKey = 3 Then
        'tmIbfSrchKey3=billYear, billMonth
        tmIbfSrchKey3.iBillYear = right(str(ilBillYear), 2)
        tmIbfSrchKey3.iBillMonth = ilBillMonth
        ilRet = btrGetGreaterOrEqual(hlIbf, tmIbf, Len(tmIbf), tmIbfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
    Else
        'invalid, exit
        Exit Sub
    End If
    
    '-----------------------------
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hlIbf, llNoRec, -1, "UC", "IBF", "") 'Set extract limits (all records)
'        If ilWhichKey = 0 Then
'            'tmIbfSrchKey0 = ibfCode
'        ElseIf ilWhichKey = 1 Then
'            'tmIbfSrchKey1 = contrNo, podCPMID
'            ilOffSet = gFieldOffset("Ibf", "IbfChfCode")
'            tlLongTypeBuff.lCode = ilIbfCode
'            ilRet = btrExtAddLogicConst(hlIbf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)
'        ElseIf ilWhichKey = 2 Then
'            'tmIbfSrchKey2 = contrNo, vefCode
'        ElseIf ilWhichKey = 3 Then
'            'tmIbfSrchKey3=billYear, billMonth
'        End If

        'all options require date test
'        If slStartDate <> "" Then
'            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
'            ilOffSet = gFieldOffset("Ibf", "IbfDate")
'            If slEndDate <> "" Then
'                ilRet = btrExtAddLogicConst(hlIbf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
'            Else
'                ilRet = btrExtAddLogicConst(hlIbf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
'            End If
'        End If
'        If slEndDate <> "" Then
'            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
'            ilOffSet = gFieldOffset("Ibf", "IbfDate")
'            ilRet = btrExtAddLogicConst(hlIbf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
'        End If
'        ilRet = btrExtAddField(hlIbf, 0, ilExtLen)  'Extract Name
'        If ilRet <> BTRV_ERR_NONE Then
'            Exit Sub
'        End If
        
        ilRet = btrExtGetNext(hlIbf, tmIbf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            Do While ilRet = BTRV_ERR_NONE
                'see if this is a Ibf type to include
                ilRet = btrExtGetNext(hlIbf, tmIbf, ilExtLen, llRecPos)
                tlIbfInfo(UBound(tlIbfInfo())).sKey = tmIbf.lCode & "|" & tmIbf.lCntrNo & "|" & tmIbf.iVefCode
                tlIbfInfo(UBound(tlIbfInfo())).tIbf = tmIbf
                ReDim Preserve tlIbfInfo(LBound(tlIbfInfo()) To UBound(tlIbfInfo()) + 1) As IBFSORTBYLINE
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlIbf, tmIbf, ilExtLen, llRecPos)
                    tlIbfInfo(UBound(tlIbfInfo())).sKey = tmIbf.lCode & "|" & tmIbf.lCntrNo & "|" & tmIbf.iVefCode
                    tlIbfInfo(UBound(tlIbfInfo())).tIbf = tmIbf
                Loop
            Loop
        End If
    End If              'ilRet <> BTRV_ERR_END_OF_FILE
    Exit Sub
End Sub

'
'           Read Sales Offices and build sales office & sales source into array
'           Assume SOF opened
'           <input> hlSof  - SOF file handle
Public Sub gObtainSOF(hlSof As Integer, tlSofList() As SOFLIST)
    Dim tlSof As SOF
    Dim ilTemp As Integer
    Dim ilRet As Integer

    ilTemp = 0
    ilRet = btrGetFirst(hlSof, tlSof, Len(tlSof), INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tlSofList(0 To ilTemp) As SOFLIST
        tlSofList(ilTemp).iSofCode = tlSof.iCode
        tlSofList(ilTemp).iMnfSSCode = tlSof.iMnfSSCode
        ilRet = btrGetNext(hlSof, tlSof, Len(tlSof), BTRV_LOCK_NONE, SETFORREADONLY)
        ilTemp = ilTemp + 1
    Loop
    Exit Sub
End Sub

Public Sub gClearUOR()
    Dim ilRet As Integer
    hmUor = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUor, "", sgDBPath & "Uor.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmUor)
        btrDestroy hmUor
        Exit Sub
    End If
    imUorRecLen = Len(tmUor)
    tmUorSrchKey.iGenDate(0) = igNowDate(0)
    tmUorSrchKey.iGenDate(1) = igNowDate(1)
    'tmUorSrchKey.iGenTime(0) = igNowTime(0)
    'tmUorSrchKey.iGenTime(1) = igNowTime(1)
    tmUorSrchKey.lGenTime = lgNowTime       '10-20-01
    ilRet = btrGetGreaterOrEqual(hmUor, tmUor, imUorRecLen, tmUorSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    'Do While (ilRet = BTRV_ERR_NONE) And (tmUor.iGenDate(0) = igNowDate(0)) And (tmUor.iGenDate(1) = igNowDate(1)) And (tmUor.iGenTime(0) = igNowTime(0)) And (tmUor.iGenTime(1) = igNowTime(1))
    Do While (ilRet = BTRV_ERR_NONE) And (tmUor.iGenDate(0) = igNowDate(0)) And (tmUor.iGenDate(1) = igNowDate(1)) And (tmUor.lGenTime = lgNowTime)
        ilRet = btrDelete(hmUor)
        ilRet = btrGetNext(hmUor, tmUor, imUorRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmUor)
    btrDestroy hmUor
End Sub

'          'Obtain the Delivery and Engineering pre-feed links using key 1 (type, startdate, vehicle
'           5-6-10
Public Function gObtainPFF(RptForm As Form, hlPff, tlPff() As PFF, slStartDate As String, ilCode As Integer, ilVefCode As Integer) As Integer
'
'    gObtainPff (hlPff, slStartDate, slEndDate, tlPff(), ilType, ilVefCode)
'           <input>  RptForm - form name source
'                    hlPff - Pff handle
'                    slStartDate - earliest date to gather spots
'                    ilCode - 1 = delivery, 2 = engineering
'                    ilVefCode -  vehicle code
'           <output> tlPff() array of Pff records
'           <return> true if valid reads
'
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim tlCharTypeBuff As POPCHARTYPE
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlIntTypeBuff As POPICODE
    
    ReDim tlPff(0 To 0) As PFF
    btrExtClear hlPff   'Clear any previous extend operation
    ilExtLen = Len(tlPff(0))  'Extract operation record size
    imPffRecLen = Len(tlPff(0))

    ilRet = btrGetFirst(hlPff, tmPff, imPffRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hlPff, llNoRec, -1, "UC", "Pff", "") '"EG") 'Set extract limits (all records)


        If ilCode = 1 Then           'delivery
            tlCharTypeBuff.sType = "D"
        Else
            tlCharTypeBuff.sType = "E"
        End If
        ilOffSet = gFieldOffset("Pff", "PffType")
        ilRet = btrExtAddLogicConst(hlPff, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)

        tlIntTypeBuff.iCode = ilVefCode
        ilOffSet = gFieldOffset("Pff", "PffVefCode")
        ilRet = btrExtAddLogicConst(hlPff, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)

        gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Pff", "PffStartDate")
        ilRet = btrExtAddLogicConst(hlPff, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        ilRet = btrExtAddField(hlPff, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainPffErr
        gBtrvErrorMsg ilRet, "gObtainPff (btrExtAddField):" & "Pff.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlPff, tmPff, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainPffErr
            gBtrvErrorMsg ilRet, "gObtainPff (btrExtGetNextExt):" & "Pff.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmPff)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlPff, tmPff, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tlPff(UBound(tlPff)) = tmPff           'save entire record
                ReDim Preserve tlPff(0 To UBound(tlPff) + 1) As PFF
                ilRet = btrExtGetNext(hlPff, tmPff, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlPff, tmPff, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainPFF = True
    Exit Function
mObtainPffErr:
    On Error GoTo 0
    MsgBox "RptExtra: gObtainPff error", vbCritical + vbOkOnly, "Pff I/O Error"
    gObtainPFF = False
    Exit Function
End Function

'               Populate Traffic and Affiliate Users (by option)
'               <input> ilWhichUsers: Users type to populate- 0 = all, 1 = traffic only, 2 = affil only
Public Function gPopAllUsers(Form As Form, ilWhichUsers As Integer, lbcCtrl As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
    Dim tlUrf As URF
    Dim tlUst As UST
    Dim ilUrf As Integer
    Dim ilPop As Integer
    Dim slStamp As String
    Dim ilRet As Integer
    Dim ilSortCode As Integer
    Dim slName As String
    Dim slUser As String
    Dim slNameCode As String
    Dim llLen As Long
        
    gPopAllUsers = BTRV_ERR_NONE
    slStamp = gFileDateTime(sgDBPath & "Urf.Btr") & gFileDateTime(sgDBPath & "Ust.mkd")

    On Error GoTo gPopAllUsersErr2
    ilRet = 0
    If ilRet <> 0 Then
        slSortCodeTag = ""
    End If
    On Error GoTo 0

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcCtrl.ListCount > 0 Then
                Exit Function
            End If
            Exit Function
        End If
    End If
    lbcCtrl.Clear
    slSortCodeTag = slStamp

    ilRet = gObtainUrf()
    ilRet = gObtainUst()
    lbcCtrl.Clear
    ilSortCode = 0
    ReDim tlSortCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)

    If ilWhichUsers = 0 Or ilWhichUsers = 1 Then     'all users or just traffic
        For ilUrf = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
            tlUrf = tgPopUrf(ilUrf)
            'ignore counterpoint user, dormant users
            If tlUrf.iCode > 1 And tlUrf.sDelete <> "Y" Then
                slUser = Trim$(tlUrf.sRept)
                If Trim$(slUser) = "" Then
                    slUser = Trim$(tlUrf.sName)
                End If
                    
                slName = Trim$(slUser) & "/Traffic" & "\" & Trim$(str$(tlUrf.iCode))
                tlSortCode(ilSortCode).sKey = slName
                ReDim Preserve tlSortCode(0 To ilSortCode + 1) As SORTCODE
                ilSortCode = ilSortCode + 1
            End If
        Next ilUrf
    End If
    
    If ilWhichUsers = 0 Or ilWhichUsers = 2 Then     'all users or just affiliate
        For ilUrf = LBound(tgUst) To UBound(tgUst) - 1 Step 1
            tlUst = tgUst(ilUrf)
            'ignore dormant and users that cannot view the activity log
            If tlUst.iState = 0 Then
                slUser = Trim$(tlUst.sReportName)
                If Trim$(slUser) = "" Then
                    slUser = Trim$(tlUst.sName)
                End If
                    
                slName = slUser & "/Affiliate" & "\" & Trim$(str$(tlUst.iCode))
                tlSortCode(ilSortCode).sKey = slName
                ReDim Preserve tlSortCode(0 To ilSortCode + 1) As SORTCODE
                ilSortCode = ilSortCode + 1
            End If
        Next ilUrf
    End If
    
    'Sort the traffic and affiliate users by name
    ReDim Preserve tlSortCode(0 To ilSortCode) As SORTCODE
    If UBound(tlSortCode) - 1 > 0 Then
        ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
    End If
    
    For ilUrf = 0 To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(ilUrf).sKey
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopAllUsers = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        llLen = 0
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcCtrl.AddItem slName  'Add ID to list box
    Next ilUrf
    Exit Function

gPopAllUsersErr2:
    ilRet = 1
    Resume Next
    Exit Function
End Function

'               Clear Arf.mkd temporary file
Public Sub gAfrClear()
    Dim ilRet As Integer
    hmAfr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAfr, "", sgDBPath & "Afr.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAfr)
        btrDestroy hmAfr
        Exit Sub
    End If
    imAfrRecLen = Len(tmAfr)
    tmAfrSrchKey.iGenDate(0) = igNowDate(0)
    tmAfrSrchKey.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmAfrSrchKey.lGenTime = lgNowTime
    ilRet = btrGetGreaterOrEqual(hmAfr, tmAfr, imAfrRecLen, tmAfrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmAfr.iGenDate(0) = igNowDate(0)) And (tmAfr.iGenDate(1) = igNowDate(1)) And (tmAfr.lGenTime = lgNowTime)
        ilRet = btrDelete(hmAfr)
        ilRet = btrGetNext(hmAfr, tmAfr, imAfrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmAfr)
    btrDestroy hmAfr

End Sub

'***************************************************************************************
'*
'*      Procedure Name:gGetSpotsByVefDateAndSort - obtain spots from SDF by Key1:
'*                        Vehicle, Date, Time, SchStatus.
'*                        Create array of spots that are sorted by chf, line id,date
'*      <input> hlSdf - SDF handle
'               ilWhichkey:  INDEXKEY0 = vef, chfcode; INDEXKEY1 = vefcode, date
'*              ilVefCode - vehicle code
'               llChfCode = single contract search
'*              slStartDate - earliest date to retrieve spots
'*              slEndDate - latest date to retrieve spots
'               tlSpotTypes() - spot type to include/exclude
'*              tlSdfInfo() - array of spots for the requested span with sortkey
'*      <return> sorted tlsdfinfo() array
'*
'*
'***************************************************************************************
Public Function gGetSpotsbyVefDateAndSort(hlSdf As Integer, ilWhichKey As Integer, ilVefCode As Integer, llChfCode As Long, slStartDate As String, slEndDate As String, tlSpotTypes As SPOTTYPES, tlSdfInfo() As SDFSORTBYLINE) As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilSdfRecLen As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE
    Dim slStr As String

    ReDim tlSdfInfo(0 To 0) As SDFSORTBYLINE
    Dim llUpper As Long
    Dim llDate As Long
    Dim ilOk As Integer

    gGetSpotsbyVefDateAndSort = True
    ilSdfRecLen = Len(tmSdf)
    btrExtClear hlSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)  'Extract operation record size
    
    If ilWhichKey = INDEXKEY1 Then      'by vef, date
        tmSdfSrchKey1.iVefCode = ilVefCode
        gPackDate slStartDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
        tmSdfSrchKey1.iTime(0) = 0
        tmSdfSrchKey1.iTime(1) = 0
        tmSdfSrchKey1.sSchStatus = ""   'slType
        ilSdfRecLen = Len(tmSdf)
        llUpper = 0
        gPackDate slStartDate, tmSdfSrchKey0.iDate(0), tmSdfSrchKey0.iDate(1)
        ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
    ElseIf ilWhichKey = INDEXKEY0 Then          'by contract
        tmSdfSrchKey0.iVefCode = ilVefCode
        gPackDate slStartDate, tmSdfSrchKey0.iDate(0), tmSdfSrchKey0.iDate(1)
        tmSdfSrchKey0.iTime(0) = 0
        tmSdfSrchKey0.iTime(1) = 0
        tmSdfSrchKey0.iLineNo = 0
        tmSdfSrchKey0.lChfCode = llChfCode
        tmSdfSrchKey0.sSchStatus = ""
        ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
    End If
    
    'If (tmSdf.iVefCode = ilVefCode) And (ilRet <> BTRV_ERR_END_OF_FILE) Then

    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Function
        End If
 
        If ilWhichKey = INDEXKEY1 Then
            tlIntTypeBuff.iType = ilVefCode
            ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        Else            'selective contract number
            tlIntTypeBuff.iType = ilVefCode
            ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
            ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, llChfCode, 4)
        End If
            
        ' Prepare to execute an extended operation.
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC", "SDF", "") '"EG") 'Set extract limits (all records)


        gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

        ' And on the records where the date is between the passed  date
        gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        ilRet = btrExtAddField(hlSdf, 0, ilExtLen) 'Extract the whole record
        If ilRet <> BTRV_ERR_NONE Then
            Exit Function
        End If
        ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Function
            End If
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilOk = True
                'filter out cancel and/or missed spots
                If tmSdf.sSchStatus = "S" And Not tlSpotTypes.iSched Then
                    ilOk = False
                ElseIf tmSdf.sSchStatus = "G" And Not tlSpotTypes.iMG Then
                    ilOk = False
                ElseIf tmSdf.sSchStatus = "O" And Not tlSpotTypes.iOutside Then
                    ilOk = False
                ElseIf tmSdf.sSchStatus = "M" And Not tlSpotTypes.iMissed Then
                    ilOk = False
                ElseIf tmSdf.sSchStatus = "C" And Not tlSpotTypes.iCancel Then
                    ilOk = False
                End If
                
                If tmSdf.sSpotType = "O" And Not tlSpotTypes.iOpen Then         'open bb
                    ilOk = False
                ElseIf tmSdf.sSpotType = "C" And Not tlSpotTypes.iClose Then    'close bb
                    ilOk = False
                ElseIf tmSdf.sSpotType = "X" And Not tlSpotTypes.iFill Then     'fills
                    ilOk = False
                End If
                
                If ilOk Then
                    'create key for sorting:  chfcode, line # and date
                    slStr = Trim$(str$(tmSdf.lChfCode))
                    Do While Len(slStr) < 9
                        slStr = "0" & slStr
                    Loop
                    
                    tlSdfInfo(llUpper).sKey = slStr & "|"
                    slStr = Trim$(str$(tmSdf.iLineNo))
                    Do While Len(slStr) < 5
                        slStr = "0" & slStr
                    Loop
                    tlSdfInfo(llUpper).sKey = slStr & "|"
    
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                    slStr = Trim$(str$(llDate))
                    Do While Len(slStr) < 6
                        slStr = "0" & slStr
                    Loop
                    tlSdfInfo(llUpper).sKey = Trim$(tlSdfInfo(llUpper).sKey) & slStr
                    
                    tlSdfInfo(llUpper).tSdf = tmSdf
                    
                    ReDim Preserve tlSdfInfo(0 To UBound(tlSdfInfo) + 1) As SDFSORTBYLINE
                    llUpper = llUpper + 1
                End If
                ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
        
        'sort the array
        If UBound(tlSdfInfo) - 1 > 0 Then
            ArraySortTyp fnAV(tlSdfInfo(), 0), UBound(tlSdfInfo), 0, LenB(tlSdfInfo(0)), 0, LenB(tlSdfInfo(0).sKey), 0
        End If
    End If
End Function

Public Sub gCBFClearWithUserID()
'*******************************************************
'*                                                     *
'*      Procedure Name:Clear the CBF table based
'               on gen date/time & user                *
'*                                                     *
'*             Created:05/29/96      By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear Contract/Proposal  data   *
'*                     for Crystal report              *
'*
'*    2-16-13 created from gCrCBFClear
'*******************************************************
    Dim ilRet As Integer
    Dim llNowTime As Long       '10-10-01 gen time
    hmCbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCbf)
        btrDestroy hmCbf
        Exit Sub
    End If
    '10-10-01
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, llNowTime
    imCbfRecLen = Len(tmCbf)
    tmCbfSrchKey1.iGenDate(0) = igNowDate(0)
    tmCbfSrchKey1.iGenDate(1) = igNowDate(1)
    '10-10-01
    tmCbfSrchKey1.lGenTime = llNowTime
    tmCbfSrchKey1.iUrfCode = tgUrf(0).iCode     '2-17-13
    
    ilRet = btrGetGreaterOrEqual(hmCbf, tmCbf, imCbfRecLen, tmCbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    '10-10-01
    Do While (ilRet = BTRV_ERR_NONE) And (tmCbf.iGenDate(0) = igNowDate(0)) And (tmCbf.iGenDate(1) = igNowDate(1)) And (tmCbf.lGenTime = llNowTime) And tmCbf.iUrfCode = tgUrf(0).iCode
        ilRet = btrDelete(hmCbf)
        'ilRet = btrGetNext(hmCbf, tmCbf, imCbfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        '8-26-13 change way in which deletions processed to avoid losing positioning
        tmCbfSrchKey1.iGenDate(0) = igNowDate(0)
        tmCbfSrchKey1.iGenDate(1) = igNowDate(1)
        tmCbfSrchKey1.lGenTime = llNowTime
        tmCbfSrchKey1.iUrfCode = tgUrf(0).iCode     '2-17-13
        ilRet = btrGetGreaterOrEqual(hmCbf, tmCbf, imCbfRecLen, tmCbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Loop
    ilRet = btrClose(hmCbf)
    btrDestroy hmCbf
End Sub

'           Determine how much a spot length is worth based on 30" units
'           <input> ilSpotLen - Spot length to test
'                   tlSpotLenRatio - table of spot lengths as associated ratios to use based on a 30" rate
'           return - ratio of 30" rate to use (50 represents .5, 100 = 1.00).  If spot length not found because it
'                   higher than the user defined table, calc the # of 30" units based on every 30 = 1 unit
'                   Return it with a negative number so calling routine can determine whether to use it or not
Public Function gDetermineSpotLenRatio(ilSpotLen As Integer, tlSpotLenRatio As SPOTLENRATIO) As Integer
    Dim ilLoop As Integer
    Dim ilLow As Integer
    Dim ilHi As Integer
    Dim ilRatio As Integer
    Dim ilFound As Integer
    
    ilRatio = 0
    ilFound = False
    For ilLoop = 0 To 9     '8 6-1-20 last element missing
        If ilLoop = 0 Then
            ilLow = 0
            ilHi = tlSpotLenRatio.iLen(0)
        End If
        If ilHi = 0 Then            'no more lengths, take the highest ratio
            ilRatio = -(ilSpotLen / 30) * 100     'carry 2 decimal places
            ilFound = True
            Exit For
        Else
            If ilSpotLen >= ilLow And ilSpotLen <= ilHi Then
                ilRatio = tlSpotLenRatio.iRatio(ilLoop)
                ilFound = True
                Exit For
            End If
        End If
        ilLow = ilHi + 1
        ilHi = tlSpotLenRatio.iLen(ilLoop + 1)
    Next ilLoop
     
    gDetermineSpotLenRatio = ilRatio
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:D. Smith       *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Function mOpenMsgFile(slMsgFile As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim ilRet As Integer
    Dim ilERet As Integer

    ilRet = 0
    'On Error GoTo mOpenMsgFileErr:
    slToFile = sgExportPath & "SlsVsPln" & CStr(tgUrf(0).iCode) & ".Txt"

    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        Kill slToFile
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        ilERet = ilRet
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & " error " & str$(ilERet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        ilERet = ilRet
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & " error " & str$(ilERet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, ""
    slMsgFile = slToFile
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilERet = ilRet
'    ilRet = 1
'    Resume Next
End Function

Sub gRvrClear()
'*******************************************************
'*                                                     *
'*      Procedure Name:gRvrClear                     *
'*                                                     *
'*             Created:10/21/96      By:D. Hosaka      *
'*                                                     *
'*            Comments:Clear Receivables Report file   *
'*                     for Crystal report              *
'*                                                     *
'*******************************************************
    Dim ilRet As Integer
    hmRvr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRvr, "", sgDBPath & "Rvr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRvr)
        btrDestroy hmRvr
        Exit Sub
    End If
    imRvrRecLen = Len(tmRvr)
    tmRvrSrchKey.iGenDate(0) = igNowDate(0)
    tmRvrSrchKey.iGenDate(1) = igNowDate(1)
    'tmRvrSrchKey.iGenTime(0) = igNowTime(0)
    'tmRvrSrchKey.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmRvrSrchKey.lGenTime = lgNowTime
    ilRet = btrGetGreaterOrEqual(hmRvr, tmRvr, imRvrRecLen, tmRvrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmRvr.iGenDate(0) = igNowDate(0)) And (tmRvr.iGenDate(1) = igNowDate(1)) And (tmRvr.lGenTime = lgNowTime)
        ilRet = btrDelete(hmRvr)
        ilRet = btrGetNext(hmRvr, tmRvr, imRvrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmRvr)
    btrDestroy hmRvr
End Sub

Function TimeStringtoEnglish(sTimeString As String) As String
    Dim slAnswer As String
    Dim slHour As String
    Dim slMin As String
    Dim ilTemp As Integer
    Dim slTemp As String
    Dim ilPos As Integer
    
    ilPos = InStr(sTimeString, ":") - 1
    
    slHour = Left$(sTimeString, ilPos)
    If CLng(slHour) <> 0 Then
        slAnswer = CLng(slHour) & " hour"
        If CLng(slHour) > 1 Then slAnswer = slAnswer & "s"
        slAnswer = slAnswer & ", "
    End If
    
    slMin = Mid$(sTimeString, ilPos + 2, 2)
    
    ilTemp = slMin
    
    If slMin = "00" Then
       slAnswer = IIF(Len(slAnswer), slAnswer & "0 minutes, and ", "")
    Else
       slTemp = IIF(ilTemp = 1, " minute", " minutes")
       slTemp = IIF(Len(slAnswer), slTemp & ", and ", slTemp & " and ")
       slAnswer = slAnswer & Format$(ilTemp, "##") & slTemp
    End If
    
    ilTemp = Val(right$(sTimeString, 2))
    slMin = Format$(ilTemp, "#0")
    slAnswer = slAnswer & slMin & " second"
    If ilTemp <> 1 Then slAnswer = slAnswer & "s"
    
    TimeStringtoEnglish = slAnswer
End Function

'           gAddDormantVehicle - The vehicle sent is dormant, add to list of dormant vehicles  if not already in list
'                   <input> ilVefCode - internal vehicle code found to be dormant
'                   <input/output> ilDormantVehicles - vehicle code appended to list
'
Public Sub gAddDormantVehicle(ilVefCode As Integer, ilDormantVehicles() As Integer)
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim blFoundDormant As Boolean

    blFoundDormant = False
    ilUpper = UBound(ilDormantVehicles)
    For ilLoop = LBound(ilDormantVehicles) To UBound(ilDormantVehicles) - 1
        If ilDormantVehicles(ilLoop) = ilVefCode Then
            blFoundDormant = True
            Exit For
        End If
                        
    Next ilLoop
    If Not blFoundDormant Then
        ilDormantVehicles(ilUpper) = ilVefCode
        ilUpper = ilUpper + 1
        ReDim Preserve ilDormantVehicles(LBound(ilDormantVehicles) To ilUpper) As Integer
    End If
    Exit Sub
End Sub

'           gBuildDormantVehicles - build listof dormant selling, conventional and game vehicles
'                                  for processing of spot data when not in the selected user list
'                       <input/output>     ilVehiclesToProcess() - list of active vehicles to add the dormant vehicles to
'
Public Sub gBuildDormantVehicles(ilVehiclesToProcess() As Integer)
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1
        If (tgMVef(ilLoop).sState = "D") And (tgMVef(ilLoop).sType = "C" Or tgMVef(ilLoop).sType = "G" Or tgMVef(ilLoop).sType = "S") Then
            ilUpper = UBound(ilVehiclesToProcess)
            ilVehiclesToProcess(ilUpper) = tgMVef(ilLoop).iCode
            ilUpper = ilUpper + 1
            'ReDim Preserve ilVehiclesToProcess(1 To ilUpper) As Integer
            ReDim Preserve ilVehiclesToProcess(0 To ilUpper) As Integer
        End If
    Next ilLoop
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gCRPlayListClear                *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear Play List                 *
'*                     for Crystal report              *
'*                                                     *
'*       7-6-15 moved from rptcrm.bas                  *
'*******************************************************
Sub gCRPlayListClear()
    Dim ilRet As Integer
    hmCpr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpr, "", sgDBPath & "Cpr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCpr)
        btrDestroy hmCpr
        Exit Sub
    End If
    ReDim tmCpr(0 To 0) As CPR
    imCprRecLen = Len(tmCpr(0))
    tmCprSrchKey.iGenDate(0) = igNowDate(0)
    tmCprSrchKey.iGenDate(1) = igNowDate(1)
    'tmCprSrchKey.iGenTime(0) = igNowTime(0)
    'tmCprSrchKey.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmCprSrchKey.lGenTime = lgNowTime
    ilRet = btrGetGreaterOrEqual(hmCpr, tmCpr(0), imCprRecLen, tmCprSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmCpr(0).iGenDate(0) = igNowDate(0)) And (tmCpr(0).iGenDate(1) = igNowDate(1)) And (tmCpr(0).lGenTime = lgNowTime)
        ilRet = btrDelete(hmCpr)
        ilRet = btrGetNext(hmCpr, tmCpr(0), imCprRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    Erase tmCpr
    ilRet = btrClose(hmCpr)
    btrDestroy hmCpr
End Sub
'
'               gFormatRequestDatesforCrystal - format start and end dates to send to crystal for report headers
'               i.e. All Dates or xx/xx/xx - xx/xx/xx, or thru xx/xx/xx or from xx/xx/xx
'               <input>  slStartDate as string
'                        slEndDate as string
'
'               <return> string of requested user dates to show in crystal reports header
Public Function gFormatRequestDatesForCrystal(slDateFrom As String, slDateTo As String) As String
    Dim slStr As String

    If Trim$(slDateFrom) = "" And Trim$(slDateTo) = "" Then             'no dates entered
        slStr = "All Dates"
    ElseIf Trim$(slDateFrom) <> "" And Trim$(slDateTo) <> "" Then       'both dates entered
        slStr = Format$(gDateValue(slDateFrom), "m/d/yy")    'make sure year included
        slStr = slStr & " - " & Format$(gDateValue(slDateTo), "m/d/yy")
    ElseIf Trim$(slDateFrom) = "" Then
        slStr = "Thru " & Format$(gDateValue(slDateTo), "m/d/yy")
    Else
        slStr = "From " & Format$(gDateValue(slDateFrom), "m/d/yy")
    End If
        
    gFormatRequestDatesForCrystal = slStr
    Exit Function
End Function

'           Option set up to exclude certain vehicles, due to more than 1/2 of vehicles included
'           Array set up to exclude the remaining vehicles so its less to test in array
'           When more than half is to be excluded, the dormant vehicles do not get tested for
'           exclusion; therefore that vehicle gets included.  Dormant vehicles are included
'           only when ALL vehicles is selected
'           8-4-17
'
'           <input> ilInclude: true to include vehicles, else false and append any dormant vehicles unless none are in table,
'                   which indicates Exclude Nothing
'                   tlVef() as Vef - array of vehicle records to test
'           <outout> updated array of vehicles to exclude
Public Sub gAddDormantVefToExclList(ilInclude As Integer, tlVef() As VEF, ilCodesToExclude() As Integer)
    Dim ilCount As Integer
    If Not ilInclude Then
        If LBound(ilCodesToExclude) = UBound(ilCodesToExclude) Then         'All vehicles checked on, so include the dormant vehicles which is only way to get reporting on them
            Exit Sub
        End If
        
        For ilCount = LBound(tlVef) To UBound(tlVef) - 1
            If tlVef(ilCount).sState = "D" Then            'dormant, add to list of exclusions
                ilCodesToExclude(UBound(ilCodesToExclude)) = tlVef(ilCount).iCode
                ReDim Preserve ilCodesToExclude(LBound(ilCodesToExclude) To UBound(ilCodesToExclude) + 1) As Integer
            End If
        Next ilCount
    End If
    Exit Sub
End Sub

'           gBinarySearchCallLetters - find the matching vehicle name string in array
'           <input> slCallLetters = vehicle name to match
'           return - index to matching vehicle entry
'                    -1 if not found
Public Function gBinarySearchCallLetters(slCallLetters As String) As Integer
    Dim ilMiddle As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    ilMin = LBound(tgAllStations)
    ilMax = UBound(tgAllStations)
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If StrComp(UCase(Trim$(slCallLetters)), Trim$(tgAllStations(ilMiddle).sCallLetters), vbBinaryCompare) = 0 Then
            'found the match
            gBinarySearchCallLetters = ilMiddle
            Exit Function
        ElseIf StrComp(UCase(Trim$(slCallLetters)), Trim$(tgAllStations(ilMiddle).sCallLetters), vbBinaryCompare) < 0 Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchCallLetters = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainAllStations                 *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgAllStations             *
'*                                                     *
'*******************************************************
Function gObtainAllStations() As Integer
'
'   ilRet = gObtainAllStations ()
'   Where:
'       tgAllStations() (O)- SHTTINFO record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Mnf date/time stamp
    Dim hlShtt As Integer        'Mnf handle
    Dim ilShttRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim tlShtt As SHTT
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperShtt As Integer
    Dim ilUpperVefShtt As Integer
    Dim ilShtt As Integer
    'Dim hlAtt As Integer
    'Dim ilAttRecLen As Integer     'Record length
    'Dim tlAtt As ATT
    'Dim tlAttSrchKey As INTKEY0
    Dim ilAdd As Integer
    Dim ilVef As Integer
    Dim ilLowLimit As Integer
    '8132 Dan moved
'    bgStationAreVehicles = False
    
    'If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) <> SPLITNETWORKS) And ((Asc(tgSpf.sUsingFeatures2) And SPLITCOPY) <> SPLITCOPY) Then
    '    ReDim tgAllStations(1 To 1) As SHTTINFO
    '    gObtainStations = True
    '    Exit Function
    'End If

    slStamp = gFileDateTime(sgDBPath & "Shtt.mkd")

    '11/26/17: Check Changed date/time
'    If Not gFileChgd("shtt.mkd") Then
'        gObtainAllStations = True
'        Exit Function
'    End If

    'On Error GoTo gObtainStationsErr2
    'ilRet = 0
    'ilLowLimit = LBound(tgAllStations)
    'If ilRet <> 0 Then
    '    sgStationsStamp = ""
    'End If
    'On Error GoTo 0
    If PeekArray(tgAllStations).Ptr <> 0 Then
        ilLowLimit = LBound(tgAllStations)
    Else
        sgAllStationsStamp = ""
        ilLowLimit = 0
    End If

    If sgAllStationsStamp <> "" Then
        If StrComp(slStamp, sgAllStationsStamp, 1) = 0 Then
            'If UBound(tgAllStations) > 1 Then
                gObtainAllStations = True
                Exit Function
            'End If
        End If
    End If
    ReDim tgAllStations(0 To 20000) As SHTTINFO
    hlShtt = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlShtt, "", sgDBPath & "Shtt.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAllStations = False
        ilRet = btrClose(hlShtt)
        btrDestroy hlShtt
        Exit Function
    End If

    ilShttRecLen = Len(tlShtt) 'btrRecordLength(hlShtt)  'Get and save record length
    sgAllStationsStamp = slStamp
    'ilUpperShtt = UBound(tgAllStations)
    ilUpperShtt = LBound(tgAllStations)
    ilExtLen = Len(tlShtt)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlShtt) 'Obtain number of records
    btrExtClear hlShtt   'Clear any previous extend operation
    ilRet = btrGetFirst(hlShtt, tlShtt, ilShttRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlShtt)
        btrDestroy hlShtt
        'ilRet = btrClose(hlAtt)
        'btrDestroy hlAtt
        gObtainAllStations = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainAllStations = False
            ilRet = btrClose(hlShtt)
            btrDestroy hlShtt
            'ilRet = btrClose(hlAtt)
            'btrDestroy hlAtt
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlShtt, llNoRec, -1, "UC", "SHTT", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlShtt, ilOffSet, ilShttRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAllStations = False
        ilRet = btrClose(hlShtt)
        btrDestroy hlShtt
        'ilRet = btrClose(hlAtt)
        'btrDestroy hlAtt
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlShtt)    'Extract record
    ilRet = btrExtGetNext(hlShtt, tlShtt, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainAllStations = False
            ilRet = btrClose(hlShtt)
            btrDestroy hlShtt
            'ilRet = btrClose(hlAtt)
            'btrDestroy hlAtt
            Exit Function
        End If
        ilExtLen = Len(tlShtt)  'Extract operation record size
        'ilRet = btrExtGetFirst(hlShtt, tgAllStations(ilUpperShtt), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlShtt, tlShtt, ilExtLen, llRecPos)
        Loop


        Do While ilRet = BTRV_ERR_NONE
            If tlShtt.iType = 0 Then
                    tgAllStations(ilUpperShtt).iCode = tlShtt.iCode
                    tgAllStations(ilUpperShtt).sCallLetters = tlShtt.sCallLetters
                    '12/28/15: Replace state with user specified state
                    'tgAllStations(ilUpperShtt).sState = tlShtt.sState
                    If (Asc(tgSaf(0).sFeatures3) And SPLITCOPYLICENSE) = SPLITCOPYLICENSE Then 'Require Station Posting Prior to Invoicing
                        tgAllStations(ilUpperShtt).sState = tlShtt.sStateLic
                    ElseIf (Asc(tgSaf(0).sFeatures3) And SPLITCOPYPHYSICAL) = SPLITCOPYPHYSICAL Then
                        tgAllStations(ilUpperShtt).sState = tlShtt.sONState
                    Else
                        tgAllStations(ilUpperShtt).sState = tlShtt.sState
                    End If
                    tgAllStations(ilUpperShtt).sTimeZone = tlShtt.sTimeZone
                    tgAllStations(ilUpperShtt).sAgreementExist = tlShtt.sAgreementExist
                    tgAllStations(ilUpperShtt).lPermStationID = tlShtt.lPermStationID
                    tgAllStations(ilUpperShtt).iMktCode = tlShtt.iMktCode
                    tgAllStations(ilUpperShtt).iFmtCode = tlShtt.iFmtCode
                    tgAllStations(ilUpperShtt).iTztCode = tlShtt.iTztCode
                    tgAllStations(ilUpperShtt).iMetCode = tlShtt.iMetCode
                    tgAllStations(ilUpperShtt).iShttVefCode = 0
                    ilUpperShtt = ilUpperShtt + 1
                    If ilUpperShtt > UBound(tgAllStations) Then
                        'ReDim Preserve tgAllStations(1 To ilUpperShtt + 1000) As SHTTINFO
                        ReDim Preserve tgAllStations(0 To ilUpperShtt + 1000) As SHTTINFO
                    End If
            End If
            ilRet = btrExtGetNext(hlShtt, tlShtt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlShtt, tlShtt, ilExtLen, llRecPos)
            Loop
        Loop
        'Dick
        ReDim Preserve tgAllStations(0 To ilUpperShtt) As SHTTINFO
    End If
    ilRet = btrClose(hlShtt)
    btrDestroy hlShtt
    If UBound(tgAllStations) - 1 > 1 Then
        ArraySortTyp fnAV(tgAllStations(), 1), UBound(tgAllStations) - 1, 0, LenB(tgAllStations(1)), 2, 40, 0
    End If
    Exit Function
gObtainAllStationsErr2:
    ilRet = 1
    Resume Next
End Function

'           mFindStationMkt - find the Station Market name from Affiliate Market file.
'           The vehicle must be a station which is defined as either a Rep vehicle type,
'           or a conventional where Insertion Orders are sent (vffOnInsertions)
'           If that doesnt exist, use the market name from vehicle groups file in traffic
'           Show the market name on the contract only if Spf indicates to use it.
'           Set cbfAffMktCode to -1 if the market name should not be shown
'           Set with affiliate market code if a market exists; otherwise leave as 0
'           and vff will be used to show the market name
'
Public Sub gFindStationMkt(ilVefInxForCallLetters As Integer, tlCbf As CBF)
    Dim ilVefIndex As Integer
    Dim ilVffInx As Integer
    Dim ilShttInx As Integer
    Dim ilMktInx As Integer
    tlCbf.iAffMktCode = -1          'default to ignore the market name , do not show on contract
    'show Market Name on BR allowed for stations? Set flag for detail and vehicle summary records
    '0 = detail, 2 = vehicle summary, 4 = ntr, 8 = ntr summary
    If ((Asc(tgSaf(0).sFeatures4) And MKTNAMEONBR) = MKTNAMEONBR) And (tlCbf.iExtra2Byte = 0 Or tlCbf.iExtra2Byte = 2 Or tlCbf.iExtra2Byte = 4 Or tlCbf.iExtra2Byte = 8) Then
'                ilVefIndex = gBinarySearchVef(tlCbf.iVefCode)
'                If ilVefIndex <> -1 Then        '10-23-06 if new package vehicle entered for proposal, the index wont exist.  The vehicle
            ilShttInx = gBinarySearchCallLetters(Trim$(tgMVef(ilVefInxForCallLetters).sName))        'if vehicle is station, can be found
            If tgMVef(ilVefInxForCallLetters).sType = "R" Then
                'make sure station name is found in affiliate file, then see if mkt code exists
                If ilShttInx <> -1 Then
                    'obtain the mkt code
                    If tgAllStations(ilShttInx).iMktCode > 0 Then          'first look for matching station in station file, and retrieve a mkt code
                        tlCbf.iAffMktCode = tgAllStations(ilShttInx).iMktCode
                    Else                                            'no mkt code, use the vehicle mkt code from traffic
                        'get the vehicle group mkt code
                        tlCbf.iAffMktCode = 0               'flag to let Crystal know to use Vehicle mkt code
                    End If
                Else
                    tlCbf.iAffMktCode = 0                   'flag to let crystal know to use vehicle group mkt code
                End If
            Else
                ilVffInx = gBinarySearchVff(tlCbf.iVefCode)         'look for vff where Insertion Order flag is stored
                If ilVffInx <> -1 Then
                    '8-15-18 test for No, field could be blank, which defaults to Yes
                    If tgVff(ilVffInx).sOnInsertions <> "N" Then     'send Insertions to this station/vehicle?
                        'look for the affiliate mkt name
                        If ilShttInx <> -1 Then
                            If tgAllStations(ilShttInx).iMktCode > 0 Then      'first look for matching station in station file, and retrieve a mkt code
                                tlCbf.iAffMktCode = tgAllStations(ilShttInx).iMktCode
                            Else                                            'no mkt code, use the vehicle mkt code from traffic
                                'get the vehicle group mkt code
                                tlCbf.iAffMktCode = 0               'flag to let Crystal know to use Vehicle mkt code
                            End If
                        Else
                            tlCbf.iAffMktCode = 0
                        End If
                    End If
                End If
            End If
'                End If
    End If
    Exit Sub
End Sub

'           get the Net value from a rate, return gross or net as required
'           gGetGrossOrNetFromRate
'           <input> llRate - price to convert to net if required
'                   slGrossOrNet :  G = gross, N = net
'                   internal agency code
'           return - converted net amt , or gross value
'
Public Function gGetGrossOrNetFromRate(llRate As Long, slGrossOrNet As String, ilAgfCode As Integer, Optional blItsNTR As Boolean = False, Optional slNTRAgyCommFlag As String = "N") As Long
    Dim llTempPrice As Long
    Dim ilCommPct As Integer
    Dim slStr As String
    Dim slAmount As String
    Dim slSharePct As String
    Dim ilInx As Integer

    If slGrossOrNet = "G" Then              'return back gross, do nothing
        llTempPrice = llRate
    Else            'net
        llTempPrice = llRate
        If ilAgfCode = 0 Then           'direct
            ilCommPct = 10000                'no commission
        Else
            ilCommPct = 8500         'default to commissionable if no agency found
            ilInx = gBinarySearchAgf(ilAgfCode)
            If ilInx >= 0 Then
                ilCommPct = (10000 - tgCommAgf(ilInx).iCommPct)
            End If
        End If
        If blItsNTR Then                '3-28-19 may need to alter the agy commission for NTR since it can be different than the air time
                                        'air time may be commission, but not the NTR
            If slNTRAgyCommFlag <> "Y" Then
                ilCommPct = 10000
            End If
        End If

        slAmount = gLongToStrDec(llTempPrice, 2)
        slSharePct = gIntToStrDec(ilCommPct, 4)
        slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
        slStr = gRoundStr(slStr, ".01", 2)
        llTempPrice = gStrDecToLong(slStr, 2) 'adjusted net
    End If
    gGetGrossOrNetFromRate = llTempPrice
    Exit Function
End Function

