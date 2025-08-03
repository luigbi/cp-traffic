Attribute VB_Name = "rptRec4"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptrec4.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
'********************************************************
'
'Research Report (RSR) file definition
'
' D. Smith 12/5/00
'*********************************************************
'RSR record layout
Type RSR
    iGenDate(0 To 1) As Integer   'Generation Date
    '10-10-01
    lGenTime As Long              'generation time
   'iGenTime(0 To 1) As Integer   'Generation Time
    lDrfCode As Long              'Demo Research Code
    sDemoType As String * 1       '
    sDataType As String * 1       '
    iPopAud As Integer            '0 = Pop Values 1 = Aud Values
    sDemoDesc(0 To 17) As String * 8      'Demo #1 Description
    sForm As String * 1           '8=18 values; 6 or blank = 16 values (test for 8)
    lDemo(0 To 17) As Long
    sUnused As String * 3        'Future Expansion Space
End Type

'Rsrkey record layout
Type RSRKEY0
    iGenDate(0 To 1) As Integer 'Generation Date
    '10-10-01
    lGenTime As Long             'generation time
    'iGenTime(0 To 1) As Integer 'Generation Time
End Type

'Type RSR
'    iGenDate(0 To 1) As Integer   'Generation Date
'    '10-10-01
'    lGenTime As Long              'generation time
'   'iGenTime(0 To 1) As Integer   'Generation Time
'    lDrfCode As Long              'Demo Research Code
'    sDemoType As String * 1       '
'    sDataType As String * 1       '
'    iPopAud As Integer            '0 = Pop Values 1 = Aud Values
'    sDemoDesc1 As String * 8      'Demo #1 Description
'    sDemoDesc2 As String * 8      'Demo #2 Description
'    sDemoDesc3 As String * 8      'Demo #3 Description
'    sDemoDesc4 As String * 8      'Demo #4 Description
'    sDemoDesc5 As String * 8      'Demo #5 Description
'    sDemoDesc6 As String * 8      'Demo #6 Description
'    sDemoDesc7 As String * 8      'Demo #7 Description
'    sDemoDesc8 As String * 8      'Demo #8 Description
'    sDemoDesc9 As String * 8      'Demo #9 Description
'    sDemoDesc10 As String * 8     'Demo #10 Description
'    sDemoDesc11 As String * 8     'Demo #11 Description
'    sDemoDesc12 As String * 8     'Demo #12 Description
'    sDemoDesc13 As String * 8     'Demo #13 Description
'    sDemoDesc14 As String * 8     'Demo #14 Description
'    sDemoDesc15 As String * 8     'Demo #15 Description
'    sDemoDesc16 As String * 8     'Demo #16 Description
'    sDemoDesc17 As String * 8     'Demo #17 Description
'    sDemoDesc18 As String * 8     'Demo #18 Description
'    sForm As String * 1           '8=18 values; 6 or blank = 16 values (test for 8)
'    lDemo1 As Long
'    lDemo2 As Long
'    lDemo3 As Long
'    lDemo4 As Long
'    lDemo5 As Long
'    lDemo6 As Long
'    lDemo7 As Long
'    lDemo8 As Long
'    lDemo9 As Long
'    lDemo10 As Long
'    lDemo11 As Long
'    lDemo12 As Long
'    lDemo13 As Long
'    lDemo14 As Long
'    lDemo15 As Long
'    lDemo16 As Long
'    lDemo17 As Long
'    lDemo18 As Long
'    sUnused As String * 3        'Future Expansion Space
'End Type
''Rsrkey record layout
'Type RSRKEY0
'    iGenDate(0 To 1) As Integer 'Generation Date
'    '10-10-01
'    lGenTime As Long             'generation time
'    'iGenTime(0 To 1) As Integer 'Generation Time
'End Type

Type RESEARCHINFO              'array of vehicles or entire contract's spots per week , quarter at a time
    iVefCode As Integer
    iLineNo As Integer
    sType As String * 1           'S = std, O = order, a = air, h = hidden
    iPkLineNo As Integer         'associated package line # reference (if stype = H)
    iPkvefCode As Integer        'associated package vef
    lPop As Long
    'iQSpots As Integer           'total spots per this vehicles qtr
    lQSpots As Long               'total spots per this vehicles qtr
    'lTotalCost As Long
    dTotalCost As Double 'TTP 10439 - Rerate 21,000,000
    iTotalAvgRating As Long
    lTotalAvgAud As Long
    lTotalGrimps As Long
    lTotalGRP As Long
    lTotalCPP As Long
    lTotalCPM As Long
    lSatelliteEst As Long       '6-1-04
'    iSpots(1 To 13) As Integer
'iSpots(1 To 14) As Integer
''    lRates(1 To 13) As Long
'lRates(1 To 14) As Long
' '   lAvgAud(1 To 13) As Long
'lAvgAud(1 To 14) As Long
'    'iWklyRating(1 To 13) As Integer
'iWklyRating(1 To 14) As Integer
'    'lWklyGrImp(1 To 13) As Long
'lWklyGrimp(1 To 14) As Long
'    'lWklyGRP(1 To 13) As Long
'lWklyGRP(1 To 14) As Long
'    'lPopEst(1 To 13) As Long
'lPopEst(1 To 14) As Long
    lSpots(0 To 13) As Long
    lRates(0 To 13) As Long
    lAvgAud(0 To 13) As Long
    iWklyRating(0 To 13) As Integer
    lWklyGrimp(0 To 13) As Long
    lWklyGRP(0 To 13) As Long
    lPopEst(0 To 13) As Long
End Type

Type AFFBYMONTH         '8-29-02 Delinquent Affidavit report & Unbillable report
    lChfCode As Long
    iVefCode As Integer
    iAdfCode As Integer
    lOrderedAmt As Long     'ordered $ from chf
    lOrderedSpots As Long   'ordered spots from clf/cff
    lAiredAmt As Long       'aired $ from SBF
    lAiredSpots As Long     'aired spots from SBF
    iShowFlag As Integer    '0 = no SBF found (show on report) 1= SBF found, aff received, dont show on report)
End Type

Type AFFORDERED 'total $ ordered per contract
    lChfCode As Long
    lCntGross As Long
    lCntSpots As Long
End Type

Type ACTIVECNTS
    sKey As String * 8         'contract code left filled with zeroes for sorting
    lChfCode As Long            'contract code
    iAdfCode As Integer         '6-17-13
    lStartDate As Long          'contract start date
    lEndDate As Long            'contract end date
    sType As String * 1         'contract type
    lPop As Long                'population ,-1 = initialized value, 0 = pop varies across spots, else population
    lContrCost As Long          'total contract cost from schedule lines
    iMnfDemo As Integer         'primary demo
    lPledged As Long            'pledged contract aud
    lContrGrimp As Long           'audience total for all spots (sum of avg aud which = gross impressions0
    lCharge As Long             'audience charged spots
    lNC As Long                 'aud no charge spots
    lFill As Long               'aud fill spots
    lADU As Long                'aud ADU spots
    lMissed As Long             'aud missed spots
    lContrGrp As Long           'total grps per contract
    lChargeGrp As Long             'audience charged spots
    lNCGrp As Long                 'aud no charge spots
    lFillGrp As Long               'aud fill spots
    lADUGrp As Long                'aud ADU spots
    lMissedGrp As Long             'aud missed spots
    lContrSpots As Long           'total Spots per contract
    lChargeSpots As Long          'Total charged spots
    lNCSpots As Long              'Total no charge spots
    lFillSpots As Long            'Total fill spots
    lADUSpots As Long             'Total ADU spots
    lMissedSpots As Long          'Total missed spots
    iBookMissing As Integer     '0 = book exists for every vehicle in contract, 1= at least 1 vehicle doesnt have a book
    iFirstPopLink As Integer     'first pointer to entry containing the population for a schdule line
    iVaryPop As Integer         '0 = use the population found or 0 for varying pop across lines; 1 = cant product grps, varying pop within same line
End Type

Type POPLINKLIST
    iLine As Integer          'line #
    iNextLink As Integer        'pointer to next entry in list, which is another schedule line to a contract
    lPop As Long                'pop for a schedule line; -1 initialized, 0 = varying pop within the line, non-zero = population for line
End Type

Type VEHICLEBOOK
    iVefCode As Integer         'vehicle code
    'iDnfFirstLink As Integer    'index into first book for this vehicle
    'iDnfLastLink As Integer     'index into last book for this vehicle
    lDnfFirstLink As Long    '1-15-08 chg to long index into first book for this vehicle
    lDnfLastLink As Long     '1-15-08 chg to long index into last book for this vehicle

End Type

Type BOOKLIST
    sKey As String * 5          'date in string form,left filled with zeroes for sort
    iDnfCode As Integer         '
    lStartDate As Long          'start date of book
End Type

Type DNFLINKLIST
    idnfInx As Integer
End Type

Type SPOTTYPESORTAD
    sKey As String * 80 'Office Advertiser Contract
    tSdf As SDF
End Type
