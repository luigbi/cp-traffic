Attribute VB_Name = "RptSubsContractBR"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptSubsContractBR.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************

Option Explicit
Option Compare Text


Dim tmVsf As VSF

'12-16-20  CPM data for line or summary output on Broadcast Contract
Type CPM_BR
    lChfCode As Long            'contract internal code
    iVefCode As Integer         'vehicle code
    sType As String * 1         'Standard, hidden, or package
    lCost As Long               'total cost
    lImpressions As Long        'total Impressions
    lCPM As Long                'total cpms
    iPodCPMID As Integer        'id #
    iPkCPMID As Integer         'package ID reference for hidden IDs
    iUniquePkgVefCode As Integer    'for more than 1 package using the same vehicle name (i.e same vehicle, different dp)
    iRdfCode As Integer         'daypart
    lRafCode As Long         'geotarget
    lCxfCode As Long            'comment
    iStartDate(0 To 1) As Integer   'start date
    iEndDate(0 To 1) As Integer     'end date
    lStartDate As Long
    lEndDate As Long
    iCopyTypeMnfCode As Integer       'mnf copy type internal code
    sHideCBS As String * 1              'Hide the Cancel Before Start ID
    iTotalMonths As Integer         'total months to spread billing across dates
    lMonthlyCost(0 To 13) As Long       '12 months + over
    sPriceType As String * 1            'Price Type C=CPM F=Flat rate, B=Baked-in
    iLen As Integer 'Boostr Phase 2: Add new "Length" column next to "Ad Location"
End Type

'1-26-10 contract Billing Summary for NTR items by vehicle (non-installment).  This information
'will be merged with the air time billing contract summary
Type NTRBILLSUMMARY
    iVefCode As Integer         'NTR vehicle
    'lMonth(1 To 13) As Long     '1 year of NTR billing, plus overage after 12 months
    'lInstallment(1 To 13) As Long   '1 year of installment billing, plus overage after 12 months
    lMonth(0 To 13) As Long     '1 year of NTR billing, plus overage after 12 months. Index zero ignored
    lInstallment(0 To 13) As Long   '1 year of installment billing, plus overage after 12 months. Index zero ignored
    lTax1 As Long               'tax1 amount
    lTax2 As Long               'tax 2 amount
    lNet As Long
    lAgyComm As Long
End Type

Public sgAirTimeGrimp As String         'air time gross impressions to combine with ad server 3-31-21
Public lgAirTimeGross As Long          'air time gross $ to combine with adserver 3-31-21

Public Sub gFutureTaxesForNTR(llProject As Long, llProjectedTax1 As Long, llProjectedTax2 As Long, ilAgyComm As Integer, ilPctTrade As Integer, llTax1Pct As Long, llTax2Pct As Long, slGrossNet As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilTemp                        slNet                                                   *
'******************************************************************************************

    Dim slAmount As String
    Dim slStr As String
    Dim slCashAgyComm As String
    Dim slTax1Pct As String
    Dim slTax2Pct As String
    Dim slPctTrade As String
    Dim slCashPortion As String
    Dim slTax1 As String
    Dim slTax2 As String

    slCashAgyComm = gIntToStrDec(ilAgyComm, 2)      'need to get net if taxed by net for canada
    slPctTrade = gIntToStrDec(ilPctTrade, 0)        'trade portion doesnt get taxed
    'slPctTrade = ".00"                              'NTR doesnt have trade, force to 0 trade in case later required
    slTax1Pct = gLongToStrDec(llTax1Pct, 4)
    slTax2Pct = gLongToStrDec(llTax2Pct, 4)


    'calc the taxes 1 month at a time; thats the way the invoice program will be done
    'tax on gross if usa, else tax on net
    'If ((Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA) Then
    If slGrossNet <> "N" Then
        'Gross for USA
        slAmount = gLongToStrDec(llProject, 2)
    Else        'calc taxes net for canada
        slAmount = gLongToStrDec(llProject, 2)
        slAmount = gRoundStr(gDivStr(gMulStr(slAmount, gSubStr("100.00", slCashAgyComm)), "100.00"), ".01", 2)
    End If
    slStr = gMulStr(slTax1Pct, slAmount)                    'amt * tax1 %
    slStr = gRoundStr(slStr, "1", 0)
    slCashPortion = gSubStr("100.", slPctTrade)

    slTax1 = gDivStr(gMulStr(slStr, slCashPortion), "100")
    slTax1 = gRoundStr(slTax1, "1", 0)
    llProjectedTax1 = Val(slTax1)

    slStr = gMulStr(slTax2Pct, slAmount)                    'gross * tax2 %
    slStr = gRoundStr(slStr, "1", 0)
    slTax2 = gDivStr(gMulStr(slStr, slCashPortion), "100")
    slTax2 = gRoundStr(slTax2, "1", 0)
    llProjectedTax2 = Val(slTax2)

    Exit Sub
End Sub

Public Sub gFutureTaxes(ilLastBilledInx As Integer, llProject() As Long, llProjectedTax1() As Long, llProjectedTax2() As Long, ilAgyComm As Integer, ilPctTrade As Integer, llTax1Pct As Long, llTax2Pct As Long, slGrossNet As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slNet                                                                                 *
'******************************************************************************************
    Dim ilTemp As Integer
    Dim slAmount As String
    Dim slStr As String
    Dim slCashAgyComm As String
    Dim slTax1Pct As String
    Dim slTax2Pct As String
    Dim slPctTrade As String
    Dim slCashPortion As String
    Dim slTax1 As String
    Dim slTax2 As String

    slCashAgyComm = gIntToStrDec(ilAgyComm, 2)      'need to get net if taxed by net for canada
    slPctTrade = gIntToStrDec(ilPctTrade, 0)        'trade portion doesnt get taxed
    slTax1Pct = gLongToStrDec(llTax1Pct, 4)
    slTax2Pct = gLongToStrDec(llTax2Pct, 4)

    For ilTemp = 1 To 12
        llProjectedTax1(ilTemp) = 0
        llProjectedTax2(ilTemp) = 0
    Next ilTemp

    'calc the taxes 1 month at a time; thats the way the invoice program will be done
    For ilTemp = ilLastBilledInx To 12
        'tax on gross if usa, else tax on net
        'If ((Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA) Then
        If slGrossNet <> "N" Then
            'Gross for USA
            slAmount = gLongToStrDec(llProject(ilTemp), 2)
        Else        'calc taxes net for canada
            slAmount = gLongToStrDec(llProject(ilTemp), 2)
            slAmount = gRoundStr(gDivStr(gMulStr(slAmount, gSubStr("100.00", slCashAgyComm)), "100.00"), ".01", 2)
        End If
        slStr = gMulStr(slTax1Pct, slAmount)                    'amt * tax1 %
        slStr = gRoundStr(slStr, "1", 0)
        slCashPortion = gSubStr("100.", slPctTrade)

        slTax1 = gDivStr(gMulStr(slStr, slCashPortion), "100")
        slTax1 = gRoundStr(slTax1, "1", 0)
        llProjectedTax1(ilTemp) = llProjectedTax1(ilTemp) + Val(slTax1)

        slStr = gMulStr(slTax2Pct, slAmount)                    'gross * tax2 %
        slStr = gRoundStr(slStr, "1", 0)
        slTax2 = gDivStr(gMulStr(slStr, slCashPortion), "100")
        slTax2 = gRoundStr(slTax2, "1", 0)
        llProjectedTax2(ilTemp) = llProjectedTax2(ilTemp) + Val(slTax2)
    Next ilTemp

    Exit Sub
End Sub

'   Insure that the drive and directory changed for mapping of server (\\)
'   3-25-03
Public Sub gChDrDir()
    If InStr(1, sgCurDir, ":") > 0 Then 'colon exists
        ChDrive Left$(sgCurDir, 2)  'Set the default drive
        ChDir sgCurDir
    End If
End Sub

'       See if anything selected in a list box
'       <return> return true if at least 1 item selected
'
'       10-8-03
Public Function gSetGenCommand(lbcSelection As Control) As Integer
    Dim ilHowMany As Integer
    gSetGenCommand = False
    ilHowMany = lbcSelection.SelCount     'see if any payess selected
    If ilHowMany > 0 Then
        gSetGenCommand = True
    End If
    Exit Function
End Function

'       Test the flag in spot (PriceType) to determine if spot should
'       be shown on invoice or not.  If the spot doesnt have a "- or +" stored
'       in pricetype field, then it has been overridden in spot screen to
'       answer on-demand; test the spot instead of advertiser to determine
'       how its to be shown
'       <input>  slPriceType = price type flag from spot (SDF)
'                ilAdfCode - advertiser code
'       return - Y = show on invoice (+), N = dont show on invoice (-)
'       1-19-04

Public Function gTestShowFill(slPriceType As String, ilAdfCode As Integer) As String
    Dim ilLoopOnAdvt As Integer

    If slPriceType <> "-" And slPriceType <> "+" Then     'neither a - or +, fill wasnt overridden then use advt to determine how to show
        ilLoopOnAdvt = gBinarySearchAdf(ilAdfCode)
        If ilLoopOnAdvt <> -1 Then
            gTestShowFill = tgCommAdf(ilLoopOnAdvt).sBonusOnInv
        Else
            gTestShowFill = "N"       'dont show on inv (-)
        End If
    Else                'was overrriden in fill screen, use spot
        If slPriceType = "-" Then
            gTestShowFill = "N"       'dont show on inv (-)
        Else
            gTestShowFill = "Y"       'show on inv (+)
        End If
    End If

   Exit Function
End Function

Public Sub gParseDelimitedFields(slDelimited As String, slCDStr As String, ilLower As Integer, slFields() As String)
'
'   gParseCDFields slDelimitedChar ,slCDStr, ilLower, slFields()
'   Where:
'       slDelimitedChar - (I) delimited character.  Special testing for vertical bar delimiter
'                       if vertical bar delimiters accept quote in strings
'       slCDStr(I)- Comma delinited string
'       ilLower(I)- True=Convert string fields characters to lower case (preceding character is A-Z)
'       slFields() (O)- fields parsed from comma delimited string
'
    Dim ilFieldNo As Integer
    Dim ilFieldType As Integer  '0=String, 1=Number
    Dim slChar As String
    Dim ilIndex As Integer
    Dim ilAscChar As Integer
    Dim ilAddToStr As Integer
    Dim slNextChar As String
    Dim slDC As String * 1

    slDC = slDelimited

    For ilIndex = LBound(slFields) To UBound(slFields) Step 1
        slFields(ilIndex) = ""
    Next ilIndex
    ilFieldNo = 1
    ilIndex = 1
    ilFieldType = -1
    Do While ilIndex <= Len(Trim$(slCDStr))
        slChar = Mid$(slCDStr, ilIndex, 1)
        If ilFieldType = -1 Then
            If slChar = slDC Then    'delimiter was followed by a comma-blank field
                ilFieldType = -1
                ilFieldNo = ilFieldNo + 1
                If ilFieldNo > UBound(slFields) Then
                    Exit Sub
                End If
            ElseIf slChar <> """" Then
                ilFieldType = 1
                slFields(ilFieldNo) = slChar
            Else
                If slDelimited = "|" Then       'if vertical delimited, ok to accept quote
                    ilFieldType = 1
                    slFields(ilFieldNo) = slChar
                Else
                    ilFieldType = 0 'Quote field
                End If
            End If
        Else
            If ilFieldType = 0 Then 'Started with a Quote
                'Add to string unless "
                ilAddToStr = True
                If slChar = """" And slDelimited <> "|" Then     'if delimiter is vertical line, treat quote as OK
                    If ilIndex = Len(Trim$(slCDStr)) Then
                        ilAddToStr = False
                    Else
                        slNextChar = Mid$(slCDStr, ilIndex + 1, 1)
                        If slNextChar = slDC Then
                            ilAddToStr = False
                        End If
                    End If
                End If
                If ilAddToStr Then
                    If (slFields(ilFieldNo) <> "") And ilLower Then
                        ilAscChar = Asc(UCase(right$(slFields(ilFieldNo), 1)))
                        If ((ilAscChar >= Asc("A")) And (ilAscChar <= Asc("Z"))) Then
                            slChar = LCase$(slChar)
                        End If
                    End If
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        Exit Sub
                    End If
                    ilIndex = ilIndex + 1   'bypass comma
                End If
            Else
                'Add to string unless ,
                If slChar <> slDC Then
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        Exit Sub
                    End If
                End If
            End If
        End If
        ilIndex = ilIndex + 1
    Loop
End Sub

Public Function gIsItPolitical(ilAdfCode As Integer) As Integer
    Dim ilIsItPolitical As Integer
    Dim ilIndex As Integer
    ilIsItPolitical = False
    ilIndex = gBinarySearchAdf(ilAdfCode)
    If ilIndex <> -1 Then               'advt found
        If tgCommAdf(ilIndex).sPolitical = "Y" Then 'see if political flag is set
            ilIsItPolitical = True
        End If
    End If
    gIsItPolitical = ilIsItPolitical
End Function

'       Find the DP sort code for a given vehicle and DP
'       tgRcf contains the Rate card for the given contract,
'       tgRif and tgRdf contain the Item and DP info
'       <input>  DP code (rdf)
'                vefcode
'       <return> sort code, return -1 if line # should be used to sort
'
'       Test Site to determine how the sort should be obtained
'       spfUsingFeature5:  bit 3 - 4:  bit 3 = use R/C items entry sort code
'                                      bit 4 = use Line #
'                                      both 3 & 4 = 0:  use DP sort code
Public Function gFindDPSort(tmMRif() As RIF, tmMRDF() As RDF, ilRdfCode As Integer, ilVefCode As Integer) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop2                                                                               *
'******************************************************************************************
    Dim ilValue As Integer
    Dim illoop As Integer

    ilValue = Asc(tgSpf.sUsingFeatures5)
    If (ilValue And CNTRINVSORTRC) = CNTRINVSORTRC Then     'use rate card items sort code
        For illoop = LBound(tmMRif) To UBound(tmMRif) - 1
            If tmMRif(illoop).iRdfCode = ilRdfCode And tmMRif(illoop).iVefCode = ilVefCode Then
                If tmMRif(illoop).iSort <> 0 Then        'if the sort code is zero, drop down to using the DP sort code
                    gFindDPSort = tmMRif(illoop).iSort
                    Exit Function
                Else
                    Exit For
                End If
            '3/29/07:  move rdf search after search of rif
            'Else
            '    For ilLoop2 = LBound(tmMRDF) To UBound(tmMRDF) - 1
            '        If tmMRDF(ilLoop2).iCode = ilRdfCode Then
            '            gFindDPSort = tmMRDF(ilLoop2).iSortCode
            '            Exit For
            '        End If
            '    Next ilLoop2
            End If
        Next illoop
        'RIF not found or rifsort = 0, then use rdf
        For illoop = LBound(tmMRDF) To UBound(tmMRDF) - 1
            If tmMRDF(illoop).iCode = ilRdfCode Then
                If tmMRDF(illoop).iSortCode = 0 Then
                    gFindDPSort = -1
                Else
                    gFindDPSort = tmMRDF(illoop).iSortCode
                End If
                Exit For
            End If
        Next illoop
    ElseIf (ilValue And CNTRINVSORTLN) = CNTRINVSORTLN Then    'use DP sort codes
        gFindDPSort = -1                        'using line sort
    Else
        For illoop = LBound(tmMRDF) To UBound(tmMRDF) - 1
            If tmMRDF(illoop).iCode = ilRdfCode Then
                If tmMRDF(illoop).iSortCode = 0 Then
                    gFindDPSort = -1
                Else
                    gFindDPSort = tmMRDF(illoop).iSortCode
                End If
                Exit For
            End If
        Next illoop
    End If
End Function

'       Build Installment $ for a contract/proposal report
'       Process only installment records (trantype = F) from SBF for the matching contract code
'       If a vehicle has already been processed for this contract, exit out.
'       This routine processes by schedule lines for the contract/proposal (BR)
'       (report list and snapshot modules)
'       <input> -
'                 ilVefCode - vehicle to gather $
'                 tlVehiclesDone() -array of vehicles already processed.  This array
'                                   should be initialized the first time from calling rtn
'                 llStartDates()   - array of start dates
'                 llProject()      - array of projection $
'                 tlInstallSBF()   - array of SBF installment records
'TTP 8410 - blIgnoreRepeat, if blIgnoreRepeat=true, it will ignore the check for iVehiclesDone.  this way we can get the installment $ for a hidden vehicle
Public Sub gBuildInstallMonths(ilVefCode As Integer, tlVehiclesDone() As Integer, llStartDates() As Long, llProject() As Long, tlInstallSBF() As SBF, Optional blIgnoreRepeat As Boolean = False)
    Dim ilVef As Integer
    Dim ilUpper As Integer
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim ilMonth As Integer
    Dim slDate As String
    Dim llDate As Long
    
    For ilMonth = 1 To 13
        llProject(ilMonth) = 0
    Next ilMonth
    'TTP 8410
    If blIgnoreRepeat = False Then
        For ilVef = LBound(tlVehiclesDone) To UBound(tlVehiclesDone) - 1
            If tlVehiclesDone(ilVef) = ilVefCode Then
                Exit Sub
            End If
        Next ilVef
    End If
    'vehicle not found, so not processed yet
    ilUpper = UBound(tlVehiclesDone)
    tlVehiclesDone(ilUpper) = ilVefCode
    ReDim Preserve tlVehiclesDone(LBound(tlVehiclesDone) To ilUpper + 1) As Integer

    For llSbf = LBound(tlInstallSBF) To UBound(tlInstallSBF) - 1
        If tlInstallSBF(llSbf).iBillVefCode = ilVefCode And tlInstallSBF(llSbf).sTranType = "F" Then
            'determine what month this installment billing goes into
            For ilMonth = 1 To 13           'max 12 month billing, then over 12 months lumped into 1 bucket
                gUnpackDate tlInstallSBF(llSbf).iDate(0), tlInstallSBF(llSbf).iDate(1), slDate
                llDate = gDateValue(slDate)
                If llDate >= llStartDates(13) Then    'beyond 13 months, lump into one bucket for remaining months
                    llProject(13) = llProject(13) + tlInstallSBF(llSbf).lGross
                    Exit For
                Else
                    If llDate >= llStartDates(ilMonth) And llDate < llStartDates(ilMonth + 1) Then
                        llProject(ilMonth) = llProject(ilMonth) + tlInstallSBF(llSbf).lGross
                        Exit For
                    End If
                End If

            Next ilMonth
        End If
    Next llSbf
    'round the values
'1-26-10 removed rounding
'    For ilMonth = 1 To 13
'        llProject(ilMonth) = llProject(ilMonth) / 100
'    Next ilMonth
    Exit Sub
End Sub

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
Public Sub gBuildFlightInfo(ilClf As Integer, llStdStartDates() As Long, ilFirstProjInx As Integer, ilMaxInx As Integer, llProject() As Long, llProjectSpots() As Long, llProjectRC() As Long, llProjectAcq() As Long, ilWkOrMonth As Integer, tlClfInp() As CLFLIST, tlCffInp() As CFFLIST)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llWhichRate                                                                           *
'******************************************************************************************
    Dim ilCff As Integer
    Dim slStr As String
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim illoop As Integer
    Dim llDate As Long
    Dim llDate2 As Long
    Dim llSpots As Long
    Dim ilTemp As Integer
    Dim llStdStart As Long
    Dim llStdEnd As Long
    Dim ilMonthInx As Integer
    Dim ilWkInx As Integer
    Dim tlCff As CFF
    Dim llAcquisitionRate As Long

    llStdStart = llStdStartDates(ilFirstProjInx)
    llStdEnd = llStdStartDates(ilMaxInx)
    ilCff = tlClfInp(ilClf).iFirstCff
    Do While ilCff <> -1
        If (tlCffInp(ilCff).iStatus = 0) Or (tlCffInp(ilCff).iStatus = 1) Then
            tlCff = tlCffInp(ilCff).CffRec

            llAcquisitionRate = tlClfInp(ilClf).ClfRec.lAcquisitionCost

            gUnpackDate tlCff.iStartDate(0), tlCff.iStartDate(1), slStr
            llFltStart = gDateValue(slStr)
            'backup start date to Monday
            'ilLoop = gWeekDayLong(llFltStart)
            'Do While ilLoop <> 0
            '    llFltStart = llFltStart - 1
            '    ilLoop = gWeekDayLong(llFltStart)
            'Loop
            gUnpackDate tlCff.iEndDate(0), tlCff.iEndDate(1), slStr
            llFltEnd = gDateValue(slStr)

            'the flight dates must be within the start and end of the projection periods,
            'not be a CAncel before start flight, and have a cost > 0
            If (llFltStart < llStdEnd And llFltEnd >= llStdStart) And (llFltEnd >= llFltStart) Then
                'backup start date to Monday
                illoop = gWeekDayLong(llFltStart)
                Do While illoop <> 0
                    llFltStart = llFltStart - 1
                    illoop = gWeekDayLong(llFltStart)
                Loop
                'only retrieve for projections, anything in the past has already
                'been invoiced and has been retrieved from history or receiv files
                'adjust the gather dates from flights: use flight start date or requested start date, whichever is later
                If llStdStart > llFltStart Then
                    llFltStart = llStdStart
                End If
                'use flight end date or requsted end date, whichever is lesser
                If llStdEnd < llFltEnd Then
                    llFltEnd = llStdEnd
                End If

                For llDate = llFltStart To llFltEnd Step 7
                    'Loop on the number of weeks in this flight
                    'calc week into of this flight to accum the spot count
                    If tlCff.sDyWk = "W" Then            'weekly
                        llSpots = tlCff.iSpotsWk + tlCff.iXSpotsWk
                    Else                                        'daily
                        If illoop + 6 < llFltEnd Then           'we have a whole week
                            llSpots = tlCff.iDay(0) + tlCff.iDay(1) + tlCff.iDay(2) + tlCff.iDay(3) + tlCff.iDay(4) + tlCff.iDay(5) + tlCff.iDay(6)
                        Else
                            llFltEnd = llDate + 6
                            If llDate > llFltEnd Then
                                llFltEnd = llFltEnd       'this flight isn't 7 days
                            End If
                            For llDate2 = llDate To llFltEnd Step 1
                                ilTemp = gWeekDayLong(llDate2)
                                llSpots = llSpots + tlCff.iDay(ilTemp)
                            Next llDate2
                        End If
                    End If
                    If ilWkOrMonth = 1 Then                     'monthly buckets
                        'determine month that this week belongs in, then accumulate the gross and net $
                        'currently, the projections are based on STandard bdcst
                        For ilMonthInx = ilFirstProjInx To ilMaxInx - 1 Step 1       'loop thru months to find the match
                            If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                                llProject(ilMonthInx) = llProject(ilMonthInx) + (llSpots * tlCff.lActPrice)
                                llProjectAcq(ilMonthInx) = llProjectAcq(ilMonthInx) + (llSpots * llAcquisitionRate)
                                llProjectRC(ilMonthInx) = llProjectRC(ilMonthInx) + (llSpots * tlCff.lPropPrice)
                                llProjectSpots(ilMonthInx) = llProjectSpots(ilMonthInx) + llSpots
                                Exit For
                            End If
                        Next ilMonthInx
                    Else                                    'weekly buckets
                        ilWkInx = (llDate - llStdStartDates(1)) \ 7 + 1
                        ''4-3-07 make sure the data isnt gathered beyond the period requested
                        'If ilWkInx > 0 And llDate >= llStdStartDates(LBound(llStdStartDates)) And llDate < llStdStartDates(ilMaxInx) Then   '1-24-08(UBound(llStdStartDates)) Then
                        If ilWkInx > 0 And llDate >= llStdStartDates(1) And llDate < llStdStartDates(ilMaxInx) Then   '1-24-08(UBound(llStdStartDates)) Then
                            llProject(ilMonthInx) = llProject(ilMonthInx) + (llSpots * tlCff.lActPrice)
                            llProjectAcq(ilMonthInx) = llProjectAcq(ilMonthInx) + (llSpots * llAcquisitionRate)
                            llProjectRC(ilMonthInx) = llProjectRC(ilMonthInx) + (llSpots * tlCff.lPropPrice)
                            llProjectSpots(ilWkInx) = llProjectSpots(ilWkInx) + llSpots
                        End If
                    End If
                Next llDate                                     'for llDate = llFltStart To llFltEnd
            End If                                          '
        End If
        ilCff = tlCffInp(ilCff).iNextCff                   'get next flight record from mem
    Loop                                            'while ilcff <> -1
End Sub

'           gIsItHardCost - given an item code, test the matching item to
'           see if it a hard cost NTR item
'
'           <input> ilMnfCode - Type "I" (NTRtype) mnfcode
'                   tlNTRMNF() array of mnf ntr item records
'           <return) true if hard cost
'
Public Function gIsItHardCost(ilMnfCode As Integer, tlNTRMNF() As MNF) As Integer
    Dim ilNTRLoop As Integer
    Dim ilIsItHardCost As Integer

    ilIsItHardCost = False          'assume this item is not a hard cost
    If ilMnfCode > 0 Then
        For ilNTRLoop = LBound(tlNTRMNF) To UBound(tlNTRMNF) - 1
            If (ilMnfCode = tlNTRMNF(ilNTRLoop).iCode) Then
                If (Trim(tlNTRMNF(ilNTRLoop).sCodeStn) = "Y") Then
                    ilIsItHardCost = True
                End If
                Exit For
            End If
        Next ilNTRLoop
    End If
    gIsItHardCost = ilIsItHardCost
End Function

'       Populate all CPM IDs by the selected contract.
'       Build a table of all unique vehicles for the podcast summary page, totalling cost, impressions and cpm
'
'       <input>  hlPcf - Podcast file handle
'                llChfCode =  internal contract code
'       <output>  Return array of all cpm records for the contract (tlPcf()) with the monthly summary $ spread across the months
'                 Return array of unique vehicles with cpm calculated (tlCPMSummary())
Public Function gBuildCPMIDs(hlPcf As Integer, tlChf As CHF, llStdStartDates() As Long, tlCPM_IDs() As CPM_BR, tlCPMSummary() As CPM_BR) As Boolean
    Dim blAllOK As Boolean
    Dim blFirstTime As Boolean
    Dim illoop As Integer           'loop on detail cpm items
    Dim ilLoopOnSum As Integer        'loop on unique vehicle array
    Dim ilUpperCPMSummary As Integer
    Dim blFoundMatch As Boolean
    ReDim ilPkgVefList(0 To 0) As Integer
    ReDim ilStdVefList(0 To 0) As Integer
    Dim ilLoopOnVef As Integer
    Dim ilLoopForMatchPkg As Integer
    Dim ilMonth As Integer
    Dim tlPcf() As PCF
    Dim llEarliestDate As Long
    Dim slEarliestDate As String
    Dim llLatestDate As Long
    Dim slLatestDate As String
    Dim ilTotalMonths As Integer
    Dim llTempEarliestDate As Long
    Dim llTempLatestDate As Long
    Dim llAvgCost As Long
    Dim llVehicleCost As Long
    Dim ilFindMatchPkgIDForVeh As Integer
    Dim llMonthlyAmt As Long 'Boostr Phase 2
    
    gBuildCPMIDs = True           'assume all reads ok
    ReDim tlPcf(0 To 0) As PCF
    ReDim tlCPM_IDs(0 To 0) As CPM_BR
    ReDim tlCPMSummary(0 To 0) As CPM_BR
    If tlChf.sAdServerDefined <> "Y" Then
        gBuildCPMIDs = False
        Exit Function
    End If
    
    blAllOK = gObtainPcf(hlPcf, tlChf.lCode, tlPcf())       'obtain all pcm for matching contract code
    If Not blAllOK Then
        gBuildCPMIDs = False
    Else
        gBuildCPMIDs = True
        'get all the unique package vehicles.  If more than 1 package exists for same vehicle, it needs to be combined
        blFoundMatch = False
        blFirstTime = True
        For illoop = LBound(tlPcf) To UBound(tlPcf) - 1
            If tlPcf(illoop).sType = "P" Then
                If blFirstTime Then
                    blFirstTime = False
                    ilPkgVefList(UBound(ilPkgVefList)) = tlPcf(illoop).iVefCode
                    ReDim Preserve ilPkgVefList(0 To UBound(ilPkgVefList) + 1) As Integer
                Else
                    blFoundMatch = False
                    For ilLoopOnVef = LBound(ilPkgVefList) To UBound(ilPkgVefList) - 1
                        If ilPkgVefList(ilLoopOnVef) = tlPcf(illoop).iVefCode Then
                            blFoundMatch = True
                        Else
                            ilPkgVefList(UBound(ilPkgVefList)) = tlPcf(illoop).iVefCode
                            ReDim Preserve ilPkgVefList(0 To UBound(ilPkgVefList) + 1) As Integer
                        End If
                    Next ilLoopOnVef
                End If
            End If
        Next illoop
        
        blFirstTime = True
        
        'this is each CPM id for the line by line item to show on report
        For illoop = LBound(tlPcf) To UBound(tlPcf) - 1
            tlCPM_IDs(illoop).lChfCode = tlPcf(illoop).lChfCode
            tlCPM_IDs(illoop).iVefCode = tlPcf(illoop).iVefCode
            tlCPM_IDs(illoop).sType = tlPcf(illoop).sType        'Standard, hidden, or package
            tlCPM_IDs(illoop).iLen = tlPcf(illoop).iLen  'Boostr Phase 2: Add new "Length" column next to "Ad Location"
            '09/27/2022 - JW - Contract report: display "Baked-in" for flat rate lines with "baked-in" flag
            If tlPcf(illoop).iDeliveryType = 1 Then
                tlCPM_IDs(illoop).sPriceType = "B" 'Baked-in
            Else
                tlCPM_IDs(illoop).sPriceType = tlPcf(illoop).sPriceType
            End If
            tlCPM_IDs(illoop).lCost = tlPcf(illoop).lTotalCost
            '09/27/2022 - JW - Contract report: display "Baked-in" for flat rate lines with "baked-in" flag
            If tlPcf(illoop).iDeliveryType = 1 Then
                tlCPM_IDs(illoop).lImpressions = 0
            Else
                tlCPM_IDs(illoop).lImpressions = tlPcf(illoop).lImpressionGoal
            End If
            tlCPM_IDs(illoop).lCPM = tlPcf(illoop).lPodCPM
            tlCPM_IDs(illoop).iPodCPMID = tlPcf(illoop).iPodCPMID
            tlCPM_IDs(illoop).iPkCPMID = tlPcf(illoop).iPkCPMID        'package ID reference for hidden IDs
            If tlCPM_IDs(illoop).iPkCPMID > 0 Then                      'find the package vehicle reference by matching on the package id
                For ilFindMatchPkgIDForVeh = LBound(tlPcf) To UBound(tlPcf) - 1
                    If tlCPM_IDs(illoop).iPkCPMID = tlPcf(ilFindMatchPkgIDForVeh).iPodCPMID Then
                        tlCPM_IDs(illoop).iUniquePkgVefCode = tlPcf(ilFindMatchPkgIDForVeh).iVefCode
                        Exit For
                    End If
                Next ilFindMatchPkgIDForVeh
            Else
                tlCPM_IDs(illoop).iUniquePkgVefCode = tlPcf(illoop).iVefCode
            End If
            'iUniquePkgVefCode As Integer    'for more than 1 package using the same vehicle name (i.e same vehicle, different dp)
            tlCPM_IDs(illoop).iRdfCode = tlPcf(illoop).iRdfCode      'daypart
            tlCPM_IDs(illoop).lRafCode = tlPcf(illoop).lRafCode            'geotarget
            tlCPM_IDs(illoop).lCxfCode = tlPcf(illoop).lCxfCode         'comment
            tlCPM_IDs(illoop).iStartDate(0) = tlPcf(illoop).iStartDate(0) 'start date
            tlCPM_IDs(illoop).iStartDate(1) = tlPcf(illoop).iStartDate(1)
            tlCPM_IDs(illoop).iEndDate(0) = tlPcf(illoop).iEndDate(0) 'start date
            tlCPM_IDs(illoop).iEndDate(1) = tlPcf(illoop).iEndDate(1)
            gUnpackDateLong tlPcf(illoop).iStartDate(0), tlPcf(illoop).iStartDate(1), tlCPM_IDs(illoop).lStartDate
            gUnpackDateLong tlPcf(illoop).iEndDate(0), tlPcf(illoop).iEndDate(1), tlCPM_IDs(illoop).lEndDate
            tlCPM_IDs(illoop).iCopyTypeMnfCode = tlPcf(illoop).iCopyTypeMnfCode      'mnf copy type internal code
            tlCPM_IDs(illoop).sHideCBS = tlPcf(illoop).sHideCBS              'Hide the Cancel Before Start ID
            'detrmine standard or calendar billing to spread across the months
            'create vehicle billing summary for podcast CPM
            'Determine # of months, std or cal
            'determine # of calendar months to average across the billing periods
            ilTotalMonths = 0
            llEarliestDate = tlCPM_IDs(illoop).lStartDate
            llLatestDate = tlCPM_IDs(illoop).lEndDate
            If tlCPM_IDs(illoop).lStartDate <= tlCPM_IDs(illoop).lEndDate Then
                llTempEarliestDate = llEarliestDate
                'TTP 10947 - Contract report monthly/quarterly summary page: flat rate line now showing in month or quarter, shown instead between Q4 and Total Cost
                Do While llTempEarliestDate <= llLatestDate
                    ilTotalMonths = ilTotalMonths + 1
                    slEarliestDate = Format$(llTempEarliestDate, "ddddd")
                    If tgChf.sBillCycle = "C" Then
                        slEarliestDate = gObtainStartCal(slEarliestDate)
                        slLatestDate = gObtainEndCal(slEarliestDate)
                    Else
                        slEarliestDate = gObtainStartStd(slEarliestDate)
                        slLatestDate = gObtainEndStd(slEarliestDate)
                    End If
                    llTempLatestDate = gDateValue(slLatestDate)
                    llTempEarliestDate = llTempLatestDate + 1
                Loop

                tlCPM_IDs(illoop).iTotalMonths = ilTotalMonths
                llTempEarliestDate = llEarliestDate
                If ilTotalMonths <> 0 Then 'TTP 10643 - Proposals/Contract report: Ad server flat rate line with $0 cost results in "overflow" error
                    llAvgCost = tlPcf(illoop).lTotalCost / ilTotalMonths
                Else
                    llAvgCost = 0
                End If
            
                llVehicleCost = tlPcf(illoop).lTotalCost
                'TTP 10947 - Contract report monthly/quarterly summary page: flat rate line now showing in month or quarter, shown instead between Q4 and Total Cost
                Do While llTempEarliestDate <= llLatestDate          'the start months of the array may not always be the same start month of the CPM ID record (due to starting on a qtr)
                    For ilMonth = 0 To 12
                        If llTempEarliestDate >= llStdStartDates(ilMonth) And llTempEarliestDate < llStdStartDates(ilMonth + 1) Then
                            ilTotalMonths = ilTotalMonths - 1
                            'Boostr Phase 2: Contract report: modify final monthly/quarterly summary page to use daily or monthly method depending on Site setting
                            If tgSpfx.iLineCostType = 1 And tlPcf(illoop).sPriceType = "F" Then
                                'Use Daily Averaging
                                llMonthlyAmt = mDeterminePeriodAmountByDaily(Format(tlCPM_IDs(illoop).lStartDate - 1, "ddddd"), Format(llStdStartDates(ilMonth), "ddddd"), Format(llStdStartDates(ilMonth + 1) - 1, "ddddd"), Format(tlCPM_IDs(illoop).lStartDate, "ddddd"), Format(tlCPM_IDs(illoop).lEndDate, "ddddd"), 0, tlPcf(illoop).lTotalCost) * 100
                                If llVehicleCost - llMonthlyAmt >= 0 Then          'last month gets remainder, not the total avg since it may not balance to penny
                                    tlCPM_IDs(illoop).lMonthlyCost(ilMonth) = llMonthlyAmt
                                Else
                                    tlCPM_IDs(illoop).lMonthlyCost(ilMonth) = llVehicleCost
                                End If
                                llVehicleCost = llVehicleCost - llMonthlyAmt
                            Else
                                'Use Monthly Averaging
                                If ilMonth > 11 Then
                                    If llVehicleCost - llAvgCost >= 0 Then          'last month gets remainder, not the total avg since it may not balance to penny
                                        tlCPM_IDs(illoop).lMonthlyCost(12) = llAvgCost
                                    Else
                                        tlCPM_IDs(illoop).lMonthlyCost(12) = llVehicleCost
                                    End If
                                    llVehicleCost = llVehicleCost - llAvgCost
                                ElseIf ilTotalMonths >= 0 Then
                                    If llVehicleCost - llAvgCost >= 0 Then          'last month gets remainder, not the total avg since it may not balance to penny
                                        tlCPM_IDs(illoop).lMonthlyCost(ilMonth) = llAvgCost
                                    Else
                                        tlCPM_IDs(illoop).lMonthlyCost(ilMonth) = llVehicleCost
                                    End If
                                    llVehicleCost = llVehicleCost - llAvgCost
                                Else            'remainder of months are 0
                                    tlCPM_IDs(illoop).lMonthlyCost(ilMonth) = 0
                                End If
                            End If
                            Exit For
                        End If
                    Next ilMonth
                    
                    slEarliestDate = Format$(llTempEarliestDate, "ddddd")
                    If tgChf.sBillCycle = "C" Then
                        slEarliestDate = gObtainStartCal(slEarliestDate)
                        slLatestDate = gObtainEndCal(slEarliestDate)
                    Else
                        slEarliestDate = gObtainStartStd(slEarliestDate)
                        slLatestDate = gObtainEndStd(slEarliestDate)
                    End If
                    llTempLatestDate = gDateValue(slLatestDate)
                    llTempEarliestDate = llTempLatestDate + 1
                Loop
                If llVehicleCost > 0 Then
                    tlCPM_IDs(illoop).lMonthlyCost(13) = llVehicleCost
                End If
            Else                'cancel before start
                ilTotalMonths = ilTotalMonths
                tlCPM_IDs(illoop).lCost = 0
                tlCPM_IDs(illoop).lImpressions = 0
                tlCPM_IDs(illoop).lCPM = 0
            End If
            ReDim Preserve tlCPM_IDs(LBound(tlCPM_IDs) To UBound(tlCPM_IDs) + 1) As CPM_BR
        Next illoop
        
        'build array of unique vehicles; accumulate the impressions and cost
        'more than 1 vehicle of same name must calculate the cpm (cost*1000)/impressions
        ilUpperCPMSummary = UBound(tlCPMSummary)
        For illoop = LBound(tlCPM_IDs) To UBound(tlCPM_IDs) - 1
            If blFirstTime Then
                blFirstTime = False
                tlCPMSummary(ilUpperCPMSummary).iVefCode = tlCPM_IDs(illoop).iVefCode
                tlCPMSummary(ilUpperCPMSummary).lCost = tlCPM_IDs(illoop).lCost
                tlCPMSummary(ilUpperCPMSummary).lImpressions = tlCPM_IDs(illoop).lImpressions
                tlCPMSummary(ilUpperCPMSummary).lCPM = tlCPM_IDs(illoop).lCPM
                tlCPMSummary(ilUpperCPMSummary).iUniquePkgVefCode = tlCPM_IDs(illoop).iVefCode
                tlCPMSummary(ilUpperCPMSummary).sType = tlCPM_IDs(illoop).sType
                tlCPMSummary(ilUpperCPMSummary).iLen = tlCPM_IDs(illoop).iLen 'Boostr Phase 2: Add new "Length" column next to "Ad Location"
                tlCPMSummary(ilUpperCPMSummary).sPriceType = tlCPM_IDs(illoop).sPriceType
                If tlCPM_IDs(illoop).sType = "P" Then
                    tlCPMSummary(ilUpperCPMSummary).lCost = 0
                    tlCPMSummary(ilUpperCPMSummary).lImpressions = 0
                    tlCPMSummary(ilUpperCPMSummary).lCPM = 0
                End If
                For ilMonth = 0 To 13
                    tlCPMSummary(ilUpperCPMSummary).lMonthlyCost(ilMonth) = tlCPM_IDs(illoop).lMonthlyCost(ilMonth)
                Next ilMonth
                ilUpperCPMSummary = ilUpperCPMSummary + 1
                ReDim Preserve tlCPMSummary(0 To ilUpperCPMSummary) As CPM_BR
            Else
                'find matching vehicle to combine for summary
                blFoundMatch = False
                For ilLoopOnSum = LBound(tlCPMSummary) To ilUpperCPMSummary - 1
                    If tlCPMSummary(ilLoopOnSum).iVefCode = tlCPM_IDs(illoop).iVefCode And tlCPMSummary(ilLoopOnSum).sType = tlCPM_IDs(illoop).sType Then
                        blFoundMatch = True
                        If tlCPM_IDs(illoop).sType <> "P" Then          'ignore accumulating the pkg values, get from hidden lines or std lines
                            tlCPMSummary(ilLoopOnSum).lCost = tlCPMSummary(ilLoopOnSum).lCost + tlCPM_IDs(illoop).lCost
                            tlCPMSummary(ilLoopOnSum).lImpressions = tlCPMSummary(ilLoopOnSum).lImpressions + tlCPM_IDs(illoop).lImpressions
                        End If
                        For ilMonth = 0 To 13
                            tlCPMSummary(ilLoopOnSum).lMonthlyCost(ilMonth) = tlCPMSummary(ilLoopOnSum).lMonthlyCost(ilMonth) + tlCPM_IDs(illoop).lMonthlyCost(ilMonth)
                        Next ilMonth
                        Exit For
                    End If
                Next ilLoopOnSum
                If Not blFoundMatch Then
                    tlCPMSummary(ilUpperCPMSummary).iVefCode = tlCPM_IDs(illoop).iVefCode
                    tlCPMSummary(ilUpperCPMSummary).lCost = tlCPM_IDs(illoop).lCost
                    tlCPMSummary(ilUpperCPMSummary).lImpressions = tlCPM_IDs(illoop).lImpressions
                    tlCPMSummary(ilUpperCPMSummary).iUniquePkgVefCode = tlCPM_IDs(illoop).iVefCode
                    tlCPMSummary(ilUpperCPMSummary).sType = tlCPM_IDs(illoop).sType
                    tlCPMSummary(ilUpperCPMSummary).sPriceType = tlCPM_IDs(illoop).sPriceType
                    If tlCPM_IDs(illoop).sType = "P" Then
                        tlCPMSummary(ilUpperCPMSummary).lCost = 0
                        tlCPMSummary(ilUpperCPMSummary).lImpressions = 0
                        tlCPMSummary(ilUpperCPMSummary).lCPM = 0
                    End If
                    For ilMonth = 0 To 13
                        tlCPMSummary(ilUpperCPMSummary).lMonthlyCost(ilMonth) = tlCPM_IDs(illoop).lMonthlyCost(ilMonth)
                    Next ilMonth
                    ilUpperCPMSummary = ilUpperCPMSummary + 1
                    ReDim Preserve tlCPMSummary(0 To ilUpperCPMSummary) As CPM_BR
                End If
            End If
        Next illoop
        'Example--
        '       Vehicle A Pkg      ID 1
        '          Veh 1 Hidden    ID 2
        '          Veh 2 Hidden    ID 3
        '       Vehicle A Pkg      ID 4
        '          Veh 1 Hidden    ID 5
        '          Veh 2 Hidden    ID 6
        'Steps:  --------> Outer loop is to loop on the unique vehicles (ilPkgVefList)
        '          |
        '          | ---------->  Loop on the detail CPM items (tlPcf) to find the matching pkg vehicle code
        '          |
        '             |
        '             | ----------> Loop  on the detail CPM items (tlPcf) to find the matching hidden lines, using the reference pkg ID, for the package vehicle code
        '             |
        '                |
        '                |   ---------->  Loop on the vehicle Summary array (tlCPMSummary) matching on vehicle to add in cost & impressions so the CPM can be calculated for package vehicles
        '                |                 when more than 1 of same vehicle exists
        '
        'Calculate the package CPM for unique vehicle
        For ilLoopOnVef = LBound(ilPkgVefList) To UBound(ilPkgVefList) - 1
            For illoop = LBound(tlCPM_IDs) To UBound(tlCPM_IDs) - 1
                If (tlCPM_IDs(illoop).iUniquePkgVefCode = ilPkgVefList(ilLoopOnVef)) And tlCPM_IDs(illoop).sType <> "P" Then      'found the matching package vehicle, look for the matching reference #
                    For ilLoopOnSum = LBound(tlCPMSummary) To UBound(tlCPMSummary) - 1      'summarize all the hidden lines for the package
                        If tlCPMSummary(ilLoopOnSum).iVefCode = ilPkgVefList(ilLoopOnVef) And tlCPMSummary(ilLoopOnSum).sType = "P" Then
                            tlCPMSummary(ilLoopOnSum).lCost = tlCPMSummary(ilLoopOnSum).lCost + tlCPM_IDs(illoop).lCost
                            tlCPMSummary(ilLoopOnSum).lImpressions = tlCPMSummary(ilLoopOnSum).lImpressions + tlCPM_IDs(illoop).lImpressions
                            Exit For
                        End If
                    Next ilLoopOnSum        'keep searching for matching vehicle
                End If
            Next illoop             'look for another package vehicle with same name
        Next ilLoopOnVef            'next unique package vehicle

        
        'loop thru summary array and calculate the cpm just in case there were multiple ids using same vehicle reference (more than 1 pkg using same name)
        For ilLoopOnSum = LBound(tlCPMSummary) To UBound(tlCPMSummary) - 1
            If tlCPMSummary(ilLoopOnSum).lImpressions > 0 Then
                'TTP 10513 - Proposals/Contract report: Overflow when ad server impressions are too high, TTP 10514 - Invoice: overflow error when impressions too high
                'Working:
                tlCPMSummary(ilLoopOnSum).lCPM = ((1000 * CSng(tlCPMSummary(ilLoopOnSum).lCost)) / tlCPMSummary(ilLoopOnSum).lImpressions)    'cpm xxxxx.xx
                'Alternate:
                'tlCPMSummary(ilLoopOnSum).lCPM = tlCPMSummary(ilLoopOnSum).lCost / (tlCPMSummary(ilLoopOnSum).lImpressions / 1000)    'cpm xxxxx.xx
            End If
        Next ilLoopOnSum
    End If
    Erase tlPcf
    Exit Function
End Function

'           Create the CPM Podcast records to output on the Contract/proposal report
'           Records created:  Type 9 = Each individual ID
'                                  10 = Research totals by vehicle (showing cpm/impressions/cost)
'                                  11 = Billing summary by vehicle ($)
'           gWriteBR_CPM
'           <Input>  hlCbf - CBF file handle
'                    hlRdf - Daypart file handle
'                    hlCbf - CBF buffer
'                    tlBR - record of report selections
'                    llSTdStartDates() - array of standard start dates for Standard Billing projection
'                    tlCPM_IDs() - array of CPM records this contract
'                    tlCPMSummary() - array of CPM vehicle totals (cost/impressions/cpm)
'
Public Sub gWriteBR_CPM(hlCbf As Integer, hlRdf As Integer, tlChf As CHF, tlCbf As CBF, ilShowProof As Integer, tlCPM_IDs() As CPM_BR, tlCPMSummary() As CPM_BR)
    Dim illoop As Integer
    Dim ilRet As Integer
    Dim tlRdf As RDF
    Dim tlRdfSrchKey As INTKEY0
    Dim llEarliestDate As Long
    Dim llLatestDate As Long
    Dim ilTotalMonths As Integer
    Dim llTempEarliestDate As Long
    Dim llTempLatestDate As Long
    Dim slEarliestDate As String
    Dim slLatestDate As String
    Dim llAvgCost As Long
    Dim ilMonth As Integer
    Dim llLineRef As Long

    'Create all the detail records to show for Podcast items (shown on a separate page, not merged on the Schedule line output)
    For illoop = LBound(tlCPM_IDs) To UBound(tlCPM_IDs) - 1
        If ((ilShowProof = True) And (tlCPM_IDs(illoop).sType = "H")) Or (tlCPM_IDs(illoop).sType = "P" Or tlCPM_IDs(illoop).sType = "S") Then                'byass the cancel before start IDs to hide
            If tlCPM_IDs(illoop).sHideCBS <> "Y" Then               'if this is a cbs, hide it?
                tlCbf.lChfCode = tlCPM_IDs(illoop).lChfCode         'internal contract code
                tlCbf.iVefCode = tlCPM_IDs(illoop).iVefCode
                tlCbf.lLineNo = tlCPM_IDs(illoop).iPodCPMID         'CPM ID #
                tlCbf.lCPM = tlCPM_IDs(illoop).lCPM
                tlCbf.lGrImp = tlCPM_IDs(illoop).lImpressions
                tlCbf.lRate = tlCPM_IDs(illoop).lCost
                tlCbf.sType = tlCPM_IDs(illoop).sType
                tlCbf.iStartDate(0) = tlCPM_IDs(illoop).iStartDate(0)
                tlCbf.iStartDate(1) = tlCPM_IDs(illoop).iStartDate(1)
                tlCbf.iEndDate(0) = tlCPM_IDs(illoop).iEndDate(0)
                tlCbf.iEndDate(1) = tlCPM_IDs(illoop).iEndDate(1)
                tlCbf.sDysTms = ""
                tlRdfSrchKey.iCode = tlCPM_IDs(illoop).iRdfCode  ' Daypart File Code
                ilRet = btrGetEqual(hlRdf, tlRdf, Len(tlRdf), tlRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    tlCbf.sDysTms = Trim$(tlRdf.sName)
                End If
                tlCbf.sPriceType = tlCPM_IDs(illoop).sPriceType         'c = cpm, f = flat rate
                tlCbf.lLineComment = tlCPM_IDs(illoop).lCxfCode         'ID comment
                tlCbf.lRafCode = tlCPM_IDs(illoop).lRafCode             'pod target name
                tlCbf.iPctDist = tlCPM_IDs(illoop).iCopyTypeMnfCode      'copy type
                tlCbf.sLineType = tlCPM_IDs(illoop).sType                'std, pkg or hidden ID
                tlCbf.lValue(0) = tlCPM_IDs(illoop).iLen 'Boostr Phase 2: Add new "Length" column next to "Ad Location"
                'header info (advt, agy, etc) has already been set into the output fields
                For ilMonth = 0 To 12
                    tlCbf.lMonth(ilMonth) = tlCPM_IDs(illoop).lMonthlyCost(ilMonth) / 100
                Next ilMonth
                If tlCPM_IDs(illoop).sType = "H" Then
                    llLineRef = tlCPM_IDs(illoop).iPkCPMID
                Else
                    llLineRef = tlCPM_IDs(illoop).iPodCPMID
                End If
                gResortLine tlCbf, tlCPM_IDs(illoop).sType, llLineRef
                tlCbf.iExtra2Byte = 9                               'cpm items
                ilRet = btrInsert(hlCbf, tlCbf, Len(tlCbf), INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Sub
                End If
            End If
        End If
        
    Next illoop
    
    'Create all the vehicle summary records to show on the Research summary page
    For illoop = LBound(tlCPMSummary) To UBound(tlCPMSummary) - 1
        If ((ilShowProof = True) And (tlCPMSummary(illoop).sType = "H")) Or (tlCPMSummary(illoop).sType = "P") Or (tlCPMSummary(illoop).sType = "S") Then                'byass the cancel before start IDs to hide
            tlCbf.iVefCode = tlCPMSummary(illoop).iVefCode
            tlCbf.lCPM = tlCPMSummary(illoop).lCPM
            tlCbf.lGrImp = tlCPMSummary(illoop).lImpressions
            tlCbf.lRate = tlCPMSummary(illoop).lCost
            tlCbf.sLineType = tlCPMSummary(illoop).sType
            If tlCPMSummary(illoop).lImpressions <> 0 Then
                'TTP 10513 - Proposals/Contract report: Overflow when ad server impressions are too high, TTP 10514 - Invoice: overflow error when impressions too high
                'Issue:
                'tlCbf.lCPM = (tlCPMSummary(ilLoop).lCost * 1000) / tlCPMSummary(ilLoop).lImpressions
                'Fix #1
                'tlCbf.lCPM = (tlCPMSummary(ilLoop).lCost) / (tlCPMSummary(ilLoop).lImpressions / 1000)
                'Fix #2
                'tlCbf.lCPM = ((1000 * CSng(tlCPMSummary(ilLoopOnSum).lCost)) / tlCPMSummary(ilLoopOnSum).lImpressions)    'cpm xxxxx.xx
                'Fix #3
                tlCbf.lCPM = tlCPMSummary(illoop).lCPM 'Already computed
            End If
            tlCbf.sPriceType = tlCPMSummary(illoop).sPriceType
            '3-31-21 total impressions from air time; total gross from airtime so combined cpm can be calculated on cpm summary
            'place this value in every adserver line so it can be retrieved in the research summary
            tlCbf.lCntGrimps = gStrDecToLong(sgAirTimeGrimp, 0)
            tlCbf.lPop = lgAirTimeGross     'air time gross (excl ntr)
            tlCbf.iExtra2Byte = 10                                      'vehicle Podcast summary
            ilRet = btrInsert(hlCbf, tlCbf, Len(tlCbf), INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Sub
            End If
        End If
    Next illoop

    'billing summary
    For illoop = LBound(tlCPMSummary) To UBound(tlCPMSummary) - 1
        If ((ilShowProof = True) And (tlCPMSummary(illoop).sType = "H")) Or (tlCPMSummary(illoop).sType = "P" Or tlCPMSummary(illoop).sType = "S") Then                'byass the cancel before start IDs to hide
            tlCbf.iVefCode = tlCPMSummary(illoop).iVefCode
            tlCbf.sLineType = tlCPMSummary(illoop).sType             'hidden, std, pkg
            For ilMonth = 0 To 11
                tlCbf.lMonth(ilMonth) = tlCPMSummary(illoop).lMonthlyCost(ilMonth + 1)
            Next ilMonth
            tlCbf.lMonth(12) = tlCPMSummary(illoop).lMonthlyCost(13)
            
            tlCbf.iExtra2Byte = 11                                      'billing summary
            ilRet = btrInsert(hlCbf, tlCbf, Len(tlCbf), INDEXKEY0)
        End If
    Next illoop
    Exit Sub
End Sub

Public Sub gResortLine(tlCbf As CBF, slType As String, llLineRef As Long)
'
'           mSetResortField - tlCbf.sResort is used in sorting the output in the proposal/contract print.
'           This has to be set for all lines, including Cancel Before start lines
'           '12-22-20  Modify to handle resort flags for both schedule lines and CPM Line IDs
    Dim slStr As String

    tlCbf.sResort = ""
    tlCbf.sResortType = ""          '5-31-05
    If Trim$(slType) = "H" Then               'hidden
        slStr = Trim$(str$(llLineRef))  'package line ref id stored in this field
        Do While Len(slStr) < 4         '5-31-05 use 4 digit line #s (vs 3 digit line #s)
        slStr = "0" & slStr
        Loop
        tlCbf.sResort = slStr '& "C"
        tlCbf.sResortType = "C"         '5-31-05
    ElseIf Trim$(slType) = "A" Or Trim$(slType) = "O" Or Trim$(slType) = "E" Or Trim$(slType) = "P" Then     'packages: P = package for cpm; all others package for spots
        slStr = Trim$(str$(llLineRef))  'package line # stored in this field
        Do While Len(slStr) < 4         '5-31-05
        slStr = "0" & slStr
        Loop
        tlCbf.sResort = slStr '& "A"
        tlCbf.sResortType = "A"         '5-31-05
    Else                                    'conventionals, all others (fall after package/hiddens)
        tlCbf.sResort = "9999"  '~"
        tlCbf.sResortType = "~"         '5-31-05
    End If

End Sub
'
'           get total months of this PCF (cpm ) ID
'           <input> CPM start date
'                   CPM end date
'           <return> # of months this cpm ID airs
Public Function gObtainMonthsOfCPMID(llCPMStartDate As Long, llCPMEndDate As Long, slCalType As String) As Integer
    Dim ilTotalMonths As Integer
    Dim llEarliestDate As Long
    Dim llLatestDate As Long
    Dim llTempEarliestDate As Long
    Dim llTempLatestDate As Long
    Dim slEarliestDate As String
    Dim slLatestDate As String

    gObtainMonthsOfCPMID = 0
    If llCPMStartDate > llCPMEndDate Then               'all dates are already invoiced, nothing left to average
        Exit Function
    End If
    'Determine # of months, std or cal
    'determine # of calendar months to average across the billing periods
    ilTotalMonths = 0
    llEarliestDate = llCPMStartDate
    llLatestDate = llCPMEndDate
    If llCPMStartDate <= llCPMEndDate Then
        llTempEarliestDate = llEarliestDate
        Do While llTempEarliestDate <= llLatestDate
            ilTotalMonths = ilTotalMonths + 1
            slEarliestDate = Format$(llTempEarliestDate, "ddddd")
            If Trim$(slCalType = "C") Then
'                    If tlChf.sBillCycle = "C" Then
                slEarliestDate = gObtainStartCal(slEarliestDate)
                slLatestDate = gObtainEndCal(slEarliestDate)
            Else                    'std
                slEarliestDate = gObtainStartStd(slEarliestDate)
                slLatestDate = gObtainEndStd(slEarliestDate)
            End If
            llTempLatestDate = gDateValue(slLatestDate)
            llTempEarliestDate = llTempLatestDate + 1
        Loop

        llTempEarliestDate = llEarliestDate
        gObtainMonthsOfCPMID = ilTotalMonths
    End If
    Exit Function
End Function


'Boostr Phase 2: Billed and Booked std broadcast: update to use new daily or monthly method depending on Site setting for future periods
'mDeterminePeriodAmountByDaily
'   Inputs:
'       slPeriodStart   - Start date of the month being reported (string Date MM/DD/YYYY)
'       slPeriodEnd     - End date of the month being reported (string Date MM/DD/YYYY)
'       slLineStartDate - Start Date of the PCF Line (string Date MM/DD/YYYY)
'       slLineEndDate   - End Date of the PCF Line (string Date MM/DD/YYYY)
'       slRemainingAmount - The Remaining amount (Line Total minus whats already been Invoiced) (string Amount ####.##)
'   Output:
'       Month Amount (double precision number ####.##)
Public Function mDeterminePeriodAmountByDaily(slLastBilledDate As String, slPeriodStart As String, slPeriodEnd As String, slLineStartDate As String, slLineEndDate As String, llBilledAmount, llTotalAmount As Long) As Double
    Dim dlDailyAmount As Double 'The daily $ Amount
    Dim ilNumberOfDaysRemaining As Integer 'How many days from Invoice Start Date to Line EndDate
    Dim ilNumberOfDaysInPeriod As Integer 'How many days are in this period
    Dim dStartDate As Date 'Temp Start Date
    Dim dEndDate As Date 'Temp End
    
    'Determine how many days remain of this line (beyond what's been billed)
    If DateValue(slLastBilledDate) + 1 > DateValue(slLineStartDate) Then
        dStartDate = DateValue(slLastBilledDate) + 1
    Else
        dStartDate = DateValue(slLineStartDate)
    End If
    dEndDate = DateValue(slLineEndDate)
    ilNumberOfDaysRemaining = DateDiff("d", dStartDate, dEndDate) + 1
    If ilNumberOfDaysRemaining <= 0 Then Exit Function
    
    'Determine how many days of this Line are being invoiced
    dStartDate = IIF(DateValue(slLineStartDate) > gDateValue(slPeriodStart), gDateValue(slLineStartDate), gDateValue(slPeriodStart))
    dEndDate = IIF(DateValue(slLineEndDate) < gDateValue(slPeriodEnd), gDateValue(slLineEndDate), gDateValue(slPeriodEnd))
    ilNumberOfDaysInPeriod = DateDiff("d", dStartDate, dEndDate) + 1
    If ilNumberOfDaysInPeriod <= 0 Then Exit Function
    
    'Determine the daily amount
    dlDailyAmount = ((llTotalAmount - llBilledAmount) / 100) / ilNumberOfDaysRemaining
    
    'Determine the amount to apply to this Period (slPeriodStart - slPeriodEnd)
    mDeterminePeriodAmountByDaily = dlDailyAmount * ilNumberOfDaysInPeriod
    
    Debug.Print "mDeterminePeriodAmountByDaily: "
    Debug.Print " -> Line Dates: " & slLineStartDate & " to " & slLineEndDate
    Debug.Print " -> RemainingAmount: " & Format((llTotalAmount - llBilledAmount) / 100, "#.00")
    Debug.Print " -> NumberOfDaysRemaining: " & ilNumberOfDaysRemaining
    Debug.Print " -> Period: " & slPeriodStart & " to " & slPeriodEnd & " = " & ilNumberOfDaysInPeriod
    Debug.Print " -> DailyAmount: " & dlDailyAmount
    Debug.Print " -> Month Amount: " & mDeterminePeriodAmountByDaily
End Function

