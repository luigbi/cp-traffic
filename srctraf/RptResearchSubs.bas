Attribute VB_Name = "RPTResearchSUBS"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
'       Release 6.0

Option Explicit
Option Compare Text
Dim tmChfSrchKey As LONGKEY0    'CHF record image
Dim tmChfSrchKey1 As CHFKEY1
Dim tmChf As CHF

Dim tmDnfSrchKey As INTKEY0
Dim tmDnf As DNF

Dim tmDrf As DRF
Dim tmDrfSrchKey1 As DRFKEY1

'Not used
'Type WKLYAUDINFO
'    lChfCode As Long
'    iLine As Integer
'    iVefCode As Integer
'    iDrfcode As Integer
'    lPop As Long
'    lPopEst As Long
'    lWklyRate(1 To 53) As Long
'    lWklySpots(1 To 53) As Long
'    lWklyAud(1 To 53) As Long
'End Type

'
'
'       gBuildBookInfo - build Research tables for faster access

'
'       tmActiveCnts() -Active contracts are gathered based on the users start/end dates.  All
'       contracts active during those dates are processed for its contracts start date
'       to its contract end date (all spots). Array is sorted by contract code.  A binary
'       search is used to find the contract associated with each spot.
'
'       tlVehicleBook() - all valid (active conventional and selling vehicles which hold
'       the vehicle code and first and last book name index pointers.  These are the
'       books associated with each vehicle, which point to tlBookList.
'
'       tlBookList() - array of books containing book start date and book code.  This array
'       is sorted by book start date to speed up search.  Each spot needs to find the
'       book closest to airing.
'
'       tlDnfLinkList() - array of indices that point to the tlBookList array associated with a vehicle.
'       tlVehicleBook points to this array.
'
'
Function gBuildBookInfo(Form As Form, hlDnf As Integer, hlDrf As Integer, ilDaysInPast As Integer, tlVehicleBook() As VEHICLEBOOK, tlBookList() As BOOKLIST, tlDnfLinkList() As DNFLINKLIST) As Boolean
Dim ilUpper As Integer
Dim ilLoop As Integer
Dim ilVehicle As Integer
Dim ilRet As Integer
Dim slStr As String
Dim slCode As String
Dim llfirstTime As Long         '1-15-08 chg to long
Dim llChfStartDate As Long
Dim llChfEndDate As Long
Dim llEnteredStartDate As Long
Dim llEnteredEndDate As Long
Dim llUpper As Long                     '1-15-08

   
    gBuildBookInfo = True
    ReDim tlVehicleBook(0 To 0) As VEHICLEBOOK
    'ilRet = gObtainVef()         'vehicles have already been gathered in global array tgMVef
    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1
        'look for Active (not dormant) and vehicle type Conventional or Selling vehicle, or game (added 5-06-08)
        If (tgMVef(ilLoop).sType = "C" Or tgMVef(ilLoop).sType = "S" Or tgMVef(ilLoop).sType = "G") Then
            ilUpper = UBound(tlVehicleBook)         '
            tlVehicleBook(ilUpper).iVefCode = tgMVef(ilLoop).iCode
            tlVehicleBook(ilUpper).lDnfFirstLink = -1
            tlVehicleBook(ilUpper).lDnfLastLink = 0
            ReDim Preserve tlVehicleBook(0 To ilUpper + 1) As VEHICLEBOOK
        End If
    Next ilLoop
    ReDim tlBookList(0 To 0) As BOOKLIST        'list of books

    For ilLoop = 0 To Form!cbcBook.ListCount - 1 Step 1
        ilUpper = UBound(tlBookList)

        slStr = tgBookNameCode(ilLoop).sKey
        ilRet = gParseItem(slStr, 2, "\", slCode)
        tlBookList(ilUpper).iDnfCode = Val(slCode)  'book code
        'get the book to store the start date
        tmDnfSrchKey.iCode = Val(slCode)
        ilRet = btrGetGreaterOrEqual(hlDnf, tmDnf, Len(tmDnf), tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
        If ilRet <> BTRV_ERR_NONE Then
            gBuildBookInfo = False
            Exit Function
        End If
        gUnpackDateLong tmDnf.iBookDate(0), tmDnf.iBookDate(1), tlBookList(ilUpper).lStartDate

        If tlBookList(ilUpper).lStartDate >= (llChfStartDate - ilDaysInPast) Then            '1-15-08 only gather the books whose book date is equal/greater than the
                                                                            'earliest contract start date minus 2 years (365 * 2)
            slStr = Trim$(str$(tlBookList(ilUpper).lStartDate))
            Do While Len(slStr) < 5        'left fill zeroes for date sort
                slStr = "0" & slStr
            Loop
            tlBookList(ilUpper).sKey = slStr
            ReDim Preserve tlBookList(0 To ilUpper + 1) As BOOKLIST
        End If
    Next ilLoop
    If ilUpper > 0 Then    'sort by book date
         ArraySortTyp fnAV(tlBookList(), 0), ilUpper + 1, 0, LenB(tlBookList(0)), 0, LenB(tlBookList(0).sKey), 0
    End If
    'Build  list of associated books with vehicles and demo research link list
    ReDim tlDnfLinkList(0 To 0) As DNFLINKLIST
    For ilVehicle = 0 To UBound(tlVehicleBook) - 1
        llfirstTime = -1            '1-15-08 chg to long
        For ilLoop = 0 To UBound(tlBookList) - 1
            tmDrfSrchKey1.iDnfCode = tlBookList(ilLoop).iDnfCode
            tmDrfSrchKey1.sDemoDataType = "D"
            tmDrfSrchKey1.iMnfSocEco = 0
            tmDrfSrchKey1.iVefCode = tlVehicleBook(ilVehicle).iVefCode
            tmDrfSrchKey1.iStartTime(0) = 0
            tmDrfSrchKey1.iStartTime(1) = 0
            tmDrfSrchKey1.sInfoType = "D"
            ilRet = btrGetGreaterOrEqual(hlDrf, tmDrf, Len(tmDrf), tmDrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
            If ilRet <> BTRV_ERR_NONE Then
                gBuildBookInfo = False
                Exit Function
            End If
            If tmDrf.iVefCode = tlVehicleBook(ilVehicle).iVefCode And tmDrf.iDnfCode = tlBookList(ilLoop).iDnfCode Then
                llUpper = UBound(tlDnfLinkList)     '1-15-08 chg to long
                If llfirstTime < 0 Then
                    tlVehicleBook(ilVehicle).lDnfFirstLink = llUpper
                End If
                llfirstTime = llUpper
               ' tlDnfLinkList(ilUpper).idnfInx = tmDrf.iDnfCode
               'The LinkList points to the associated tlBookList entry of this vehicle
                tlDnfLinkList(llUpper).idnfInx = ilLoop
                ReDim Preserve tlDnfLinkList(0 To llUpper + 1) As DNFLINKLIST
            End If
        Next ilLoop
        'No more books for the current vehicle, set its last index
        tlVehicleBook(ilVehicle).lDnfLastLink = llUpper
    Next ilVehicle

    Exit Function
End Function


