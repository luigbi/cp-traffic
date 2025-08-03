Attribute VB_Name = "RptcrRR"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptcrrr.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text


'Type VEHICLEBOOK
'    iVefCode As Integer         'vehicle code
'    iDnfFirstLink As Integer    'index into first book for this vehicle
'    iDnfLastLink As Integer     'index into last book for this vehicle
'End Type
'Type BOOKLIST
'    sKey As String * 5          'date in string form,left filled with zeroes for sort
'    iDnfCode As Integer         '
'    lStartDate As Long          'start date of book
'End Type
'Type DNFLINKLIST
'    idnfInx As Integer
'End Type
'Type SPOTTYPESORTAD
'    sKey As String * 80 'Office Advertiser Contract
'    tSdf As SDF
'End Type
'

Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0    'CHF record image
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChf As CHF
Dim tmChfAdvtExt() As CHFADVTEXT
Dim hmClf As Integer            'Contract line file handle
Dim tmClfSrchKey As CLFKEY0     'CLF record image
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmDnf As Integer            'Book Name file handle
Dim imDnfRecLen As Integer      '
Dim tmDnf As DNF
Dim tmDnfSrchKey As INTKEY0
Dim hmDrf As Integer            'Demo Research data
Dim imDrfRecLen As Integer
Dim tmDrf As DRF
Dim tmDrfSrchKey1 As DRFKEY1
Dim hmDpf As Integer            'Demo Plus Research data
Dim imDpfRecLen As Integer
Dim tmDpf As DPF
Dim hmDef As Integer
Dim hmRaf As Integer
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim hmSdf As Integer            'Spots file handle
Dim imSdfRecLen As Integer      '
Dim tmSdf As SDF
Dim tmPLSdf() As SPOTTYPESORTAD
Dim tmSdfExt() As SDFEXT
Dim hmSmf As Integer            'Spots file handle
Dim imSmfRecLen As Integer      '
Dim tmSmf As SMF
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVsf As Integer
Dim tmVsf As VSF
Dim imVsfRecLen As Integer
Dim tmVehicleBook() As VEHICLEBOOK   'array of all vehicle and their associated books
Dim tmBookList() As BOOKLIST        'array of book names
Dim tmDnfLinkList() As DNFLINKLIST  'list of book indices associated with a vehicle



'********************************************************************************************
'
'       5-5-00        gCrResearchRev - Prepass for Research Revenue report
'
'
'       produce the pre-pass to show research revenue of active contracts for a selected date span.
'       Allow selectivity to use original research book, latest book or closes book to each spot.
'      Include all demos for each contract, plus base demo or select 1st, 2nd, 3rd, or 4 demos of contract.
'      Show EAch contracts rating, audience, grps, grimps, cpp, cpm & avg unit rate
'
'********************************************************************************************
Sub gCrResearchRev()
ReDim ilNowTime(0 To 1) As Integer    'end time of run
Dim slStr As String
Dim ilRet As Integer
'user entered parametrs
Dim llStart As Long             'active Start Date entered
Dim slStart As String           'active start date entered
Dim llEnd As Long               'active End date entered
Dim slEnd As String             'active end date entered
Dim ilBook As Integer           '0 = closest book to air dates, 1 = specific, 2 = schedule line book
'end of user entered parameters
Dim ilVehicle As Integer        'loop to process spots: gather by one vehicle at a time
Dim llContrCode As Long
'Dim ilSpotLoop As Integer
Dim llSpotLoop As Long
Dim ilOk As Integer
Dim ilMin As Integer
Dim ilMax As Integer
Dim ilFound As Integer
Dim ilDay As Integer
Dim ilActiveCntInx As Integer
Dim llTime As Long              'avail time
Dim llOvStartTime As Long
Dim llOvEndTime As Long
Dim llPop As Long               'pop of demo to be used for calculations
Dim llDemoPop As Long           'pop of 1 spot
Dim llAvgAud As Long            'aud for demo
ReDim ilInputDays(0 To 6) As Integer    'days of week flags for avgaud rtn
Dim illoop As Integer
Dim ilDemoLoop As Integer
Dim ilBookInx As Integer
Dim ilDnfCode As Integer        'book to use for spot obtained
Dim llDnfDate As Long
Dim llDate As Long
Dim slPrice As String
Dim llPrice As Long
Dim ilCkcAll As Integer
Dim llGross As Long
Dim llSpots As Long
Dim ilBookMissing As Integer
Dim ilDemoOK As Integer

Dim ilDemos(0 To 4) As Integer      'each entry represents the 4 demos in hdr, 5th is for base demo (reallcation demo in Site)
                                    'true if OK to include it

Dim llPopEst As Long

'Dim ilSpotSortedInx As Integer
Dim llSpotSortedInx As Long
Dim tlLongTypeBuff As POPLCODENAME
Dim ilUpper As Integer
Dim ilFirstTimeSpot As Integer
ReDim tlSdfExtSort(0 To 0) As SDFEXTSORT
Dim ilAudFromSource As Integer
Dim llAudFromCode As Long
Dim slGrossNet As String * 1       '3-2-20

Dim slContractTypes As String       'Date: 4/3/2020 include/exclude contract types
Dim slHOStatus As String            'Date: 4/7/2020 include/exclude contract status Hold/Orders

'****** following required for gAvgAudToLnResearch, calculated for every spot





    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)
    hmDnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDnf)
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imDnfRecLen = Len(tmDnf)
    hmDrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imDrfRecLen = Len(tmDrf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVsf)
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imVsfRecLen = Len(tmVsf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)

    hmDpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDpf)
        btrDestroy hmDpf
        btrDestroy hmMnf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imDpfRecLen = Len(tmDpf)

    hmDef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmDef
        btrDestroy hmDpf
        btrDestroy hmMnf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmRaf
        btrDestroy hmDef
        btrDestroy hmDpf
        btrDestroy hmMnf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '7-23-01 setup global variable to determine if demo plus info exists
    lgDpfNoRecs = btrRecords(hmDpf)
    If lgDpfNoRecs = 0 Then
        lgDpfNoRecs = -1
    End If

    For illoop = 0 To 4        'assume to include all demos from hdr, plus the base demo (Site)
        ilDemos(illoop) = True
    Next illoop
    If Not RptSelRR!ckcDemo1(0).Value = vbChecked Then    'include primary demo?
        ilDemos(0) = False
    End If
    If Not RptSelRR!ckcDemo234(1).Value = vbChecked Then  'include demos 2, 3 & 4?
        ilDemos(1) = False
        ilDemos(2) = False
        ilDemos(3) = False
    End If
    If Not RptSelRR!ckcDemoBase(2).Value = vbChecked Then 'include base demo from site?
        ilDemos(4) = False
    End If

'    slStr = RptSelRR!edcSelCFrom.Text               'Active Start Date
    slStr = RptSelRR!CSI_CalFrom.Text               'Active Start Date  12-16-19 change to use csi calendar control
    If slStr = "" Then
        slStr = "1/1/1970"                          'get everything for the selected contracts - this is never used for "All advt"
    End If
    llStart = gDateValue(slStr)                     'gather contracts thru this date
    slStr = RptSelRR!edcSelCTo.Text               'Active Start Date
    ilDay = Val(slStr)

    llEnd = llStart + ilDay - 1
    slStart = Format$(llStart, "m/d/yy")   'insure the year is in the format, may not have been entered with the date input
    slEnd = Format$(llEnd, "m/d/yy")

    ilBook = 0                  'default to use book closest to airing
    If RptSelRR!rbcBook(1).Value = True Then        'default book (last used)
        ilBook = 1
    ElseIf RptSelRR!rbcBook(2).Value = True Then        'sched line book
        ilBook = 2
    End If

    If RptSelRR!ckcAll.Value = vbChecked Then
        ilCkcAll = True
    Else
        ilCkcAll = False
    End If
    
    If RptSelRR!rbcGrossNet(0).Value = True Then        '3-2-20 gross
        slGrossNet = "G"
    Else
        slGrossNet = "N"
    End If
   
    ilRet = mBuildTables(ilCkcAll, RptSelRR!lbcSelection(0), RptSelRR!lbcCntrCode, slStart, slEnd, tmChfAdvtExt(), tmVehicleBook(), tmBookList(), tmDnfLinkList())      'build table of active contracts, vehicles, books and the list of associated books with each vehicle
    If ilRet <> 0 Then
        Erase tmVehicleBook, tmDnfLinkList, tmPLSdf


        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmDnf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmMnf)
        Exit Sub
    End If

    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)

    'Date: 4/3/2020 check if include/exclude selected contract types "CVTRQSM"
    If RptSelRR!ckcSelC5(0).Value = vbChecked Then  ' Standard
        slContractTypes = "C"
    End If
    If RptSelRR!ckcSelC5(1).Value = vbChecked Then  ' Reserved
        If slContractTypes = "" Then
            slContractTypes = "V"
        Else
            slContractTypes = slContractTypes + "V"
        End If
    End If
    If RptSelRR!ckcSelC5(2).Value = vbChecked Then  ' Remnant
        If slContractTypes = "" Then
            slContractTypes = "T"
        Else
            slContractTypes = slContractTypes & "T"
        End If
    End If
    If RptSelRR!ckcSelC5(3).Value = vbChecked Then  ' Dr
        If slContractTypes = "" Then
            slContractTypes = "R"
        Else
            slContractTypes = slContractTypes & "R"
        End If
    End If
    If RptSelRR!ckcSelC5(4).Value = vbChecked Then  ' Per Inquiry
        If slContractTypes = "" Then
            slContractTypes = "Q"
        Else
            slContractTypes = slContractTypes & "Q"
        End If
    End If
    If RptSelRR!ckcSelC5(5).Value = vbChecked Then  ' PSA
        If slContractTypes = "" Then
            slContractTypes = "S"
        Else
            slContractTypes = slContractTypes & "S"
        End If
    End If
    If RptSelRR!ckcSelC5(6).Value = vbChecked Then  ' Promo
        If slContractTypes = "" Then
            slContractTypes = "M"
        Else
            slContractTypes = slContractTypes & "M"
        End If
    End If
    
    'Date: 4/7/2020 added to include/exclude contract status: Hold and Orders
    
    slHOStatus = ""
    If RptSelRR!ckcSelC5(7).Value = vbChecked Then  ' include/exclude scheduled/unscheduled Hold
        slHOStatus = "HG"
    End If
    If RptSelRR!ckcSelC5(8).Value = vbChecked Then  ' include/exclude scheduled/unscheduled Orders
        If slHOStatus = "" Then
            slHOStatus = "ON"
        Else
            slHOStatus = slHOStatus & "ON"
        End If
    End If
    
    tmClf.lChfCode = 0          'initialize for first time thru and re-entrant
    For ilActiveCntInx = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1

        llContrCode = tmChfAdvtExt(ilActiveCntInx).lCode
        tmChfSrchKey.lCode = llContrCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> 0 Then
            Exit Sub
        End If
        ilOk = True
        'find all the spots for this contract for all vehicles between selected dates

        ilRet = gObtainCntrSpot(-1, False, llContrCode, -1, "S", slStart, slEnd, tlSdfExtSort(), tmSdfExt(), 1, False)  'search for spots between requested user dates


        For ilDemoLoop = 0 To 4     'up to 5 demos can be reported (4 from contr hdr & the base demo from site)
            ilDemoOK = True
            If ilDemoLoop < 4 Then     'test one of 4 demos from contr
                If (ilDemos(ilDemoLoop) = True And tmChf.iMnfDemo(ilDemoLoop) = 0) Or Not (ilDemos(ilDemoLoop)) Then
                    ilDemoOK = False
                End If
            Else
                If (ilDemoLoop = 4 And tgSpf.iReallMnfDemo = 0) Or Not (ilDemos(ilDemoLoop)) Then
                    ilDemoOK = False
                End If
            End If
            If RptSelRR!rbcSortBy(2).Value = True And tmChf.iMnfComp(0) = 0 Then  'test if product protect exists to report it if selected
                ilDemoOK = False
            End If
            If RptSelRR!rbcSortBy(3).Value = True And tmChf.iMnfBus = 0 Then   'test if business catgory exits to report it if selected
                ilDemoOK = False
            End If
            If ilDemoOK Then
                If ilDemoLoop = 4 Then                             'base demo from site
                'use tmgrf.icode to send to research routines
                    tmGrf.iCode2 = tgSpf.iReallMnfDemo
                Else
                    tmGrf.iCode2 = tmChf.iMnfDemo(ilDemoLoop)          'demo from cnt
                End If

                'initialize variables for each time new contract
                ilFirstTimeSpot = True
                tmClf.iLine = -1
                tmClf.iVefCode = -1
                llPop = -1
                ilBookMissing = 0
                tmGrf.iVefCode = -1
                llGross = 0
                llSpots = 0
                'ReDim llWklyspots(1 To 1) As Integer
                'ReDim llWklyAvgAud(1 To 1) As Long
                'ReDim llWklyRates(1 To 1) As Long
                'ReDim llWklyPopEst(1 To 1) As Long
                ReDim llWklyspots(0 To 0) As Long
                ReDim llWklyAvgAud(0 To 0) As Long
                ReDim llWklyRates(0 To 0) As Long
                ReDim llWklyPopEst(0 To 0) As Long

                'For ilSpotLoop = 0 To UBound(tlSdfExtSort) - 1      'loop thru the spots gathered, bypass missed spots & psa/promo/per inq/ DR/ reservations/remnant
                For llSpotLoop = 0 To UBound(tlSdfExtSort) - 1      'loop thru the spots gathered, bypass missed spots & psa/promo/per inq/ DR/ reservations/remnant
                    'filter out unwanted contract types .  Spot List (tmPlSdf is in Contract code, line & date order)
                    'ilSpotSortedInx = tlSdfExtSort(ilSpotLoop).iSdfExtIndex
                    llSpotSortedInx = tlSdfExtSort(llSpotLoop).lSdfExtIndex

                    ilOk = True     'assume ok
                    
                    'Date: 4/6/2020 include/exclude selected contract types (CVTRQSM): 4/7/2020 include/exclude selected contract status (Hold/Orders "HGON")
                    If ((InStr(1, slContractTypes, tmChfAdvtExt(ilActiveCntInx).sType) <= 0) Or (InStr(1, slHOStatus, tmChfAdvtExt(ilActiveCntInx).sStatus) <= 0) Or (tmSdfExt(llSpotSortedInx).sSchStatus = "M")) Then
                         ilOk = False
                    End If
                    
                    'always ignore: psa, promo, reserveration, PI, DR, and Remnants & missed spots
'                    If (tmChfAdvtExt(ilActiveCntInx).sType = "S") Or (tmChfAdvtExt(ilActiveCntInx).sType = "M") Or (tmChfAdvtExt(ilActiveCntInx).sType = "V") Or (tmChfAdvtExt(ilActiveCntInx).sType = "T") Or (tmChfAdvtExt(ilActiveCntInx).sType = "R") Or (tmChfAdvtExt(ilActiveCntInx).sType = "Q") Or (tmSdfExt(llSpotSortedInx).sSchStatus = "M") Then
'                        ilOk = False
'                    End If
                    If ilOk Then            'only test to reread line of contract is OK to use (not an excluded type such as PI, DR, etc)
                        ilRet = 0
                        'read the spot to get the spot price info
                        'tlLongTypeBuff.lCode = tmSdfExt(ilSpotSortedInx).lCode
                        tlLongTypeBuff.lCode = tmSdfExt(llSpotSortedInx).lCode
                        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tlLongTypeBuff, INDEXKEY3, BTRV_LOCK_NONE)

                        If ilRet <> BTRV_ERR_NONE Then
                            'No spot found for the code
                        End If
                        If tmGrf.iVefCode <> tmSdf.iVefCode Then
                            If ilFirstTimeSpot Then         'get the research for the vehicle
                                ilFirstTimeSpot = False
                            Else
                                If RptSelRR!rbcSortBy(1).Value = True Then          'option by vehicle
                                     'write out the vehicle totals if sort by vehicle; otherwise
                                     'combine all vehicles data for the contract

                                     'llPOP contains the population to use for the vehicle
                                     mAvgResearch llPop, llWklyPopEst(), llWklyspots(), llWklyRates(), llWklyAvgAud(), llSpots

                                     'initialize for the next vehicle if by vehicle; otherwise next contract will be processed
                                     llSpots = 0
                                     llGross = 0
                                     llPop = -1
                                    'ReDim llWklyspots(1 To 1) As Integer
                                    'ReDim llWklyAvgAud(1 To 1) As Long
                                    'ReDim llWklyRates(1 To 1) As Long
                                    'ReDim llWklyPopEst(1 To 1) As Long
                                    ReDim llWklyspots(0 To 0) As Long
                                    ReDim llWklyAvgAud(0 To 0) As Long
                                    ReDim llWklyRates(0 To 0) As Long
                                    ReDim llWklyPopEst(0 To 0) As Long

                                End If
                            End If
                            'get the correctline for this spot
                            'tmClfSrchKey.lChfCode = tmSdfExt(ilSpotSortedInx).lChfCode
                            tmClfSrchKey.lChfCode = tmSdfExt(llSpotSortedInx).lChfCode
                            'tmClfSrchKey.iLine = tmSdfExt(ilSpotSortedInx).iLineNo
                            tmClfSrchKey.iLine = tmSdfExt(llSpotSortedInx).iLineNo
                            tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                            tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                            Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                                ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop

                            tmGrf.iVefCode = tmSdf.iVefCode
                        Else                    'same vehicle, check to see if schedule line changed
                            'If tmClf.iLine <> tmSdfExt(ilSpotSortedInx).iLineNo Then
                            If tmClf.iLine <> tmSdfExt(llSpotSortedInx).iLineNo Then
                                'get the correctline for this spot
                                'tmClfSrchKey.lChfCode = tmSdfExt(ilSpotSortedInx).lChfCode
                                tmClfSrchKey.lChfCode = tmSdfExt(llSpotSortedInx).lChfCode
                                'tmClfSrchKey.iLine = tmSdfExt(ilSpotSortedInx).iLineNo
                                tmClfSrchKey.iLine = tmSdfExt(llSpotSortedInx).iLineNo
                                tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                                tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                                    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            End If
                        End If

                        'If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdfExt(ilSpotSortedInx).lChfCode) And (tmClf.iLine = tmSdfExt(ilSpotSortedInx).iLineNo) Then
                        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdfExt(llSpotSortedInx).lChfCode) And (tmClf.iLine = tmSdfExt(llSpotSortedInx).iLineNo) Then

                            ilRet = gGetFlightPrice(tmSdf, tmClf, hmCff, hmSmf, slPrice)  'get spot price here so that the flight can be accessed in cff
                            gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), True, llOvStartTime
                            llOvEndTime = llOvStartTime           'use the avail time for overrides to determine audience
                            For illoop = 0 To 6                 'init all days to not airing, setup for research results later
                                ilInputDays(illoop) = False
                            Next illoop
                            'set day of week aired
                            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                            slStr = Format(llDate, "m/d/yy")
                            illoop = gWeekDayStr(slStr)     'day index

                            ilInputDays(illoop) = True     'set day of week in week pattern
                            'Determine whether to use the default book or the book closest to airing
                            'Set Default book of vehicle in case not found
                            ilDnfCode = 0

                            illoop = gBinarySearchVef(tmClf.iVefCode)
                            If illoop <> -1 Then
                                ilDnfCode = tgMVef(illoop).iDnfCode
                            End If

                            'find the vehicle to see which book should be used
                            For ilVehicle = LBound(tmVehicleBook) To UBound(tmVehicleBook) - 1
                                If tmVehicleBook(ilVehicle).iVefCode = tmClf.iVefCode Then
                                    Exit For
                                End If
                            Next ilVehicle
                            'default book has been set in case there isnt a book closest to airing found
                            If ilBook = 0 Then      'use closest to airing (vs default book)
                                ilFound = False
                                ilMin = tmVehicleBook(ilVehicle).lDnfFirstLink  'iDnfFirstLink
                                ilMax = tmVehicleBook(ilVehicle).lDnfLastLink      'iDnfLastLink
                                'The LinkList points to the associated tmBookList entry of this vehicle
                                If ilMin <> -1 Then          '-1 indicates no books found
                                    For illoop = ilMax To ilMin Step -1
                                    'For ilLoop = ilMin To ilMax
                                        ilBookInx = tmDnfLinkList(illoop).idnfInx
                                        llDnfDate = tmBookList(ilBookInx).lStartDate
                                        If llDate > llDnfDate And llDnfDate <> 0 Then
                                            ilDnfCode = tmBookList(ilBookInx).iDnfCode
                                            ilFound = True
                                            Exit For
                                        End If
                                    Next illoop
                                End If
                                If Not ilFound Then
                                    ilBookMissing = 1   'Flag error, 1 vehiclemissing an associated book
                                End If
                            ElseIf ilBook = 2 Then          'use the line book for debugging, only if one exists; else use the default book
                                If tmClf.iDnfCode <> 0 Then
                                    ilDnfCode = tmClf.iDnfCode
                                    If tmClf.iStartTime(0) = 1 And tmClf.iStartTime(1) = 0 Then
                                        llOvStartTime = 0
                                        llOvEndTime = 0
                                    Else
                                        'override times exist
                                        gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llOvStartTime
                                        gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llOvEndTime
                                    End If

                                    If tgPriceCff.sDyWk = "W" Then            'weekly
                                        For ilDay = 0 To 6 Step 1
                                            If tgPriceCff.iDay(ilDay) > 0 Or tgPriceCff.sXDay(ilDay) = "1" Then
                                                ilInputDays(ilDay) = True
                                            End If
                                        Next ilDay
                                    Else                                        'daily
                                        For ilDay = 0 To 6 Step 1
                                            If tgPriceCff.iDay(ilDay) > 0 Then
                                                ilInputDays(ilDay) = True
                                            End If
                                        Next ilDay
                                    End If
                                Else
                                    ilBookMissing = 1   'Flag error, 1 vehiclemissing an associated book
                                End If
                            End If

                            llPrice = 0     'init incase a decimal number isnt in price field (its adu, nc, fill,etc.)
                            If (InStr(slPrice, ".") <> 0) Then        'found spot cost
                                llPrice = gStrDecToLong(slPrice, 2)
                            End If
                            '3-2-20 option gross or net
                            llPrice = gGetGrossOrNetFromRate(llPrice, slGrossNet, tmChf.iAgfCode)

                            llSpots = llSpots + 1
                            llGross = llGross + llPrice
                            ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, 0, tmGrf.iCode2, llDemoPop)
                            'Set if varying populations within/across schedule lines in contract

                             If llPop < 0 Then       'never been set up, save population first time
                                If llDemoPop <> 0 Then
                                    llPop = llDemoPop
                                End If
                            ElseIf llPop <> 0 Then
                                If llPop <> llDemoPop And llDemoPop <> 0 Then
                                    llPop = 0
                                End If
                            End If
                            '************ DEBUGGING ONLY   *******************
                            'llOvStartTime = 0         'use daypart times only, no overrides
                            'llOvEndTime = 0
                            ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, tmClf.iVefCode, 0, tmGrf.iCode2, llDate, llDate, tmClf.iRdfCode, llOvStartTime, llOvEndTime, ilInputDays(), tmClf.sType, tmClf.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)

                            'need to get grp from gAvgAudToLnResearch
                            ilUpper = UBound(llWklyspots)
                            llWklyspots(ilUpper) = 1
                            llWklyAvgAud(ilUpper) = llAvgAud
                            llWklyRates(ilUpper) = llPrice
                            llWklyPopEst(ilUpper) = llPopEst
                            ilUpper = ilUpper + 1
                            'ReDim Preserve llWklyspots(1 To ilUpper) As Integer
                            'ReDim Preserve llWklyAvgAud(1 To ilUpper) As Long
                            'ReDim Preserve llWklyRates(1 To ilUpper) As Long
                            'ReDim Preserve llWklyPopEst(1 To ilUpper) As Long
                            ReDim Preserve llWklyspots(0 To ilUpper) As Long
                            ReDim Preserve llWklyAvgAud(0 To ilUpper) As Long
                            ReDim Preserve llWklyRates(0 To ilUpper) As Long
                            ReDim Preserve llWklyPopEst(0 To ilUpper) As Long
                        End If
                    End If
                Next llSpotLoop

                tmGrf.iVefCode = tmClf.iVefCode
                mAvgResearch llPop, llWklyPopEst(), llWklyspots(), llWklyRates(), llWklyAvgAud(), llSpots
                'initialize for the next vehicle if by vehicle; otherwise next contract will be processed
                llSpots = 0
                llGross = 0
                llPop = -1
                'ReDim llWklyspots(1 To 1) As Integer
                'ReDim llWklyAvgAud(1 To 1) As Long
                'ReDim llWklyRates(1 To 1) As Long
                'ReDim llWklyPopEst(1 To 1) As Long
                ReDim llWklyspots(0 To 0) As Long
                ReDim llWklyAvgAud(0 To 0) As Long
                ReDim llWklyRates(0 To 0) As Long
                ReDim llWklyPopEst(0 To 0) As Long
            End If                          'ildemos(ildemoloop) = true
        Next ilDemoLoop                     'for ilDemoloop = 0 to 4
    Next ilActiveCntInx



    'debugging only for time program took to run
    slStr = Format$(gNow(), "h:mm:ssAM/PM")       'end time of run
    gPackTime slStr, ilNowTime(0), ilNowTime(1)
    gUnpackTimeLong ilNowTime(0), ilNowTime(1), False, llTime
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, llPop   'start time of run
    llPop = llPop - llTime              'time in seconds in runtime
    ilRet = gSetFormula("RunTime", llPop)  'show how long report generated

    sgCntrForDateStamp = ""     'initialize contract routine next time thru
    Erase tmVehicleBook, tmDnfLinkList, tmBookList, tmPLSdf
    Erase tmChfAdvtExt
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmDrf)
    ilRet = btrClose(hmDnf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmDpf)
    btrDestroy hmRaf
    btrDestroy hmDef
    btrDestroy hmDpf
    btrDestroy hmMnf
    btrDestroy hmVsf
    btrDestroy hmSmf
    btrDestroy hmSdf
    btrDestroy hmVef
    btrDestroy hmDrf
    btrDestroy hmDnf
    btrDestroy hmCHF
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    Exit Sub
mTerminate:
    On Error GoTo 0
    Exit Sub
End Sub
'
'
'           mBinarySearch - find the spots contract code in the list of active
'            contracts to process
'
'           <input> llChfcode - contract code to match against list
'                   tlActiveCnts - list to perform search on
'           <output> ilmin - starting point of list
'                    ilmax - ending point of list
'                    mBinarySearch - index of matching entry
Function mBinarySearch(llChfCode As Long, tlActiveCnts() As ACTIVECNTS, ilMin As Integer, ilMax As Integer) As Integer
Dim ilMiddle As Integer
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If llChfCode = tlActiveCnts(ilMiddle).lChfCode Then
            'found the match
            mBinarySearch = ilMiddle
            Exit Function
        ElseIf llChfCode < tlActiveCnts(ilMiddle).lChfCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearch = -1
End Function
'
'
'       mBuildTAbles - build tables required to process the Delivery information
'
'       Created 5/7/00
'
'       tlActiveCnts() -Active contracts are gathered based on the users start/end dates.  All
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
'       <input>  ilDemo() array of 1-5 representing to include demo1, demo 2 demo 3 and demo 4 of hdr, plus
'                       include base demo from site
'                slstart - user entered active start date
'                slend - user enetered active end date
'       <output>
'                tlChfAdvtExt- array of active contrcts
'                tlVehicleBook - array of books by vehicle
'                tlBookList - array of all books and their start dates
'                tlDnfLinkList
'       <return> mBuildTables - 0 = OK, 1 = error
'
Function mBuildTables(ilCkcAll As Integer, lbcSelection As Control, lbcCntrCode As Control, slStart As String, slEnd As String, tlChfAdvtExt() As CHFADVTEXT, tlVehicleBook() As VEHICLEBOOK, tlBookList() As BOOKLIST, tlDnfLinkList() As DNFLINKLIST) As Integer
 Dim ilUpper As Integer
Dim illoop As Integer
Dim ilVehicle As Integer
Dim ilRet As Integer
Dim slCntrTypes As String
Dim slCntrStatus As String
Dim ilHOState As Integer
Dim slStr As String
Dim slCode As String
Dim ilfirstTime As Integer
Dim llChfStartDate As Long
Dim llChfEndDate As Long
Dim llEnteredStartDate As Long
Dim llEnteredEndDate As Long
Dim tl1DnfLinkList() As DNFLINKLIST         'temp list
Dim ilNewMin As Integer
Dim ilMin As Integer
Dim ilMax As Integer
Dim ilInUseOnly As Integer
Dim ilBookInx As Integer
Dim llDnfDate As Long
Dim llFirstInClosest As Long
Dim ilUpperChf As Integer

    mBuildTables = 0
    llEnteredStartDate = gDateValue(slStart)
    llEnteredEndDate = gDateValue(slEnd)
    'Gather all contracts for previous year and current year whose effective date entered
    'is prior to the effective date that affects either previous year or current year
    
    'Date: 4/13/2020 commented out for gBuildCntTypes doesn't include for PSA and Promo when creating tlChfAdvtExt() array
    'slCntrTypes = gBuildCntTypes()
    slCntrTypes = "CVTRQSM"             'grab all contract types "CVTRQSM" to pass QA testing including PSA and PROMO
    
    slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order.
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)

    'if selective advertiser, get the contracts for that advt
    If Not (ilCkcAll) Then
    
        'ReDim tlChfAdvtExt(1 To 1) As CHFADVTEXT
        ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
        'ilUpperChf = 1
        ilUpperChf = 0
        'Loop on the contract list for the selective contracts
        For illoop = 0 To lbcSelection.ListCount - 1 Step 1
            If lbcSelection.Selected(illoop) Then
                slStr = lbcCntrCode.List(illoop)
                ilRet = gParseItem(slStr, 2, "\", slCode)
                tmChfSrchKey.lCode = Val(slCode)
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> 0 Then
                    mBuildTables = 1
                    Exit Function
                End If
                'ReDim tlChfAdvtExt(1 To 2) As CHFADVTEXT
                'If user date entered, the contract must span it
                gUnpackDateLong tmChf.iStartDate(0), tmChf.iStartDate(1), llChfStartDate
                gUnpackDateLong tmChf.iEndDate(0), tmChf.iEndDate(1), llChfEndDate
                If llChfEndDate >= llEnteredStartDate And llChfStartDate <= llEnteredEndDate Then
                    
                    tlChfAdvtExt(ilUpperChf).lCode = slCode
                    tlChfAdvtExt(ilUpperChf).lCntrNo = tmChf.lCntrNo    'Date: 4/9/2020 wrong assignment for contract number --> tmChf.lCode
                    tlChfAdvtExt(ilUpperChf).sType = tmChf.sType
                    tlChfAdvtExt(ilUpperChf).sStatus = tmChf.sStatus    'Date: 4/9/2020 added Status assignment
                    tlChfAdvtExt(ilUpperChf).iMnfDemo0 = tmChf.iMnfDemo(0)
                    tlChfAdvtExt(ilUpperChf).iStartDate(0) = tmChf.iStartDate(0)
                    tlChfAdvtExt(ilUpperChf).iStartDate(1) = tmChf.iStartDate(1)
                    tlChfAdvtExt(ilUpperChf).iEndDate(0) = tmChf.iEndDate(0)
                    tlChfAdvtExt(ilUpperChf).iEndDate(1) = tmChf.iEndDate(1)
                    'ReDim Preserve tlChfAdvtExt(1 To ilUpperChf + 1) As CHFADVTEXT
                    ReDim Preserve tlChfAdvtExt(0 To ilUpperChf + 1) As CHFADVTEXT
                    ilUpperChf = ilUpperChf + 1
                End If
            End If
        Next illoop
    Else        'all contracts, retrieve active ones based on the dates entered
        ilRet = gObtainCntrForDate(RptSelRR, slStart, slEnd, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
        If ilRet <> 0 Then
            mBuildTables = 1
            Exit Function
        End If
    End If

    ReDim tlVehicleBook(0 To 0) As VEHICLEBOOK
    'ilRet = gObtainVef()         'vehicles have already been gathered in global array tgMVef
    For illoop = LBound(tgMVef) To UBound(tgMVef) - 1
        'look for Active (not dormant) and vehicle type Conventional or Selling vehicle
        If (tgMVef(illoop).sType = "C" Or tgMVef(illoop).sType = "S") Then
            ilUpper = UBound(tlVehicleBook)
            tlVehicleBook(ilUpper).iVefCode = tgMVef(illoop).iCode
            tlVehicleBook(ilUpper).lDnfFirstLink = -1       'iDnfFirstLink
            tlVehicleBook(ilUpper).lDnfLastLink = 0         'iDnfLastLink
            ReDim Preserve tlVehicleBook(0 To ilUpper + 1) As VEHICLEBOOK
        End If
    Next illoop
    ReDim tlBookList(0 To 0) As BOOKLIST        'list of books

    For illoop = 0 To RptSelRR!cbcBook.ListCount - 1 Step 1
        ilUpper = UBound(tlBookList)

        slStr = tgBookNameCode(illoop).sKey
        ilRet = gParseItem(slStr, 2, "\", slCode)
        tlBookList(ilUpper).iDnfCode = Val(slCode)  'book code
        'get the book to store the start date
        tmDnfSrchKey.iCode = Val(slCode)
        ilRet = btrGetGreaterOrEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
        If ilRet <> BTRV_ERR_NONE Then
            mBuildTables = 1
            Exit Function
        End If

        'exclude the book if its past the end date requested
        gUnpackDateLong tmDnf.iBookDate(0), tmDnf.iBookDate(1), tlBookList(ilUpper).lStartDate

        If tlBookList(ilUpper).lStartDate <= llEnteredEndDate Then
            slStr = Trim$(str$(tlBookList(ilUpper).lStartDate))
            Do While Len(slStr) < 5        'left fill zeroes for date sort
                slStr = "0" & slStr
            Loop
            tlBookList(ilUpper).sKey = slStr
            ReDim Preserve tlBookList(0 To ilUpper + 1) As BOOKLIST
        Else
            ilRet = ilRet
        End If
    Next illoop
    If ilUpper > 0 Then    'sort by book date
         ArraySortTyp fnAV(tlBookList(), 0), ilUpper + 1, 0, LenB(tlBookList(0)), 0, LenB(tlBookList(0).sKey), 0
    End If

    If RptSelRR!rbcBook(0).Value = False Then         'not option to use book closest to airing, dont need to build the table
        Exit Function
    End If

    'Build  list of associated books with vehicles
    ReDim tlDnfLinkList(0 To 0) As DNFLINKLIST
    For ilVehicle = 0 To UBound(tlVehicleBook) - 1          'loop thru list of vehicles to set up link list of valid books for this vehicle within date
        ReDim tl1DnfLinkList(0 To 0) As DNFLINKLIST

        ilfirstTime = -1
        For illoop = 0 To UBound(tlBookList) - 1            'valid books for dates
            tmDrfSrchKey1.iDnfCode = tlBookList(illoop).iDnfCode
            tmDrfSrchKey1.sDemoDataType = "D"
            tmDrfSrchKey1.iMnfSocEco = 0
            tmDrfSrchKey1.iVefCode = tlVehicleBook(ilVehicle).iVefCode
            tmDrfSrchKey1.iStartTime(0) = 0
            tmDrfSrchKey1.iStartTime(1) = 0
            tmDrfSrchKey1.sInfoType = "D"
            ilRet = btrGetGreaterOrEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
            If ilRet <> BTRV_ERR_NONE Then
                mBuildTables = 1
                Exit Function
            End If
            If tmDrf.iDnfCode = tlBookList(illoop).iDnfCode Then
                If tmDrf.iVefCode = tlVehicleBook(ilVehicle).iVefCode Then
                    ilUpper = UBound(tl1DnfLinkList)
                    If ilfirstTime < 0 Then
                        tlVehicleBook(ilVehicle).lDnfFirstLink = ilUpper        'iDnfLastLink
                    End If
                    ilfirstTime = ilUpper
                   ' tlDnfLinkList(ilUpper).idnfInx = tmDrf.iDnfCode
                   'The LinkList points to the associated tlBookList entry of this vehicle
                    tl1DnfLinkList(ilUpper).idnfInx = illoop
                    ReDim Preserve tl1DnfLinkList(0 To ilUpper + 1) As DNFLINKLIST
                End If
            Else        'books no longer match
                Exit For
            End If
        Next illoop
        'No more books for the current vehicle, set its last index
        tlVehicleBook(ilVehicle).lDnfLastLink = ilUpper     'iDnfLastLink

        'All books have been built for this vehicle; now get the closest so that all do not
        'have to be searched.  Once the closest date is found, gather all the books for that same date.
        'tl1DnfLinkList contains ALL books forthe vheicle;
        'tlDnfLinkList will contain only the closest books to the airing date for the vehicle
        ilInUseOnly = UBound(tlDnfLinkList)
        ilMin = tlVehicleBook(ilVehicle).lDnfFirstLink      'iDnfFirstLink
        ilMax = tlVehicleBook(ilVehicle).lDnfLastLink       'iDnfLastLink
        llFirstInClosest = -1
        'The LinkList points to the associated tmBookList entry of this vehicle
        If ilMin <> -1 Then          '-1 indicates no books found
            For ilNewMin = ilMax To ilMin Step -1
            'For ilLoop = ilMin To ilMax
                ilBookInx = tl1DnfLinkList(ilNewMin).idnfInx
                llDnfDate = tlBookList(ilBookInx).lStartDate
                '11-23-11 use the end date of requested period to determine which book
                If llEnteredEndDate >= llDnfDate And llDnfDate <> 0 Then   'find the first closest book
                'If llEnteredStartDate > llDnfDate And llDnfDate <> 0 Then   'find the first closest book
                    If llFirstInClosest = -1 Then
                        llFirstInClosest = llDnfDate
                        'ilmax is highest, illoop = lowest
                        'Copy the number of books for this vehicle into list of books by vehicle to process
                        tlVehicleBook(ilVehicle).lDnfFirstLink = UBound(tlDnfLinkList)      'new min   iDnfFirstLink
                        tlDnfLinkList(ilInUseOnly) = tl1DnfLinkList(ilNewMin)
                        ilInUseOnly = ilInUseOnly + 1
                        ReDim Preserve tlDnfLinkList(LBound(tlDnfLinkList) To ilInUseOnly) As DNFLINKLIST
                        'Exit For
                    Else            'get all the books with the same date as the first one it found
                        If llFirstInClosest = llDnfDate And llDnfDate <> 0 Then   'matching book date as the book closest to airing date
                            tlDnfLinkList(ilInUseOnly) = tl1DnfLinkList(ilNewMin)
                            ilInUseOnly = ilInUseOnly + 1
                            ReDim Preserve tlDnfLinkList(LBound(tlDnfLinkList) To ilInUseOnly) As DNFLINKLIST
                        Else
                            Exit For
                        End If
                    End If
                Else
                    'none exist
                    ilNewMin = -1
                    ilMax = 0
                    Exit For
                End If
            Next ilNewMin
            If llFirstInClosest <> -1 Then
                'tlVehicleBook(ilVehicle).iDnfFirstLink = ilNewMin
                tlVehicleBook(ilVehicle).lDnfLastLink = UBound(tlDnfLinkList) - 1   'new max        iDnfLastLink
            End If
        End If

    Next ilVehicle

    Exit Function

    On Error GoTo 0
    mBuildTables = 1
    Exit Function
End Function

Sub mAvgResearch(llPop As Long, llWklyPopEst() As Long, llWklyspots() As Long, llWklyRates() As Long, llWklyAvgAud() As Long, llSpots As Long)
Dim ilRet As Integer
'Dim llGross As Long         'total cost (or spot cost)
Dim dlGross As Double         'total cost (or spot cost)'TTP 10439 - Rerate 21,000,000
Dim llTotalAvgAud As Long      'avg aud per week
Dim ilAVgRtg As Integer        'avg rating
Dim llTotalGrImp As Long        'total grimps
Dim llTotalGRP As Long          'total grps
Dim llTotalCPP As Long          'total CPPS
Dim llTotalCPM As Long          'Total CPMS
Dim ilUpper As Integer
Dim llPopEst As Long

    If llSpots <> 0 Then
        ilUpper = UBound(llWklyspots)
        'ReDim ilWklyRtg(1 To ilUpper) As Integer
        'ReDim llWklyGrimp(1 To ilUpper) As Long
        'ReDim llWklyGRP(1 To ilUpper) As Long
        ReDim ilWklyRtg(0 To ilUpper) As Integer
        ReDim llWklyGrimp(0 To ilUpper) As Long
        ReDim llWklyGRP(0 To ilUpper) As Long
        
        '10-30-14 default to use 1 place rating regardless of agency flag
        'gAvgAudToLnResearch "1", True, llPop, llWklyPopEst(), llWklyspots(), llWklyRates(), llWklyAvgAud(), llGross, llTotalAvgAud, ilWklyRtg(), ilAVgRtg, llWklyGrimp(), llTotalGrImp, llWklyGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst
        gAvgAudToLnResearch "1", True, llPop, llWklyPopEst(), llWklyspots(), llWklyRates(), llWklyAvgAud(), dlGross, llTotalAvgAud, ilWklyRtg(), ilAVgRtg, llWklyGrimp(), llTotalGrImp, llWklyGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
        'write the out Prepass record
        ilRet = mWRiteGrf(ilAVgRtg, llTotalAvgAud, llTotalGRP, llTotalGrImp, llSpots, CLng(dlGross * 100), llTotalCPP, llTotalCPM)
    End If
End Sub
Function mWRiteGrf(ilAVgRtg As Integer, llTotalAvgAud As Long, llTotalGRP As Long, llTotalGrImp As Long, llSpots As Long, llGross As Long, llTotalCPP As Long, llTotalCPM As Long) As Integer
Dim ilRet As Integer
'    tmGrf.lDollars(1) = ilAVgRtg
'    tmGrf.lDollars(2) = llTotalAvgAud
'    tmGrf.lDollars(3) = llTotalGRP
'    tmGrf.lDollars(4) = llTotalGrImp
'    tmGrf.lDollars(5) = llSpots
'    tmGrf.lDollars(6) = llGross
'    tmGrf.lDollars(7) = llTotalCPP
'    tmGrf.lDollars(8) = llTotalCPM
    tmGrf.lDollars(0) = ilAVgRtg
    tmGrf.lDollars(1) = llTotalAvgAud
    tmGrf.lDollars(2) = llTotalGRP
    tmGrf.lDollars(3) = llTotalGrImp
    tmGrf.lDollars(4) = llSpots
    tmGrf.lDollars(5) = llGross
    tmGrf.lDollars(6) = llTotalCPP
    tmGrf.lDollars(7) = llTotalCPM
    tmGrf.lChfCode = tmChf.lCode          'contract code
    tmGrf.iRdfCode = tmChf.iMnfComp(0)       'prod category
    tmGrf.iSlfCode = tmChf.iMnfBus                'bus category

    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)


End Function
