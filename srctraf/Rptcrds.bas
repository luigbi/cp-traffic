Attribute VB_Name = "RPTCRDS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrds.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmSmfSrchKey                                                                          *
'******************************************************************************************

Option Explicit
Option Compare Text
'Public igYear As Integer                'budget year used for filtering
'Public igMonthOrQtr As Integer          'entered month or qtr
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
Dim lmSingleCntr As Long        'single contract #
Dim imSelLists() As Integer   'array of advt, agy, mnf (product protection), or slsp codes
Dim imSelList As Integer      '0 = all, 1 = selective adv, 2 = agy, 3= prod prot, 4 = slsp
Dim lmSTime As Long           'user entered start time (or 12AM)
Dim lmETime As Long           'user entered end time (or 12M)
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVsf As Integer            'Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim tmClfSrchKey As CLFKEY0     'CLF record image
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmSdf As Integer            'Spot detail file handle
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSdf As SDF
Dim tmSdfSrchKey3 As LONGKEY0     'SDF record image (SDF code as keyfield)
Dim hmSsf As Integer
Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim hmSmf As Integer            'MG file handle
Dim tmSmf As SMF                'SMF record image
Dim tmSmfSrchKey2 As LONGKEY0
Dim imSmfRecLen As Integer        'SMF record length
'Log Calendar
Dim hmLcf As Integer            'Log Calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim imLcfRecLen As Integer        'LCF record length
Dim tmCbf As CBF
Dim hmCbf As Integer
Dim imCbfRecLen As Integer        'CBF record length
Dim tmCxf As CXF
Dim hmCxf As Integer
Dim tmCxfSrchKey  As LONGKEY0         'Gen date and time
Dim imCxfRecLen As Integer        'CXF record length
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim hmRdf As Integer            'Dayparts file handle
Dim imRdfRecLen As Integer      'RD record length
Dim tmRdfSrchKey As INTKEY0     'RDF key image
Dim tmRdf As RDF

Dim tmFsf As FSF
Dim hmFsf As Integer            'Feed Spots  file handle
Dim imFsfRecLen As Integer      'FSF record length

Dim tmAnf As ANF
Dim hmAnf As Integer            'Feed Spots  file handle
Dim imAnfRecLen As Integer
'*******************************************************************
'*                                                                 *
'*      Procedure Name:gCRQtrlyBookSpots                           *
'*                                                                 *
'*             Created:12/29/97      By:D. Hosaka                  *
'*            Modified:              By:                           *
'*                                                                 *
'*            Comments: Generate Daily Spot Report                 *
'*                                                                 *
'*            DH: Created 10/9/2000
'
'*                                                                 *
'*******************************************************************
Sub gCreateDS()
'
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim llEDate As Long         'end date entered
    Dim slTime As String
    Dim ilVehicle As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim llEffDate As Long
    Dim llDateEntered As Long                   'orig user date entered (backed up to a Monday)
    Dim tlCntTypes As CNTTYPES
    Dim ilCkcAllOthers As Integer               'true if selectivity on agy, adv, prod protection or slsp

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmCbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCbf
        btrDestroy hmCHF
    Exit Sub
    End If
    imCbfRecLen = Len(tmCbf)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmCbf
        btrDestroy hmCHF
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCbf
        btrDestroy hmCHF
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCbf
        btrDestroy hmCHF
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
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
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmSsf
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
        btrDestroy hmLcf
        btrDestroy hmSsf
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    imLcfRecLen = Len(tmLcf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmSmf
        btrDestroy hmLcf
        btrDestroy hmSsf
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCbf
        btrDestroy hmCHF
        Exit Sub
    End If
    imVsfRecLen = Len(tmVsf)
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmRdf
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCbf
        btrDestroy hmCHF
        Exit Sub
    End If
    imRdfRecLen = Len(tmRdf)

    hmCxf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCxf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCxf
        btrDestroy hmRdf
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCbf
        btrDestroy hmCHF
        Exit Sub
    End If
    imCxfRecLen = Len(tmCxf)

    hmFsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmCxf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmFsf
        btrDestroy hmCxf
        btrDestroy hmRdf
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCbf
        btrDestroy hmCHF
        Exit Sub
    End If
    imFsfRecLen = Len(tmFsf)

    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmCxf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAnf
        btrDestroy hmFsf
        btrDestroy hmCxf
        btrDestroy hmRdf
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmCbf
        btrDestroy hmCHF
        Exit Sub
    End If
    imAnfRecLen = Len(tmAnf)


    tlCntTypes.iHold = gSetCheck(RptSelDS!ckcCType(0).Value)
    tlCntTypes.iOrder = gSetCheck(RptSelDS!ckcCType(1).Value)
    tlCntTypes.iNetwork = gSetCheck(RptSelDS!ckcCType(2).Value)
    tlCntTypes.iStandard = gSetCheck(RptSelDS!ckcCType(3).Value)
    tlCntTypes.iReserv = gSetCheck(RptSelDS!ckcCType(4).Value)
    tlCntTypes.iRemnant = gSetCheck(RptSelDS!ckcCType(5).Value)
    tlCntTypes.iDR = gSetCheck(RptSelDS!ckcCType(6).Value)
    tlCntTypes.iPI = gSetCheck(RptSelDS!ckcCType(7).Value)
    tlCntTypes.iPSA = gSetCheck(RptSelDS!ckcCType(8).Value)
    tlCntTypes.iPromo = gSetCheck(RptSelDS!ckcCType(9).Value)
    tlCntTypes.iTrade = gSetCheck(RptSelDS!ckcCType(10).Value)
    'tlCntTypes.iMissed = gSetCheck(RptSelDS!ckcSpots(0).Value)
    tlCntTypes.iMG = gSetCheck(RptSelDS!ckcSpots(0).Value)      '10-22-10 index 0 is replaced with MG as it was unused
    tlCntTypes.iCharge = gSetCheck(RptSelDS!ckcSpots(1).Value)
    tlCntTypes.iZero = gSetCheck(RptSelDS!ckcSpots(2).Value)
    tlCntTypes.iADU = gSetCheck(RptSelDS!ckcSpots(3).Value)
    tlCntTypes.iBonus = gSetCheck(RptSelDS!ckcSpots(4).Value)
    tlCntTypes.iXtra = gSetCheck(RptSelDS!ckcSpots(5).Value)
    tlCntTypes.iFill = gSetCheck(RptSelDS!ckcSpots(6).Value)
    tlCntTypes.iNC = gSetCheck(RptSelDS!ckcSpots(7).Value)
    tlCntTypes.iRecapturable = gSetCheck(RptSelDS!ckcSpots(8).Value)
    tlCntTypes.iSpinoff = gSetCheck(RptSelDS!ckcSpots(9).Value)
    tlCntTypes.iFixedTime = gSetCheck(RptSelDS!ckcRank(0).Value)
    tlCntTypes.iSponsor = gSetCheck(RptSelDS!ckcRank(1).Value)
    tlCntTypes.iDP = gSetCheck(RptSelDS!ckcRank(2).Value)
    tlCntTypes.iROS = gSetCheck(RptSelDS!ckcRank(3).Value)
    tlCntTypes.iRated = gSetCheck(RptSelDS!ckcEvents(0).Value)     'include program events
    tlCntTypes.iNonRAted = gSetCheck(RptSelDS!ckcEvents(1).Value)  'include comments
    tlCntTypes.iSuburban = gSetCheck(RptSelDS!ckcEvents(2).Value)  'show open avails only

    'build the spot lengths
    tlCntTypes.iLenHL(0) = Val(RptSelDS!edcLength(0))
    tlCntTypes.iLenHL(1) = Val(RptSelDS!edcLength(1))
    For ilLoop = 0 To 6
        tlCntTypes.iValidDays(ilLoop) = True
        If Not gSetCheck(RptSelDS!ckcDays(ilLoop).Value) Then
            tlCntTypes.iValidDays(ilLoop) = False
        End If
    Next ilLoop

    'get all the dates needed to work with
'    slDate = RptSelDS!edcSelCFrom.Text               'start date entred
    slDate = RptSelDS!csi_CalFrom.Text               'start date entred
    llEffDate = gDateValue(slDate)
    llDateEntered = llEffDate               'save orig date to calculate week index

'    slDate = RptSelDS!edcSelCFrom1.Text               'end date entered
    slDate = RptSelDS!csi_CalTo.Text               'end date entered
    If slDate = "" Then
        llEDate = llDateEntered
    Else
        llEDate = gDateValue(slDate)
    End If

    'Get the Start and End Times
    slTime = RptSelDS!edcSTime.Text
    If slTime = "" Then
        slTime = "12M"
        lmSTime = gTimeToLong(slTime, False)
    Else            'use entered start time
        lmSTime = gTimeToLong(slTime, False)
    End If
    slTime = RptSelDS!edcETime.Text
    If slTime = "" Then
        slTime = "12M"
        lmETime = gTimeToLong(slTime, True)
    Else              'use entered end time
        lmETime = gTimeToLong(slTime, True)
    End If
    slCode = RptSelDS!edcContr
    If slCode = "" Then
        lmSingleCntr = 0
    Else
        lmSingleCntr = CLng(slCode)
    End If
    ReDim imSelLists(0 To 0) As Integer         'array of selective lists
    imSelList = 0
    If RptSelDS!rbcShow(1).Value Then   'advt selection
        imSelList = 1
    ElseIf RptSelDS!rbcShow(2).Value Then   'agy selection
        imSelList = 2
    ElseIf RptSelDS!rbcShow(3).Value Then   'product protection selection
        imSelList = 3
    Else
        imSelList = 4
    End If
    ilCkcAllOthers = False
    If imSelList > 0 Then
        If Not (RptSelDS!ckcAllOthers.Value = vbChecked) Then
            ilCkcAllOthers = True
            For ilLoop = 0 To RptSelDS!lbcSelection(imSelList).ListCount - 1 Step 1
                If (RptSelDS!lbcSelection(imSelList).Selected(ilLoop)) Then
                    If imSelList = 1 Then       'advt
                        slNameCode = tgRptAdvertiserCode(ilLoop).sKey
                    ElseIf imSelList = 2 Then   'agy
                        slNameCode = tgRptAgencyCode(ilLoop).sKey
                    ElseIf imSelList = 3 Then   'prod protection
                        slNameCode = tgRptNameCode(ilLoop).sKey
                    Else                        'slsp
                        slNameCode = tgRptSalespersonCode(ilLoop).sKey
                    End If
                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                    ilRet = gParseItem(slName, 3, "|", slName)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    imSelLists(UBound(imSelLists)) = Val(slCode)
                    ReDim Preserve imSelLists(UBound(imSelLists) + 1) As Integer
                End If
            Next ilLoop
        End If
    End If
    tmVef.iCode = 0
    For ilVehicle = 0 To RptSelDS!lbcSelection(0).ListCount - 1 Step 1
        If (RptSelDS!lbcSelection(0).Selected(ilVehicle)) Then
            slNameCode = tgCSVNameCode(ilVehicle).sKey 'RptSelSP!lbcCSVNameCode.List(ilVehicle)
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
                mBuildDetail ilVefCode, ilVpfIndex, llEffDate, llEDate, tlCntTypes
            End If                              'vpfindex > 0
        End If                                  'vehicle selected
    Next ilVehicle                              'For ilvehicle = 0 To RptSelDS!lbcSelection(0).ListCount - 1
    ilRet = btrClose(hmCxf)
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmCbf)
    ilRet = btrClose(hmLcf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmFsf)
    ilRet = btrClose(hmAnf)
    btrDestroy hmCxf
    btrDestroy hmRdf
    btrDestroy hmSdf
    btrDestroy hmVef
    btrDestroy hmSsf
    btrDestroy hmLcf
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmCbf
    btrDestroy hmCHF
    btrDestroy hmMnf
    btrDestroy hmFsf
    btrDestroy hmAnf
    Erase imSelLists
    Exit Sub
End Sub
'*****************************************************************
'*                                                               *
'*                                                               *
'*                                                               *
'*      Procedure Name:gGetSpotCounts                            *
'*                                                               *
'*      Created:9/27/00       By:D. Hosaka                       *
'*                                                               *
'*                                                               *
'*****************************************************************
Sub mBuildDetail(ilVefCode As Integer, ilVpfIndex As Integer, llSDate As Long, llEDate As Long, tlCntTypes As CNTTYPES)
'
'   Where:
'
'   ilVefCode (I) - vehicle code to process
'   ilVpfIndex (I) - vehicle options pointer
'   llSDate (I) - start date to begin searching Avails
'   llEDate (I) - end date to stop searching avails
'   tlCntTypes (I) - contract and spot types to include in search
'
'   Note: Remnants; Direct Response; per Inquiry; PSA and Promos are not
'         saved with a miss status
'         For scheduled spots the rank is used to determine if it is one
'         of the above (Direct reponse=1010; Remnant=1020; per Inquiry= 1030;
'         PSA=1060; Promo=1050.
'
                '   CbfFields
                '   3 types of records:  1 = program, 2 = avail, 3 = spot
                '      all types have the folliwng keys:
                '      GenDate - Generation date
                '      GenTime - Generation Time
                '      VefCode - vehicle code
                '      DtFrstBkt - SDF Date
                '      Time - time of event
                '      iPropVer - 0 = program event, 1 = all others (avail & spots).  all spots within an avail get "boxed" together; this field
                '                 required for sorting in Crystal
                '      sLineType =  1 = program, 2 = avail, 3 = spot
                '      Extra2Byte - Seq # for multiple spots in the same avail
                '  Program records (type 1)
                '      agfCode - Event name code (program name)
                '      iPropOrdTime- End time of program
                '      lIntComment -program comment
                '      sLineType = 1
                '      iExtra2Byte = 0
                '  Avail records (type 2)
                '      agfCode  - named avail code
                '      lIntComment - avail comment
                '      Value(1) - seconds rem in avail
                '           (2) - units rem in avail
                '           (3) = inventory seconds
                '           (4) = inventory units
                '      sLineType = 2
                '      iExtra2Byte = 0
                '  Spot records (type 3)
                '      adfCode - advertiser code
                '      Product - Product description
                '      DysTms - Daypart description (with override)
                '      CurrModSpots - # spots/week
                '      SortField2 - Daily/Weekly string
                '      dnfCode - package vehicle code
                '      chfCode - contract code  (from there retrieve the product protection & program exclusion codes)
                '                               (also get contract header end date, and contract type)
                '      ContrNo - contract #  (7-26-04 replaced with Feed contr code)
                '      Status = MG(G) or outside (O)
                '      Line = Line #
                '      sSurvey - if MG or Out, Missed Date & Time string
                '      RdfDPSort - if MG or Out, orig. vehicle missed
                '      SortField1 - Price String (ADU, N/C, $ value, etc)
                '      Rate - price field (for totals, same as sortfield 1 but a number
                '      lIntComment - line comment code
                '      slineType = 3
                '      iExtra2Byte = 1 thru "N"
                '      sCBS = contains + if contract altered and requires scheduling (shown to left of contract #)
                '      sSnapshot - contains * if theres a DP override
                '

    'Dim slType As String
    Dim ilType As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilEvt As Integer
    Dim ilRet As Integer
    Dim ilSpot As Integer
    Dim llTime As Long
    Dim ilLoop As Integer
    Dim llCefCode As Long               'event (program, avail or line)comment
    Dim ilEnfCode As Integer            'event (program name) code
    Dim ilPass As Integer
    Dim ilSpotOK As Integer
    Dim llLoopDate As Long
    Dim ilWeekDay As Integer
    Dim llLatestDate As Long
    Dim ilIndex As Integer
    Dim ilRemSec As Integer     'time in seconds of avail, each spot length subtracted to get remaining seconds
    Dim ilRemUnits As Integer   '# of units of avail, each spot length subtracted to get remaining units
    Dim ilCTypes As Integer       'bit map of cnt types to include starting lo order bit
                                  '0 = unused, 1= unused, 2 = network, 3 = std, 4 = Reserved, 5 = remanant, 6 = DR
                                  '7 = PI, 8 = psa, 9 = promo, 10 = trade
                                  'bit 0 and 1 previously hold & order; but it conflicts with other contract types
    Dim ilSpotTypes As Integer    'bit map of spot types to include starting lo order bit:
                                  '0 = missed (not used in this report), 1 = charge, 2 = 0.00, 3 = adu, 4 = bonus, 5 = extra
                                  '6 = fill, 7 = n/c 8 = recapturable, 9 = spinoff, 0= MG, 10-22-10 was missed but unused in this report
    Dim ilRanks As Integer        '0 = fixed time, 1 = sponsorship, 2 = DP, 3= ROS
    Dim ilFilterDay As Integer
    ReDim ilStartTime(0 To 1) As Integer
    ReDim ilEndTime(0 To 1) As Integer
    Dim slTime As String
    ReDim ilEvtType(0 To 14) As Integer
    Dim slPrice As String
    Dim ilStartPass As Integer
    Dim ilCreateAvail As Integer
    ReDim ilTempTime(0 To 1) As Integer
    Dim slProduct As String
    Dim ilVefIndex As Integer
    Dim slXMid As String
    Dim ilFoundLLC As Integer

    'slType = "O"
    ilType = 0
    llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
    'set the type of events to get fro the day
    For ilLoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(ilLoop) = False
    Next ilLoop
    If tlCntTypes.iRated Then           'include program events?
        ilEvtType(1) = True     'retrieve programs
    End If
    ilEvtType(2) = True     'retrieve contract avails
    imSdfRecLen = Len(tmSdf)
    imCHFRecLen = Len(tmChf)
    tmCbf.iGenDate(0) = igNowDate(0)
    tmCbf.iGenDate(1) = igNowDate(1)
    'tmCbf.iGenTime(0) = igNowTime(0)
    'tmCbf.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmCbf.lGenTime = lgNowTime
    tmCbf.iVefCode = ilVefCode
    ilVefIndex = gBinarySearchVef(ilVefCode)
    For llLoopDate = llSDate To llEDate Step 1
        ilFilterDay = gWeekDayLong(llLoopDate)
        If tlCntTypes.iValidDays(ilFilterDay) Then      'Has this day of the week been selected?
            slDate = Format$(llLoopDate, "m/d/yy")
            gPackDate slDate, ilDate0, ilDate1
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            'tmSsfSrchKey.sType = slType
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
            'If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.sType <> slType) Or (tmSsf.iVefCode <> ilVefCode Or (tmSsf.iDate(0) <> ilDate0) And (tmSsf.iDate(1) = ilDate1)) Then
                'If (llLoopDate > llLatestDate) Then
                    'Always build the events within the library so that the associated program
                    'name and avail information can be retrieved (they are not in SSF)
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
                    If (llLoopDate > llLatestDate) Then     'if there are not spots scheduled this far in future, fake out an SSF entry
                                                            'so that empty avails can still be shown
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
                    End If
                    ilRet = BTRV_ERR_NONE
                'End If
            'End If

            'Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slType) And (tmSsf.iVefcode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
                gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llDate
                ilEvt = 1

                Do While ilEvt <= tmSsf.iCount
                    tmCbf.iDtFrstBkt(0) = ilDate0  'spot date
                    tmCbf.iDtFrstBkt(1) = ilDate1
                    llCefCode = 0           'init event comment code
                    ilEnfCode = 0           'init event name code
                    tmCbf.iExtra2Byte = 0       'init seq # for multiple events of the same time (spots & avail)
                   LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)

                    If tmProg.iRecType = 1 Then    'Program (not working for nested prog)
                        gUnpackTimeLong tmProg.iStartTime(0), tmProg.iStartTime(1), False, llTime
                        If llTime >= lmSTime And llTime < lmETime And tlCntTypes.iRated Then      'event within entered time parameters
                            'find the associated time in the LLC array to obtain the event name and comment
                            ilFoundLLC = False
                            For ilLoop = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
                                'match time and length
                                If tlLLC(ilLoop).iEtfCode = 1 Then
                                    ilFoundLLC = True
                                    gPackTime tlLLC(ilLoop).sStartTime, ilStartTime(0), ilStartTime(1)
                                    If (ilStartTime(0) = tmProg.iStartTime(0)) And (ilStartTime(1) = tmProg.iStartTime(1)) Then
                                        gAddTimeLength tlLLC(ilLoop).sStartTime, tlLLC(ilLoop).sLength, "A", "1", slTime, slXMid

                                        gPackTime slTime, ilEndTime(0), ilEndTime(1)
                                        'Program name code
                                        ilEnfCode = tlLLC(ilLoop).iEnfCode
                                        'Program comments
                                        llCefCode = tlLLC(ilLoop).lCefCode
                                        'program end time

                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                            If Not ilFoundLLC Then
                                gUnpackTime tmProg.iEndTime(0), tmProg.iEndTime(1), "A", "2", slTime
                                gPackTime slTime, ilEndTime(0), ilEndTime(1)
                            End If
                            'create the program record
                            tmCbf.iTime(0) = ilStartTime(0)  'program start time
                            tmCbf.iTime(1) = ilStartTime(1)
                            tmCbf.iPropOrdTime(0) = ilEndTime(0) 'program end time
                            tmCbf.iPropOrdTime(1) = ilEndTime(1)
                            tmCbf.iPropVer = 0              'program event are 0, all others are 1 for sorting in Crysta
                            tmCbf.sLineType = "1"           'Distinguish a program record from avail or spot in Crystal
                            tmCbf.iAgfCode = ilEnfCode      'event name code
                            If tlCntTypes.iNonRAted Then       'include  comments
                                tmCbf.lIntComment = llCefCode   'program comment
                            Else
                                tmCbf.lIntComment = 0
                            End If
                            ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                        End If                  'within time parameters
                    ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 2) Then 'Contract Avails only
                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                        If llTime >= lmSTime And llTime < lmETime Then      'event within entered time parameters

                            ilStartTime(0) = tmAvail.iTime(0)
                            ilStartTime(1) = tmAvail.iTime(1)
                            ilEnfCode = tmAvail.ianfCode        'avail name
                            For ilLoop = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
                                'Match start time and length (contract avails only) Test start time of avail only
                                If (tlLLC(ilLoop).iEtfCode >= 2) And (tlLLC(ilLoop).iEtfCode <= 2) Then
                                    gPackTime tlLLC(ilLoop).sStartTime, ilTempTime(0), ilTempTime(1)
                                    If (ilTempTime(0) = tmAvail.iTime(0)) And (ilTempTime(1) = tmAvail.iTime(1)) Then
                                        llCefCode = tlLLC(ilLoop).lCefCode
                                        Exit For
                                    End If
                                End If
                            Next ilLoop

                            ilRemSec = tmAvail.iLen
                            ilRemUnits = tmAvail.iAvInfo And &H1F
                            ilEnfCode = tmAvail.ianfCode        'named avail Code

                            If tlCntTypes.iSuburban Then        'show only open avails
                                ilStartPass = 1
                            Else
                                ilStartPass = 2
                                ilCreateAvail = True
                            End If

                            'Make two passes thru the spots:
                            'Pass 1 determines if there are any open avails if only open avails is the user option
                            'Pass 2 spits out the data
                            For ilPass = ilStartPass To 2
                                For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                   LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                                    ilSpotOK = True                             'assume spot is OK to include
                                    ilSpotTypes = 0
                                    ilCTypes = 0
                                    ilRanks = 0

                                    gFilterSpotTypes ilCTypes, ilSpotTypes, ilSpotOK, tlCntTypes, tmSpot    'determine to include/exclude trades, psa,promo, remnants, PI, DR & extra spots
                                    If ilSpotOK Then                            'continue testing other filters
                                        mGetSdfChf ilSpotOK, tlCntTypes   'get the sdf record from ssf

                                        mFilterLine tlCntTypes, ilSpotOK, ilCTypes, ilSpotTypes, slPrice
                                        If ilSpotOK Then
                                            If tmSdf.lChfCode = 0 Then          'feed spot
                                                slProduct = ""
                                            Else
                                                'contract spot
                                                If (tmClf.lChfCode <> tmSdf.lChfCode) Then     'got the matching line reference?
                                                    ilSpotOK = False
                                                Else
                                                    mFilterDPRanks ilRanks, ilSpotOK, tlCntTypes   'Filter out spot ranks (Fixed Time, Sponsorship, Daypart and ROS)
                                                    slProduct = Trim$(tmChf.sProduct)

                                                End If
                                            End If
                                        End If
                                    End If

                                    If ilSpotOK And ilPass = 2 Then   'only create recd if pass
                                        'create the spot record

                                        tmCbf.lLineNo = tmSdf.iLineNo
                                        tmCbf.iLen = tmSdf.iLen
                                        tmCbf.lRate = tlCntTypes.lRate
                                        tmCbf.iTime(0) = ilStartTime(0)     'avail time
                                        tmCbf.iTime(1) = ilStartTime(1)
                                        tmCbf.iPropVer = 1              'program event are 0, all others are 1 for sorting in Crysta;
                                        tmCbf.sLineType = "3"           'Distinguish a program record from avail or spot in Crystal
                                        tmCbf.iExtra2Byte = tmCbf.iExtra2Byte + 1  'seq # for multiple spots within avail
                                        tmCbf.iAdfCode = tmSdf.iAdfCode 'advertiser code
                                        tmCbf.sProduct = slProduct  'product description
                                        tmCbf.lChfCode = tmSdf.lChfCode        'contract code
                                        tmCbf.lContrNo = tmSdf.lFsfCode          'feed code (if applicable)
                                        tmCbf.sSortField1 = slPrice         'price as a string (recapturable, adu, n/c, .00, etc)
                                        mDPDesc                         'format daypart description with overrides


                                        'check if MG
                                        mMGOut          'setup missed date & time & orig vehicle if mg or outside

                                        mGetPkgLine                 'get the package vehicle name
                                        mGetLineComments tlCntTypes
                                        ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                                    End If

                                    If ilSpotOK Then
                                        ilRemSec = ilRemSec - (tmSpot.iPosLen And &HFFF)  'keep running total of whats remaining in avail based on spots used in stats
                                        ilRemUnits = ilRemUnits - 1
                                    End If
                                Next ilSpot                                 'loop from ssf file for # spots in avail

                                If ilPass = 1 Then              'if pass was to see if theres any open avail, and there is--reestablish the orig. inventory to create the records on disk
                                    If ilRemSec > 0 And ilRemUnits > 0 Then  'both units & sec must be available to show if only opens are to be shown
                                        ilRemSec = tmAvail.iLen
                                        ilRemUnits = tmAvail.iAvInfo And &H1F
                                        ilCreateAvail = True
                                    Else
                                        ilPass = 2
                                        ilCreateAvail = False
                                        Exit For
                                    End If
                                End If
                            Next ilPass
                            'create the avail record
                            If ilCreateAvail Then                       'create avail because either theres open avails or user is requesting everything
                                tmCbf.iTime(0) = ilStartTime(0)         'avail time
                                tmCbf.iTime(1) = ilStartTime(1)
                                tmCbf.iPropVer = 1              'program event are 0, all others are 1 for sorting in Crysta;
                                tmCbf.sLineType = "2"           'Distinguish a program record from avail or spot in Crystal
                                tmCbf.iExtra2Byte = 0           'avails go before spots
                                tmCbf.iAgfCode = ilEnfCode      'named avail code
                                If tlCntTypes.iNonRAted Then    'include comments
                                    tmCbf.lIntComment = llCefCode   'avail comment
                                Else
                                    tmCbf.lIntComment = 0
                                End If
                                'tmCbf.lValue(1) = ilRemSec      'seconds remaining in avail
                                'tmCbf.lValue(2) = ilRemUnits    'units remaining in avail

                                'tmCbf.lValue(3) = tmAvail.iLen  'inventory seconds of avail
                                'tmCbf.lValue(4) = tmAvail.iAvInfo And &H1F  'inventory units of avail
                                tmCbf.lValue(0) = ilRemSec      'seconds remaining in avail
                                tmCbf.lValue(1) = ilRemUnits    'units remaining in avail

                                tmCbf.lValue(2) = tmAvail.iLen  'inventory seconds of avail
                                tmCbf.lValue(3) = tmAvail.iAvInfo And &H1F  'inventory units of avail
                                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                            End If
                        End If              'endif within user entered time parameters
                        ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                    End If
                    ilEvt = ilEvt + 1   'Increment to next event
                Loop                                                        'do while ilEvt <= tmSsf.iCount
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                If tgMVef(ilVefIndex).sType = "G" Then
                    ilType = tmSsf.iType
                End If
            Loop
        End If
    Next llLoopDate
    Erase ilEvtType
    Erase tlLLC
End Sub
'
'
'       Create Daypart Description
'
'       tmRdf defined in module declaration
'
Sub mDPDesc()
Dim ilLoop As Integer
Dim ilShowOVDays As Integer
Dim ilShowOVTimes As Integer
Dim slStartTime As String
Dim slEndTime As String
Dim ilTempStartTime(0 To 1) As Integer
Dim ilTempEndTime(0 To 1) As Integer
Dim slTempDyWk As String * 1
Dim ilTempDays(0 To 6) As Integer

    ilShowOVDays = False
    ilShowOVTimes = False
    tmCbf.sDysTms = " "
    If tmSdf.lChfCode = 0 Then
        slTempDyWk = tmFsf.sDyWk
        For ilLoop = 0 To 6
            ilTempDays(ilLoop) = tmFsf.iDays(ilLoop)
        Next ilLoop
        ilTempStartTime(0) = tmFsf.iStartTime(0)
        ilTempStartTime(1) = tmFsf.iStartTime(1)
        ilTempEndTime(0) = tmFsf.iEndTime(0)
        ilTempEndTime(1) = tmFsf.iEndTime(1)
    Else
        slTempDyWk = tgPriceCff.sDyWk
        For ilLoop = 0 To 6
            ilTempDays(ilLoop) = tgPriceCff.iDay(ilLoop)
        Next ilLoop
        ilTempStartTime(0) = tmClf.iStartTime(0)
        ilTempStartTime(1) = tmClf.iStartTime(1)
        ilTempEndTime(0) = tmClf.iEndTime(0)
        ilTempEndTime(1) = tmClf.iEndTime(1)
    End If


    If slTempDyWk = "D" Then
        For ilLoop = 0 To 6 Step 1
            'If tmRdf.sWkDays(7, ilLoop + 1) = "Y" Then             'is DP is a valid day
            If tmRdf.sWkDays(6, ilLoop) = "Y" Then             'is DP is a valid day
                If ilTempDays(ilLoop) >= 0 Then         'for daily, its # spots/day
                    ilShowOVDays = True
                    Exit For
                Else
                    ilShowOVDays = False
                End If
            End If
        Next ilLoop
    Else
        For ilLoop = 0 To 6 Step 1
            'If tmRdf.sWkDays(7, ilLoop + 1) = "Y" Then             'is DP is a valid day
            If tmRdf.sWkDays(6, ilLoop) = "Y" Then             'is DP is a valid day
                If ilTempDays(ilLoop) = 0 Then         'is flight a valid day? 0=invalid day
                    ilShowOVDays = True
                    Exit For
                Else
                    ilShowOVDays = False
                End If
            End If
        Next ilLoop
    End If
    'Times
    ilShowOVTimes = False
    If ((ilTempStartTime(0) <> 1) Or (ilTempStartTime(1) <> 0)) And ((ilTempEndTime(0) <> 1) Or (ilTempEndTime(1) <> 0)) Then
        gUnpackTime ilTempStartTime(0), ilTempStartTime(1), "A", "1", slStartTime
        gUnpackTime ilTempEndTime(0), ilTempEndTime(1), "A", "1", slEndTime
        ilShowOVTimes = True
    Else
        'Add times
        For ilLoop = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1 'Row
            If (tmRdf.iStartTime(0, ilLoop) <> 1) Or (tmRdf.iStartTime(1, ilLoop) <> 0) Then
                gUnpackTime tmRdf.iStartTime(0, ilLoop), tmRdf.iStartTime(1, ilLoop), "A", "1", slStartTime
                gUnpackTime tmRdf.iEndTime(0, ilLoop), tmRdf.iEndTime(1, ilLoop), "A", "1", slEndTime
                Exit For
            End If
        Next ilLoop
    End If


    If ilShowOVDays Or ilShowOVTimes Then
        'tmCbf.sDysTms = Trim$(tmCbf.sSortField2) & " " & Trim$(slStartTime) & "-" & Trim$(slEndTime)
         tmCbf.sDysTms = Trim$(slStartTime) & "-" & Trim$(slEndTime)

        tmCbf.sSnapshot = "*"   'flag on report as override

    Else
        tmCbf.sDysTms = tmRdf.sName
        tmCbf.sSnapshot = ""    'no override flag for DP
    End If
End Sub
'
'
'           Filter out spot ranks (Fixed Time, Sponsorship, Daypart and ROS)
'
'
'           <input>
'                  ilRanks     '0 = fixed time, 1 = sponsorship, 2 = DP, 3= ROS
'                  tlCntTypes as CNTTYPES - Ranks to include/exclude
'           <output> ilSpotOK - false if exclude spot
'
Sub mFilterDPRanks(ilRanks As Integer, ilSpotOK As Integer, tlCntTypes As CNTTYPES)
Dim ilFound As Integer
Dim ilRet As Integer
Dim llStartTime As Long
Dim llEndTime As Long
Dim llTime As Long
    'Determine the Rank from the Daypart
    tmRdfSrchKey.iCode = tmClf.iRdfCode
    ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching r/c recd
    If ilRet <> BTRV_ERR_NONE Then
        ilSpotOK = False
    End If
    'Fixed time is <= 15M on Daypart or line Override times .  If Fixed time & sponsorship, add counts to fixed time
    ilFound = False
    If ((tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0)) And ((tmClf.iEndTime(0) <> 1) Or (tmClf.iEndTime(1) <> 0)) Then
        gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llStartTime
        gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llEndTime
    Else
        'gUnpackTimeLong tmRdf.iStartTime(0, 7), tmRdf.iStartTime(1, 7), False, llStartTime
        'gUnpackTimeLong tmRdf.iEndTime(0, 7), tmRdf.iEndTime(1, 7), True, llEndTime
        gUnpackTimeLong tmRdf.iStartTime(0, 6), tmRdf.iStartTime(1, 6), False, llStartTime
        gUnpackTimeLong tmRdf.iEndTime(0, 6), tmRdf.iEndTime(1, 6), True, llEndTime
    End If

    gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
    If llTime >= llStartTime And llTime <= llEndTime Then   'spot is scheduled within the override and/or DP times; if its not, consider it a DP count
        If (llEndTime - llStartTime) <= 900 Then          '15M
            ilFound = True
            ilRanks = &H1
            If Not tlCntTypes.iFixedTime Then
                ilSpotOK = False
            End If
        End If
    Else        'spot outside its DP or override times
        ilFound = True
        ilRanks = &H4           'DP count
        If Not tlCntTypes.iDP Then
            ilSpotOK = False
        End If
    End If

    'test for Sponsor
    If Not ilFound Then
        If tmRdf.sInOut = "I" Then   'Book into
            If tmAvail.ianfCode = tmRdf.ianfCode Then
                ilFound = True
                ilRanks = &H2
                If Not tlCntTypes.iSponsor Then
                    ilSpotOK = False
                End If
            Else            'sponsship spot not in matching avail, force to DP count
                ilFound = True
                ilRanks = &H4
                If Not tlCntTypes.iDP Then
                    ilSpotOK = False
                End If
            End If
        End If
        If Not ilFound Then      'not sponsorship or fixed time
            'Test for ROS
            If tmRdf.sBase <> "Y" Then   'ROS
                ilRanks = &H8
                If Not tlCntTypes.iROS Then
                    ilSpotOK = False
                End If
            Else
                ilRanks = &H4
                If Not tlCntTypes.iDP Then
                    ilSpotOK = False
                End If
            End If
        End If
    End If
End Sub
'
'
'
'               mFilterSpot - Test header and line exclusions for user request.
'
'               <input>
'                       tlCntTypes - structure of inclusions/exclusions of contract types & status
'              <output> ilSpotOk - true if spot is OK, else false to ignore spot
'                       ilCTypes - set bit to 1 for the matching type
'                       ilSpotTypes - set bit to 1 for matching type
'
'   ilCTypes        'bit map of cnt types to include starting lo order bit
'                    '0 = unused, 1= unused, 2 = network, 3 = std, 4 = Reserved, 5 = remanant, 6 = DR
'                    '7 = PI, 8 = psa, 9 = promo, 10 = trade
'                    'bit 0 and 1 previously hold & order; but it conflicts with other contract types
'
'   ilSpotTypes     'bit map of spot types to include starting lo order bit:
                    '0 = missed (not used in this report), 1 = charge, 2 = 0.00, 3 = adu, 4 = bonus, 5 = extra
                    '6 = fill, 7 = n/c 8 = recapturable, 9 = spinoff,  0= MG, 10-22-10 was missed but unused in this report
'
'   3-24-03 Change testing of Fill vs Extra

Sub mFilterLine(tlCntTypes As CNTTYPES, ilSpotOK As Integer, ilCTypes As Integer, ilSpotTypes As Integer, slPrice As String)
Dim ilRet As Integer
'Dim slDaysOfWk As String * 14
Dim ilLoop As Integer
Dim slStr As String
Dim slTempDays As String
Dim slStrip As String
Dim slXDay(0 To 6) As String * 1

    If ilSpotOK Then
        If tmSdf.lChfCode = 0 Then
            slPrice = "Feed"
            tmCbf.sSortField2 = ""              'days of the week string
            'create the daily or weekly days string
            If tmFsf.sDyWk = "D" Then     'daily
                tmCbf.lCurrModSpots = 0
                'iday represents valid # spots / day
                tmCbf.sSortField2 = ""
                For ilLoop = 0 To 6
                    tmCbf.lCurrModSpots = tmCbf.lCurrModSpots + tmFsf.iDays(ilLoop)
                    If tmFsf.iDays(ilLoop) > 9 Then
                        tmCbf.sSortField2 = RTrim(tmCbf.sSortField2) & "+ "
                     Else
                    'always make the spot count value followed by space
                        tmCbf.sSortField2 = RTrim(tmCbf.sSortField2) & RTrim(str$(tmFsf.iDays(ilLoop))) & " "
                    End If
                Next ilLoop
            Else                                'weekly
                tmCbf.lCurrModSpots = tmFsf.iNoSpots  'spots per week

                slTempDays = gDayNames(tmFsf.iDays(), slXDay(), 2, slStr)            'slstr not needed when returned
                slStr = ""
                For ilLoop = 1 To Len(slTempDays) Step 1    'strip out blanks and commas
                    slStrip = Mid$(slTempDays, ilLoop, 1)
                    If slStrip <> " " And slStrip <> "," Then
                        slStr = Trim$(slStr) & Trim$(slStrip)
                    End If
                Next ilLoop
                'slStr contains the stripped out version of the days of the wek
                tmCbf.sSortField2 = Trim$(slStr)
            End If
            If (tlCntTypes.iLenHL(0) = 0 And tlCntTypes.iLenHL(1) = 0) Or ((tlCntTypes.iLenHL(0) > 0 And tlCntTypes.iLenHL(0) = tmFsf.iLen) Or (tlCntTypes.iLenHL(1) > 0 And tlCntTypes.iLenHL(1) = tmFsf.iLen)) Then
                'if both spot length highlights are not selected, show all spot lengths; otherwise
                'the spot lengths must match the input
            Else
             ilSpotOK = False
            End If
        Else
             'slDaysOfWk = "MoTuWeThFrSaSu"
             'Test header exclusions (types of contrcts and statuses)
    
             If tmChf.sStatus = "H" Then
                 If Not tlCntTypes.iHold Then
                     ilSpotOK = False
                 End If
             ElseIf tmChf.sStatus = "O" Then
    
                 If Not tlCntTypes.iOrder Then
                     ilSpotOK = False
                 End If
             End If
    
             If tmChf.sType = "C" Then          '3-16-10 wrong flag previously tested (S--->C for standard)
                 ilCTypes = &H8
                 If Not tlCntTypes.iStandard Then       'include Standard types?
                     ilSpotOK = False
                 End If
    
             ElseIf tmChf.sType = "V" Then
                 ilCTypes = &H10
                 If Not tlCntTypes.iReserv Then      'include reservations ?
                     ilSpotOK = False
                 End If
    
             ElseIf tmChf.sType = "R" Then
                 ilCTypes = &H20
                 If Not tlCntTypes.iDR Then       'include DR?
                     ilSpotOK = False
                 End If
             End If
             '3-28-02 Reread line if either contract or line # is different.
             If (tmClf.lChfCode <> tmSdf.lChfCode) Or (tmClf.iLine <> tmSdf.iLineNo) Then     'got the matching line reference?
                 tmClfSrchKey.lChfCode = tmSdf.lChfCode
                 tmClfSrchKey.iLine = tmSdf.iLineNo
                 tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
                 tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
                 ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
             End If
             If (tmClf.lChfCode <> tmSdf.lChfCode) Then     'got the matching line reference?
                 ilSpotOK = False
             End If
             If (tlCntTypes.iLenHL(0) = 0 And tlCntTypes.iLenHL(1) = 0) Or (tlCntTypes.iLenHL(0) > 0 And tlCntTypes.iLenHL(0) = tmClf.iLen) Or (tlCntTypes.iLenHL(1) > 0 And tlCntTypes.iLenHL(1) = tmClf.iLen) Then
                'if both spot length highlights are not selected, show all spot lengths; otherwise
                'the spot lengths must match the input
             Else
                 ilSpotOK = False
             End If
             'If tlCntTypes.iLenHL(1) <> 0 And tlCntTypes.iLenHL(1) <> tmClf.iLen Then
             '    ilSpotOK = False
             'End If
    
             'Retrieve spot cost from flight ; flight not returned if spot type is Extra/Fill
             'otherwise flight returned in tgPriceCff
             ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
             tlCntTypes.lRate = 0                'init the spot rate until something found
             tmCbf.sSortField2 = ""              'days of the week string
             'look for inclusion of spot types
             'If Trim$(slPrice) <> "Extra" And Trim$(slPrice) <> "Fill" Then
             If Trim$(slPrice) <> "+ Fill " And Trim$(slPrice) <> "- Fill" Then
                 'create the daily or weekly days string
                 If tgPriceCff.iSpotsWk = 0 Then     'daily
                     tmCbf.lCurrModSpots = 0
                     'iday represents valid # spots / day
                     tmCbf.sSortField2 = ""
                     For ilLoop = 0 To 6
                        tmCbf.lCurrModSpots = tmCbf.lCurrModSpots + tgPriceCff.iDay(ilLoop)
                         If tgPriceCff.iDay(ilLoop) > 9 Then
                             tmCbf.sSortField2 = RTrim(tmCbf.sSortField2) & "+ "
                         Else
                         'always make the spot count value followed by space
                             tmCbf.sSortField2 = RTrim(tmCbf.sSortField2) & RTrim(str$(tgPriceCff.iDay(ilLoop))) & " "
                         End If
                     Next ilLoop
                 Else                                'weekly
                     tmCbf.lCurrModSpots = tgPriceCff.iSpotsWk   'spots per week
    
                     slTempDays = gDayNames(tgPriceCff.iDay(), tgPriceCff.sXDay(), 2, slStr)            'slstr not needed when returned
                     slStr = ""
                     For ilLoop = 1 To Len(slTempDays) Step 1    'strip out blanks and commas
                         slStrip = Mid$(slTempDays, ilLoop, 1)
                         If slStrip <> " " And slStrip <> "," Then
                             slStr = Trim$(slStr) & Trim$(slStrip)
                         End If
                     Next ilLoop
                     'slStr contains the stripped out version of the days of the wek
                     tmCbf.sSortField2 = Trim$(slStr)
                     'For ilLoop = 0 To 6
                     '    If tgPriceCff.iDay(ilLoop) = 1 Then     'valid day of week
                     ''        tmCbf.sSortField2 = Trim$(tmCbf.sSortField2) & Mid(slDaysOfWk, ((ilLoop + 1) * 2) - 1, 2)
                     '    End If
                     'Next ilLoop
                 End If
             End If
             If (InStr(slPrice, ".") <> 0) Then    'found spot cost
                 tlCntTypes.lRate = gStrDecToLong(slPrice, 2)    'get the actual spot value
                 'is it a .00?
                 If gCompNumberStr(slPrice, "0.00") = 0 Then     'its a .00 spot
                     ilSpotTypes = &H4
                     If Not tlCntTypes.iZero Then
                         ilSpotOK = False
                     End If
                 Else
                     ilSpotTypes = &H2
                     If Not tlCntTypes.iCharge Then           'exclude charged spots
                         ilSpotOK = False
                     End If
                 End If
             ElseIf Trim$(slPrice) = "ADU" Then
                 ilSpotTypes = &H8
                 If Not tlCntTypes.iADU Then
                     ilSpotOK = False
                 End If
             ElseIf Trim$(slPrice) = "Bonus" Then
                 ilSpotTypes = &H10
                 If Not tlCntTypes.iBonus Then
                     ilSpotOK = False
                 End If
             'ElseIf Trim$(slPrice) = "Extra" Then
             ElseIf Trim$(slPrice) = "+ Fill" Then
                 ilSpotTypes = &H20
                 If Not tlCntTypes.iXtra Then
                     ilSpotOK = False
                 End If
             'ElseIf Trim$(slPrice) = "Fill" Then
             ElseIf Trim$(slPrice) = "- Fill" Then
                 ilSpotTypes = &H40
                 If Not tlCntTypes.iFill Then
                     ilSpotOK = False
                 End If
             ElseIf Trim$(slPrice) = "N/C" Then
                 ilSpotTypes = &H80
                 If Not tlCntTypes.iNC Then
                     ilSpotOK = False
                 End If
             ElseIf Trim$(slPrice) = "Recapturable" Then
                 ilSpotTypes = &H100
                 If Not tlCntTypes.iRecapturable Then
                     ilSpotOK = False
                 End If
             ElseIf Trim$(slPrice) = "Spinoff" Then
                 ilSpotTypes = &H200
                 If Not tlCntTypes.iSpinoff Then
                     ilSpotOK = False
                 End If
            ElseIf Trim$(slPrice) = "MG" Then       '10-22-10 option to incl/excl spot rate of MG
                 ilSpotTypes = &H1
                 If Not tlCntTypes.iMG Then
                     ilSpotOK = False
                 End If
             End If
             
             '10-22-10 if MG/out spot, check option to include/excl mg. The outside cant be a fill spot
             If ((tmSdf.sSchStatus = "G" Or tmSdf.sSchStatus = "O") And (tmSdf.sSpotType <> "X")) Then
                If Not tlCntTypes.iMG Then
                    ilSpotOK = False
                End If
            End If
            
        End If                              'tmsdf.lchfcode = 0
    End If                                  'ilspotOK
End Sub
'       4-24-08 replaced with global subroutine in rptsubs.bas
'
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
'Sub mFilterSpotTypes(ilCTypes As Integer, ilSpotTypes As Integer, ilSpotOK As Integer, tlCntTypes As CNTTYPES)
'    If (tmSpot.iRank And RANKMASK) = 1020 Then
'        ilCTypes = &H10
'        If Not tlCntTypes.iRemnant Then
'            ilSpotOK = False
'        End If
'
'    ElseIf (tmSpot.iRank And RANKMASK) = 1030 Then
'        ilCTypes = &H800
'        If Not tlCntTypes.iPI Then
'            ilSpotOK = False
'        End If
'
'    ElseIf (tmSpot.iRank And RANKMASK) = 1040 Then
'        ilCTypes = &H400
'        If Not tlCntTypes.iTrade Then
'            ilSpotOK = False
'        End If
'
'    'ElseIf tmSpot.iRank = 1045 Then
'    '    ilSpotTypes = &H20
'    '    If Not tlCntTypes.iXtra Then
'    '        ilSpotOK = False
'    '    End If
'
'    ElseIf (tmSpot.iRank And RANKMASK) = 1050 Then
'        ilCTypes = &H200
'        If Not tlCntTypes.iPromo Then
'            ilSpotOK = False
'        End If
'
'    ElseIf (tmSpot.iRank And RANKMASK) = 1060 Then
'        ilCTypes = &H100
'        If Not tlCntTypes.iPSA Then
'            ilSpotOK = False
'        End If
'    End If
'    If (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
'        ilSpotOK = False
'    End If
'End Sub
'
'
'           mGetLineComments - obtain the Line comments from CXF and see if it should be shown
'           on the Log.  If so, show it on the Spot Report
'
'           <input> tlCntTypes - contract types / ranks, etc to be included/excluded
'
Sub mGetLineComments(tlCntTypes As CNTTYPES)
Dim ilRet As Integer
    tmCbf.lIntComment = 0
    tmCxfSrchKey.lCode = tmClf.lCxfCode      'comment  code
    imCxfRecLen = Len(tmCxf)
    ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching comment recd
    If ilRet = BTRV_ERR_NONE Then
        If tlCntTypes.iNonRAted Then         'include comments
            If tmCxf.sShSpot = "Y" Then         'show comment on spot screen
                tmCbf.lIntComment = tmClf.lCxfCode
            Else
                tmCbf.lIntComment = 0
            End If
        End If
    End If
End Sub
'
'
'               mGetPkgLine - if spot is hidden, retrieve package line to
'               show its package vehicle name
'
'
Sub mGetPkgLine()
Dim ilRet As Integer
Dim tlClf As CLF
    If tmClf.sType = "H" Then
        tmClfSrchKey.lChfCode = tmSdf.lChfCode
        tmClfSrchKey.iLine = tmClf.iPkLineNo
        tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
        tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
        ilRet = btrGetGreaterOrEqual(hmClf, tlClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet = BTRV_ERR_NONE And tlClf.iLine = tmClf.iPkLineNo And tmChf.lCode = tlClf.lChfCode Then
            tmCbf.iDnfCode = tlClf.iVefCode
        End If
    Else
        tmCbf.iDnfCode = 0          'show the package vehicle forthis hidden
    End If
End Sub
'
'
'           Obtain Sdf record from the SSF
'
'       <Output> - ilSpotOK : False if error
'
'   3-24-03 change way to test for fill vs extra (now test from advt vs spot)
'   1-19-04 change way to test for fill again.  If overridden in fill screen not to use the advt default,
'           then test the spot to determine to show or not show
Sub mGetSdfChf(ilSpotOK As Integer, tlCntTypes As CNTTYPES)
Dim ilRet As Integer
Dim ilLoop As Integer
Dim ilCode1 As Integer
Dim ilCode2 As Integer
Dim ilFound As Integer
Dim tlChf As CHF
Dim slShowOnInv As String * 1   '3-24-03
Dim ilSelectedField(0 To 1) As Integer

    ilSelectedField(0) = 0
    ilSelectedField(1) = 0
    tmSdfSrchKey3.lCode = tmSpot.lSdfCode
    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
    If tmSpot.lSdfCode = tmSdf.lCode And ilRet = BTRV_ERR_NONE Then
        If tmSdf.lChfCode = 0 Then                           'feed spot vs contr
            tmChfSrchKey.lCode = tmSdf.lFsfCode
            ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If imSelList = 1 Then                   'adv
                ilCode1 = tmFsf.iAdfCode
            ElseIf imSelList = 2 Then               'agy
                ilCode1 = 0
            ElseIf imSelList = 3 Then               'prod protection
                ilCode1 = tmFsf.iMnfComp1
                ilCode2 = tmFsf.iMnfComp2
            Else                                    'slsp
                ilCode1 = 0
            End If
            If Not tlCntTypes.iNetwork Then
                ilSpotOK = False
            End If
        Else
            If tmSdf.lChfCode <> tmChf.lCode Then               'if already in mem, don't reread
                tmChfSrchKey.lCode = tmSdf.lChfCode

                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                'test for selective contr entered

                tmChfSrchKey1.lCntrNo = tmChf.lCntrNo
                tmChfSrchKey1.iCntRevNo = 32000      'look for latest revision
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tlChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)

                tmCbf.sCBS = ""
                If tlChf.lCntrNo = tmChf.lCntrNo And ilRet = BTRV_ERR_NONE Then
                    If tlChf.sSchStatus = "A" Then  'contract altered, flag as needs scheduling on report
                        tmCbf.sCBS = "+"
                    End If
                End If
                If lmSingleCntr > 0 And lmSingleCntr <> tmChf.lCntrNo Then
                    ilSpotOK = False
                End If
                If imSelList = 1 Then                   'adv
                    ilCode1 = tmChf.iAdfCode
                ElseIf imSelList = 2 Then               'agy
                    ilCode1 = tmChf.iAgfCode
                ElseIf imSelList = 3 Then               'prod protection
                    ilCode1 = tmChf.iMnfComp(0)
                    ilCode2 = tmChf.iMnfComp(1)
                Else                                    'slsp
                    ilCode1 = tmChf.iSlfCode(0)
                End If
                '2-4-05 if same contract, still need to test the spot
            Else
                ilRet = BTRV_ERR_NONE
            End If
        End If

        If tmSdf.sSpotType = "X" Then
            slShowOnInv = gTestShowFill(tmSdf.sPriceType, tmSdf.iAdfCode)       '1-19-04 see if spot should be shown/not shown on inv
            If slShowOnInv = "N" And Not tlCntTypes.iFill Then
                ilSpotOK = False
            End If
            If slShowOnInv = "Y" And Not tlCntTypes.iXtra Then
                ilSpotOK = False
            End If

        End If


        ilFound = True          'assume everything to be included
        'test for selective agy, adv, product protection and slsp codes (only use primary)
        If UBound(imSelLists) > 0 And imSelList > 0 Then          'some selectivity
            ilCode2 = 0
            ilFound = False
            For ilLoop = LBound(imSelLists) To UBound(imSelLists) - 1
                'If imSelList = 1 Then                   'adv
                '    ilCode1 = tmChf.iAdfCode
                'ElseIf imSelList = 2 Then               'agy
                '    ilCode1 = tmChf.iAgfCode
                'ElseIf imSelList = 3 Then               'prod protection
                '    ilCode1 = tmChf.iMnfComp(0)
                '    ilCode2 = tmChf.iMnfComp(1)
                'Else                                    'slsp
                '    ilCode1 = tmChf.iSlfCode(0)
                'End If
                If imSelLists(ilLoop) = ilCode1 Or imSelLists(ilLoop) = ilCode2 Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
        End If
        If Not ilFound Then
            ilSpotOK = False
        End If
    End If
    'Else       2-4-05 move this code to above
    '    ilRet = BTRV_ERR_NONE
    'End If

    If ilRet <> BTRV_ERR_NONE Then
        ilSpotOK = False
    End If
End Sub
'
'
'               mMGOut - if Makegood or Outside spot, setup string of
'               missed date and time and original vehicle
'               12-18-00 Show "M for N" in status column for multi-makegoods
'               11-30-04 access smf by key2 instead of key0 for speed
Sub mMGOut()
Dim llDate As Long
Dim slTime As String
Dim ilRet As Integer
    tmCbf.sSurvey = "Missed"      'orig date & time missed string
    tmCbf.iRdfDPSort = tmClf.iVefCode
    tmCbf.sStatus = ""
    If (tmSdf.sSchStatus = "G" Or tmSdf.sSchStatus = "O") And tmSdf.sSpotType <> "X" Then  'mg or outside (not extra or fill), need to show orig date & time & vehicle missed from
    '12-18-00 Show "M for N" if its multiple makegoods
        'tmSmfSrchKey.lChfCode = tmSdf.lChfCode
        'tmSmfSrchKey.lFsfCode = tmSdf.lFsfCode
        'tmSmfSrchKey.iLineNo = tmSdf.iLineNo
        'tmSmfSrchKey.iMissedDate(0) = 0 'ilDate0
        'tmSmfSrchKey.iMissedDate(1) = 0 'ilDate1
        'ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        '11-30-04 access smf by key2 instead of key0 for speed
        tmSmfSrchKey2.lCode = tmSdf.lCode
        ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation

        Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tmSdf.lChfCode) And (tmSmf.iLineNo = tmSdf.iLineNo) And (tmSmf.lFsfCode = tmSdf.lFsfCode)
            If (tmSmf.lSdfCode = tmSdf.lCode) Then
                gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llDate
                tmCbf.sSurvey = Trim$(tmCbf.sSurvey) & " " & Trim$(Format$(llDate, "m/d/yy"))
                gUnpackTime tmSmf.iMissedTime(0), tmSmf.iMissedTime(1), "A", "2", slTime
                tmCbf.sSurvey = Trim$(tmCbf.sSurvey) & " @" & Trim$(slTime)
                tmCbf.iRdfDPSort = tmSmf.iOrigSchVef    'orig vehicle code
                If tmSmf.lMtfCode > 0 Then      'its an "M for N" makegood
                    tmCbf.sStatus = "N"
                    tmCbf.sSurvey = ""
                    tmCbf.iRdfDPSort = tmSdf.iVefCode
                Else                    'not an "M for N", so there shouldnt be orig missed date & time info
                    tmCbf.sStatus = tmSdf.sSchStatus
                End If
                Exit Do
            Else
                tmCbf.sSurvey = tmCbf.sSurvey & " Unknown"
            End If
            ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        Loop
    End If
End Sub
