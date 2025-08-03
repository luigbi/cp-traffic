Attribute VB_Name = "RptCrSR"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptcrsr.bas on Fri 3/12/10 @ 11:00 AM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmStations                                                                            *
'******************************************************************************************

Option Explicit
Option Compare Text

Dim hmGrf As Integer
Dim tmGrf As GRF
Dim imGrfRecLen As Integer

Dim hmRaf As Integer
Dim tmRaf() As RAF

Dim imRafRecLen As Integer
Dim tmRafSrchKey0 As LONGKEY0

Dim hmSef As Integer
Dim tmSef As SEF
Dim imSefRecLen As Integer
Dim tmSefSrchKey1 As SEFKEY1  'SEF key record image

Dim hmSHTT As Integer
Dim tmSHTT As SHTT
Dim imSHTTRecLen As Integer

Dim imInclCatMkt As Integer             'include DMA markets category
Dim imInclCatMSAMkt As Integer          '12-23-09 msa replaced owner,include MSA markets category
'10/29: Darlene- I removed Owner
'Dim ilInclCatOwner As Integer           'include owners category
Dim imInclCatStation As Integer         'includ station category
'10/29: Darlene- I removed Owner
'Dim ilInclCatZip As Integer             'include zip code category
Dim imInclCatState As Integer           'include state category
Dim imInclCatFormat As Integer          'include format category
'10/29:Darlene- I added CatTime
Dim imInclCatTime As Integer            'include time category

Dim tmRegionDefinition() As REGIONDEFINITION
Dim tmSplitCategoryInfo() As SPLITCATEGORYINFO


'
'
'           Create Split Network/split Copy region report
'           Dump the contents of the regions and categories files
'           show the stations that belong to each region and category
'           9-19-06
'
'           Sort (major to minor) by ADvt:
'               Type (network or copy split), adv, region, category name, call letterss
'           Sort (major to minor) by Region:
'               Type (network or copy split), region, advt, category name, call letterss
'
'           6-25-07 Remove association of advertiser with split network
'                   if split network disallow advt sort and selectivity as they are no longer associated
'                   Add Format selectivity
'
'           TWO forms of this report exists:  if using Regional Copy, RAF.rpt is displayed (CR_COPYREGION)
'                           If using Split Copy and/or Split Networks, then splitregion.rpt is displayed
'                           (CR_SPLITREGION)
'
'
'           7-8-10 Change the intent of the CopySplit report to show only the stations that meet the criteria
'           that make up region.  This works for both the Network Splits and Copy Splits report.
'
Sub gCreateSplitRegions()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Removed)                                                                 *
'*  gCreateSplitRegionErr                                                                 *
'******************************************************************************************

Dim ilRet As Integer
Dim ilInclRegCopy As Integer            'include Regional copy
Dim ilInclSplitNet As Integer           'include split networks
Dim ilInclSplitCopy As Integer          'include split copy
Dim ilLoopOnRAF As Integer              'index to loop thru RAF records
Dim ilLoopOnCat As Integer              'index to loop thru the RAF categories
'Dim ilInclCatMkt As Integer             'include markets category
''10/29: Darlene- I removed Owner
''Dim ilInclCatOwner As Integer           'include owners category
'Dim ilInclCatStation As Integer         'includ station category
''10/29: Darlene- I removed Owner
''Dim ilInclCatZip As Integer             'include zip code category
'Dim ilInclCatState As Integer           'include state category
'Dim ilInclCatFormat As Integer          'include format category
''10/29:Darlene- I added CatTime
'Dim ilInclCatTime As Integer            'include time category
Dim ilFound As Integer
Dim ilInclDormant As Integer            'include dormant with active
Dim ilIncludeCodes As Integer           'true to include codes, false to exclude codes
ReDim ilUseCodes(0 To 0) As Integer             'advt codes to include or exclude
Dim ilOKList As Integer
Dim ilTemp As Integer
Dim tlSef() As SEF
Dim ilLoopOnSEF As Integer
Dim ilFoundSEF As Integer
Dim slCategory As String * 1
Dim slInclExcl As String * 1
Dim ilDetail As Integer
Dim llFromDate As Long
Dim llToDate As Long
Dim llCreationDate As Long
Dim ilListIndex As Integer
Dim llFromCode As Long
Dim llToCode As Long
Dim slStr As String
Dim slNameCode As String
Dim slCode As String
Dim ilAdfCode As Integer
Dim llRafCode As Long
Dim llUpperSplit As Long
Dim ilShttInx As Integer
Dim ilShttCode As Integer
Dim ilMktCode As Integer
Dim ilMSAMktCode As Integer
Dim slState As String
Dim ilFmtCode As Integer
Dim ilTztCode As Integer
Dim llRegionIndex As Long
Dim slGroupInfo As String
Dim ilLoopOnStations As Integer
Dim slSplitCatDesc() As String
Dim ilWhichBox As Integer
Dim tlRegionDefinition() As REGIONDEFINITION
Dim tlSplitCategoryInfo() As SPLITCATEGORYINFO


    ilListIndex = RptSelSR!lbcRptType.ListIndex
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmRaf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmRaf)
        btrDestroy hmGrf
        btrDestroy hmRaf
        Exit Sub
    End If
    imRafRecLen = Len(tmRaf(0))

    hmSef = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSef, "", sgDBPath & "Sef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSef)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmRaf)
        btrDestroy hmSef
        btrDestroy hmGrf
        btrDestroy hmRaf
        Exit Sub
    End If
    imSefRecLen = Len(tmSef)
    
    hmSHTT = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()


    ilInclSplitNet = False
    ilInclSplitCopy = False
    ilInclRegCopy = False
    imInclCatMkt = False
    imInclCatMSAMkt = False            '12-23-09
    '10/29: Darlene- I removed Owner
    'ilInclCatOwner = False
    imInclCatStation = False
    imInclCatState = False
    '10/29: Darlene- I removed Owner
    'ilInclCatZip = False
    imInclCatFormat = False
    '10/29: Darlene- I added Time zone
    imInclCatTime = False
    ilInclDormant = False

    '7-2-10 The options for category selection have been removed.  Include everything
    'The controls still exist and they are defaulted to include
    If ilListIndex = CR_SPLITREGION Then
        If RptSelSR!ckcCat(0).Value = vbChecked Then        'include dma markets
            imInclCatMkt = True
        End If
        '10/29: Darlene- I removed Owner
        'If RptSelSR!ckcCat(1).Value = vbChecked Then        'include owners
        '    ilInclCatOwner = True
        'End If
        '12-23-09 Replace Owner with MSA MKT
        If RptSelSR!ckcCat(1).Value = vbChecked Then        'include msa markets
            imInclCatMSAMkt = True
        End If

        If RptSelSR!ckcCat(2).Value = vbChecked Then        'include state
            imInclCatState = True
        End If
        If RptSelSR!ckcCat(3).Value = vbChecked Then        'include stations
            imInclCatStation = True
        End If
        If RptSelSR!ckcCat(4).Value = vbChecked Then        'include zip
            'ilInclCatZip = True
            imInclCatTime = True
        End If
        If RptSelSR!ckcCat(5).Value = vbChecked Then        'include format
            imInclCatFormat = True
        End If

        'include split network or both?
        If RptSelSR!rbcWhichSplit(0).Value = True Or RptSelSR!rbcWhichSplit(2).Value = True Then      'include split network or both split net & copy
            ilInclSplitNet = True
            ilInclRegCopy = False
        End If
        'include split copy or both
        If RptSelSR!rbcWhichSplit(1).Value = True Or RptSelSR!rbcWhichSplit(2).Value = True Then      'incl split copy or both split net & split copy
            ilInclSplitCopy = True
            ilInclRegCopy = False
        End If

    ElseIf ilListIndex = CR_COPYREGION Then
        slStr = RptSelSR!edcCodeFrom.Text
        If slStr = "" Then
            slStr = "0"
        End If
        llFromCode = Val(slStr)
        slStr = RptSelSR!edcCodeTo.Text
        If slStr = "" Then
            slStr = "999999999"
        End If
        llToCode = Val(slStr)
        ilInclRegCopy = True            'include regional copy
    End If

    If RptSelSR!ckcDormant.Value = vbChecked Then        'include dormant
        ilInclDormant = True
    End If
    
    '7-2-10  Dates have been removed for CR_SPLITREGION, defaulted to include all dates
    slStr = RptSelSR!edcDateFrom.Text
    If slStr = "" Then                  'if no start date entered, use min date to get all
        slStr = "1/1/1970"
    End If
    llFromDate = gDateValue(slStr)      'convert fromdate to long
    slStr = RptSelSR!edcDateTo.Text
    If slStr = "" Then                  'if no end date, use max date to get all
        slStr = "12/31/2020"
    End If
    llToDate = gDateValue(slStr)        'convert to date to long

    ilRet = gObtainStations()


    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    tmGrf.lGenTime = lgNowTime
    
    'tmGrf.iSofCode     owner (ARTT) code reference if category is O
    'tmGrf.iSlfCode     tation (SHTT) code reference if category is S
    'tmGrf.iRdfcode     market (MKT) code reference if category is M
    'tmgrf.ivefcode     msa market (MET) code reference if category is A
    'tmgrf.iCode2       Time zone (TZT) code reference if category = T
    'tmGrf.iAdfCode     advertiser code
    'tmGrf.sGenDesc     Copy Region: Description name for category zip code or state name (if category = "N")
    '                   Copy Splits - In the subrecords (tmGrfPerGenl(1) = 1): Split category description
    'tmGrf.iPerGenl(1)  COPY REGION rpt:  0 = header info, 1 = detail station info
    '                   COPY SPLITS RPT:  0 = station records, 1 = subrecords to link to the category that make up the region for the station
    'tmGrf.sBktType     Type of category ('M = market, s = station, O = owner, Z = zip, n = state name)
    'tmgrf.lCode4       RAF code
    'tmGrf.sDateType    Include/Exclude
    'tmGrf.lLong        COYP SPLITS RPT: Sort Code for unique splits that make up the valid station for a region (used to sort in crystal)
                        'so that the same splits are together (i.e. IF12|IM3438|ES9413  :Include format code 12|Include Mkt code 9413|Exclude Station code 9413)
    
    If ilListIndex = CR_COPYREGION Then     'regional copy
        'build list of advt codes to include or exclude for faster processing
        gObtainCodesForMultipleLists 3, tgAdvertiser(), ilIncludeCodes, ilUseCodes(), RptSelSR

        'gather all the regions
        ilRet = gObtainRAFByType(RptSelSR, hmRaf, tmRaf(), ilInclRegCopy, ilInclSplitNet, ilInclSplitCopy)
        
        'Loop for each of the regions and create a header record and multiple detail records.
        'The header record will show the region name, advt, incl/excl flag, category, and show on reports flags.
        'The detail record shows all the stations included or excluded for the given category associated with the region
        For ilLoopOnRAF = LBound(tmRaf) To UBound(tmRaf) - 1        'outer  loop, loop on regions
            ilOKList = False
            If ilIncludeCodes Then
                For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                    If ilUseCodes(ilTemp) = tmRaf(ilLoopOnRAF).iAdfCode Then
                        ilOKList = True
                        Exit For
                    End If
                Next ilTemp
            Else
                ilOKList = True        ' when more than half selected, selection fixed
                For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                    If ilUseCodes(ilTemp) = tmRaf(ilLoopOnRAF).iAdfCode Then
                        ilOKList = False
                        Exit For
                    End If
                Next ilTemp
            End If
    
            'verify if within requested dates
            gUnpackDateLong tmRaf(ilLoopOnRAF).iDateEntrd(0), tmRaf(ilLoopOnRAF).iDateEntrd(1), llCreationDate
            If llCreationDate < llFromDate Or llCreationDate > llToDate Then
                ilOKList = False
            End If
    
            If ilListIndex = CR_COPYREGION Then     'regional copy
                If tmRaf(ilLoopOnRAF).lRegionCode < llFromCode Or tmRaf(ilLoopOnRAF).lRegionCode > llToCode Then
                    ilOKList = False
                End If
            End If
    
            'if Regional copy, see if codes fall within the requested paramters
            If ilOKList And ((ilInclDormant = True And tmRaf(ilLoopOnRAF).sState = "D") Or (tmRaf(ilLoopOnRAF).sState <> "D")) Then
                tmGrf.sBktType = tmRaf(ilLoopOnRAF).sCategory   'M = market, s = station, O = owner, Z = zip, n = state name
                ilFound = False                                 'flag to find market
    
    '            Select Case tmRaf(ilLoopOnRAF).sCategory        'determine if the category should be included
    '                Case "M"            'market
    '                    If ilInclCatMkt Then
    '                        ilFound = True
    '                    End If
    '                '10/29: Darlene- I removed Owner and Zip
    '                'Case "O"    'Owner
    '                '    If ilInclCatOwner Then
    '                '        ilFound = True
    '                '    End If
    '                'Case "Z"    'Zip Code
    '                '    If ilInclCatZip Then
    '                '        ilFound = True
    '                '    End If
    '                Case "N"    'State
    '                    If ilInclCatState Then
    '                        ilFound = True
    '                    End If
    '                Case "S"    'Station
    '                    If ilInclCatStation Then
    '                        ilFound = True
    '                    End If
    '                 Case "F"    'Format
    '                    If ilInclCatFormat Then
    '                        ilFound = True
    '                    End If
    '                 Case "T"    'Time Zone
    '                    If ilInclCatTime Then
    '                        ilFound = True
    '                    End If
    '            End Select
    
                'found the matching category type, and if the region is dormant test if user selected dormants
                'If ((ilInclDormant = True And tmRaf(ilLoopOnRAF).sState = "D") Or (tmRaf(ilLoopOnRAF).sState <> "D")) Then
                    ReDim tlSef(0 To 0) As SEF
                    'gather all the categories for this region
                    tmSefSrchKey1.lRafCode = tmRaf(ilLoopOnRAF).lCode
                    tmSefSrchKey1.iSeqNo = 0
                    ilRet = btrGetGreaterOrEqual(hmSef, tmSef, imSefRecLen, tmSefSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSef.lRafCode = tmRaf(ilLoopOnRAF).lCode)
                        tlSef(UBound(tlSef)) = tmSef
                        ReDim Preserve tlSef(0 To UBound(tlSef) + 1) As SEF
                        ilRet = btrGetNext(hmSef, tmSef, imSefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
    
                    'determine if any category for the region has been requested for inclusion
                    'if so, need to show all categories within region
                    ilFoundSEF = False
                    For ilLoopOnSEF = 0 To UBound(tlSef) - 1
                        slCategory = tlSef(ilLoopOnSEF).sCategory
                        If Trim$(slCategory) = "" Then             'one of the original records before new design or its a Network split
                            slCategory = tmRaf(ilLoopOnRAF).sCategory
                        End If
                        ilFoundSEF = mTestCategory(slCategory)
                        If ilFoundSEF Then      'found a valid category, show all in the region regardless of user category selectivity
                            Exit For
                        End If
                    Next ilLoopOnSEF
    
                    If ilFoundSEF Then
                        tmGrf.lCode4 = tmRaf(ilLoopOnRAF).lCode             'RAF Code
                        tmGrf.iAdfCode = tmRaf(ilLoopOnRAF).iAdfCode     'advertiser code
                        tmGrf.iSofCode = 0                  'init owner (ARTT) code reference
                        tmGrf.iSlfCode = 0                  'init station (SHTT) code reference
                        tmGrf.iRdfCode = 0                  'init market (MKT) code reference
                        tmGrf.iCode2 = 0                    'time zone (TZT) code reference
                        tmGrf.iYear = 0                     'format code (FMT) reference
                        For ilLoopOnSEF = LBound(tlSef) To UBound(tlSef) - 1
                            tmSef = tlSef(ilLoopOnSEF)
                    'Do While (ilRet = BTRV_ERR_NONE) And (tmSef.lRafCode = tmRaf(ilLoopOnRAF).lCode)
                            tmGrf.sGenDesc = tmSef.sName
                            'tmGrf.iPerGenl(1) = 1               'detail
                            tmGrf.iPerGenl(0) = 1               'detail
                            'ReDim tltempstations(1 To 1) As INTKEY0
                            ReDim tltempstations(0 To 0) As INTKEY0
                            slCategory = tlSef(ilLoopOnSEF).sCategory
                            If Trim$(slCategory) = "" Then             'one of the original records before new design or its a Network split
                                slCategory = tmRaf(ilLoopOnRAF).sCategory
                            End If
                            'include/exclude flag, new design stored in SEF unless its an old record in which case test RAF
                            slInclExcl = tlSef(ilLoopOnSEF).sInclExcl
                            If Trim$(slInclExcl) = "" Then             'one of the original records before new design or its a Network split
                                slInclExcl = tmRaf(ilLoopOnRAF).sInclExcl
                            End If
                            Select Case Trim$(slCategory)          'determine which category to gather the station list
                                Case "M" 'market
                                    tmGrf.iRdfCode = tmSef.iIntCode     'DMA market code
                                    ilRet = gBuildStationsByIntCategory(tgMktSort(), tmSef.iIntCode, tltempstations())
    
                                Case "A"    'MSA
                                    tmGrf.iRdfCode = tmSef.iIntCode     'MSA market code
                                    ilRet = gBuildStationsByIntCategory(tgMSAMktSort(), tmSef.iIntCode, tltempstations())
    
                                '10/29: Darlene- I removed Owner and Zip
                                'Case "O"    'Owner
                                '    tmGrf.iSofCode = tmSef.iIntCode     'owner code (ARTT)
                                '    ilRet = gBuildStationsByIntCategory(tgOwnerSort(), tmSef.iIntCode, tlTempStations())
                                'Case "Z"    'Zip Code
                                '    ilRet = gBuildStationsByStrCategory(tgZipSort(), tmSef.sName, tlTempStations())
                               Case "" 'default blank category to blank, possibly previous update bug in older records
                                    tmGrf.iRdfCode = tmSef.iIntCode     'market code
                                    ilRet = gBuildStationsByIntCategory(tgMktSort(), tmSef.iIntCode, tltempstations())
                                Case "N"    'State
                                    ilRet = gBuildStationsByStrCategory(tgStateSort(), tmSef.sName, tltempstations())
                                '10/29: Darlene-I added Format
                                Case "F"    'format
                                    tmGrf.iYear = tmSef.iIntCode     'Format code
                                    ilRet = gBuildStationsByIntCategory(tgFmtSort(), tmSef.iIntCode, tltempstations())
                                '10/29: Darelen-I added time zone
                                Case "T"    'time zone
                                    tmGrf.iCode2 = tmSef.iIntCode     'Time zone code
                                    ilRet = gBuildStationsByIntCategory(tgTztSort(), tmSef.iIntCode, tltempstations())
                                Case "S"    'Station
                                    tmGrf.iSlfCode = tmSef.iIntCode     'Station Code
                                    tltempstations(UBound(tltempstations)).iCode = tmSef.iIntCode
                                    ReDim Preserve tltempstations(LBound(tltempstations) To UBound(tltempstations) + 1) As INTKEY0
                            End Select
    
                            'got the array of stations for the selected category, create a detail record
                            'for each station in the regions category
                            tmGrf.sDateType = slInclExcl            'incl/excl flag
                            tmGrf.sBktType = slCategory             'category type
                            For ilLoopOnCat = LBound(tltempstations) To UBound(tltempstations) - 1
                                tmGrf.iSlfCode = tltempstations(ilLoopOnCat).iCode     'station code
                                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                            Next ilLoopOnCat
                            'ilRet = btrGetNext(hmSef, tmSef, imSefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    'Loop                'loop on next category in region
                        Next ilLoopOnSEF
                    End If        'ilfoundSEF
    
                '3-15-10 if copy region report, no need for the categories, just the region names
                'if Split copy report, need a category
                If ilFoundSEF Or ilListIndex = CR_COPYREGION Then          'got at least one category for the region
                    'write out the header (network vs copy split)
                    tmGrf.lCode4 = tmRaf(ilLoopOnRAF).lCode             'RAF Code
                    tmGrf.iAdfCode = tmRaf(ilLoopOnRAF).iAdfCode     'advertiser code
                    'tmGrf.iPerGenl(1) = 0               'header
                    tmGrf.iPerGenl(0) = 0               'header
                    'tmGrf.sBktType = slCategory   'M = market, s = station, O = owner, Z = zip, n = state name
                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                End If
            End If              'ilOKList
        Next ilLoopOnRAF        'loop on next region
        Erase tltempstations, tmRaf
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmRaf)
        ilRet = btrClose(hmSef)
         btrDestroy hmGrf
        btrDestroy hmRaf
        btrDestroy hmSef
     End If
    
    If ilListIndex = CR_SPLITREGION Then        'split copy or split network
        ilRet = gObtainFormats()
        ilRet = gObtainMarkets()
        ilRet = gObtainMSAMarkets()
        ilRet = gObtainStates()
        ilRet = gObtainTimeZones()

'        ReDim tlRegionDefinition(0 To 1) As REGIONDEFINITION
'        ReDim tlSplitCategoryInfo(0 To 500) As SPLITCATEGORYINFO
'        ReDim tmRegionDefinition(0 To 0) As REGIONDEFINITION
'        ReDim tmSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
'        ReDim tmRaf(0 To 0) As RAF
'        'ReDim slSplitCatDesc(1 To 1) As String
'        ReDim slSplitCatDesc(0 To 0) As String
        'get the advertiser and the region selected for that adverter

        For ilTemp = 0 To RptSelSR!lbcSelection(0).ListCount - 1
            If RptSelSR!lbcSelection(0).Selected(ilTemp) Then
                slNameCode = tgAdvertiser(RptSelSR!lbcSelection(0).ListIndex).sKey   'Traffic!lbcAdvertiser.List(lbcAdvt.ListIndex - 1)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilAdfCode = Val(Trim$(slCode))
                Exit For
            End If
        Next ilTemp
        
        ilWhichBox = 1
        If RptSelSR!rbcWhichSplit(0).Value = True Then
            ilWhichBox = 2
        End If
        For ilTemp = 0 To RptSelSR!lbcSelection(ilWhichBox).ListCount - 1
            If RptSelSR!lbcSelection(ilWhichBox).Selected(ilTemp) Then
                slNameCode = tgRegionCode(ilTemp).sKey
                'slNameCode = tgRegionCode(RptSelSR!lbcSelection(ilWhichBox).ListIndex).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                llRafCode = Val(Trim$(slCode))
    '                Exit For
    '            End If
    '        Next ilTemp
                ReDim tlRegionDefinition(0 To 1) As REGIONDEFINITION
                ReDim tlSplitCategoryInfo(0 To 500) As SPLITCATEGORYINFO
                ReDim tmRegionDefinition(0 To 0) As REGIONDEFINITION
                ReDim tmSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
                ReDim tmRaf(0 To 0) As RAF
                'ReDim slSplitCatDesc(1 To 1) As String
                ReDim slSplitCatDesc(0 To 0) As String
                
                tmRafSrchKey0.lCode = llRafCode
                ilRet = btrGetEqual(hmRaf, tmRaf(0), imRafRecLen, tmRafSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If (ilRet = BTRV_ERR_NONE) And (tmRaf(0).lCode = llRafCode) Then
        
                    tlRegionDefinition(0).lRotNo = 0
                    tlRegionDefinition(0).lRafCode = tmRaf(0).lCode
                    tlRegionDefinition(0).sCategory = Trim$(tmRaf(0).sCategory)
                    tlRegionDefinition(0).sInclExcl = tmRaf(0).sInclExcl
                    tlRegionDefinition(0).sRegionName = tmRaf(0).sName
                    tlRegionDefinition(0).lFormatFirst = -1
                    tlRegionDefinition(0).lOtherFirst = -1
                    tlRegionDefinition(0).lExcludeFirst = -1
                    tlRegionDefinition(0).sPtType = ""
                    tlRegionDefinition(0).lCopyCode = 0
                    tlRegionDefinition(0).lCrfCode = 0
                    tlRegionDefinition(0).lRsfCode = 0
                    
                    llUpperSplit = 0
                    gBuildSplitCategoryInfo hmSef, llUpperSplit, tlRegionDefinition(0), tlSplitCategoryInfo(), RptSelSR
                    ReDim Preserve tlSplitCategoryInfo(0 To llUpperSplit) As SPLITCATEGORYINFO
                    gSeparateRegions tlRegionDefinition(), tlSplitCategoryInfo(), tmRegionDefinition(), tmSplitCategoryInfo()
                    
                    'For ilLoopOnStations = 1 To UBound(tgStations) - 1
                    For ilLoopOnStations = LBound(tgStations) To UBound(tgStations) - 1
                        ilShttInx = gBinarySearchStation(tgStations(ilLoopOnStations).iCode)
                        If ilShttInx = -1 Then
                            Exit Sub
                        End If
                        ilShttCode = tgStations(ilShttInx).iCode
                        
                        ilMktCode = tgStations(ilShttInx).iMktCode
                        ilMSAMktCode = tgStations(ilShttInx).iMetCode
                        slState = tgStations(ilShttInx).sState
                        ilFmtCode = tgStations(ilShttInx).iFmtCode
                        ilTztCode = tgStations(ilShttInx).iTztCode
                        ilFound = gRegionTestDefinition(ilShttCode, ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, tmRegionDefinition(), tmSplitCategoryInfo(), llRegionIndex, slGroupInfo, RptSelSR)
                        If ilFound Then
                        
                            tmGrf.lCode4 = tmRaf(0).lCode             'RAF Code
                            tmGrf.iAdfCode = tmRaf(0).iAdfCode     'advertiser code
                            tmGrf.iSofCode = 0                  'init owner (ARTT) code reference
                            tmGrf.iSlfCode = 0                  'init station (SHTT) code reference
                            tmGrf.iRdfCode = 0                  'init market (MKT) code reference
                            tmGrf.iCode2 = 0                    'time zone (TZT) code reference
                            tmGrf.iYear = 0                     'format code (FMT) reference
                            tmGrf.sGenDesc = ""
                            'tmGrf.iPerGenl(1) = 1               'sub records
                            tmGrf.iPerGenl(0) = 1               'sub records
                                   
                            'because the concatenation of the categories can exceed 40 char (field in prepass record), create
                            'a sequence number so all like concatenation of categories fall together for sorting purposes in crystal
                            ilFound = False
                            For ilLoopOnCat = LBound(slSplitCatDesc) To UBound(slSplitCatDesc) - 1
                                If StrComp(UCase(Trim$((slGroupInfo))), UCase(Trim$(slSplitCatDesc(ilLoopOnCat))), vbBinaryCompare) = 0 Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoopOnCat
                            If Not ilFound Then
                                slSplitCatDesc(UBound(slSplitCatDesc)) = Trim$(slGroupInfo)
                                ilLoopOnCat = UBound(slSplitCatDesc)
                                ReDim Preserve slSplitCatDesc(LBound(slSplitCatDesc) To UBound(slSplitCatDesc) + 1) As String
                            End If
        
                            If Not ilFound Then
                                tmGrf.lLong = ilLoopOnCat + tmRaf(0).lCode          '6-5-17 keep different region names apart in crystal
                                'create the prepass record to form the categories that make up valid station
                                mFormRegionCategory ilShttCode, ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, slGroupInfo
                            End If
                            
                            tmGrf.sGenDesc = ""
                            tmGrf.iSofCode = 0                  ' owner (ARTT) code reference
                            tmGrf.iSlfCode = ilShttCode         ' station (SHTT) code reference
                            tmGrf.iRdfCode = ilMktCode          ' market (MKT) code reference
                            tmGrf.iCode2 = ilTztCode             'time zone (TZT) code reference
                            tmGrf.iYear = ilFmtCode              'format code (FMT) reference
                            tmGrf.sGenDesc = slState            'postal name
                            'tmGrf.iPerGenl(1) = 0               'detail for station info
                            tmGrf.iPerGenl(0) = 0               'detail for station info
                            tmGrf.iVefCode = ilMSAMktCode       'msa market
                            tmGrf.lLong = ilLoopOnCat + tmRaf(0).lCode               '6-5-17 keep different region names apart in crystal
                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                        End If
                    Next ilLoopOnStations
                End If
'            Exit For
            End If      'selected
        Next ilTemp
        Erase tmRegionDefinition, tmSplitCategoryInfo, tlRegionDefinition, tlSplitCategoryInfo

        ilRet = btrClose(hmSef)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmRaf)
        btrDestroy hmSef
        btrDestroy hmGrf
        btrDestroy hmRaf
    End If
    Exit Sub

End Sub

Public Function mTestCategory(slCategory As String) As Integer
Dim ilFound As Integer
            ilFound = False
            Select Case slCategory       'determine if the category should be included
                Case "M"            'market
                    If imInclCatMkt Then
                        ilFound = True
                    End If
                '10/29: Darlene- I removed Owner and Zip
                'Case "O"    'Owner
                '    If ilInclCatOwner Then
                '        ilFound = True
                '    End If
                'Case "Z"    'Zip Code
                '    If ilInclCatZip Then
                '        ilFound = True
                '    End If
                Case "A"            'dma market
                    If imInclCatMSAMkt Then
                        ilFound = True
                    End If
                Case "N"    'State
                    If imInclCatState Then
                        ilFound = True
                    End If
                Case "S"    'Station
                    If imInclCatStation Then
                        ilFound = True
                    End If
                 Case "F"    'Format
                    If imInclCatFormat Then
                        ilFound = True
                    End If
                 Case "T"    'Time Zone
                    If imInclCatTime Then
                        ilFound = True
                    End If
            End Select
            mTestCategory = ilFound
End Function
'
'           Create prepass records to form the categories that make up the
'           valid station for the region
Private Sub mFormRegionCategory(ilShttCode As Integer, ilMktCode As Integer, ilMSAMktCode As Integer, slState As String, ilFmtCode As Integer, ilTztCode As Integer, slInGroupInfo As String)
    'Translate the User defined region into region definitions
    'User enters:
    'Urban and California
    '(Format, State, Time zone and Market.  For stations, the call letters are used)
    'Symbols:  ^ = And; ~ = Not
    Dim ilPos As Integer
    Dim slGroupInfo As String
    Dim slStr As String
    Dim slInclExcl As String
    Dim slCategory As String
    Dim slvalue As String
    Dim ilValue As Integer
    Dim ilSnt As Integer
    Dim slAddress As String
    Dim ilRet As Integer
    Dim slGroupName As String
    Dim ilShtt As Integer
    Dim ilFind As Integer
    Dim slChar As String
    Dim llSerialNo1 As Long
    Dim llTestSerialNo1 As Long
    Dim ilImport As Integer
    Dim ilLoop As Integer
    
    On Error GoTo mFormRegionCategoryErr:
    slAddress = ""
    slGroupInfo = slInGroupInfo
    ilPos = 1
    Do
        ilPos = InStr(1, slGroupInfo, "|", vbTextCompare)
        If ilPos = 0 Then
            If Len(slGroupInfo) = 0 Then
                Exit Do
            Else
                ilPos = Len(slGroupInfo) + 1
            End If
        End If
        slStr = Left(slGroupInfo, ilPos - 1)
        slGroupInfo = Mid$(slGroupInfo, ilPos + 1)
        slInclExcl = Left$(slStr, 1)
        slCategory = Mid$(slStr, 2, 1)
        slvalue = Trim$(Mid$(slStr, 3))
        If slCategory <> "N" Then
            ilValue = Val(slvalue)
        End If
        tmGrf.sBktType = ""
        Select Case slCategory
            Case "M"    'DMA Market
                ilRet = gBinarySearchMkt(CLng(ilValue))
                If ilRet <> -1 Then
                    slGroupName = Trim$(tgMarkets(ilRet).sName)
                End If
                tmGrf.sBktType = "M"
            Case "A"    'MSA Market
                ilRet = gBinarySearchMSAMkt(CLng(ilValue))
                If ilRet <> -1 Then
                    slGroupName = Trim$(tgMSAMarkets(ilRet).sName)
                End If
                tmGrf.sBktType = "A"
            Case "N"    'State Name
                For ilSnt = 0 To UBound(tgStates) - 1 Step 1
                    If StrComp(Trim$(tgStates(ilSnt).sPostalName), slvalue, vbTextCompare) = 0 Then
                        slGroupName = Trim$(tgStates(ilSnt).sName)
                        Exit For
                    End If
                Next ilSnt
                tmGrf.sBktType = "N"
            Case "F"    'Format
                ilRet = gBinarySearchFmt(ilValue)
                If ilRet <> -1 Then
                    slGroupName = Trim$(tgFormats(ilRet).sName)
                End If
                tmGrf.sBktType = "F"
            Case "T"    'Time zone
                ilRet = gBinarySearchTzt(ilValue)
                If ilRet <> -1 Then
                    slGroupName = Trim$(tgTimeZones(ilRet).sName)
                End If
                tmGrf.sBktType = "T"
            Case "S"    'Station
               ilRet = gBinarySearchStation(ilValue)
               slGroupName = Trim$(tgStations(ilRet).sCallLetters)
                tmGrf.sBktType = "S"
        End Select
        If slInclExcl = "E" Then
            slGroupName = "Not " & slGroupName
            tmGrf.sBktType = "X"            'excluded stations
        End If
'        If slCategory <> "S" Then   'category not station
'            If slAddress = "" Then
'                slAddress = slGroupName
'            Else
'                slAddress = slAddress & "^" & slGroupName
'            End If
'        Else
'            If slAddress = "" Then
'                slAddress = slGroupName
'            Else
'                slAddress = slAddress & "^" & slGroupName
'            End If
'        End If

'       write out the category to form the valid station
        tmGrf.sGenDesc = Trim$(slGroupName)
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
    Loop While Len(slGroupInfo) > 0
    
    Exit Sub
mFormRegionCategoryErr:
    Resume Next
End Sub
