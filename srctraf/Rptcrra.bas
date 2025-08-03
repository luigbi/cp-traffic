Attribute VB_Name = "RPTCRRA"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrra.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
Dim hmAsf As Integer
Dim tmAsf As ASF
Dim imAsfRecLen As Integer
Dim hmPvf As Integer
Dim tmPvf As PVF
Dim imPvfRecLen As Integer
Dim tmLongSrchKey As LONGKEY0
Dim hmGrf As Integer
Dim tmGrf As GRF
Dim imGrfRecLen As Integer
Type STDPKGINFO
    iPkgVef As Integer
    iHiddenVef As Integer
    iRdfCode As Integer 'Rate Card daypart code
    iNoSpot As Integer   'Number of spots
    iPctRate As Integer '% Rate Split (for vefStdPrice = 2) (xxx.xx)
End Type
Type RDFANF
    iRdfCode As Integer
    ianfCode As Integer
End Type




'
'
'           Create Prepass for package avails
'           Gather all the valid vehicles for each Standard package
'
'           Created  1/23/01    d.hosaka
'
'
'            tmGrf.iGenDate - Generation date
'            tmGrf.iGenTime - Generation time
'            tmGrf.iPerGenl(1) - tmAsf.iVefAutoCode      'summary vehicle code
'            tmGrf.iPerGenl(2) - tmAsf.iRcfCode          'rate card that generated this summary
'            tmGrf.ivefCode = tlPkgVef(ilFindVef).iPkgVef    'package vehicle name
'            tmGrf.lCode4 = tmAsf.lCode                   'auto increment code
'            tmGrf.iPerGenl(3) = tlPkgVef(ilFindVef).iNoSpot    '# spots from Standard package defined with the hidden vehicle
'            tmGrf.iPerGenl(4) = flag to indicate whether a 30/60 avail summary exists, 10" avail summary exists, both exists or none exists
'                               0 = only zeroes exist, 1 = only 30/60 summary exist, 2 = only 10 summary exists, 3 = both summaries exist
'
'
Sub gCreatePkgAvails()
Dim llNoRec As Long
Dim ilRet  As Integer
Dim ilLoop As Integer
Dim slNameCode As String
Dim ilTemp As Integer
Dim slCode As String
Dim ilUpper As Integer
Dim ilFindVef As Integer
Dim llPvfCode As Long
Dim ilFound As Integer
Dim ilExtLen As Integer
Dim llRecPos As Long
Dim ilLoopAvail As Integer
Dim ilLoopHidden As Integer
Dim ilRifShowLoop As Integer
Dim il10 As Integer                 'do 10s exist
Dim il30 As Integer                 'do 30s exist
Dim il60 As Integer                 'do 60s exist
Dim ilCount10 As Integer
Dim ilCount30 As Integer
Dim ilCount60 As Integer
Dim ilOnlyDPNamedAvails As Integer  'true if user selected to included only those dayparts defined to
                                    'book into a specific named avail

Dim ilAllAvails As Integer          'only applicable if ilOnlyDPNamedAvails is true. Then the user
                                    'may have selected named avails from the selection list
ReDim tlPkgVef(0 To 0) As STDPKGINFO
Dim tlRcf As RCF
ReDim tlRif(0 To 0) As RIF              'items within rate card
ReDim tlRdf(0 To 0) As RDF              'all Daypart records in rate card
ReDim tlRifShow(0 To 0) As RIF          'only daypart records that are "show on report = Y"
ReDim tlRdfAnf(0 To 0) As RDFANF        'array of dayparts and named avails designated to "book into" (i.e. book into News)
ReDim ilSelectedAnf(0 To 0) As Integer  'array of selected named avail codes
    hmAsf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAsf, "", sgDBPath & "Asf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAsf)
        btrDestroy hmAsf
        Exit Sub
    End If
    imAsfRecLen = Len(tmAsf)
    hmPvf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPvf, "", sgDBPath & "Pvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPvf)
        btrDestroy hmPvf
        Exit Sub
    End If
    imPvfRecLen = Len(tmPvf)

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmAsf)
        btrDestroy hmGrf
        btrDestroy hmAsf
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    ilOnlyDPNamedAvails = False                 'assume not selecting only those DP with "Book Into" named avails
    ilAllAvails = True
    If RptSelRA!ckcAllAvails.Value = vbChecked Then         'only look for those dayparts booking into a selected named avail
        ilOnlyDPNamedAvails = True
        If Not RptSelRA!ckcBookInto.Value = vbChecked Then
            ilAllAvails = False
        End If
    End If

    If igRARcfCode = 0 Then           'no date exists
        Erase tlRif, tlRdf, tlPkgVef, ilSelectedAnf
        Erase tlRifShow, tlRdfAnf
        sgRCStamp = ""              'init for re-entrant problem
        ilRet = btrClose(hmAsf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmPvf)
        btrDestroy hmAsf
        btrDestroy hmGrf
        btrDestroy hmPvf
        Exit Sub
    End If

    gRCRead RptSelRA, igRARcfCode, tlRcf, tlRif(), tlRdf() 'obtain the lastest r/c and its dayparts & items

    'create a smaller array of the Dayparts that are only "book into" a specific named avail
    ilUpper = 0
    If ilOnlyDPNamedAvails Then 'user wants to filter out all dayparts except those with a book into named avail
        'build array of all selected named avail codes
        For ilLoopAvail = 0 To RptSelRA!lbcSelection(2).ListCount - 1 Step 1
            If RptSelRA!lbcSelection(2).Selected(ilLoopAvail) Then
                slNameCode = tgNamedAvail(ilLoopAvail).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilSelectedAnf(ilUpper) = Val(slCode)
                ilUpper = ilUpper + 1
                ReDim Preserve ilSelectedAnf(0 To ilUpper) As Integer
            End If
        Next ilLoopAvail
    End If
    'build array of the dayparts that are defined as "book into" named avails
    ilUpper = 0
    For ilLoopAvail = 0 To UBound(tlRdf)
        If tlRdf(ilLoopAvail).sInOut = "I" Then   'Book into
            'is this "book into" match any of the selected ones? If so, build this daypart in array
            For ilLoop = 0 To UBound(ilSelectedAnf) - 1
                If tlRdf(ilLoopAvail).ianfCode = ilSelectedAnf(ilLoop) Then
                    tlRdfAnf(ilUpper).iRdfCode = tlRdf(ilLoopAvail).iCode
                    tlRdfAnf(ilUpper).ianfCode = tlRdf(ilLoopAvail).ianfCode
                    ilUpper = ilUpper + 1
                    ReDim Preserve tlRdfAnf(0 To ilUpper) As RDFANF
                    Exit For
                End If
            Next ilLoop
        End If
    Next ilLoopAvail
    'build array of items in rate card that are defined to "show on report"
    ilUpper = 0
    For ilLoopAvail = 0 To UBound(tlRif)
        If tlRif(ilLoopAvail).sRpt = "Y" Then       'show on report
            tlRifShow(ilUpper) = tlRif(ilLoopAvail)
            ilUpper = ilUpper + 1
            ReDim Preserve tlRifShow(0 To ilUpper) As RIF
        End If
    Next ilLoopAvail
    ilUpper = 0
    'Build array of all the hidden vehicles for each standard package
    For ilLoop = 0 To RptSelRA!lbcSelection(1).ListCount - 1 Step 1
        If RptSelRA!lbcSelection(1).Selected(ilLoop) Then
            slNameCode = tgSellNameCodeRA(ilLoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            'see what hidden vehicles are associated with this standard package
            llPvfCode = 0
            'For ilFindVef = LBound(tgMVef) To UBound(tgMVef)
            '    If tgMVef(ilFindVef).iCode = Val(slCode) Then
                ilFindVef = gBinarySearchVef(Val(slCode))
                If ilFindVef <> -1 Then
                    llPvfCode = tgMVef(ilFindVef).lPvfCode          'vehicles defined in std pkg
            '        Exit For
                End If
            'Next ilFindVef
            Do While llPvfCode > 0
                tmLongSrchKey.lCode = llPvfCode
                ilRet = btrGetEqual(hmPvf, tmPvf, imPvfRecLen, tmLongSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                On Error GoTo gCreatPkgAvailsErr
                'Pvf has the defined vehicles within the package
                For ilLoopHidden = 0 To 24
                    If tmPvf.iVefCode(ilLoopHidden) > 0 Then
                        tlPkgVef(ilUpper).iPkgVef = Val(slCode)
                        tlPkgVef(ilUpper).iHiddenVef = tmPvf.iVefCode(ilLoopHidden)
                        tlPkgVef(ilUpper).iRdfCode = tmPvf.iRdfCode(ilLoopHidden)
                        tlPkgVef(ilUpper).iNoSpot = tmPvf.iNoSpot(ilLoopHidden)
                        ReDim Preserve tlPkgVef(0 To ilUpper + 1)
                        ilUpper = ilUpper + 1
                    Else
                        Exit For
                    End If
                Next ilLoopHidden
                llPvfCode = tmPvf.lLkPvfCode
            Loop
        End If
    Next ilLoop

    'Read all Avails summary and see what standard package the vehicle belongs.  Print only those vehicles with dayparts that should
    'be shown on Avails reports.

    btrExtClear hmAsf   'Clear any previous extend operation
    ilExtLen = Len(tmAsf) 'Extract operation record size

    ilRet = btrGetFirst(hmAsf, tmAsf, imAsfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hmAsf, llNoRec, -1, "UC", "ASF", "") '"EG") 'Set extract limits (all records)
        ilRet = btrExtAddField(hmAsf, 0, ilExtLen)  'Extract the whole record
        On Error GoTo gCreatPkgAvailsErr
        gBtrvErrorMsg ilRet, "gObtainAsf (btrExtAddField):" & "Asf.Btr", RptSelRA
        On Error GoTo 0
        ilRet = btrExtGetNext(hmAsf, tmAsf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gCreatPkgAvailsErr
            gBtrvErrorMsg ilRet, "gCreatePkgAvails (btrExtGetNextExt):" & "Asf.Btr", RptSelRA
            On Error GoTo 0
            ilExtLen = Len(tmAsf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmAsf, tmAsf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                '1) Find all matching standard packages that this vehicles avail summary record belongs in .
                    'The summary records daypart must also match the hidden lines daypart
                '2) if not found, the vehicles summary doesnt belong in any standard package--ignore it
                '3) if found, loop thru the items (RIF) for the matching vehicle/daypart to see if the Show on Report is YES
                '4) if found, did the user ask for only those dayparts that are defined as 'book into' named avails
                    'if yes, was the named avail selected to include
                '5) if all above true, create GRF to report results
                For ilFindVef = 0 To UBound(tlPkgVef)
                    'test all matching packages that this avail summary record belongs with
                    'If tmAsf.iVefAutoCode = tlPkgVef(ilFindVef).iHiddenVef And tmAsf.iRdfAutoCode = tlPkgVef(ilFindVef).iRdfcode Then
                    If tmAsf.iVefCode = tlPkgVef(ilFindVef).iHiddenVef And tmAsf.iRdfCode = tlPkgVef(ilFindVef).iRdfCode Then
                        For ilRifShowLoop = LBound(tlRifShow) To UBound(tlRifShow) - 1
                            'find the matching vehicle & daypart in RIF to see if it is marked to Show on Report
                            'Determine if this vehicle has dayparts that shouldnt be shown
                            'If tlRifShow(ilRifShowLoop).iVefCode = tmAsf.iVefAutoCode And tlRifShow(ilRifShowLoop).iRdfcode = tmAsf.iRdfAutoCode Then
                            If tlRifShow(ilRifShowLoop).iVefCode = tmAsf.iVefCode And tlRifShow(ilRifShowLoop).iRdfCode = tmAsf.iRdfCode Then
                                'found a matching daypart & vehicle that should be shown, test to see
                                'if only certain named avails to be included
                                ilFound = True
                                'Include only dayparts that is defined to "book into" a specific avail?  If so, check to
                                'see which named avails they want to include
                                If ilOnlyDPNamedAvails And Not ilAllAvails Then
                                    'find the daypart to see which named avail to book into
                                    'loop thru all the named avails
                                    ilFound = False
                                    For ilLoopAvail = 0 To UBound(tlRdfAnf) - 1
                                        If tlRdfAnf(ilLoopAvail).iRdfCode = tlRifShow(ilRifShowLoop).iRdfCode Then
                                            'see if this dayparts "BookInto" named avail code matches one that was selected
                                            For ilTemp = 0 To UBound(ilSelectedAnf) - 1
                                                If tlRdfAnf(ilLoopAvail).ianfCode = ilSelectedAnf(ilTemp) Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilTemp
                                        End If
                                    Next ilLoopAvail
                                End If
                                If ilFound Then
                                    tmGrf.iGenDate(0) = igNowDate(0)
                                    tmGrf.iGenDate(1) = igNowDate(1)
                                    'tmGrf.iGenTime(0) = igNowTime(0)
                                    'tmGrf.iGenTime(1) = igNowTime(1)
                                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                    tmGrf.lGenTime = lgNowTime
                                    'tmGrf.iPerGenl(1) = tmAsf.iVefCode  'tmAsf.iVefAutoCode      'summary vehicle code
                                    'tmGrf.iPerGenl(2) = tmAsf.iRcfCode          'rate card that generated this summary
                                    tmGrf.iPerGenl(0) = tmAsf.iVefCode  'tmAsf.iVefAutoCode      'summary vehicle code
                                    tmGrf.iPerGenl(1) = tmAsf.iRcfCode          'rate card that generated this summary
                                    tmGrf.iVefCode = tlPkgVef(ilFindVef).iPkgVef    'package vehicle name
                                    tmGrf.lCode4 = tmAsf.lCode
                                    'tmGrf.iPerGenl(3) = tlPkgVef(ilFindVef).iNoSpot
                                    tmGrf.iPerGenl(2) = tlPkgVef(ilFindVef).iNoSpot
                                    'Determine if there are 30/60, 10s or both so that Crystal knows which lines to show
                                    il10 = False
                                    il30 = False
                                    il60 = False
'                                    For ilLoop = 1 To 26 Step 2         'look at the odd values
                                    For ilLoop = 1 To 25                  '9-28-18
                                        ilCount30 = Asc(tmAsf.sAvail30(ilLoop)) Mod 256
                                        ilCount60 = Asc(tmAsf.sAvail60(ilLoop)) Mod 256
                                        ilCount10 = Asc(tmAsf.sAvail10(ilLoop)) Mod 256
                                        If ilCount30 <> 0 Then
                                            il30 = True
                                        End If
                                        If ilCount60 <> 0 Then
                                            il60 = True
                                        End If
                                        If ilCount10 <> 0 Then
                                            il10 = True
                                        End If
                                    Next ilLoop
'                                    For ilLoop = 2 To 26 Step 2         'look at the even values
                                     For ilLoop = 0 To 24 Step 1        'look at the even values
                                        ilCount30 = Asc(tmAsf.sAvail30(ilLoop)) \ 256
                                        ilCount60 = Asc(tmAsf.sAvail60(ilLoop)) \ 256
                                        ilCount10 = Asc(tmAsf.sAvail10(ilLoop)) \ 256
                                        If ilCount30 <> 0 Then
                                            il30 = True
                                        End If
                                        If ilCount60 <> 0 Then
                                            il60 = True
                                        End If
                                        If ilCount10 <> 0 Then
                                            il10 = True
                                        End If
                                    Next ilLoop
'                                    tmGrf.iPerGenl(4) = 0
'                                    If (il30 Or il60) And il10 Then     'both 30/60 & 10 exist; show both lines of avails
'                                        tmGrf.iPerGenl(4) = 3
'                                    ElseIf Not (il30 Or il60) And il10 Then
'                                        tmGrf.iPerGenl(4) = 2           '10 counts only
'                                    Else
'                                        tmGrf.iPerGenl(4) = 1           'show 30/60 counts only
'                                    End If
                                    tmGrf.iPerGenl(3) = 0
                                    If (il30 Or il60) And il10 Then     'both 30/60 & 10 exist; show both lines of avails
                                        tmGrf.iPerGenl(3) = 3
                                    ElseIf Not (il30 Or il60) And il10 Then
                                        tmGrf.iPerGenl(3) = 2           '10 counts only
                                    Else
                                        tmGrf.iPerGenl(3) = 1           'show 30/60 counts only
                                    End If

                                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        ilRet = btrClose(hmAsf)
                                        ilRet = btrClose(hmGrf)
                                        btrDestroy hmAsf
                                        btrDestroy hmGrf
                                    End If
                                End If
                            End If      'tlRifShow(ilRifShowLoop).ivefCode = tmAsf.iVefAutoCode And tlRifShow(ilRifShowLoop).irdfCode = tmAsf.iRdfAutoCode
                        Next ilRifShowLoop
                    End If              'tmAsf.iVefAutoCode = tlPkgVef(ilFindVef).iHiddenVef And tmAsf.iRdfAutoCode = tlPkgVef(ilFindVef).iHiddenVef
                Next ilFindVef
                'ilRet = btrGetNext(hmAsf, tmAsf, imAsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                ilRet = btrExtGetNext(hmAsf, tmAsf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmAsf, tmAsf, ilExtLen, llRecPos)
                Loop
            Loop    'Do While ilRet = BTRV_ERR_NONE
        End If      'if (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT)
    End If          'if ilRet <> BTRV_ERR_END_OF_FILE
    Erase tlRif, tlRdf, tlPkgVef, ilSelectedAnf
    Erase tlRifShow, tlRdfAnf
    sgRCStamp = ""              'init for re-entrant problem
    ilRet = btrClose(hmAsf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmPvf)
    btrDestroy hmAsf
    btrDestroy hmGrf
    btrDestroy hmPvf
    Exit Sub
gCreatPkgAvailsErr:
    On Error GoTo 0
    ilRet = btrClose(hmAsf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmPvf)
    btrDestroy hmAsf
    btrDestroy hmGrf
    btrDestroy hmPvf
    Exit Sub
End Sub
