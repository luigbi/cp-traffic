Attribute VB_Name = "RptCrFD"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptcrfd.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
'Public igYear As Integer                'budget year used for filtering
'Public igMonthOrQtr As Integer          'entered month or qtr
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer

Dim hmFsf As Integer            'Feed file handle
Dim tmFsf As FSF                'FSF record image
Dim imFsfRecLen As Integer       'FSF record length

Dim hmFdf As Integer            'Feed data file handle
Dim tmFdf As FDF                'FSF record image
Dim imFdfRecLen As Integer       'FDF record length

Dim hmFpf As Integer            'Feed headerfile handle
Dim tmFpf As FPF                'FPF record image
Dim imFpfRecLen As Integer       'FPF record length

Dim hmPrf As Integer            'Product file handle
Dim tmPrf As PRF                'PRF record image
Dim imPrfRecLen As Integer       'PRF record length


Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length

Dim tmPFF As PFF
Dim hmPff As Integer
Dim imPffRecLen As Integer        'PFF record length

Dim tmPbf As PBF
Dim hmPbf As Integer
Dim imPbfRecLen As Integer        'PBF record length
Dim tmPbfSrchKey1 As LONGKEY0


'********************************************************************************************
'
'                   gCrFeedRecap - Prepass for Feed Recap report
'                   Gather feed spots by date/time spans and create
'                   a prepass of all valid spot in GRF.
'
'                   User selectivity:  STart/End DAtes
'                                      Start/End Times
'                                      Feed Names
'
'                   Created:  8/4/04 D. Hosaka
'********************************************************************************************
Sub gCrFeedRecap()
Dim ilRet As Integer                    '
Dim ilFoundOne As Integer               'Found a matching  office built into mem
Dim ilLoop As Integer                   'temp loop variable
Dim slNameCode As String
Dim slCode As String
Dim ilUpper As Integer
Dim slStartTime As String           'user entered start time
Dim slEndTime As String             'user entered end time
Dim llStartTime As Long             'user entered start time
Dim llEndTime As Long               'user entered end time
Dim slStartDate As String           'user entered start date
Dim slEndDate As String             'user entered end date
Dim llStartDate As Long
Dim llEndDate As Long
Dim ilFsf As Integer
Dim tlFsf() As FSF

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmPrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmGrf)
        btrDestroy hmPrf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imPrfRecLen = Len(tmPrf)

    hmFsf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmGrf)
        btrDestroy hmFsf
        btrDestroy hmPrf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imFsfRecLen = Len(tmFsf)

    ReDim ilSelectedVehicles(0 To 0) As Integer
    ilUpper = 0
    If (Not rptSelFD!ckcAllVehicles.Value = vbChecked) Then                               'slsp, check if any of the split slsp should be excluded
        For ilLoop = 0 To rptSelFD!lbcSelection(1).ListCount - 1 Step 1
            If rptSelFD!lbcSelection(1).Selected(ilLoop) Then              'selected slsp
                slNameCode = tgVehicle(ilLoop).sKey        'pick up slsp code
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilSelectedVehicles(ilUpper) = Val(slCode)
                ilUpper = ilUpper + 1
                ReDim Preserve ilSelectedVehicles(0 To ilUpper)
            End If
        Next ilLoop
    End If

    ReDim ilSelectedFeedNames(0 To 0) As Integer
    ilUpper = 0
    If (Not rptSelFD!ckcAll.Value = vbChecked) Then                               'slsp, check if any of the split slsp should be excluded
        For ilLoop = 0 To rptSelFD!lbcSelection(0).ListCount - 1 Step 1
            If rptSelFD!lbcSelection(0).Selected(ilLoop) Then              'selected slsp
                slNameCode = tgRptNameCode(ilLoop).sKey        'pick up slsp code
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilSelectedFeedNames(ilUpper) = Val(slCode)
                ilUpper = ilUpper + 1
                ReDim Preserve ilSelectedFeedNames(0 To ilUpper)
            End If
        Next ilLoop
    End If

    slStartTime = rptSelFD!edcFromTime.Text        'start time
    llStartTime = gTimeToLong(slStartTime, False)
    slEndTime = rptSelFD!edcToTime.Text            'end time
    llEndTime = gTimeToLong(slEndTime, True)

    slStartDate = rptSelFD!edcFromDate.Text   'Start date
    llStartDate = gDateValue(slStartDate)
    slEndDate = rptSelFD!edcToDate.Text   'End date
    llEndDate = gDateValue(slEndDate)

    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for retrieval/removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime

    ReDim tlFsf(0 To 0) As FSF
    ilRet = gObtainFSF(rptSelFD, hmFsf, tlFsf(), slStartDate, slEndDate, slStartTime, slEndTime)
    For ilFsf = LBound(tlFsf) To UBound(tlFsf) - 1
        tmFsf = tlFsf(ilFsf)

        ilFoundOne = False
        If rptSelFD!ckcAll = vbChecked Then             'check for selected feed name
            ilFoundOne = True
        Else
            For ilLoop = 0 To UBound(ilSelectedFeedNames) - 1 Step 1
                If ilSelectedFeedNames(ilLoop) = tmFsf.iFnfCode Then
                    ilFoundOne = True
                    Exit For
                End If
            Next ilLoop
        End If

        If ilFoundOne Then              'got valid feed name, now check for valid vahicle
            ilFoundOne = False
            If rptSelFD!ckcAllVehicles = vbChecked Then         'all vehicles selected
                ilFoundOne = True
            Else
                For ilLoop = 0 To UBound(ilSelectedVehicles) - 1 Step 1
                    If ilSelectedVehicles(ilLoop) = tmFsf.iVefCode Then
                        ilFoundOne = True
                        Exit For
                    End If
                Next ilLoop
            End If
        End If

        If ilFoundOne Then
            tmGrf.iAdfCode = tmFsf.iAdfCode
            tmGrf.lChfCode = tmFsf.lCode
            tmGrf.iCode2 = tmFsf.iFnfCode
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        End If
    Next ilFsf

    ilRet = btrClose(hmFsf)
    ilRet = btrClose(hmPrf)
    ilRet = btrClose(hmGrf)
    btrDestroy hmFsf
    btrDestroy hmPrf
    btrDestroy hmGrf
End Sub
'
'
'                gCrFeedPledge - create prepass to report on the Pledge
'                data for Logs needing conversions
'

'
'           gCrFeedPledge - generate pre-pass for the Pledge report
'           Gather FPF headers to determine selected start/end dates,
'           vehicles or feed names.  Then get associated pledge data
'           and show feed/pledge days/times.
'
Public Sub gCrFeedPledge()
Dim ilRet As Integer
Dim ilLoop As Integer                   'temp loop variable
Dim slNameCode As String
Dim slCode As String
Dim ilUpper As Integer
Dim slStartDate As String           'user entered start date
Dim slEndDate As String             'user entered end date
Dim llStartDate As Long
Dim llEndDate As Long
Dim ilFNFLoop As Integer
Dim ilVefLoop As Integer
Dim ilFnfCode As Integer
Dim ilVefCode As Integer
Dim slErrMsg As String
Dim ilOk As Integer
Dim tlFPF() As FPF

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmFdf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFdf, "", sgDBPath & "Fdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFdf)
        ilRet = btrClose(hmGrf)
        btrDestroy hmFdf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imFdfRecLen = Len(tmFdf)

    hmFpf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFpf, "", sgDBPath & "Fpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFpf)
        ilRet = btrClose(hmFdf)
        ilRet = btrClose(hmGrf)
        btrDestroy hmFpf
        btrDestroy hmFdf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imFpfRecLen = Len(tmFpf)

    ReDim ilSelectedVehicles(0 To 0) As Integer
    ilUpper = 0
    If (Not rptSelFD!ckcAllVehicles.Value = vbChecked) Then                               'slsp, check if any of the split slsp should be excluded
        For ilLoop = 0 To rptSelFD!lbcSelection(1).ListCount - 1 Step 1
            If rptSelFD!lbcSelection(1).Selected(ilLoop) Then              'selected slsp
                slNameCode = tgVehicle(ilLoop).sKey        'pick up slsp code
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilSelectedVehicles(ilUpper) = Val(slCode)
                ilUpper = ilUpper + 1
                ReDim Preserve ilSelectedVehicles(0 To ilUpper)
            End If
        Next ilLoop
    End If

    ReDim ilSelectedFeedNames(0 To 0) As Integer
    ilUpper = 0
    For ilLoop = 0 To rptSelFD!lbcSelection(0).ListCount - 1 Step 1
        If rptSelFD!lbcSelection(0).Selected(ilLoop) Then              'selected slsp
            slNameCode = tgRptNameCode(ilLoop).sKey        'pick up slsp code
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilSelectedFeedNames(ilUpper) = Val(slCode)
            ilUpper = ilUpper + 1
            ReDim Preserve ilSelectedFeedNames(0 To ilUpper)
        End If
    Next ilLoop

    slStartDate = rptSelFD!edcFromDate.Text   'Start date
    llStartDate = gDateValue(slStartDate)
    slEndDate = rptSelFD!edcToDate.Text   'End date
    llEndDate = gDateValue(slEndDate)

    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for retrieval/removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime


    For ilFNFLoop = 0 To UBound(ilSelectedFeedNames) - 1
        ReDim tlFPF(0 To 0) As FPF
        ilFnfCode = ilSelectedFeedNames(ilFNFLoop)
        If rptSelFD!ckcAllVehicles.Value = vbChecked Then           'all vehicles selected
            ilVefCode = -1
            ilRet = gObtainFPF(rptSelFD, hmFpf, tlFPF(), slStartDate, slEndDate, ilFnfCode, ilVefCode)
            If Not ilRet Then
                slErrMsg = "Error reading FPF in gCrFeedPledge: gObtainFPF"
                GoTo gCrFeedPledgeErr
            End If
            ilOk = mBuildPledge(tlFPF())
            If Not ilOk Then            'error inserting record
                slErrMsg = "Error inserting GRF or reading FDF in gCrFeedPledge: mBuildPledge"
                GoTo gCrFeedPledgeErr
            End If
        Else
            For ilVefLoop = 0 To UBound(ilSelectedVehicles) - 1   'gather only selected vehicles for the feed name
                ilVefCode = ilSelectedVehicles(ilVefLoop)
                ilRet = gObtainFPF(rptSelFD, hmFpf, tlFPF(), slStartDate, slEndDate, ilFnfCode, ilVefCode)
                If Not ilRet Then
                    slErrMsg = "Error reading FPF in gCrFeedPledge: gObtainFPF"
                    GoTo gCrFeedPledgeErr
                End If
                ilOk = mBuildPledge(tlFPF())
                If Not ilOk Then            'error inserting record
                    slErrMsg = "Error inserting GRF or reading FDF in gCrFeedPledge: mBuildPledge"
                    GoTo gCrFeedPledgeErr
                End If
            Next ilVefLoop
        End If
        If Not ilOk Then            'error inserting record
            Exit For
        End If
    Next ilFNFLoop

    ilRet = btrClose(hmFpf)
    ilRet = btrClose(hmFdf)
    ilRet = btrClose(hmGrf)
    btrDestroy hmFpf
    btrDestroy hmFdf
    btrDestroy hmGrf
    Exit Sub
gCrFeedPledgeErr:
    MsgBox slErrMsg
    ilRet = btrClose(hmFpf)
    ilRet = btrClose(hmFdf)
    ilRet = btrClose(hmGrf)
    btrDestroy hmFpf
    btrDestroy hmFdf
    btrDestroy hmGrf
    Exit Sub
End Sub
'
'
'           mBuildPledge - build the pledge record in GRF .
'           <input> tlFpf() - array of feed headers
'           <return> - true if valid inserts
'
'   GRF fields:
'       grfGenTime - generation time
'       grfGenDate - generation date
'       grfvefcode - vehicle code
'       grfrdfcode - FDF code
'       grfStartDate - FPF effective start date
'       grfEndDate - FDF effective end date
'
Public Function mBuildPledge(tlFPF() As FPF) As Integer
Dim llFPFLoop As Long
Dim ilOk  As Integer
Dim ilRet As Integer
Dim llFdfLoop As Long

    ReDim tlFdf(0 To 0) As FDF
    ilOk = True
    For llFPFLoop = 0 To UBound(tlFPF) - 1
        tmFpf = tlFPF(llFPFLoop)
        ilRet = gObtainFDFByCode(rptSelFD, hmFdf, tlFdf(), tmFpf.iCode)
        If ilRet <> True Then
            Exit Function
        End If
        tmGrf.iVefCode = tmFpf.iVefCode
        tmGrf.iStartDate(0) = tmFpf.iEffStartDate(0)
        tmGrf.iStartDate(1) = tmFpf.iEffStartDate(1)
        tmGrf.iDate(0) = tmFpf.iEffEndDate(0)
        tmGrf.iDate(1) = tmFpf.iEffEndDate(1)
        For llFdfLoop = 0 To UBound(tlFdf) - 1          'loop on all detail records and build a crystal report record            tmGrf.iVefCode = tmFpf.iVefCode         'vehicle code
            tmGrf.iRdfCode = tlFdf(llFdfLoop).iCode           'FDF code
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                ilOk = False
                Exit For
            End If
        Next llFdfLoop
    Next llFPFLoop
    mBuildPledge = ilOk
    Exit Function
End Function
'
'           Create Dump of Delivery and Engineering Pre-feed links
'           Files retreived PFF for Delivery
'                           PFF & PBF for Engineering
'           5-6-10 Create prepass GRf file
'   grfGenDate - generation date
'   grfGenTime - generation time
'   grfCode2 - 1 = delivery, 2 = engineering
'   grfVefCode - vehicle code
'   grfCode4 - PFF code
'   grfChfCode - PBF code
'   grfPerGenl(3) - sequence number
Public Sub gCrPreFeed()
Dim ilSelectedVehicles() As Integer
Dim ilUpper As Integer
Dim ilLoop As Integer
Dim slNameCode As String
Dim slCode As String
Dim ilRet As Integer
Dim ilLoopOnVehicle As Integer
Dim ilLoopOnReport As Integer
Dim ilMinType As Integer
Dim ilMaxType As Integer
Dim ilVefCode As Integer
Dim slMsgErr As String
Dim slStartDate As String
Dim llDate As Long
Dim ilLoopOnPff As Integer
Dim tlPff() As PFF
Dim slErrMsg As String

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    
    hmPff = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPff, "", sgDBPath & "Pff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPff)
        btrDestroy hmPff
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imPffRecLen = Len(tmPFF)

    hmPbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPbf, "", sgDBPath & "Pbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPbf)
        btrDestroy hmPbf
        btrDestroy hmPff
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imPbfRecLen = Len(tmPbf)
    
    'get active date entered
    slStartDate = rptSelFD!edcFromDate.Text
    llDate = gDateValue(slStartDate)
    slStartDate = Format$(llDate, "m/d/yy")
   
    'build array of vehicles to include
    ReDim ilSelectedVehicles(0 To 0) As Integer
    ilUpper = 0
    For ilLoop = 0 To rptSelFD!lbcSelection(1).ListCount - 1 Step 1
        If rptSelFD!lbcSelection(1).Selected(ilLoop) Then              'selected slsp
            slNameCode = tgVehicle(ilLoop).sKey        'pick up slsp code
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilSelectedVehicles(ilUpper) = Val(slCode)
            ilUpper = ilUpper + 1
            ReDim Preserve ilSelectedVehicles(0 To ilUpper)
        End If
    Next ilLoop
    
    If rptSelFD!ckcReport(0).Value = vbChecked And rptSelFD!ckcReport(1).Value = vbChecked Then
        ilMinType = 1
        ilMaxType = 2
    ElseIf rptSelFD!ckcReport(0).Value = vbChecked Then
        ilMinType = 1
        ilMaxType = 1
    Else
        ilMinType = 2
        ilMaxType = 2
    End If
    
    For ilLoopOnReport = ilMinType To ilMaxType
        For ilLoopOnVehicle = LBound(ilSelectedVehicles) To UBound(ilSelectedVehicles) - 1
            ilVefCode = ilSelectedVehicles(ilLoopOnVehicle)
            ilRet = gObtainPFF(rptSelFD, hmPff, tlPff(), slStartDate, ilLoopOnReport, ilVefCode)
            If Not ilRet Then
                slErrMsg = "Error reading Pff in gCrPreFeed: gObtainPFF"
                GoTo gCrPreFeedErr
            End If
            For ilLoopOnPff = LBound(tlPff) To UBound(tlPff) - 1
                'test for day selectivity
                If (tlPff(ilLoopOnPff).sAirDay = "0" And rptSelFD!ckcDay(0).Value = vbChecked) Or (tlPff(ilLoopOnPff).sAirDay = "6" And rptSelFD!ckcDay(1).Value = vbChecked) Or (tlPff(ilLoopOnPff).sAirDay = "7" And rptSelFD!ckcDay(2).Value = vbChecked) Then
                    tmGrf.iGenDate(0) = igNowDate(0)
                    tmGrf.iGenDate(1) = igNowDate(1)
                    tmGrf.lGenTime = lgNowTime
                    'tmGrf.iPerGenl(3) = 1           'seq #
                    tmGrf.iPerGenl(2) = 1           'seq #
                    tmGrf.iVefCode = ilVefCode
                    tmGrf.iCode2 = ilLoopOnReport       '1 = delivery, 2 = engineering
                    tmGrf.lCode4 = tlPff(ilLoopOnPff).lCode     'Prefeed code
                    tmGrf.lChfCode = 0                  'init the bus pointer
                    If tlPff(ilLoopOnPff).sType = "D" Then      'if delivery, create record
                                                                'if engineering, create record with the bus info
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                        If ilRet <> BTRV_ERR_NONE Then
                            slErrMsg = "Error writing GRF: gCrPreFeedErr"
                            GoTo gCrPreFeedErr
                        End If
                    End If
                    'if Engineering, get the matching bus
                    If tlPff(ilLoopOnPff).sType = "E" Then
                        'tmGrf.iPerGenl(3) = 0     'init the seq # to 0 since each sucessive record for the same link has its seq # incremented
                        tmGrf.iPerGenl(2) = 0     'init the seq # to 0 since each sucessive record for the same link has its seq # incremented
                        tmPbfSrchKey1.lCode = tlPff(ilLoopOnPff).lCode
                        ilRet = btrGetGreaterOrEqual(hmPbf, tmPbf, imPbfRecLen, tmPbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                        Do While ilRet = BTRV_ERR_NONE And tlPff(ilLoopOnPff).lCode = tmPbf.lPffCode
                            'tmGrf.iPerGenl(3) = tmGrf.iPerGenl(3) + 1       'increment seq #
                            tmGrf.iPerGenl(2) = tmGrf.iPerGenl(2) + 1       'increment seq #
                            tmGrf.lChfCode = tmPbf.lCode        'bus information
                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                            ilRet = btrGetNext(hmPbf, tmPbf, imPbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        Loop
                    End If
                End If
            Next ilLoopOnPff
        Next ilLoopOnVehicle
    Next ilLoopOnReport
    Erase tlPff
    Erase ilSelectedVehicles
    Exit Sub
    
gCrPreFeedErr:
    MsgBox slErrMsg
    ilRet = btrClose(hmPff)
    ilRet = btrClose(hmPbf)
    ilRet = btrClose(hmGrf)
    btrDestroy hmPff
    btrDestroy hmPbf
    btrDestroy hmGrf
    Exit Sub
End Sub
