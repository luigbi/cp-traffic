Attribute VB_Name = "RptCrNoSel"
Option Explicit

Dim hmAvr As Integer
Dim imAvrRecLen As Integer        'AVR record length
Dim tmAvr As AVR

Dim hmIvr As Integer
Dim imIvrRecLen As Integer        'IVR record length
Dim tmIvr As IVR

Dim hmUrf As Integer
Dim imUrfRecLen As Integer        'Urf record length
Dim tmUrf As URF

Dim hmSaf As Integer
Dim imSafRecLen As Integer        'Saf record length
Dim tmSaf As SAF
Dim tmSafSrchKey1 As INTKEY0

Dim hmSite As Integer
Dim imSiteRecLen As Integer        'Site record length
Dim tmSite As SITE
Dim tmSiteSrchKey As LONGKEY0

Dim hmNrf As Integer
Dim imNrfRecLen As Integer        'Nrf record length
Dim tmNrf As NRF

Dim hmSdf As Integer
Dim imSdfRecLen As Integer        'SDF spot record length
Dim tmSdf As SDF

Dim imTerminate As Integer  'True = terminating task, False= OK
Private Sub mGetLastBkup()
    Dim slBasePath As String
    Dim slFileName As String
    Dim slLastDate As String
    Dim llBkupSize As Long
    Dim ilPos As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slBackupHour As String
    Dim slCSIServerINIFile As String
    Dim slBUWeekDays As String
    On Error GoTo mGenErr:

    'Build the base path to the Save directory
    ilPos = InStrRev(sgExePath, "\", Len(sgExePath) - 1)
    slBasePath = Left$(sgExePath, ilPos) & "SaveData\"

    slCSIServerINIFile = sgExePath & "\CSI_Server.ini"
    If Not gLoadINIValue(slCSIServerINIFile, "MainSettings", "LastBackupFileName", slFileName) Then
        tmIvr.sEDIComment = "     *** No Backup Exists ***"
        tmIvr.sRAmount = "0KB"
        Exit Sub
    End If

    If gLoadINIValue(slCSIServerINIFile, "Backup", "StartTime", slBackupHour) Then
        tmIvr.sATime = slBackupHour
    End If
    If gLoadINIValue(slCSIServerINIFile, "Backup", "WeekDays", slBUWeekDays) Then
        If Len(slBUWeekDays) = 7 Then
            tmIvr.sODays = slBUWeekDays
'            For ilLoop = 1 To 7
'                If Mid(smBUWeekDays, ilLoop, 1) = 1 Then
'                    chkDOW(ilLoop - 1).Value = 1
'                End If
'            Next
        End If
    End If
    
    'On Error GoTo mIgnoreErr
    ilRet = 0
    'tmIvr.sADayDate = FileDateTime(slBasePath & "\" & slFileName)
    'If ilRet = -1 Then
    ilRet = gFileExist(slBasePath & "\" & slFileName)
    If ilRet <> 0 Then
        tmIvr.sEDIComment = "     *** No Backup Exists ***"
        tmIvr.sRAmount = "0KB"
        Exit Sub
    End If
    tmIvr.sADayDate = gFileDateTime(slBasePath & "\" & slFileName)
    llBkupSize = Round(FileLen(slBasePath & "\" & slFileName) / 1024)
    tmIvr.sRAmount = CStr(llBkupSize) & " KB"
    tmIvr.sEDIComment = slFileName
    'tmIvr.sAddr(1) = Left$(sgExePath, ilPos) & "SaveData\"
    tmIvr.sAddr(0) = Left$(sgExePath, ilPos) & "SaveData\"
    Exit Sub

mIgnoreErr:
    ilRet = -1
    Resume Next

mGenErr:
    MsgBox "An error has occured in mGetLastBkup."
End Sub



'**************************************************************************************
'
'       gCreateSite
'               Created:  11-12-09
'       Create a prepass to show all Site options and features
'               Since there a bit flags representing options, plus various features are stored
'               in files other than SPF (MNF, SAF, URF, CXF, SITE, NRF), the AVR temporary file
'               will hold the bit flag (unpacked), plus pointers to the other supporting files
'
'               1 temporary record in IVR will also be created to store the Backup information
'               since it is only stored in the csi_server.ini file
'**************************************************************************************
Sub gCreateSite()

    Dim ilRet As Integer    'Return Status
    Dim ilValue As Integer  'bit flags
    Dim ilValue2 As Integer 'bit flags
    Dim llDate As Long

        'spf is already open and in tgSpf buffer
        hmAvr = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmAvr, "", sgDBPath & "AVR.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gCreateSiteErr
        gBtrvErrorMsg ilRet, "gCreateSite (btrOpen: Avr.Btr)", RptNoSel
        On Error GoTo 0
        imAvrRecLen = Len(tmAvr)
        
        hmIvr = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmIvr, "", sgDBPath & "Ivr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gCreateSiteErr
        gBtrvErrorMsg ilRet, "gCreateSite (btrOpen: IVR.Btr)", RptNoSel
        On Error GoTo 0
        imIvrRecLen = Len(tmIvr)
    
        hmSaf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmSaf, "", sgDBPath & "SAF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gCreateSiteErr
        gBtrvErrorMsg ilRet, "gCreateSite (btrOpen: Saf.Btr)", RptNoSel
        On Error GoTo 0
        imSafRecLen = Len(tmSaf)
        
        hmSite = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmSite, "", sgDBPath & "SITE.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gCreateSiteErr
        gBtrvErrorMsg ilRet, "gCreateSite (btrOpen: Site.mkd)", RptNoSel
        On Error GoTo 0
        imSiteRecLen = Len(tmSite)
            
        hmNrf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmNrf, "", sgDBPath & "NRF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gCreateSiteErr
        gBtrvErrorMsg ilRet, "gCreateSite (btrOpen: Nrf.Btr)", RptNoSel
        On Error GoTo 0
        imNrfRecLen = Len(tmNrf)
        
        hmSdf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmSdf, "", sgDBPath & "SDF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gCreateSiteErr
        gBtrvErrorMsg ilRet, "gCreateSite (btrOpen: SDf.Btr)", RptNoSel
        On Error GoTo 0
        imSdfRecLen = Len(tmSdf)

  
        tmSafSrchKey1.iCode = 0
        ilRet = btrGetEqual(hmSaf, tmSaf, imSafRecLen, tmSafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        
        tmSiteSrchKey.lCode = 1
        ilRet = btrGetEqual(hmSite, tmSite, imSiteRecLen, tmSiteSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)

        llDate = gGetEarliestTrafSpotDate(hmSdf, -1)
        tmAvr.iQStartDate(0) = 0
        tmAvr.iQStartDate(1) = 0
        If llDate <> -1 Then
            gPackDateLong llDate, tmAvr.iQStartDate(0), tmAvr.iQStartDate(1)
        End If


        'General scheme is to place most fields into AVR because of the bit manipulation which cannot be done in crystal (easily).
        'The fields are parsed and the put into an AVR field. Once in Crystal, it is placed in a SHARED variable, to be passed
        'to subreports from the main report.
        'There are not many fields remaining for storage in AVR:
        'search for reserved avr fields; they may be used when new bits (features) are defined in spf or saf.
        'Other unused fields (verify before using)
        
        'avrVefCode
        'avrDay
        'avrDPStartTime
        'avrDPEndTime
        'avrDays
        'avrNot30or60
        'avrrdfsortcode
        'avrRate4 - avrRate4
        'avrRate7 - avrRate14
        'avrMonth1 - avrMonth3
        
        
        'determine necessary pointers to place into AVR to link up to supporting tables
        'SPF (site options from traffic) has only 1 reocrd
        tmAvr.i30Prop(0) = 1            'spfcode
        tmAvr.i30Prop(1) = tgSpf.iMnfClientAbbr     'this will be for the MNF pointer for client name abbreviation
        'Network/rep links has only 1 record
        tmAvr.i30Prop(2) = 1            'nrfcode
        'Find SAF for system features; other records may exist which supports scheduling parameters for vehicles
        tmAvr.i30Prop(3) = tmSaf.iCode            'safcode
        'Site from Affiliate has only 1 record
        tmAvr.lRate(4) = tmSite.lCode            'sitecode
        
        If (((Asc(tgSpf.sAutoType2)) And RN_REP) <> RN_REP) And (((Asc(tgSpf.sAutoType2)) And RN_NET) <> RN_NET) Then
            tmAvr.lRate(5) = 0          'no NRFs exist
        Else
            ilRet = btrGetFirst(hmNrf, tmNrf, imNrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            Do While ilRet = BTRV_ERR_NONE
                If (((Asc(tgSpf.sAutoType2)) And RN_REP) = RN_REP) Then
                    If tmNrf.sType = "R" Then
                        tmAvr.lRate(5) = tmNrf.iCode
                        Exit Do
                    End If
                ElseIf (((Asc(tgSpf.sAutoType2)) And RN_NET) = RN_NET) Then
                    If tmNrf.sType = "N" Then
                        tmAvr.lRate(5) = tmNrf.iCode
                        Exit Do
                    End If
                End If
                ilRet = btrGetNext(hmNrf, tmNrf, imNrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If

        
        'tgSpf.sUsingFeatures
        ilValue = Asc(tgSpf.sUsingFeatures)  'Option Fields in Orders/Proposals
        If (ilValue And MATRIXEXPORT) = MATRIXEXPORT Then
            tmAvr.i30Count(0) = 1
        Else
            tmAvr.i30Count(0) = 0
        End If
        If (ilValue And REVENUEEXPORT) = REVENUEEXPORT Then
            tmAvr.i30Count(1) = 1
        Else
            tmAvr.i30Count(1) = 0
        End If
        If (ilValue And LIVECOPY) = LIVECOPY Then
            tmAvr.i30Count(2) = 1
        Else
            tmAvr.i30Count(2) = 0
        End If
        If (ilValue And MULTIMEDIA) = MULTIMEDIA Then
            tmAvr.i30Count(3) = 1
        Else
            tmAvr.i30Count(3) = 0
        End If
        If (ilValue And USINGLIVELOG) = USINGLIVELOG Then
            tmAvr.i30Count(4) = 1
        Else
            tmAvr.i30Count(4) = 0
        End If
        'Part in sOverrideOptions
        If (ilValue And BESTFITWEIGHT) = BESTFITWEIGHT Then
            tmAvr.i30Count(5) = 2   'For non-matching DP give weight to Times
        ElseIf (Asc(tgSpf.sOverrideOptions) And BESTFITWEIGHTNONE) <> BESTFITWEIGHTNONE Then
            tmAvr.i30Count(5) = 1    'For non-matching DP give weight to Days
        End If
        If (ilValue And HIDDENOVERRIDE) = HIDDENOVERRIDE Then
            tmAvr.i30Count(6) = 1         'for hidden line reserach, ignore overrides
        Else
            tmAvr.i30Count(6) = 0
        End If
        If (ilValue And USINGREP) = USINGREP Then
            tmAvr.i30Count(7) = 1
        Else
            tmAvr.i30Count(7) = 0
        End If

        'tgSpf.sOptionFields
        ilValue = Asc(tgSpf.sOptionFields)  'Option Fields in Orders/Proposals
        If (ilValue And &H1) = &H1 Then 'Projections
            tmAvr.i30Count(8) = 1
        Else
            tmAvr.i30Count(8) = 0
        End If
        If (ilValue And &H2) = &H2 Then 'Business Category
            tmAvr.i30Count(9) = 1
        Else
            tmAvr.i30Count(9) = 0
        End If
        If (ilValue And &H4) = &H4 Then 'Share
            tmAvr.i30Count(10) = 1
        Else
            tmAvr.i30Count(10) = 0
        End If
        If (ilValue And &H8) = &H8 Then 'Revenue Set
            tmAvr.i30Count(11) = 1
        Else
            tmAvr.i30Count(11) = 0
        End If
        If (ilValue And &H10) = &H10 Then   'Delivery Guarantee %
            tmAvr.i30Count(12) = 1
        Else
            tmAvr.i30Count(12) = 0
        End If
        If (ilValue And &H20) = &H20 Then   'Calendar Cycle
            tmAvr.i30Count(13) = 1
        Else
            tmAvr.i30Count(13) = 0
        End If
        If (ilValue And &H40) = &H40 Then   'Co-op Billing
            tmAvr.i60Count(0) = 1
        Else
            tmAvr.i60Count(0) = 0
        End If
        If (ilValue And &H80) = &H80 Then   'Research
            tmAvr.i60Count(1) = 1
        Else
            tmAvr.i60Count(1) = 0
        End If
 
        'tgSpf.sOverrideOptions
        ilValue = Asc(tgSpf.sOverrideOptions)  'Option Fields in Orders/Proposals
        If (ilValue And SPALLOCATION) = SPALLOCATION Then 'Allocation %
            tmAvr.i60Count(2) = 1
        Else
            tmAvr.i60Count(2) = 0
        End If
        '6/7/15: replaced acquisition from site override with Barter in system options
        'If (ilValue And SPACQUISITION) = SPACQUISITION Then 'Acquisition Cost
        If ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
            tmAvr.i60Count(3) = 1
        Else
            tmAvr.i60Count(3) = 0
        End If
        If (ilValue And SP1STPOSITION) = SP1STPOSITION Then '1st Position
            tmAvr.i60Count(4) = 1
        Else
            tmAvr.i60Count(4) = 0
        End If
        If (ilValue And SPPREFERREDDT) = SPPREFERREDDT Then 'Preferred Days/Times
            tmAvr.i60Count(5) = 1
        Else
            tmAvr.i60Count(5) = 0
        End If
        If (ilValue And SPSOLOAVAIL) = SPSOLOAVAIL Then   'Solo Avails
            tmAvr.i60Count(6) = 1
        Else
            tmAvr.i60Count(6) = 0
        End If
        
        'Also part in sUsingFeatures
        If (ilValue And BESTFITWEIGHTNONE) = BESTFITWEIGHTNONE Then
            tmAvr.i30Count(5) = 0    'None
        End If
        If (ilValue And BBSAMELINE) = BBSAMELINE Then   'BBs on Same Line
            tmAvr.i60Count(8) = 1
        Else
            tmAvr.i60Count(8) = 0
        End If
 
        If (ilValue And BYPASSHIPRI) = BYPASSHIPRI Then   'Bypass Hi Priority - unused at this time
            tmAvr.i60Count(9) = 1
        Else
            tmAvr.i60Count(9) = 0
        End If
        
        
        'tgSpf.sUsingFeatures2
        ilValue = Asc(tgSpf.sUsingFeatures2)
        If (ilValue And REGIONALCOPY) = REGIONALCOPY Then
            tmAvr.i60Count(10) = 1
        Else
            tmAvr.i60Count(10) = 0
        End If
        If (ilValue And SPLITCOPY) = SPLITCOPY Then
            tmAvr.i60Count(11) = 1
        Else
            tmAvr.i60Count(11) = 0
        End If
        If (ilValue And SPLITNETWORKS) = SPLITNETWORKS Then
            tmAvr.i60Count(12) = 1
        Else
            tmAvr.i60Count(12) = 0
        End If
        If (ilValue And BARTER) = BARTER Then
            tmAvr.i60Count(13) = 1
        Else
            tmAvr.i60Count(13) = 0
        End If
        If (ilValue And STRONGPASSWORD) = STRONGPASSWORD Then
            tmAvr.i30InvCount(0) = 1
        Else
            tmAvr.i30InvCount(0) = 0
        End If
        If (ilValue And MERCHPROMOBYDOLLAR) = MERCHPROMOBYDOLLAR Then
            tmAvr.i30InvCount(1) = 1
        Else
            tmAvr.i30InvCount(1) = 0
        End If
        'If (ilValue And MIXAIRTIMEANDREP) = MIXAIRTIMEANDREP Then
        '    ckcMixAirTimeAndRep.Value = vbChecked
        'Else
            'ckcMixAirTimeAndRep.Value = vbUnchecked
            tmAvr.i30InvCount(2) = 0
        'End If
        If (ilValue And GREATPLAINS) = GREATPLAINS Then
            tmAvr.i30InvCount(3) = 1
        Else
            tmAvr.i30InvCount(3) = 0
        End If
        
        
        'tgSpf.sUsingFeatures3
        ilValue = Asc(tgSpf.sUsingFeatures3)
        If (ilValue And USINGHUB) = USINGHUB Then
            tmAvr.i30InvCount(4) = 1
        Else
            tmAvr.i30InvCount(4) = 0
        End If
        If (ilValue And TAXONAIRTIME) = TAXONAIRTIME Then
            tmAvr.i30InvCount(5) = 1
        Else
            tmAvr.i30InvCount(5) = 0
        End If
        If (ilValue And TAXONNTR) = TAXONNTR Then
            tmAvr.i30InvCount(6) = 1
        Else
            tmAvr.i30InvCount(6) = 0
        End If
        If (ilValue And PROMOCOPY) = PROMOCOPY Then
            tmAvr.i30InvCount(7) = 1
        Else
            tmAvr.i30InvCount(7) = 0
        End If
        If (ilValue And MEDIACODEBYVEH) = MEDIACODEBYVEH Then
            tmAvr.i30InvCount(8) = 1
        Else
            tmAvr.i30InvCount(8) = 0
        End If
        If (ilValue And INCMEDIACODEAUDIOVAULT) = INCMEDIACODEAUDIOVAULT Then
            tmAvr.i30InvCount(9) = 1
        Else
            tmAvr.i30InvCount(9) = 0
        End If
        If tgSpf.sSchdPSA = "N" Then            'manually sch psa
            tmAvr.i30InvCount(10) = 0
        Else                    'auto sch psa, book into contract avails?
            tmAvr.i30InvCount(10) = 1
            If (ilValue And PSAINTOCONTRACTAVAILS) = PSAINTOCONTRACTAVAILS Then
                tmAvr.i30InvCount(11) = 1
            Else
                tmAvr.i30InvCount(11) = 0
            End If
        End If
        If tgSpf.sSchdPromo = "N" Then            'manually sch Promo
            tmAvr.i30InvCount(12) = 0
        Else
            tmAvr.i30InvCount(12) = 1
            If (ilValue And PROMOINTOCONTRACTAVAILS) = PROMOINTOCONTRACTAVAILS Then
                tmAvr.i30InvCount(13) = 1
            Else
                tmAvr.i30InvCount(13) = 0
            End If
        End If
    
        'tgSpf.sUsingFeatures4
        ilValue = Asc(tgSpf.sUsingFeatures4)
        If (ilValue And LOCKBOXBYVEHICLE) = LOCKBOXBYVEHICLE Then
            tmAvr.i60InvCount(0) = 1
        Else
            tmAvr.i60InvCount(0) = 0
        End If
        'tmAvr.i60InvCount(2): BYPASSDEMOGRAPHIC = &H2 (unused), Reservee field
        'tmAvr.i60InvCount(3): BYPASSDATENEEDED = &H4 (unused), Reservee field
        If (ilValue And ALLOWMOVEONTODAY) = ALLOWMOVEONTODAY Then
            tmAvr.i60InvCount(3) = 1
        Else
            tmAvr.i60InvCount(3) = 0
        End If
        If (ilValue And CHGBILLEDPRICE) = CHGBILLEDPRICE Then
            tmAvr.i60InvCount(4) = 1
        Else
            tmAvr.i60InvCount(4) = 0
        End If
        tmAvr.i60InvCount(5) = 0                  'no taxes
        If (ilValue And TAXBYUSA) = TAXBYUSA Then
            tmAvr.i60InvCount(5) = 1
        ElseIf (ilValue And TAXBYCANADA) = TAXBYCANADA Then
            tmAvr.i60InvCount(5) = 2
        End If
        If (ilValue And INVSORTBYVEHICLE) = INVSORTBYVEHICLE Then
            tmAvr.i60InvCount(6) = 1
        Else
            tmAvr.i60InvCount(6) = 0
        End If
        
    
        'tgSpf.sUsingFeatures5
        ilValue = Asc(tgSpf.sUsingFeatures5)  'Option Fields in Orders/Proposals
        If (ilValue And REMOTEEXPORT) = REMOTEEXPORT Then
            tmAvr.i60InvCount(7) = 1
        Else
            tmAvr.i60InvCount(7) = 0
        End If
        If (ilValue And REMOTEIMPORT) = REMOTEIMPORT Then
            tmAvr.i60InvCount(8) = 1
        Else
            tmAvr.i60InvCount(8) = 0
        End If
        If (ilValue And COMBINEAIRNTR) = COMBINEAIRNTR Then
            tmAvr.i60InvCount(9) = 1
        Else
            tmAvr.i60InvCount(9) = 0
        End If
        If (ilValue And CNTRINVSORTRC) = CNTRINVSORTRC Then
            tmAvr.i60InvCount(10) = 0           'sort orders by rate card
        ElseIf (ilValue And CNTRINVSORTLN) = CNTRINVSORTLN Then
            tmAvr.i60InvCount(10) = 1           'sort orders by line
        Else
            tmAvr.i60InvCount(10) = 2           'sort orders by DP
        End If
        If (ilValue And STATIONINTERFACE) = STATIONINTERFACE Then
            tmAvr.i60InvCount(11) = 1
        Else
            tmAvr.i60InvCount(11) = 0
        End If
        If (ilValue And RADAR) = RADAR Then
            tmAvr.i60InvCount(12) = 1
        Else
            tmAvr.i60InvCount(12) = 0
        End If
        If (ilValue And SUPPRESSTIMEFORM1) = SUPPRESSTIMEFORM1 Then
            tmAvr.i60InvCount(13) = 1
        Else
            tmAvr.i60InvCount(13) = 0
        End If
    
        'tgSpf.sUsingFeatures6
        ilValue = Asc(tgSpf.sUsingFeatures6)
        If (ilValue And BBNOTSEPARATELINE) = BBNOTSEPARATELINE Then    'this option isnt shown on site screen
            tmAvr.i30Hold(0) = 1
        Else
            tmAvr.i30Hold(0) = 0
        End If
        If (ilValue And BBCLOSEST) = BBCLOSEST Then
            tmAvr.i30Hold(1) = 1
        Else
            tmAvr.i30Hold(1) = 0
        End If
        If (ilValue And INSTALLMENT) = INSTALLMENT Then
            tmAvr.i30Hold(2) = 1
            If (ilValue And INSTALLMENTREVENUEEARNED) = INSTALLMENTREVENUEEARNED Then
                tmAvr.i30Hold(3) = 1
            Else
                tmAvr.i30Hold(3) = 0
            End If
        Else
            tmAvr.i30Hold(2) = 0
        End If
        If (ilValue And GETPAIDEXPORT) = GETPAIDEXPORT Then
            tmAvr.i30Hold(4) = 1
        Else
            tmAvr.i30Hold(4) = 0
        End If
        If (ilValue And DIGITALCONTENT) = DIGITALCONTENT Then
            tmAvr.i30Hold(5) = 1
        Else
            tmAvr.i30Hold(5) = 0
        End If
        If (ilValue And GUARBYGRIMP) = GUARBYGRIMP Then
            tmAvr.i30Hold(6) = 1
        Else
            tmAvr.i30Hold(6) = 0
        End If
        '9-24-19 tmavr30hold(7) is not used in .rpt; remove
'        If (ilValue And INVEXPORTPARAMETERS) = INVEXPORTPARAMETERS Then
'            tmAvr.i30Hold(7) = 1
'        Else
'            tmAvr.i30Hold(7) = 0
'        End If
        
        'tgSpf.sUsingFeatures7
        ilValue = Asc(tgSpf.sUsingFeatures7)
        If (ilValue And CSIBACKUP) = CSIBACKUP Then
            tmAvr.i30Hold(8) = 1
        Else
            tmAvr.i30Hold(8) = 0
        End If
        If (ilValue And BONUSCOMM) = BONUSCOMM Then
            tmAvr.i30Hold(9) = 1
        Else
            tmAvr.i30Hold(9) = 0
        End If
        tmAvr.i30Hold(10) = 0       'assume std
        If (ilValue And COMMFISCALYEAR) = COMMFISCALYEAR Then
            tmAvr.i30Hold(10) = 1
        End If
        tmAvr.i30Hold(11) = 0       'assume comm by a/e
        If tgSpf.sSubCompany = "Y" Then 'using sub?
            tmAvr.i30Hold(11) = 1
        End If
        If tgSpf.sCommByCntr = "Y" Then 'comm by contract
            tmAvr.i30Hold(11) = 2
        End If
            
        If (ilValue And EXPORTREVENUE) = EXPORTREVENUE Then
            tmAvr.i30Hold(12) = 1
        Else
            tmAvr.i30Hold(12) = 0
        End If
        If (ilValue And XDIGITALISCIEXPORT) = XDIGITALISCIEXPORT Then
            tmAvr.i30Hold(13) = 1
        Else
            tmAvr.i30Hold(13) = 0
        End If
        If (ilValue And WEGENEREXPORT) = WEGENEREXPORT Then
            tmAvr.i60Hold(0) = 1
        Else
            tmAvr.i60Hold(0) = 0
        End If
        If (ilValue And OLAEXPORT) = OLAEXPORT Then
            tmAvr.i60Hold(1) = 1
        Else
            tmAvr.i60Hold(1) = 0
        End If
        'Allow Mixed Split Network Spot Lengths to combine
        If tmAvr.i60Count(12) = 1 Then      'using split networks?
            If (ilValue And REGIONMIXLEN) = REGIONMIXLEN Then
                tmAvr.i60Hold(2) = 1
            Else
                tmAvr.i60Hold(2) = 0
            End If
        Else
            tmAvr.i60Hold(2) = 0
        End If
    
    
        ilValue = Asc(tgSpf.sUsingFeatures8)
        If (ilValue And LRMANDATORY) = LRMANDATORY Then     'live/recorded mandatory
            tmAvr.i60Hold(3) = 1
        Else
            tmAvr.i60Hold(3) = 0
        End If
        If (ilValue And SHOWCMMTONDETAILPAGE) = SHOWCMMTONDETAILPAGE Then
            tmAvr.i60Hold(4) = 1
        Else
            tmAvr.i60Hold(4) = 0
        End If
        If (ilValue And ALLOWMSASPLITCOPY) = ALLOWMSASPLITCOPY Then       '12-17-09 using Metro Split Copy
            tmAvr.i60Hold(5) = 1
        Else
            tmAvr.i60Hold(5) = 0
        End If
        If (ilValue And RIVENDELLEXPORT) = RIVENDELLEXPORT Then       '12-28-11 using rivendell automation
            tmAvr.i60Hold(6) = 1
        Else
            tmAvr.i60Hold(6) = 0
        End If
        If (ilValue And XDIGITALBREAKEXPORT) = XDIGITALBREAKEXPORT Then       '12-28-11 Using xdigital by break
            tmAvr.i60Hold(7) = 1
        Else
            tmAvr.i60Hold(7) = 0
        End If
        If (ilValue And ISCIEXPORT) = ISCIEXPORT Then       '12-28-11  isci export (not xdigital)
            tmAvr.i60Hold(8) = 1
        Else
            tmAvr.i60Hold(8) = 0
        End If
        If (ilValue And PREFEEDDEF) = PREFEEDDEF Then       '12-28-11 pre-feed defined
            tmAvr.i60Hold(9) = 1
        Else
            tmAvr.i60Hold(9) = 0
        End If
        
        If (ilValue And REPBYDT) = REPBYDT Then       '12-28-11 rep by date
            tmAvr.i60Hold(10) = 1
        Else
            tmAvr.i60Hold(10) = 0
        End If
        
        
        'tmAvr.i60Avail(1-8) reserved for usingfeatures9 byte
        ilValue = Asc(tgSpf.sUsingFeatures9)
        If (ilValue And AFFILIATECRM) = AFFILIATECRM Then   '12-28-11 Using Affiliate CRM
            tmAvr.i60Avail(0) = 1
        Else
            tmAvr.i60Avail(0) = 0
        End If
       
        If (ilValue And PC1STPOS) = PC1STPOS Then       '12-28-11 default pre-recorded copy 1st position
        Else
        End If
       
        If (ilValue And PROPOSALXML) = PROPOSALXML Then '12-28-11 using proposal XML
            tmAvr.i60Avail(2) = 1
        Else
            tmAvr.i60Avail(2) = 0
        End If
        If (ilValue And LIMITISCI) = LIMITISCI Then        '12-28-11 limit isci to 15 char
            tmAvr.i60Avail(3) = 1                           'limit to 15
        Else
            tmAvr.i60Avail(3) = 0                           'isci can be 20 char
        End If
        
        If (ilValue And WEEKLYBILL) = WEEKLYBILL Then        'weekly billing cycle
            tmAvr.i60Avail(4) = 1
        Else
            tmAvr.i60Avail(4) = 0
        End If
        
        'tmAvr.i60Avail(6) = unused, reserved for bit 5
        'tmavr.i60avail(7) = unused, reserved for bit 6
        'tmavr.i60avail(8) = unused, reserved for bit 7
        
        'tgSpf.sUsingFeatures10
                
         ilValue = Asc(tgSpf.sUsingFeatures10)
         
         If (ilValue And PKGLNRATEONBR) = PKGLNRATEONBR Then
            tmAvr.i60Prop(8) = 1
        Else
            tmAvr.i60Prop(8) = 0
        End If
        
         If (ilValue And ADDADVTTOISCI) = ADDADVTTOISCI Then
            tmAvr.i60Prop(9) = 1
        Else
            tmAvr.i60Prop(9) = 0
        End If
         If (ilValue And MIDNIGHTBASEDHOUR) = MIDNIGHTBASEDHOUR Then
            tmAvr.i60Prop(10) = 1
        Else
            tmAvr.i60Prop(10) = 0
        End If
        '9114
        If (ilValue And UNITIDBYASTCODEFORBREAK) = UNITIDBYASTCODEFORBREAK Then
            tmAvr.i60Prop(11) = 1
        Else
            tmAvr.i60Prop(11) = 0
        End If

         If (ilValue And VCREATIVEEXPORT) = VCREATIVEEXPORT Then
            tmAvr.i60Prop(12) = 1
        Else
            tmAvr.i60Prop(12) = 0
        End If
        
        If (ilValue And WegenerIPump) = WegenerIPump Then
            tmAvr.i60Prop(13) = 1
        Else
            tmAvr.i60Prop(13) = 0
        End If
        
        'Automation bit flags
        'tgSpf.sAutoType & tgSpf.sAutoType2
        ilValue = Asc(tgSpf.sAutoType)  'Automation Equipment
        ilValue2 = Asc(tgSpf.sAutoType2)    'continuation of automation equipment types
        If (ilValue And DALET) = DALET Then 'Dalet)
            tmAvr.i60Hold(11) = 1
        Else
            tmAvr.i60Hold(11) = 0
        End If
        If (ilValue And PROPHETNEXGEN) = PROPHETNEXGEN Then 'Prophet
            tmAvr.i60Hold(12) = 1
        Else
            tmAvr.i60Hold(12) = 0
        End If
        If (ilValue And SCOTT) = SCOTT Then 'Scott
            tmAvr.i60Hold(13) = 1
        Else
            tmAvr.i60Hold(13) = 0
        End If
        If (ilValue And DRAKE) = DRAKE Then 'Drake
            tmAvr.i30Reserve(0) = 1
        Else
            tmAvr.i30Reserve(0) = 0
        End If
        If (ilValue And RCS4DIGITCART) = RCS4DIGITCART Then   'RCS
            tmAvr.i30Reserve(1) = 1
        Else
            tmAvr.i30Reserve(1) = 0
        End If
        If (ilValue And PROPHETWIZARD) = PROPHETWIZARD Then   'Prophet Wizard
            tmAvr.i30Reserve(2) = 1
        Else
            tmAvr.i30Reserve(2) = 0
        End If
        If (ilValue And PROPHETMEDIASTAR) = PROPHETMEDIASTAR Then   'Prophet MediaStar
            tmAvr.i30Reserve(3) = 1
        Else
            tmAvr.i30Reserve(3) = 0
        End If
        If (ilValue And IMEDIATOUCH) = IMEDIATOUCH Then   'iMediaTouch
            tmAvr.i30Reserve(4) = 1
        Else
            tmAvr.i30Reserve(4) = 0
        End If
        
        
        'tgSpf.sAutoType2
        If (ilValue2 And AUDIOVAULT) = AUDIOVAULT Then   '8-10-05 Audio Vault Sat
            tmAvr.i30Reserve(5) = 1
        Else
            tmAvr.i30Reserve(5) = 0
        End If
        If (ilValue2 And WIREREADY) = WIREREADY Then   '6/6/06 Wire Ready
            tmAvr.i30Reserve(6) = 1
        Else
            tmAvr.i30Reserve(6) = 0
        End If
        If (ilValue2 And ENCO) = ENCO Then   '9-12-06
            tmAvr.i30Reserve(7) = 1
        Else
            tmAvr.i30Reserve(7) = 0
        End If
        If (ilValue2 And RN_REP) = RN_REP Then
            tmAvr.i30Reserve(8) = 1
        Else
            tmAvr.i30Reserve(8) = 0
        End If
        If (ilValue2 And RN_NET) = RN_NET Then
            tmAvr.i30Reserve(9) = 1
        Else
            tmAvr.i30Reserve(9) = 0
        End If
        
        If (ilValue2 And SIMIAN) = SIMIAN Then         '8-22-08
            tmAvr.i30Reserve(10) = 1
        Else
            tmAvr.i30Reserve(10) = 0
        End If
        If (ilValue2 And RCS5DIGITCART) = RCS5DIGITCART Then   'RCS
            tmAvr.i30Reserve(11) = 1
        Else
            tmAvr.i30Reserve(11) = 0
        End If
       
        'tmAvr.i30Reserve(13) = 0
        If (ilValue2 And AUDIOVAULTRPS) = AUDIOVAULTRPS Then   '12-28-11 audio vault rps
            tmAvr.i30Reserve(12) = 1
        Else
            tmAvr.i30Reserve(12) = 0
        End If
        
        'automation types field #3
        'reserve tmavr.i30avail(4) - tmAvr.i30avail(11)
        ilValue = Asc(tgSpf.sAutoType3)
        If (ilValue And AUDIOVAULTAIR) = AUDIOVAULTAIR Then   '12-28-11 audio vault Air
            tmAvr.i30Avail(3) = 1
        Else
            tmAvr.i30Avail(3) = 0
        End If
        
        If (ilValue And SCOTT_V5) = SCOTT_V5 Then 'Scott
            tmAvr.i30Avail(4) = 1
        Else
            tmAvr.i30Avail(4) = 0
        End If
        
        If (ilValue And WIDEORBIT) = WIDEORBIT Then
            tmAvr.i30Avail(5) = 1
        Else
            tmAvr.i30Avail(5) = 0
        End If
        
        If (ilValue And ENCOESPN) = ENCOESPN Then
            tmAvr.i60Count(7) = 1
        Else
            tmAvr.i60Count(7) = 0
        End If
 


        
        'Sports
        ilValue = Asc(tgSpf.sSportInfo)  'Option Fields in Orders/Proposals
        If (ilValue And USINGSPORTS) = USINGSPORTS Then 'Using Sports
            tmAvr.i30Reserve(13) = 1
            If (ilValue And PREEMPTREGPROG) = PREEMPTREGPROG Then
                tmAvr.i60Reserve(0) = 1
            Else
                tmAvr.i60Reserve(0) = 0
            End If
            If (ilValue And USINGFEED) = USINGFEED Then
                tmAvr.i60Reserve(1) = 1
            Else
                tmAvr.i60Reserve(1) = 0
            End If
            If (ilValue And USINGLANG) = USINGLANG Then
                tmAvr.i60Reserve(2) = 1
            Else
                tmAvr.i60Reserve(2) = 0
            End If
        Else
            tmAvr.i30Reserve(13) = 0
        End If
        
        'Sports bits 4-7 unused
        'Reserve output fields
        tmAvr.i60Reserve(3) = 0
        tmAvr.i60Reserve(4) = 0
        tmAvr.i60Reserve(5) = 0
        tmAvr.i60Reserve(6) = 0
        
        'COPY/RESEARCH
        ilValue = Asc(tgSpf.sMOFCopyAssign)  'Mg Copy Assignment
        If (ilValue And MGORIGVEHONLY) = MGORIGVEHONLY Then     'mg & outside cpy rotation assignment:  Original vehicle only
            tmAvr.i60Reserve(7) = 0     '
        ElseIf (ilValue And MGSCHVEHONLY) = MGSCHVEHONLY Then   'scheduled vehicle only
            tmAvr.i60Reserve(7) = 1
        Else
            tmAvr.i60Reserve(7) = 2                     'scheduled veh or original vehicle
        End If
        If (ilValue And FILLORIGVEHONLY) = FILLORIGVEHONLY Then     'Fill copy rotation assignment by:
            tmAvr.i60Reserve(9) = 0                                'original veh only
        ElseIf (ilValue And FILLSCHVEHONLY) = FILLSCHVEHONLY Then   'schedule vehicle
            tmAvr.i60Reserve(9) = 1
        Else
            tmAvr.i60Reserve(9) = 2                       'scheduled vehicle or original vehicle
        End If
        If (ilValue And MGRULESINCOPY) = MGRULESINCOPY Then
            tmAvr.i60Reserve(8) = 1                 'Ask above rule in copy
        Else
            tmAvr.i60Reserve(8) = 0              'use above rules
        End If
        If (ilValue And RSCHCUSTDEMO) = RSCHCUSTDEMO Then
            tmAvr.i60Reserve(10) = 1
        Else
            tmAvr.i60Reserve(10) = 0
        End If
        'Reserved fields for open bit on MOFCOPYASSIGN string
        'tmAvr.i60Reserve(12) = 0
        'tmAvr.i60Reserve(13) = 0
        'tmAvr.i60Reerve(14) = 0
        'tmAvr.i30Avail(1) = 0
        
        If (Asc(tmSaf.sFeatures1) And SHOWPRICEONINSERTIONWITHACQUISTION) = SHOWPRICEONINSERTIONWITHACQUISTION Then 'Show Spot Prices on Insertion Orders if Acquistion Exist
            tmAvr.i30Prop(4) = 1
        Else
            tmAvr.i30Prop(4) = 0
        End If
        
        If (Asc(tgSpf.sUsingFeatures9) And WORDWRAPVEHICLE) = WORDWRAPVEHICLE Then
            tmAvr.i30Prop(5) = 1
        Else
            tmAvr.i30Prop(5) = 0
        End If

        If (Asc(tgSpf.sUsingFeatures10) And CONTRACTVERIFY) = CONTRACTVERIFY Then       'contract verification
            tmAvr.imnfMajorCode = 1
         Else
             tmAvr.imnfMajorCode = 0
        End If
        
        If (Asc(tgSpf.sUsingFeatures10) And REPLACEDELWKWITHFILLS) = REPLACEDELWKWITHFILLS Then
             tmAvr.imnfMinorCode = 1
         Else
             tmAvr.imnfMinorCode = 0
         End If
         
         If (Asc(tgSpf.sUsingFeatures9) And PRINTEDI) = PRINTEDI Then
             tmAvr.ianfCode = 1
         Else
             tmAvr.ianfCode = 0
         End If
         
        If (Asc(tgSpf.sOverrideOptions) And SPNTRACQUISITION) = SPNTRACQUISITION Then 'NTR Acquisition Cose
             tmAvr.iRdfCode = 1
         Else
             tmAvr.iRdfCode = 0
         End If
         
        'SAF fields
        If tmSaf.sInvISCIForm = "R" Then     'for isci on inv, truncate right, truncate left, wrap around
            tmAvr.i30Avail(1) = 0               'truncate right
        ElseIf tmSaf.sInvISCIForm = "L" Then
            tmAvr.i30Avail(1) = 1
        Else
            tmAvr.i30Avail(1) = 2
        End If
        
        If tmSaf.sReSchdXCal = "Y" Then     '12-28-11 Question on screen is "Prohibit resch across calendar months", but stored as Reschedule cross calendar Months (y/n)
            tmAvr.i30Avail(2) = 1
        Else
            tmAvr.i30Avail(2) = 0
        End If
        
        tmAvr.iWksInQtr = tmSaf.iNoDaysRetainUAF        '# days to retain user activity
        
        tmAvr.sBucketType = tmSaf.sCreditLimitMsg       'cutoff proposal/orders when credit limti reached
        
        'avr.i30Prop(7) - avr.i30Prop(14)
         ilValue = Asc(tmSaf.sFeatures1)
         If (ilValue And MATRIXCAL) = MATRIXCAL Then 'Matrix-Cal
             tmAvr.i30Avail(6) = 1
         Else
             tmAvr.i30Avail(6) = 0
         End If
         If (ilValue And ENGRHIDEMEDIACODE) = ENGRHIDEMEDIACODE Then 'Engineering Export: Hide Media Code
             tmAvr.i30Avail(7) = 1
         Else
             tmAvr.i30Avail(7) = 0
         End If
         If (ilValue And SHOWAUDIOTYPEONBR) = SHOWAUDIOTYPEONBR Then 'Show audio type on proposal/order
             tmAvr.i30Avail(8) = 1
         Else
             tmAvr.i30Avail(8) = 0
         End If
         If (ilValue And SHOWPRICEONINSERTIONWITHACQUISTION) = SHOWPRICEONINSERTIONWITHACQUISTION Then 'Show Spot Prices on Insertion Orders if Acquistion Exist
             tmAvr.i30Avail(9) = 1
         Else
             tmAvr.i30Avail(9) = 0
         End If
         If (ilValue And SALESFORCEEXPORT) = SALESFORCEEXPORT Then 'Sales Force
             tmAvr.i30Avail(10) = 1
         Else
             tmAvr.i30Avail(10) = 0
         End If
         If (ilValue And EFFICIOEXPORT) = EFFICIOEXPORT Then 'Efficio export
             tmAvr.i30Avail(11) = 1
         Else
             tmAvr.i30Avail(11) = 0
         End If
         If (ilValue And JELLIEXPORT) = JELLIEXPORT Then 'Jelli export
             tmAvr.i30Avail(12) = 1
         Else
             tmAvr.i30Avail(12) = 0
         End If
         If (ilValue And COMPENSATION) = COMPENSATION Then 'COMPENSATION
             tmAvr.i30Avail(13) = 1
         Else
             tmAvr.i30Avail(13) = 0
         End If
         
         
         ilValue = Asc(tmSaf.sFeatures2)
         If (ilValue And EVENTREVENUE) = EVENTREVENUE Then 'Event Revenue
             tmAvr.i60Prop(0) = 1
         Else
             tmAvr.i60Prop(0) = 0
         End If
         If (ilValue And HIDEHIDDENLINES) = HIDEHIDDENLINES Then 'Event Revenue
             tmAvr.i60Prop(1) = 1
         Else
             tmAvr.i60Prop(1) = 0
         End If
         'bit 2 tmsaf.sFeatures2 is unused
         'tmAvr.i60Prop(3) reserved for tmsaf.sfeatures2 (bit2)
         
         If (ilValue And EMAILDISTRIBUTION) = EMAILDISTRIBUTION Then 'E-Mail distribution system
             tmAvr.i60Prop(3) = 1
         Else
             tmAvr.i60Prop(3) = 0
         End If
         If (ilValue And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then 'Acquisition Commissionable
             tmAvr.i60Prop(4) = 1
         Else
             tmAvr.i60Prop(4) = 0
         End If
         If (ilValue And PAYMENTONCOLLECTION) = PAYMENTONCOLLECTION Then 'Payment on Collection
             tmAvr.i60Prop(5) = 1
         Else
             tmAvr.i60Prop(5) = 0
         End If
         If (ilValue And TABLEAUEXPORT) = TABLEAUEXPORT Then 'Tableau
             tmAvr.i60Prop(6) = 1
         Else
             tmAvr.i60Prop(6) = 0
         End If
         If (ilValue And TABLEAUCAL) = TABLEAUCAL Then 'Tableau
             tmAvr.i60Prop(7) = 1
         Else
             tmAvr.i60Prop(7) = 0
         End If
         
        tmAvr.lGenTime = lgNowTime
        tmAvr.iGenDate(0) = igNowDate(0)
        tmAvr.iGenDate(1) = igNowDate(1)
        ilRet = btrInsert(hmAvr, tmAvr, imAvrRecLen, INDEXKEY0)
        
'        tmIvr.iInvStartDate(0) = tmSaf.iVCreativeDate(0)        'last retrieval date for vcreative
'        tmIvr.iInvStartDate(1) = tmSaf.iVCreativeDate(1)
'        tmIvr.iType = 2             'copy info stored in this record type
'        tmIvr.lGenTime = lgNowTime
'        tmIvr.iGenDate(0) = igNowDate(0)
'        ilRet = btrInsert(hmIvr, tmIvr, imIvrRecLen, INDEXKEY0)
'        tmIvr.iGenDate(1) = igNowDate(1)
    
        mGetLastBkup
        tmIvr.iLineNo = tmAvr.i30Hold(8)
        tmIvr.lGenTime = lgNowTime
        tmIvr.iGenDate(0) = igNowDate(0)
        tmIvr.iGenDate(1) = igNowDate(1)
        tmIvr.iType = 1                         'Backup type information stored in IVR in this record type
        ilRet = btrInsert(hmIvr, tmIvr, imIvrRecLen, INDEXKEY0)

        ilRet = btrClose(hmAvr)
        ilRet = btrClose(hmIvr)
        ilRet = btrClose(hmSaf)
        ilRet = btrClose(hmNrf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSite)
        btrDestroy hmAvr
        btrDestroy hmIvr
        btrDestroy hmSaf
        btrDestroy hmNrf
        btrDestroy hmSdf
        btrDestroy hmSite

    Exit Sub
gCreateSiteErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub


