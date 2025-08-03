Attribute VB_Name = "SAFRec"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SAFREC.BAS on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmSafSrchKey                                                                          *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Schedule.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the schedule function
'
'
Option Explicit
Option Compare Text
Type SAF
    iCode                 As Integer
    sSchRNG               As String * 1      ' Scheduling: Use T=Clock As Seed;
                                             ' F=Fixed Value As Seed
    sSchMaL               As String * 1      ' Major Loop on: H=Hour; D=Day;
                                             ' Q=Quarter Hour
    sSchMdL                  As String * 1      ' Middle Loop on: H=Hour; D=Day;
                                             ' Q=Quarter Hour
    sSchMiL                  As String * 1      ' Minor Loop on: H=Hour; D=Day;
                                             ' Q=Quarter Hour
    sSchCycle             As String * 1      ' Schedule cycle pattern: 1=Iterate
                                             ' on equal counts (major; middle;
                                             ' minor); 2=Iterate on minor in
                                             ' count order
    sSchMove              As String * 1      ' Schedule Move Pass: N=N/A; A=each
                                             ' iteration; B=Each Count; C=Minor
                                             ' Loop; D=Middle Loop; E=Major Loop
    sSchCompact           As String * 1      ' Schedule Compact Pass: N=N/A;
                                             ' A=each iteration; B=Each Count;
                                             ' C=Minor Loop; D=Middle Loop;
                                             ' E=Major Loop
    sSchPreempt           As String * 1      ' Schedule Preempt Pass: N=N/A;
                                             ' A=each iteration; B=Each Count;
                                             ' C=Minor Loop; D=Middle Loop;
                                             ' E=Major Loop
    sSchHour              As String * 1      ' For Hour Loop: N=N/A; R=Randomize
                                             ' Counts; S=Sort Counts Only;
                                             ' B=Sort Counts and randomize =
                                             ' counts
    sSchDay               As String * 1      ' For Day Loop: N=N/A; R=Randomize
                                             ' Counts; S=Sort Counts Only;
                                             ' B=Sort Counts and randomize =
                                             ' counts
    sSchQH                As String * 1      ' For Quarter Hour Loop: N=N/A;
                                             ' R=Randomize Counts; S=Sort Counts
                                             ' Only; B=Sort Counts and randomize
                                             ' = counts
    iImptSeqNo            As Integer         ' Contract Import Sequence Number
    sName                 As String * 20
    iCurSeqNo             As Integer         ' Current Sequence number so that
                                             ' move can be reimported
    lLastMoveID           As Long            ' ID of last move ID successfully
                                             ' completed
    iVefCode              As Integer
    lLowPrice             As Long
    'lLevelToPrice(1 To 11)        As Long
    lLevelToPrice(0 To 10)        As Long
    lHighPrice            As Long
    iPreferredPct         As Integer         ' Prefereed percentage of spots to
                                             ' try to book into the preferred
                                             ' days/times (min # spots will be
                                             ' 1)
    iWk1stSoloIndex       As Integer         ' Last Week, 1st Position and Solo
                                             ' Avail Priority index multiplier
    sInvISCIForm          As String * 1      ' Show ISCI on Invoice: R=Truncate
                                             ' Right; L=Truncate Left; W=Wrap
                                             ' around
    lGGEMailSentDate      As Long            ' Last date E-Mail sent to service about GG
    iLastArchRunDate(0 To 1) As Integer          'Rotation Archived through date
    lGPBatchNo            As Long            'Great Plain Batch Number
    sGPCustomerNo         As String * 10     'Great Plain Customer Number
    sGPPrefixChar         As String * 1      'Great Plain Customer number Prefix
    iRptLenDefault(0 To 4) As Integer        'Report Default Spot Lengths
    iRetainTrafProj       As Integer         'Number of months from last invoice date to retain projection (use rollover date for comparison)
    iEarliestAffSpot(0 To 1) As Integer      'Earliest affiliate spot date set in DB_Archive.
    iRetainCntr           As Integer         'Number of months from last invoice date to retain contracts
    iRetainPayRevHist     As Integer         'Number of months from last invoice date to retain Payment and Revenue history
    sReSchdXCal           As String * 1      ' Reschedule missed spots across
                                             ' calendar month boundaries
    iMaxPctWklyOnDay      As Integer         ' Max % of spots to schedule on day
                                             ' for weekly buys
    iThreshold            As Integer         ' Min # of weekly buy spots to
                                             ' incorporate safMaxPctWklyOnDay
    iNoDaysRetainUAF      As Integer         ' Number of days to retain User
                                             ' Activity in past
    iBLastWeeklyDate(0 To 1) As Integer      ' Last Weekly Billed date
    sFinalLogDisplay      As String * 1      ' Allow Final Logs to be Displayed (Y/N).  Test for Y
    sProdProtMan          As String * 1      ' Product Production Mandatory
                                             ' (Y/N). Test for Y
    sAvailGreenBar        As String * 1      ' Open avail green bar (Y/N).  Test
                                             ' for Y. N= Don't show. Y= % sold
                                             ' 0-50:Don't Show;50-85:Show
                                             ' Narrow; 85+ Show Large
    sIPumpZone            As String * 1      ' Wegener iPump time zone (E, C, M
                                             ' or P)
    iVCreativeDate(0 To 1) As Integer        ' Last date vCreative checked for
                                             ' completed Copy
    sInvoiceSort          As String * 1      ' Invoice Sort by Payee (P) or
                                             ' Sales Source (S).  Test for S
    sGenAutoFileWOSpt     As String * 1      ' Generate Automation files without
                                             ' Spots (Y/N). Test for Y
    sFeatures1            As String * 1      ' Bit Map: Bit 0= Matrix-Calendar;
                                             ' Bit 1=Hide Media Code in Engr
                                             ' Export (NY export); Bit 2=Show
                                             ' Audio Type from override on
                                             ' proposal/order; Bit 3=Show Spot
                                             ' Prices on Insertion Orders when
                                             ' Acquistion Cost Exist; Bit
                                             ' Bit 6=Jelli; Bit 7=Compensation
    sFeatures2            As String * 1      ' Bit Map: Bit 0=Event Revenue;1=Hide Hidden Lines; 2= Cancel Clause Mandatory; 3=E-Mail Distribution; 4=Acq Commissionable; 5=Pay on Collection; Export Tableau
    sEventTitle1          As String * 15     ' Event Title 2 (Ex: Visiting Team)
    sEventTitle2          As String * 15     ' Event Title 2 (Ex: Home Team)
    sEventSubtotal1       As String * 15     ' Event Subtotal 1 Column Title
    sEventSubtotal2       As String * 15     ' Event Subtotal 2 Column Title
    sCreditLimitMsg       As String * 1      ' W=Warning Message; C=Cutoff message
    sXMidSpotsBill        As String * 1      ' Bill spots that cross midnight:
                                             ' A=In Aired Month; O or Blank = In
                                             ' Ordered Month
    sHideDemoOnBR         As String * 1      ' Hide demo titles on Contract
                                             ' print out (BR). Y=Yes, N=No. Test
                                             ' for Y.
    sClientSentToEDS      As String * 1      ' Client Name Sent To EDS: Y=Yes;
                                             ' N=No.  Test for Y
    sXDSHeadEndZone       As String * 1      ' X-Digital Head End Time zone (E,
                                             ' C, M or P). Blank = E
    iXDSLastImptDate(0 To 1) As Integer      ' X-Digital last Import date of
                                             ' spots
    sSyncCopyInRot        As String * 1      ' Synxhronize Copy within same
                                             ' rotation across vehicle (Y/N).
                                             ' Test for Y
    sFeatures3            As String * 1      ' Bit Map: Bit 0=Suppress Acq Net/Commission on Insertion order;1=Station Posting Required to Generate Invoice;
                                             ' 2 & 3= Split Copy Station State (0=Mailing; 1=License; 2=Physical)
                                             ' 5=Activate Inv E-Mail; 6=Auto send Inv E-Mail
    lStatementComment     As Long            ' Statement comment
    sAudByPackage         As String * 1      ' Audience data by standard
                                             ' packages (Y=Yes, N or Blank =
                                             ' No). Test for Y
    sFeatures4            As String * 1      ' Bit Map: Bit 0=Filemaker Contract Import; Bit 1=ACT1 Lineup; Bit 2=Show Market Name on BR; Bit 3: Compress Transactions; Bits 4-7 Avail Count UInclude: Resedrvation; Renmant; DR; PI
    sFeatures5            As String * 1      ' Bit Map (Programmatic Buy): bit 0=Allowed; Bit 1=Show Avail Count; bit 2=Show CPP tab; bit 3=Show CPM tab; bit 4=Show Price tab; Bit 5=Pod cast vehices audience by %; Bit 6=Show Day Dropdown
    sFeatures6            As String * 1      ' Bit Map: Bit 0=Insertion Order include Monthly Billed Summary; Bit 1=X-Digital: substitute astCode for the Unit ID for Isci Method; Bit 2=Signature Line on Proposal; Bit 3=Calculate Research on each line change
                                             ' 4=Using EDIClient;5=Unused was AdvanceAvails; 6=RAB Calendar; 7=Suppress Pre-emption
    lCitationComment      As Long            ' Citation comment
    sExcludeAudioTypeR    As String * 1      ' Exclude Autio Type Recorded
                                             ' Commercial (RC: R). (Y/N, Test
                                             ' for Y. Treat Blank same as N)
    sExcludeAudioTypeL    As String * 1      ' Exclude Autio Type Live
                                             ' Commercial (LC: L). (Y/N, Test
                                             ' for Y. Treat Blank same as N)
    sExcludeAudioTypeM    As String * 1      ' Exclude Autio Type Live Promo
                                             ' (LP: M). (Y/N, Test for Y. Treat
                                             ' Blank same as N)
    sExcludeAudioTypeS    As String * 1      ' Exclude Autio Type Recorded Promo
                                             ' (RP: S). (Y/N, Test for Y. Treat
                                             ' Blank same as N)
    sExcludeAudioTypeP    As String * 1      ' Exclude Autio Type Pre-Recorded
                                             ' Live Commercial (PC: P). (Y/N,
                                             ' Test for Y. Treat Blank same as
                                             ' N)
    sExcludeAudioTypeQ    As String * 1      ' Exclude Autio Type Pre-Recorded
                                             ' Live Promo (PP: Q). (Y/N, Test
                                             ' for Y. Treat Blank same as N)
    sInvExpDelimiter      As String * 2      ' Invoice Export delimiter.  If blank use comma (,).
    sAdvanceAvail         As String * 1      ' Show Advance Avail Tab on the
                                             ' contract screen (Y/N). Test for Y
    sFeatures7            As String * 1      ' Bit Map: Bit 0=RAB Std; Bit 1=RAB Cal-Spots; Bit 2=iMediaTouch Replace COM with Media Code;3 = invoice email air only; 4 = invoice email ntr only;5= Std/Cal Custom Rev;6= Bill Over-Delivered CPM Impression
    sFeatures8            As String * 1      ' Add Bits: 0=AirTime; 1=Podcast Spots; 2=Podcast Ad Server; 3=Ad Server Tab View Only; 4=ShowWk Of
    sUnused               As String * 17     ' Unused
End Type

'Saf key record layout- use INTKEY0
'Type SAFKEY0
'    iCode As Integer
'End Type
'Saf key record layout- use INTKEY0

Type SAFKEY1
    iVefCode As Integer
End Type
Dim tmSaf As SAF
Dim hmSaf As Integer
Dim imSafRecLen As Integer
Dim tmSafSrchKey1 As SAFKEY1    'Vef key record image
Dim tmAdf As ADF        'ADF record image
Dim tmAdfSrchKey1 As ADFKEY1    'ADF key record image
Dim imAdfRecLen As Integer        'ADF record length
Dim tmAgf As AGF        'AGF record image
'10148 altered keyname and field name
Dim tmAgfSrchKey1 As AGFKEY1    'AGF key record image
Dim imAgfRecLen As Integer        'AGF record length



'*******************************************************
'*                                                     *
'*      Procedure Name:gGetSchParameters               *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get the schedule parameters     *
'*                     defined in Traffic.Ini          *
'*                                                     *
'*******************************************************
Sub gGetSchParameters()
    Dim ilRet As Integer
    Dim ilFound As Integer  'Traffic.Ini found in current directory
    Dim llNoRec As Long
    ilFound = True
    'On Error GoTo mGetSchParametersErr
    'slFileName = CurDir$ & "\Traffic.Ini"
    sgInsertPass = "A"
    'When running in debug mode- the error is not occurring-so test return string
    'ilRet = GetPrivateProfileString("Schedule", "Random", "Not Found", slReturn, 128, slFileName)
    'If Left$(slReturn, ilRet) = "Not Found" ThenFimMajor
    '    ilFound = False
    'End If
    'If ilFound Then
    '    slName = Left$(slReturn, ilRet)
    '    If StrComp(slName, "F", 1) = 0 Then
    '        igRnd = False
    '    Else
    '        igRnd = True
    '    End If
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Major", "H", slReturn, 128, slFileName)
    '    sgMajor = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Middle", "D", slReturn, 128, slFileName)
    '    sgMiddle = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Minor", "Q", slReturn, 128, slFileName)
    '    sgMinor = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Cycle Pattern", "1", slReturn, 128, slFileName)
    '    igCyclePattern = Val(Left$(slReturn, ilRet))
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Move Pass", "A", slReturn, 128, slFileName)
    '    sgMovePass = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Compact Pass", "A", slReturn, 128, slFileName)
    '    sgCompPass = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Preempt Pass", "A", slReturn, 128, slFileName)
    '    sgPreemptPass = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Hour: Sort/Random", "B", slReturn, 128, slFileName)
    '    sgSortRndHour = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Day: Sort/Random", "B", slReturn, 128, slFileName)
    '    sgSortRndDay = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "QH: Sort/Random", "B", slReturn, 128, slFileName)
    '    sgSortRndQH = Left$(slReturn, ilRet)
    'Else
    '    slFileName = "Traffic.Ini"
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Random", "T", slReturn, 128, slFileName)
    '    slName = Left$(slReturn, ilRet)
    '    If StrComp(slName, "F", 1) = 0 Then
    '        igRnd = False
    '    Else
    '        igRnd = True
    '    End If
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Major", "H", slReturn, 128, slFileName)
    '    sgMajor = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Middle", "D", slReturn, 128, slFileName)
    '    sgMiddle = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Minor", "Q", slReturn, 128, slFileName)
    '    sgMinor = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Cycle Pattern", "1", slReturn, 128, slFileName)
    '    igCyclePattern = Val(Left$(slReturn, ilRet))
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Move Pass", "A", slReturn, 128, slFileName)
    '    sgMovePass = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Compact Pass", "A", slReturn, 128, slFileName)
    '    sgCompPass = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Preempt Pass", "A", slReturn, 128, slFileName)
    '    sgPreemptPass = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Hour: Sort/Random", "B", slReturn, 128, slFileName)
    '    sgSortRndHour = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "Day: Sort/Random", "B", slReturn, 128, slFileName)
    '    sgSortRndDay = Left$(slReturn, ilRet)
    '    On Error Resume Next
    '    ilRet = GetPrivateProfileString("Schedule", "QH: Sort/Random", "B", slReturn, 128, slFileName)
    '    sgSortRndQH = Left$(slReturn, ilRet)
    'End If
    hmSaf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        imSafRecLen = Len(tmSaf) 'btrRecordLength(hmSaf)  'Get and save record length
        llNoRec = btrRecords(hmSaf) 'Obtain number of records
        If llNoRec = 0 Then
            btrDestroy hmSaf
            hmSaf = CBtrvTable(TWOHANDLES)
            ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                'Schedule parameters missing
                tmSaf.sSchRNG = "T"
                tmSaf.sSchMaL = "H"
                tmSaf.sSchMdL = "D"
                tmSaf.sSchMiL = "Q"
                tmSaf.sSchCycle = "1"
                tmSaf.sSchMove = "N"
                tmSaf.sSchCompact = "N"
                tmSaf.sSchPreempt = "E"
                tmSaf.sSchHour = "B"
                tmSaf.sSchDay = "B"
                tmSaf.sSchQH = "B"
                tmSaf.sName = ""
                tmSaf.iVefCode = 0
                tmSaf.lLowPrice = 0
                tmSaf.lLevelToPrice(0) = 0
                tmSaf.lLevelToPrice(1) = 0
                tmSaf.lLevelToPrice(2) = 0
                tmSaf.lLevelToPrice(3) = 0
                tmSaf.lLevelToPrice(4) = 0
                tmSaf.lLevelToPrice(5) = 0
                tmSaf.lLevelToPrice(6) = 0
                tmSaf.lLevelToPrice(7) = 0
                tmSaf.lLevelToPrice(8) = 0
                tmSaf.lLevelToPrice(9) = 0
                tmSaf.lLevelToPrice(10) = 0
                'tmSaf.lLevelToPrice(11) = 0
                tmSaf.lHighPrice = 0
                tmSaf.iPreferredPct = 0
                tmSaf.iWk1stSoloIndex = 0
                tmSaf.sInvISCIForm = "R"
                tmSaf.lGGEMailSentDate = 0
                tmSaf.iRetainTrafProj = 24
                tmSaf.iRetainCntr = 60
                tmSaf.iRetainPayRevHist = 60
                tmSaf.iNoDaysRetainUAF = 5
                gPackDate "1/1/1990", tmSaf.iLastArchRunDate(0), tmSaf.iLastArchRunDate(1)
                gPackDate "1/1/1990", tmSaf.iEarliestAffSpot(0), tmSaf.iEarliestAffSpot(1)
                gPackDate "1/1/1990", tmSaf.iBLastWeeklyDate(0), tmSaf.iBLastWeeklyDate(1)
                gPackDate "1/1/1970", tmSaf.iXDSLastImptDate(0), tmSaf.iXDSLastImptDate(1)
                tmSaf.sFinalLogDisplay = "N"
                tmSaf.sProdProtMan = "N"
                tmSaf.sAvailGreenBar = "N"
                tmSaf.sIPumpZone = ""
                gPackDate "1/1/1990", tmSaf.iVCreativeDate(0), tmSaf.iVCreativeDate(1)
                tmSaf.sInvoiceSort = "P"
                tmSaf.sGenAutoFileWOSpt = "N"
                tmSaf.sFeatures1 = Chr(0)
                tmSaf.sFeatures2 = Chr(0)
                tmSaf.sEventTitle1 = ""
                tmSaf.sEventTitle2 = ""
                tmSaf.sEventSubtotal1 = ""
                tmSaf.sEventSubtotal2 = ""
                tmSaf.sCreditLimitMsg = "W"
                tmSaf.sXMidSpotsBill = "O"
                tmSaf.sHideDemoOnBR = "N"
                tmSaf.sClientSentToEDS = "N"
                tmSaf.sXDSHeadEndZone = "E"
                tmSaf.sSyncCopyInRot = "N"
                tmSaf.sFeatures3 = Chr(0)
                tmSaf.lStatementComment = 0
                tmSaf.sAudByPackage = "N"
                tmSaf.sFeatures4 = Chr(0)
                tmSaf.sFeatures5 = Chr(0)
                tmSaf.sFeatures6 = Chr(0)
                tmSaf.lCitationComment = 0
                tmSaf.sExcludeAudioTypeR = "N"       ' Exclude Autio Type Recorded
                                                         ' Commercial (RC: R). (Y/N, Test
                                                         ' for Y. Treat Blank same as N)
                tmSaf.sExcludeAudioTypeL = "N"         ' Exclude Autio Type Live
                                                         ' Commercial (LC: L). (Y/N, Test
                                                         ' for Y. Treat Blank same as N)
                tmSaf.sExcludeAudioTypeM = "N"         ' Exclude Autio Type Live Promo
                                                         ' (LP: M). (Y/N, Test for Y. Treat
                                                         ' Blank same as N)
                tmSaf.sExcludeAudioTypeS = "N"         ' Exclude Autio Type Recorded Promo
                                                         ' (RP: S). (Y/N, Test for Y. Treat
                                                         ' Blank same as N)
                tmSaf.sExcludeAudioTypeP = "N"         ' Exclude Autio Type Pre-Recorded
                                                         ' Live Commercial (PC: P). (Y/N,
                                                         ' Test for Y. Treat Blank same as
                                                         ' N)
                tmSaf.sExcludeAudioTypeQ = "N"        ' Exclude Autio Type Pre-Recorded
                tmSaf.sInvExpDelimiter = ","            'Invoice export delimiter
                tmSaf.sFeatures7 = Chr(0)
                tmSaf.sUnused = ""
                tmSaf.iCode = 0  'Autoincrement
                ilRet = btrInsert(hmSaf, tmSaf, imSafRecLen, INDEXKEY0)
            End If
        Else
            ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                'Schedule parameters missing
                tmSaf.sSchRNG = "T"
                tmSaf.sSchMaL = "H"
                tmSaf.sSchMdL = "D"
                tmSaf.sSchMiL = "Q"
                tmSaf.sSchCycle = "1"
                tmSaf.sSchMove = "N"
                tmSaf.sSchCompact = "N"
                tmSaf.sSchPreempt = "E"
                tmSaf.sSchHour = "B"
                tmSaf.sSchDay = "B"
                tmSaf.sSchQH = "B"
            End If
        End If
    Else
        'Schedule parameters missing
        tmSaf.sSchRNG = "T"
        tmSaf.sSchMaL = "H"
        tmSaf.sSchMdL = "D"
        tmSaf.sSchMiL = "Q"
        tmSaf.sSchCycle = "1"
        tmSaf.sSchMove = "N"
        tmSaf.sSchCompact = "N"
        tmSaf.sSchPreempt = "E"
        tmSaf.sSchHour = "B"
        tmSaf.sSchDay = "B"
        tmSaf.sSchQH = "B"
    End If
    ilRet = btrClose(hmSaf)
    btrDestroy hmSaf
    If StrComp(tmSaf.sSchRNG, "F", 1) = 0 Then
        igRnd = False
    Else
        igRnd = True
    End If
    sgMajor = UCase$(tmSaf.sSchMaL)
    sgMiddle = UCase$(tmSaf.sSchMdL)
    sgMinor = UCase$(tmSaf.sSchMiL)
    igCyclePattern = Val(tmSaf.sSchCycle)
    sgMovePass = UCase$(tmSaf.sSchMove)
    sgCompPass = UCase$(tmSaf.sSchCompact)
    sgPreemptPass = UCase$(tmSaf.sSchPreempt)
    sgSortRndHour = UCase$(tmSaf.sSchHour)
    sgSortRndDay = UCase$(tmSaf.sSchDay)
    sgSortRndQH = UCase$(tmSaf.sSchQH)
    ilRet = gObtainSAF()
    Exit Sub

    ilFound = False
    Resume Next
End Sub

Public Function gGetNextGPNo(hlAdf As Integer, hlAgf As Integer, hlSaf As Integer) As String
    Dim ilRet As Integer
    Dim slNextNo As String

    If (Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) <> GREATPLAINS Then
        gGetNextGPNo = ""
        Exit Function
    End If
    imSafRecLen = Len(tmSaf)
    tmSafSrchKey1.iVefCode = 0
    ilRet = btrGetEqual(hlSaf, tmSaf, imSafRecLen, tmSafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        gGetNextGPNo = ""
        Exit Function
    End If
    slNextNo = Trim$(tmSaf.sGPCustomerNo)
    If slNextNo = "" Then
        slNextNo = "1"
    End If
    imAdfRecLen = Len(tmAdf)
    imAgfRecLen = Len(tmAgf)
    Do
        tmAdfSrchKey1.sCodeRep = slNextNo
        ilRet = btrGetEqual(hlAdf, tmAdf, imAdfRecLen, tmAdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmAgfSrchKey1.sCodeRep = slNextNo
            ilRet = btrGetEqual(hlAgf, tmAgf, imAgfRecLen, tmAgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                gGetNextGPNo = slNextNo
                slNextNo = gAddStr(slNextNo, "1")
                tmSaf.sGPCustomerNo = slNextNo
                ilRet = btrUpdate(hlSaf, tmSaf, imSafRecLen)
                Exit Function
            End If
        End If
        slNextNo = gAddStr(slNextNo, "1")
    Loop

End Function

Public Function gGPNoOk(slGPNoToTest As String, ilAdfCode As Integer, ilAgfCode As Integer, hlAdf As Integer, hlAgf As Integer) As Integer
    Dim ilRet As Integer

    If (Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) <> GREATPLAINS Then
        gGPNoOk = True
        Exit Function
    End If
    If Trim$(slGPNoToTest) = "" Then
        gGPNoOk = False
        Exit Function
    End If
    imAdfRecLen = Len(tmAdf)
    imAgfRecLen = Len(tmAgf)
    tmAdfSrchKey1.sCodeRep = slGPNoToTest
    ilRet = btrGetEqual(hlAdf, tmAdf, imAdfRecLen, tmAdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        tmAgfSrchKey1.sCodeRep = slGPNoToTest
        ilRet = btrGetEqual(hlAgf, tmAgf, imAgfRecLen, tmAgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            gGPNoOk = True
            Exit Function
        Else
            If tmAgf.iCode = ilAgfCode Then
                gGPNoOk = True
                Exit Function
            End If
        End If
    Else
        If tmAdf.iCode = ilAdfCode Then
            gGPNoOk = True
            Exit Function
        End If
    End If
    gGPNoOk = False
End Function
Public Sub gSetRptGGFlag(tlSaf As SAF)
    Dim slStr As String
    Dim llDate As Long
    Dim ilRptDays As Integer
    Dim llNow As Long
    
    If igGGFlag = 0 Then
        igRptGGFlag = 0
        If Len(Trim$(tlSaf.sName)) > 16 Then
            slStr = Mid$(tlSaf.sName, 2, 5)
            llDate = Val(slStr)
            slStr = Mid$(tlSaf.sName, 17)
            ilRptDays = Val(slStr)
            llNow = gDateValue(Format$(gNow(), "m/d/yy"))
            If llDate + ilRptDays > llNow Then
                igRptGGFlag = 1
            End If
        End If
    Else
        igRptGGFlag = 1
    End If
End Sub
