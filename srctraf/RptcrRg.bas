Attribute VB_Name = "RptcrRg"
Option Explicit
Option Compare Text
Dim tmSdfSrchKey1 As SDFKEY1            'SDF by vehicle, date
Dim tmSdfSrchKey3 As LONGKEY0           'SDF by internal code

Dim hmSdf As Integer
Dim tmSdf As SDF
Dim imSdfRecLen As Integer

Dim hmSmf As Integer
Dim tmSmf As SMF                'Spot Makegood
Dim imSmfRecLen As Integer

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length

Dim hmCHF As Integer            'Contract Header file handle
Dim tmChf As CHF                'CHF record image
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChfSrchKey1 As CHFKEY1

Dim hmAdf As Integer            'Advertiser  file handle
Dim tmAdf As ADF                'ADF record image
Dim imAdfRecLen As Integer      'ADF record length

Dim hmGrf As Integer            'Temp  file handle
Dim tmGrf As GRF                'Temp file record image
Dim imGrfRecLen As Integer      'Temp file record length

Dim hmCif As Integer            'Inventory  file handle
Dim tmCif As CIF                'Inventory file record image
Dim imCifRecLen As Integer      'Inventory file record length

Dim hmRaf As Integer            'Region name  file handle
Dim tmRaf As RAF                'Region file record image
Dim imRafRecLen As Integer      'Region file record length

Dim hmRsf As Integer            'Region copy  file handle
Dim tmRsf As RSF                'Region copy file record image
Dim imRsfRecLen As Integer      'Region copy file record length
Dim tmRsfSrchKey1 As LONGKEY0

Dim hmCrf As Integer            'Rotation  file handle
Dim tmCrf As CRF                'Rotation file record image
Dim imCrfRecLen As Integer      'Rotation file record length
Dim tmCrfSrchKey1 As CRFKEY1

Dim hmMnf As Integer            'Multiname  file handle
Dim tmMnf As MNF                'Multiname file record image
Dim imMnfRecLen As Integer      'Multiname file record length

Dim hmTzf As Integer            'Time zone copy  file handle
Dim tmTzf As TZF                'Time zone file record image
Dim imTzfRecLen As Integer      'Multiname file record length
Dim tmTzfSrchKey0 As LONGKEY0

Dim hmClf As Integer
Dim tmClf As CLF
Dim imClfRecLen As Integer

Dim hmCff As Integer
Dim tmCff As CFF
Dim imCffRecLen As Integer

Dim hmVsf As Integer
Dim tmVsf As VSF
Dim imVsfRecLen As Integer

Dim imInclAdvtCodes As Integer
Dim imUseAdvtCodes() As Integer
Dim imUseVehCodes() As Integer
Dim lmSingleCntr As Long
Dim tmSdfInfo() As SDFSORTBYLINE
Dim tmClfSrchKey As CLFKEY0



'
'           Open files required for Split Network Avails
'           Return - error flag = true for open error
'
Private Function mOpenRegionFiles() As Integer
Dim ilRet As Integer
Dim slTable As String * 3
Dim ilError As Integer

    ilError = False
    On Error GoTo mOpenRegionFilesErr

    slTable = "Chf"
    hmCHF = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Function
    End If
    imCHFRecLen = Len(tmChf)

    slTable = "Grf"
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Exit Function
    End If
    imGrfRecLen = Len(tmGrf)

    slTable = "Mnf"
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        Exit Function
    End If
    imMnfRecLen = Len(tmMnf)
    
    slTable = "Vef"
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Function
    End If
    imVefRecLen = Len(tmVef)
    
    slTable = "Sdf"
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmSdf)
        btrDestroy hmSdf
        Exit Function
    End If
    imSdfRecLen = Len(tmSdf)
        
    slTable = "Smf"
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmSmf)
        btrDestroy hmSmf
        Exit Function
    End If
    imSmfRecLen = Len(tmSmf)
    
    slTable = "Raf"
    hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmRaf)
        btrDestroy hmRaf
        Exit Function
    End If
    imRafRecLen = Len(tmRaf)

    slTable = "Cif"
    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmCif)
        btrDestroy hmCif
        Exit Function
    End If
    imCifRecLen = Len(tmCif)
    
    slTable = "Crf"
    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmCrf)
        btrDestroy hmCrf
        Exit Function
    End If
    imCrfRecLen = Len(tmCrf)

    slTable = "Adf"
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmAdf)
        btrDestroy hmAdf
        Exit Function
    End If
    imAdfRecLen = Len(tmAdf)
    
    slTable = "Tzf"
    hmTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmTzf)
        btrDestroy hmTzf
        Exit Function
    End If
    imTzfRecLen = Len(tmTzf)
    
    slTable = "Rsf"
    hmRsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRsf, "", sgDBPath & "Rsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmRsf)
        btrDestroy hmRsf
        Exit Function
    End If
    imRsfRecLen = Len(tmRsf)
    
    slTable = "Clf"
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        Exit Function
    End If
    imClfRecLen = Len(tmClf)
    
    slTable = "Cff"
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        Exit Function
    End If
    imCffRecLen = Len(tmCff)
    
    slTable = "Vsf"
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRegionFiles = True
        ilRet = btrClose(hmVsf)
        btrDestroy hmVsf
        Exit Function
    End If
    imVsfRecLen = Len(tmVsf)
    
    Exit Function
    
mOpenRegionFilesErr:
    ilError = err.Number
    gBtrvErrorMsg ilRet, "mOpenRegionFiles (OpenError) #" & str(ilError) & ": " & slTable, RptSelSN

    Resume Next
End Function
'
'       Generate prepass for Regional copy and Generic copy
'       for spots within a span of dates for selective advertiseres,
'       contracts, and vehicles
'
Public Sub gGenRegionalCopy()
Dim ilError As Integer
Dim ilLoopOnVeh As Integer  'loop variable for vehicles to process
Dim ilVefCode As Integer
Dim slStart As String       'earliest start date user request
Dim llStart As Long
Dim slEnd As String         'latest end date user request
Dim llEnd As Long
Dim ilWhichKey As Integer
Dim llSdfCodes() As Long
Dim tlSpotTypes As SPOTTYPES        'SDF spot types to include/exclude
Dim ilRet As Integer
Dim llLoopOnSpots As Long
Dim ilOk As Integer
Dim llSelectedChfCodes() As Long
Dim ilLoop As Integer
Dim slNameCode As String
Dim slCode As String
Dim slName As String
Dim ilInclMissed As Integer         'include Missed spots (user option)
Dim slType As String
Dim slBonusOnInv As String * 1
Dim slPrice As String
Dim llSpotRate As Long

        ilError = mOpenRegionFiles()
        If ilError Then
            Exit Sub            'at least 1 open error
        End If

'        slStart = RptSelRg!edcStartDate.Text
        slStart = RptSelRg!CSI_calStart.Text        '8-23-19 use csi calendar control vs edit box

        llStart = gDateValue(slStart)
        slStart = Format$(llStart, "m/d/yy")
'        slEnd = RptSelRg!edcEndDate.Text     'get the default year in case not entered
        slEnd = RptSelRg!CSI_CalEnd.Text     'get the default year in case not entered

        llEnd = gDateValue(slEnd)
        slEnd = Format$(llEnd, "m/d/yy")     'get the default year in case not entered
        
        ilWhichKey = INDEXKEY1              'use key 1 vehicle, date search

        ReDim llSelectedChfCodes(0 To 0) As Long
        lmSingleCntr = Val(RptSelRg!edcContract.Text)
        'determine if there is a single contract to retrieve
        llSelectedChfCodes(0) = 0
        If lmSingleCntr > 0 Then            'get the contracts internal code
            tmChfSrchKey1.lCntrNo = lmSingleCntr
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            If lmSingleCntr = tmChf.lCntrNo Then
                llSelectedChfCodes(0) = tmChf.lCode
                ReDim Preserve llSelectedChfCodes(0 To 1) As Long
            End If
        Else            'check for all advt and/or selected contracts for that advt
            For ilLoop = 0 To RptSelRg!lbcSelection(0).ListCount - 1 Step 1
                If RptSelRg!lbcSelection(0).Selected(ilLoop) Then
                    slNameCode = RptSelRg!lbcCntrCode.List(ilLoop)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                    llSelectedChfCodes(UBound(llSelectedChfCodes)) = Val(slCode)
                    ReDim Preserve llSelectedChfCodes(0 To UBound(llSelectedChfCodes) + 1) As Long
                End If
            Next ilLoop
        End If
        'setup array of codes to include or exclude, which is less for speed
        gObtainCodesForMultipleLists 1, tgAdvertiser(), imInclAdvtCodes, imUseAdvtCodes(), RptSelRg

        ReDim imUseVehCodes(0 To 0) As Integer
        For ilLoop = 0 To RptSelRg!lbcSelection(2).ListCount - 1 Step 1
            If (RptSelRg!lbcSelection(2).Selected(ilLoop)) Then
                slNameCode = tgVehicle(ilLoop).sKey
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slName, 3, "|", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                imUseVehCodes(UBound(imUseVehCodes)) = Val(slCode)
                ReDim Preserve imUseVehCodes(0 To UBound(imUseVehCodes) + 1) As Integer
            End If
        Next ilLoop
        'spot types to retrieve from SDF
        tlSpotTypes.iSched = True
        If RptSelRg!ckcInclMissed.Value = vbChecked Then     'include missed by user option
            tlSpotTypes.iMissed = True
        Else
            tlSpotTypes.iMissed = False
        End If
        tlSpotTypes.iMG = True
        tlSpotTypes.iOutside = True
        tlSpotTypes.iHidden = False
        tlSpotTypes.iCancel = False
        tlSpotTypes.iFill = True
        tlSpotTypes.iOpen = True        '1-18-11  previously always included open/close; added field for testing
                                        'due to exclusion required in B & B
        tlSpotTypes.iClose = True       '1-18-11

        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrf.lGenTime = lgNowTime
        tmGrf.iGenDate(0) = igNowDate(0)
        tmGrf.iGenDate(1) = igNowDate(1)

        'loop on array of selected vehicles
        For ilLoopOnVeh = 0 To UBound(imUseVehCodes) - 1
            ilVefCode = imUseVehCodes(ilLoopOnVeh)
            ReDim llSdfCodes(0 To 0) As Long
            ReDim tmSdfInfo(0 To 0) As SDFSORTBYLINE            'initialize for next vehicle
            'retrieving by vehicle with key1, only the array of sdfcodes will be used
            gObtainSDFByKey hmSdf, ilVefCode, slStart, slEnd, llSelectedChfCodes(0), ilWhichKey, llSdfCodes(), tmSdfInfo(), tlSpotTypes
            For llLoopOnSpots = LBound(llSdfCodes) To UBound(llSdfCodes) - 1
                tmSdfSrchKey3.lCode = llSdfCodes(llLoopOnSpots)
                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    gLogMsg "Invalid Spot ID: " & Trim$(str(llSdfCodes(llLoopOnSpots))), "Messages.txt", False
                    MsgBox "Invalid SpotID: " & Trim$(str(llSdfCodes(llLoopOnSpots)))
                    'continue with next spot
                Else
                    If tmSdf.lCode = 482 Then
                        ilRet = ilRet
                    End If
                    'filter out selection:  advt/cntr
                    ilOk = True             'assume to include, then filter advt selectivity
                    If Not gFilterLists(tmSdf.iAdfCode, imInclAdvtCodes, imUseAdvtCodes()) Then
                        ilOk = False
                    Else                'valid advt, what about selective contracts for the advt
                        'if all Advt selected, there are no contracts in the selected contract list;
                        'or, since it passed the valid advt if the all Contracts box is checked the contracts must be valid
                        If (RptSelRg!ckcAllAdvt.Value = vbChecked) Or (RptSelRg!ckcAllContracts.Value = vbChecked) Or (mTestSelectedContracts(llSelectedChfCodes(), tmSdf.lChfCode) = True) Then         'only check if not all advt checked, or not all contracts selected
                            'if All contracts selected, but single entered, need to test for single
                            ilOk = False
                            If mTestSelectedContracts(llSelectedChfCodes(), tmSdf.lChfCode) Then
                            
                                tmGrf.iVefCode = tmSdf.iVefCode
                                tmGrf.iDate(0) = tmSdf.iDate(0)     'sch date
                                tmGrf.iDate(1) = tmSdf.iDate(1)
                                tmGrf.iTime(0) = tmSdf.iTime(0)     'sch time
                                tmGrf.iTime(1) = tmSdf.iTime(1)
                                tmGrf.iCode2 = tmSdf.iLineNo        'sch line #
                                tmGrf.lChfCode = tmSdf.lChfCode     'internal contract code
                                'tmGrf.iPerGenl(1) = tmSdf.iLen      'spot length
                                tmGrf.iPerGenl(0) = tmSdf.iLen      'spot length
                                If tmSdf.sSpotType = "X" Then
                                    slBonusOnInv = "Y"
                                    ilLoop = gBinarySearchAdf(tmSdf.iAdfCode)
                                    If ilLoop <> -1 Then            'found the matching advertisr
                                        slBonusOnInv = tgCommAdf(ilLoop).sBonusOnInv
                                    End If
                                    If tmSdf.sPriceType = "+" Then
                                        tmGrf.sDateType = "+"
                                    ElseIf tmSdf.sPriceType = "-" Then
                                        tmGrf.sDateType = "-"
                                    Else
                                        If slBonusOnInv <> "N" Then     'blank (assume Yes) or Y
                                            tmGrf.sDateType = "+"
                                        Else
                                            tmGrf.sDateType = "-"
                                        End If
                                    End If

                                ElseIf tmSdf.sSchStatus = "G" Then       'makegood
                                    tmGrf.sDateType = "G"
                                ElseIf tmSdf.sSchStatus = "O" Then      'outside
                                    tmGrf.sDateType = "O"
                                ElseIf tmSdf.sSchStatus = "M" Then      'missed
                                    tmGrf.sDateType = "M"
                                ElseIf tmSdf.sSchStatus = "S" Then
                                    tmGrf.sDateType = ""            'normal scheduled spot
                                End If
                                
                                llSpotRate = 0
                        
'8-23-16 can uncomment out this section to get spot rate if debugging
'                                If tmSdf.lChfCode <> tmClf.lChfCode Then
'                                    tmClfSrchKey.lChfCode = tmSdf.lChfCode
'                                    tmClfSrchKey.iLine = tmSdf.iLineNo
'                                    tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
'                                    tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
'                                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
'                                End If
'
'                                If ilRet = BTRV_ERR_NONE Then
'                                    ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
'                                    If (InStr(slPrice, ".") <> 0) Then        'found spot cost
'                                        'is it a .00?
'                                        If gCompNumberStr(slPrice, "0.00") = 0 Then       'its a .00 spot
'                                            llSpotRate = 0
'                                        Else
'                                            llSpotRate = gStrDecToLong(slPrice, 2)
'                                        End If
'                                    Else
'                                        'its a bonus, recap, n/c, etc. which is still $0
'                                        llSpotRate = 0
'                                    End If
'                                End If
' ****
                                
                                tmGrf.lChfCode = tmSdf.lChfCode 'contract internal code
                                'tmGrf.lDollars(1) = tmSdf.lCode 'Internal spot code
                                'tmGrf.lDollars(2) = llSpotRate                      'spot rate from line
                                tmGrf.lDollars(0) = tmSdf.lCode 'Internal spot code
                                tmGrf.lDollars(1) = llSpotRate                      'spot rate from line

                                'find and setup generic copy
                                mGenericCopy (ilVefCode)
                          
                            End If
                        End If
                    End If

                End If

            Next llLoopOnSpots
        Next ilLoopOnVeh
        
        'close all files
        
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        ilRet = btrClose(hmSdf)
        btrDestroy hmSdf
        ilRet = btrClose(hmSmf)
        btrDestroy hmSmf
        ilRet = btrClose(hmRaf)
        btrDestroy hmRaf
        ilRet = btrClose(hmCif)
        btrDestroy hmCif
        ilRet = btrClose(hmCrf)
        btrDestroy hmCrf
        ilRet = btrClose(hmAdf)
        ilRet = btrClose(hmTzf)
        btrDestroy hmTzf
        ilRet = btrClose(hmRsf)
        btrDestroy hmRsf
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        ilRet = btrClose(hmVsf)
        btrDestroy hmVsf
        
        Exit Sub
End Sub
'
'         Find and create Generic copy for a spot
'
'       GRF variables:
'       grfgendate - generation date for filter
'       grfgentime - generation time for filter
'       grfDate     - date of spot
'       grfTime     - time of spot
'       grfCode2    - spot length
'       grfLong     - CRF (rotation header) internal code
'       grfChfCode  - internal contract code
'       grfPerGenl(1)- line #
'       grfPerGenl(2)- rot #
'       grfPerGenl(3)- Seq # (unused for now)
'       grfPerGenl(4)- Time zone flag (0 = All zones, 1 = EST, 2 = PST, 3 = CST, 4 = MST)
'       grfDollars(1)- Internal spot code (used to keep same spot info together when time zone copy used)
'       grfvefCode  - vehicle name that spot is in
'       grfBktType  - G = generic copy (vs REgional)
'       grfCode4    - CIF (inventory internal code)
'       grfSofCode  - rotation vehicle
'       grfDateType - M = Missed, G = MG, O = OUtside, + = show on inv fill, - is do not show on inv. fill
'       grfrdfcode  - 0= no regional exists, non zero = at least 1 regional.  Only required because bug in Crystal where it wont
'                     suppress the section if none exists.  Need to format the section with this flag.
Public Sub mGenericCopy(ilVefCode As Integer)
Dim slType As String
Dim ilRet As Integer
Dim ilZone As Integer
Dim slTimeZones As String * 12
Dim ilPos As Integer
Dim ilLoop As Integer

        slTimeZones = "ESTPSTCSTMST"
        tmGrf.sBktType = "G"            'flag as generic copy vs regional
        'tmGrf.iPerGenl(3) = 0           'seq # for generic/timezone copy
        tmGrf.iPerGenl(2) = 0           'seq # for generic/timezone copy
        tmGrf.lCode4 = 0                'initalize copy inventory code
        'tmGrf.iPerGenl(4) = 0           'time zone flags
        tmGrf.iPerGenl(3) = 0           'time zone flags
        'tmGrf.iPerGenl(2) = 0           'init rotation #
        tmGrf.iPerGenl(1) = 0           'init rotation #
        tmGrf.lLong = 0                 'init CRF internal code
        tmGrf.iRdfCode = 0
        If tmSdf.lCopyCode > 0 Then
            'copy exists, see if any regional exists for this spot
            tmRsfSrchKey1.lCode = tmSdf.lCode
            ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                'found at least 1 matching regional copy.  Rquired to test in Crystal to suppress regional copy subreport if none exists
                tmGrf.iRdfCode = 1
            End If
            
            If (tmGrf.iRdfCode > 0 And RptSelRg!ckcExclSpotsLackReg.Value = vbChecked) Or (RptSelRg!ckcExclSpotsLackReg.Value = vbUnchecked) Then
                If tmSdf.sPtType = "1" Then         'regular copy inventory
                    'tmGrf.iPerGenl(3) = 1               'seq #
                    'tmGrf.iPerGenl(2) = tmSdf.iRotNo        'Rotation #
                    tmGrf.iPerGenl(2) = 1               'seq #
                    tmGrf.iPerGenl(1) = tmSdf.iRotNo        'Rotation #
            
                    If tmSdf.sSpotType = "O" Then       'open bb
                        slType = "O"
                    ElseIf tmSdf.sSpotType = "C" Then       'closed bb
                        slType = "C"
                    Else
                        slType = "A"
                    End If
                    tmCrfSrchKey1.sRotType = slType
                    tmCrfSrchKey1.iEtfCode = 0
                    tmCrfSrchKey1.iEnfCode = 0
                    tmCrfSrchKey1.iAdfCode = tmSdf.iAdfCode
                    tmCrfSrchKey1.lChfCode = tmSdf.lChfCode
                    tmCrfSrchKey1.lFsfCode = 0
                    tmCrfSrchKey1.iVefCode = 0   'ilVefCode
                    tmCrfSrchKey1.iRotNo = tmSdf.iRotNo
                    ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iAdfCode = tmSdf.iAdfCode) And (tmCrf.lChfCode = tmSdf.lChfCode) And (tmCrf.lFsfCode = tmSdf.lFsfCode)
                        If tmCrf.iRotNo = tmSdf.iRotNo Then
                            'found the rotation pattern
                            tmGrf.lCode4 = tmSdf.lCopyCode
                            tmGrf.iSofCode = tmCrf.iVefCode 'rotation vehicle
                            tmGrf.lLong = tmCrf.lCode       'rotation internal code
                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                        End If
                        ilRet = btrGetNext(hmCrf, tmCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        
                    Loop
                ElseIf tmSdf.sPtType = "3" Then             'time zone copy
                    tmTzfSrchKey0.lCode = tmSdf.lCopyCode
                    'tmGrf.iPerGenl(3) = 1               'seq #
                    tmGrf.iPerGenl(2) = 1               'seq #
    
                    ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        For ilZone = 1 To 6 Step 1
                            tmGrf.lLong = 0             'init each zones crf internal code
                            If (Trim$(tmTzf.sZone(ilZone - 1)) <> "") And (tmTzf.lCifZone(ilZone - 1) <> 0) Then
                                If tmSdf.sSpotType = "O" Then       'open bb
                                    slType = "O"
                                ElseIf tmSdf.sSpotType = "C" Then       'closed bb
                                    slType = "C"
                                Else
                                    slType = "A"
                                End If
                                tmCrfSrchKey1.sRotType = slType
                                tmCrfSrchKey1.iEtfCode = 0
                                tmCrfSrchKey1.iEnfCode = 0
                                tmCrfSrchKey1.iAdfCode = tmSdf.iAdfCode
                                tmCrfSrchKey1.lChfCode = tmSdf.lChfCode
                                tmCrfSrchKey1.lFsfCode = 0
                                tmCrfSrchKey1.iVefCode = 0   'dont know if package or non-pkg vehicle to obtain
                                tmCrfSrchKey1.iRotNo = tmTzf.iRotNo(ilZone - 1)
                                ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                                Do While (ilRet = BTRV_ERR_NONE) And (tmCrf.sRotType = slType) And (tmCrf.iAdfCode = tmSdf.iAdfCode) And (tmCrf.lChfCode = tmSdf.lChfCode) And (tmCrf.lFsfCode = tmSdf.lFsfCode)
                                    If tmCrf.iRotNo = tmTzf.iRotNo(ilZone - 1) Then
                                        'found the rotation pattern
                                        tmGrf.lCode4 = tmTzf.lCifZone(ilZone - 1)
                                        tmGrf.iSofCode = tmCrf.iVefCode 'rotation vehicle
                                        'tmGrf.iPerGenl(2) = tmTzf.iRotNo(ilZone)        'Rotation #
                                        'tmGrf.iPerGenl(3) = tmGrf.iPerGenl(3) + ilZone      'seq # for each time zone within generic copy
                                        tmGrf.iPerGenl(1) = tmTzf.iRotNo(ilZone - 1)      'Rotation #
                                        tmGrf.iPerGenl(2) = tmGrf.iPerGenl(2) + ilZone      'seq # for each time zone within generic copy
                                        tmGrf.lLong = tmCrf.lCode       'rotation internal code
                                        'determine which time zone for flag to send to crystal
                                        ilPos = InStr(1, slTimeZones, tmTzf.sZone(ilZone - 1))
                                        If ilPos > 0 Then
                                            'tmGrf.iPerGenl(4) = (ilPos \ 3) + 1
                                            tmGrf.iPerGenl(3) = (ilPos \ 3) + 1
                                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                            ilRet = BTRV_ERR_END_OF_FILE            'force exit of dowhile
                                        End If
                                    Else
                                        ilRet = btrGetNext(hmCrf, tmCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    End If
                                Loop
                            Else
                                Exit For            'no more time zones in this record
                            End If
                        Next ilZone
                        'End If
                    End If          'err_btrv_none
                End If              'tmsdf.sPtType
            End If                  'RptSelRg.ckcExclSpotsLackReg
        Else                     'no copy exists
            If RptSelRg!ckcExclSpotsLackReg.Value = vbUnchecked Then        'excl spots lacking regional copy? Unchecked - get all
                tmGrf.iSofCode = 0          'no rotation vehicle exists
                tmGrf.lCode4 = 0            'no copy exists
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            End If
        End If
        Exit Sub
End Sub
'
'           Filter out any selected contracts
'           <input> llSelectedChfCodes - array of contract codes to includes
'
'           <return>  true = include
'                     false = exclude
Public Function mTestSelectedContracts(llSelectedChfCodes() As Long, llChfCode As Long) As Integer
Dim ilLoop As Integer

        mTestSelectedContracts = False
        If LBound(llSelectedChfCodes) = UBound(llSelectedChfCodes) Then     'nothing in list, defaults to ALL
            mTestSelectedContracts = True
            Exit Function
        End If
        
        For ilLoop = LBound(llSelectedChfCodes) To UBound(llSelectedChfCodes) - 1
            If llSelectedChfCodes(ilLoop) = llChfCode Then
                mTestSelectedContracts = True
                Exit Function
            End If
        Next ilLoop
        Exit Function
End Function
