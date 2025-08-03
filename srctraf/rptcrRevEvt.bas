Attribute VB_Name = "RptCrRevEvt"
Option Explicit
Option Compare Text


Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length

Dim hmVsf As Integer            'Vehicle file handle
Dim tmVsf As VSF                'VEF record image
Dim imVsfRecLen As Integer       'VEF record length

Dim hmVtf As Integer            'Vehicle text file handle
Dim tmVtf As VTF                'vtf record image
Dim imVtfRecLen As Integer       'vtf record length

Dim hmCHF As Integer            'Contract Header file handle
Dim tmChf As CHF                'CHF record image
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChfSrchKey1 As CHFKEY1    'key by contract#
Dim tmChfSrchKey As LONGKEY0

Dim hmAdf As Integer            'Advertiser  file handle
Dim tmAdf As ADF                'ADF record image
Dim imAdfRecLen As Integer      'ADF record length

Dim hmGhf As Integer            'Temp  file handle
Dim tmGhf As GHF                'Temp file record image
Dim imGhfRecLen As Integer      'Temp file record length
Dim tmGhfSrchKey1 As GHFKEY1

Dim hmGsf As Integer            'Temp  file handle
Dim tmGsf As GSF                'Temp file record image
Dim imGsfRecLen As Integer      'Temp file record length
Dim tmGsfSrchKey1 As GSFKEY1

Dim hmMnf As Integer            'Multi-list file handle
Dim tmMnf As MNF                '
Dim imMnfRecLen As Integer

Dim hmRvf As Integer            'Receivables  file handle
Dim tmRvf As RVF                'RVF file record image
Dim imRvfRecLen As Integer      'RVF file record length
Dim hmPhf As Integer            'History  file handle

Dim hmCbf As Integer            'Generic temporary file handle
Dim tmCbf As CBF                'Cbf file record image
Dim imCbfRecLen As Integer      'Cbf file record length

Dim hmSbf As Integer            'Special Billing  file handle
Dim tmSbf As SBF                'SBF file record image
Dim imSbfRecLen As Integer      'SBF file record length

Dim imUsevefcodes() As Integer        'array of vehicle codes to include/exclude
Dim imInclVefCodes As Integer               'flag to incl or exclude vehicle codes
Dim imUseAdvtCodes() As Integer        'array of advt codes to include/exclude
Dim imInclAdvtCodes As Integer               'flag to incl or exclude advt codes

Dim imSort1 As Integer
Dim imSort2 As Integer
Dim imSort3 As Integer

Dim lmSingleCntr As Long

Dim smStartDate As String
Dim smEndDate As String

Dim tmTranTypes As TRANTYPES
Dim tmNTRTypes As SBFTypes

'If adding or changing order of sort/selection list boxes, change these constants and also
'see rptvfyRevEvt for any further tests.
Const SORT_ADVT = 1
Const SORT_TITLE1 = 2
Const SORT_TITLE2 = 3
Const SORT_SUBT1 = 4
Const SORT_SUBT2 = 5
Const SORT_VEHICLE = 6

'
'           Open files required for Spot Business Booked
'           Return - error flag = true for open error
'
Private Function mOpenRevEvtFiles() As Integer
Dim ilRet As Integer
Dim slTable As String * 3
Dim ilError As Integer

    ilError = False
    On Error GoTo mOpenRevEvtFilesErr

    slTable = "Chf"
    hmCHF = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRevEvtFiles = True
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Function
    End If
    imCHFRecLen = Len(tmChf)
       
    slTable = "Rvf"
    hmRvf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRevEvtFiles = True
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        Exit Function
    End If
    imRvfRecLen = Len(tmRvf)

    slTable = "Ghf"
    hmGhf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRevEvtFiles = True
        ilRet = btrClose(hmGhf)
        btrDestroy hmGhf
        Exit Function
    End If
    imGhfRecLen = Len(tmGhf)


    slTable = "Gsf"
    hmGsf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRevEvtFiles = True
        ilRet = btrClose(hmGsf)
        btrDestroy hmGsf
        Exit Function
    End If
    imGsfRecLen = Len(tmGsf)
    
    slTable = "Mnf"
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRevEvtFiles = True
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        Exit Function
    End If
    imMnfRecLen = Len(tmMnf)
    
    slTable = "Vef"
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRevEvtFiles = True
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Function
    End If
    imVefRecLen = Len(tmVef)
    
    slTable = "Vsf"
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRevEvtFiles = True
        ilRet = btrClose(hmVsf)
        btrDestroy hmVsf
        Exit Function
    End If
    imVsfRecLen = Len(tmVsf)
   
    slTable = "Adf"
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRevEvtFiles = True
        ilRet = btrClose(hmAdf)
        btrDestroy hmAdf
        Exit Function
    End If
    imAdfRecLen = Len(tmAdf)
    
    slTable = "Phf"
    hmPhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPhf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRevEvtFiles = True
        ilRet = btrClose(hmPhf)
        btrDestroy hmPhf
        Exit Function
    End If
    imRvfRecLen = Len(tmRvf)
    
    slTable = "Cbf"
    hmCbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRevEvtFiles = True
        ilRet = btrClose(hmCbf)
        btrDestroy hmCbf
        Exit Function
    End If
    imCbfRecLen = Len(tmCbf)
    
    slTable = "Sbf"
    hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenRevEvtFiles = True
        ilRet = btrClose(hmSbf)
        btrDestroy hmSbf
        Exit Function
    End If
    imSbfRecLen = Len(tmSbf)


    Exit Function
    
mOpenRevEvtFilesErr:
    ilError = Err.Number
    gBtrvErrorMsg ilRet, "mOpenRevEvtFiles (OpenError) #" & str(ilError) & ": " & slTable, RptSelSN

    Resume Next
End Function
'
'       Generate prepass for Revenue by Event/Game for span of dates
'       Data will be obtained from Receivables/History for Invoice(IN) and
'       Adjustment(AN) transaction types
'
Public Sub gGenRevenueByEvent()

Dim ilError As Integer
Dim ilVefCode As Integer
Dim llStart As Long
Dim llEnd As Long

Dim ilRet As Integer
Dim ilOk As Integer
Dim ilLoop As Integer
Dim slType As String
Dim slNameCode As String
Dim slCode As String
Dim ilFoundCntr As Integer
Dim llLoopOnRvf As Long
Dim blOk As Boolean
Dim ilIndex As Integer
Dim ilTemp As Integer
Dim slEventTitle1 As String
Dim slEventTitle2 As String
Dim llChfCode As Long
Dim ilVefInx As Integer
Dim tlRvf() As RVF



        ilError = mOpenRevEvtFiles()
        If ilError Then
            Exit Sub            'at least 1 open error
        End If
                
        mObtainSelectivity
        llChfCode = mSingleContract()       'get selective contract header
        If llChfCode < 0 Then               'illegal contract
            Exit Sub
        End If
                    
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmCbf.lGenTime = lgNowTime
        tmCbf.iGenDate(0) = igNowDate(0)
        tmCbf.iGenDate(1) = igNowDate(1)
        ReDim tlRvf(0 To 0) As RVF
        ilRet = gObtainPhfRvf(RptSelRevEvt, smStartDate, smEndDate, tmTranTypes, tlRvf(), 0)        'get phf/rvf by trandate

        For llLoopOnRvf = LBound(tlRvf) To UBound(tlRvf) - 1
            tmRvf = tlRvf(llLoopOnRvf)
            If (tmRvf.lCntrNo = lmSingleCntr) Or (lmSingleCntr = 0) Then
                blOk = mFilterAllLists(tmRvf.iAdfCode, tmRvf.iAirVefCode)       'filter advertiser & vehicles
                If blOk Then
                
                    ilVefInx = gBinarySearchVef(tmRvf.iAirVefCode)      'reference vehicle to get its vehicle type.  For air time, must be sports vehicle.  For NTR , can be any vehicle type
                    
                    'test for air time, NTR
                    If (tmTranTypes.iNTR = True And tmRvf.lSbfCode > 0) Or (tmTranTypes.iAirTime = True And tmRvf.lSbfCode = 0 And tgMVef(ilVefInx).sType = "G") Then
                        gGetEventTitles tmRvf.iAirVefCode, slEventTitle1, slEventTitle2     'event titles could be down to the vehicle level
                        tmCbf.sSortField1 = Trim$(slEventTitle1)
                        tmCbf.sSortField2 = Trim$(slEventTitle2)
                        tmCbf.iVefCode = tmRvf.iAirVefCode
                        tmCbf.lContrNo = tmRvf.lCntrNo                       'contr # (not code)
                        tmCbf.iAdfCode = tmRvf.iAdfCode
                        tmCbf.lExtra4Byte = tmRvf.lGsfCode                        'game schedule
                        tmCbf.iStartDate(0) = tmRvf.iTranDate(0)
                        tmCbf.iStartDate(1) = tmRvf.iTranDate(1)
                        'gPDNToLong tmRvf.sGross, tmCbf.lValue(1)
                        'gPDNToLong tmRvf.sNet, tmCbf.lValue(2)
                        gPDNToLong tmRvf.sGross, tmCbf.lValue(0)
                        gPDNToLong tmRvf.sNet, tmCbf.lValue(1)
                        ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                    End If
                End If
            End If
        Next llLoopOnRvf
        
        
        'close all files
        mCloseRevEvtFiles
        Exit Sub
End Sub
'           mObtainSelectivity - gather all selectivity entered and place
'           in common variables
'
Private Sub mObtainSelectivity()
Dim slStart As String
Dim llStart As Long
Dim ilDay As Integer
Dim slStamp As String
Dim ilRet As Integer

        imSort1 = (RptSelRevEvt!cbcSort1.ListIndex) + 1     '0 will indicate no sort for other levels of sort
        imSort2 = RptSelRevEvt!cbcSort2.ListIndex
        imSort3 = RptSelRevEvt!cbcSort3.ListIndex
        
        smStartDate = RptSelRevEvt!calStart.Text
        smEndDate = RptSelRevEvt!calEnd.Text
        
        'Selective contract #
        lmSingleCntr = Val(RptSelRevEvt!edcContract.Text)

        'If including adjustments from receivables, look for AN transaction types for NTR, Air Time or Hard Cost (also by option)
        tmTranTypes.iAdj = True
        tmTranTypes.iInv = True
        tmTranTypes.iWriteOff = False
        tmTranTypes.iPymt = False
        tmTranTypes.iNTR = True
        tmTranTypes.iAirTime = True
        
        tmTranTypes.iHardCost = True
        tmTranTypes.iCash = True
        tmTranTypes.iMerch = False
        tmTranTypes.iPromo = False
        tmTranTypes.iTrade = True
        
        'NTR types; SBF file has installment records and import records.  Only NTR (Hardcost) of interest
        tmNTRTypes.iImport = False
        tmNTRTypes.iInstallment = False
        tmNTRTypes.iNTR = True
        
        If RptSelRevEvt!rbcAirTimeNTR(1).Value = True Then          'ntr only, disable air time
            tmTranTypes.iAirTime = False
        End If
        If RptSelRevEvt!rbcAirTimeNTR(0).Value = True Then          'air time only, disable ntr
            tmTranTypes.iNTR = False
            tmNTRTypes.iNTR = False
            tmTranTypes.iHardCost = False
        End If
   
        mGetCodesFromList               'get list of adv and vehicles
        Exit Sub
End Sub
Private Sub mCloseRevEvtFiles()
Dim ilRet As Integer

        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        ilRet = btrClose(hmGhf)
        btrDestroy hmGhf
        ilRet = btrClose(hmGsf)
        btrDestroy hmGsf
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        ilRet = btrClose(hmVsf)
        btrDestroy hmVsf
        ilRet = btrClose(hmAdf)
        btrDestroy hmAdf
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        ilRet = btrClose(hmPhf)
        btrDestroy hmPhf
        ilRet = btrClose(hmCbf)
        btrDestroy hmCbf
        ilRet = btrClose(hmSbf)
        btrDestroy hmSbf
        
        Erase imUsevefcodes
        Erase imUseAdvtCodes
   
    Exit Sub
    
End Sub
'                       mFilterAllLists - filter user selections (list boxes) from header
'                       <input> ilAdfCode - advertiser code
'
Private Function mFilterAllLists(ilAdfCode As Integer, ilVefCode As Integer) As Boolean
Dim blOk As Boolean
Dim ilLoop As Integer
Dim ilSlspOK As Integer
Dim ilVefInx As Integer

        blOk = True
        If Not gFilterLists(ilAdfCode, imInclAdvtCodes, imUseAdvtCodes()) Then
            blOk = False
        End If
        
        If Not gFilterLists(ilVefCode, imInclVefCodes, imUsevefcodes()) Then
            blOk = False
        End If
        
        mFilterAllLists = blOk
        Exit Function
End Function
'
'               Obtain contract and filter selectivity
'               <input> ilByCodeOrNUmber: 0 = use code to retrieve contract
'                                         1 = use contract # (receivables dont have chfcodes)
'                       llContractKey: contract code or Number
Private Function mFilterContract(ilByCodeOrNumber As Integer, llContractKey As Long, ilfirstTime As Integer) As Integer
Dim ilOk As Integer
Dim ilFoundCntr As Integer
Dim ilIsItPolitical As Integer
Dim ilRet As Integer
Dim ilOKForUser As Integer

        ilOk = True
        If ilByCodeOrNumber = 0 Then            'get the contract by code
            If llContractKey <> tmChf.lCode Then
                tmChfSrchKey.lCode = llContractKey
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            End If
            If tmChf.sDelete = "Y" Then     'deleted header, shouldnt be used
                ilOk = False
            End If
        Else                                    'get the contract by #
            If llContractKey <> tmChf.lCntrNo Then
                tmChfSrchKey1.lCntrNo = llContractKey
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llContractKey)
                    If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) And (tmChf.sDelete <> "Y") Then
                        ilFoundCntr = True
                        Exit Do
                    End If
                    ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If Not ilFoundCntr Then
                    ilOk = False
                End If
            End If
        End If
        'contract found
        If (ilOk) Then
            ilOKForUser = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode())
            If Not ilOKForUser Then
                ilOk = False
            End If
            
        End If

        mFilterContract = ilOk
        Exit Function
End Function
'
'               mSingleContract - determine if single contract # entered
'                                 & retrieve it
'               <return> Single selection contract code
Private Function mSingleContract() As Long
Dim llChfCode As Long
Dim ilFoundCntr As Integer
Dim ilRet As Integer

        llChfCode = -1
        'determine if there is a single contract to retrieve
        ilFoundCntr = False
        If lmSingleCntr > 0 Then            'get the contracts internal code
            tmChfSrchKey1.lCntrNo = lmSingleCntr
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = lmSingleCntr)
                If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) And (tmChf.sDelete <> "Y") Then
                    ilFoundCntr = True
                    llChfCode = tmChf.lCode
                    Exit Do
                End If
                ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If Not ilFoundCntr Then
                mCloseRevEvtFiles
                mSingleContract = llChfCode
                Exit Function
            End If
        Else
            llChfCode = 0
        End If
        mSingleContract = llChfCode
End Function
'
'                   mGetCodesFromList - build array of each of the list boxes
'                   for faster testing
'                   Codes are built in arrays for advertisers, vehicles,

Private Sub mGetCodesFromList()
Dim ilLoop As Integer
Dim slNameCode As String
Dim slCode As String
Dim ilRet As Integer
Dim ilIndex As Integer
'ReDim imUsevefcodes(1 To 1) As Integer
ReDim imUsevefcodes(0 To 0) As Integer

        'setup array of codes to include or exclude, which is less for speed
        gObtainCodesForMultipleLists 0, tgAdvertiser(), imInclAdvtCodes, imUseAdvtCodes(), RptSelRevEvt
        
        'ReDim ilVehiclesToProcess(1 To 1) As Integer   'Not used
        'build array of vehicles to include or exclude
        'gObtainCodesForMultipleLists 5, tgVehicle(), imInclVefCodes, imUsevefcodes(), RptSelRevEvt

        imInclVefCodes = True               'need to get the entire list as the generalized filter list wont test for sports only vehicles
        For ilLoop = 0 To RptSelRevEvt!lbcSelection(5).ListCount - 1 Step 1
        slNameCode = tgVehicle(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If RptSelRevEvt!lbcSelection(5).Selected(ilLoop) Then               'selected ?
            imUsevefcodes(UBound(imUsevefcodes)) = Val(slCode)
            ReDim Preserve imUsevefcodes(LBound(imUsevefcodes) To UBound(imUsevefcodes) + 1)
        End If
        Next ilLoop

        Exit Sub
End Sub


