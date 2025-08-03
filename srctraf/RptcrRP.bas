Attribute VB_Name = "RptcrRP"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptcrRP.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text


Dim imTrade As Integer  'true = include trade contracts
Dim imCash As Integer
Dim imMerchant As Integer   'true = include merchandise transactions
Dim imPromotion As Integer      'true =include promotions transactions
'  Receivables File
Dim hmRvf As Integer        'receivables file handle
Dim tmRvf As RVF            'RVF record image
Dim imRvfRecLen As Integer  'RVF record length
'  Receivables Report File
Dim hmRvr As Integer        'receivables report file handle
Dim tmRvr As RVR            'RVR record image
Dim imRvrRecLen As Integer  'RVR record length

Dim hmVef As Integer        'Vehicle table
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer  'VEF record length
'
'
'   Generate prepass file (RVR) gathering all transactions, excluding "IN", for the
'   Remote Posting (dual posting) between a span of Transaction Entry Dates and market
'   selectivity
'
'   Created: 11/5/02    d.hosaka
'
'
Sub gCRRPGen()
    Dim ilRet As Integer
    Dim ilLoop  As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilLoopOnFile As Integer             '2 passes, 1 for History, then Receivables
    Dim slStr As String
    Dim llDate As Long
    Dim llEarliestDate As Long              'start date of data to retrieve from PRF or RVF
    Dim llLatestDate As Long                'end date of data to retrieve from PRF or RVF
    Dim ilTransFound As Integer
    Dim ilIncludeH As Integer
    Dim ilIncludeI As Integer
    Dim ilIncludeP As Integer
    Dim ilIncludeA As Integer
    Dim ilIncludeW As Integer
    Dim ilStartFile As Integer      'if retrieving PHF, set to 1, otherwise RVF only and set to 2
    Dim ilEndFile As Integer        'if retrieving PHF only, set to 1, otherwise RVF and set to 2

    hmRvr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRvr, "", sgDBPath & "Rvr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRvr)
        btrDestroy hmRvr
        Exit Sub
    End If
    imRvrRecLen = Len(tmRvr)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read Vehicle files using RVF handles and buffers
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        btrDestroy hmRvr
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)

    imTrade = True                                         'assume trades should be included
    imCash = True
    imMerchant = False                                      'merchandising transactions
    imPromotion = False                                         'promotions transactions
    ilIncludeH = False                                     'include HI (inv history) transactions
    ilIncludeI = False                                     'exclude all I (invoice) transactions
    ilIncludeP = True                                     'include all P (payment) transactions
    ilIncludeA = True                                     'include all A (adjustment) transactions
    ilIncludeW = True                                     'include all W (Write off) transactions

    ilStartFile = 2                             'currently, force only RVF
    ilEndFile = 2                               'force to RVF only

'    slStr = RptSelRP!edcSelCFrom.Text               'Latest date to retrieve from PRF or RVF
    slStr = RptSelRP!CSI_CalFrom.Text               'Latest date to retrieve from PRF or RVF, 8-26-19 use csi cal control vs edit box
    llEarliestDate = gDateValue(slStr)
'    slStr = RptSelRP!edcSelCTo.Text
    slStr = RptSelRP!CSI_CalTo.Text
    llLatestDate = gDateValue(slStr)

    If ilStartFile = 1 Then
        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            btrDestroy hmRvr
            btrDestroy hmVef
            Exit Sub
        End If
        imRvfRecLen = Len(tmRvf)
    Else
        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read only RVF file (not PHF)
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            btrDestroy hmRvr
            btrDestroy hmVef
            Exit Sub
        End If
        imRvfRecLen = Len(tmRvf)
    End If


    For ilLoopOnFile = ilStartFile To ilEndFile Step 1                 '2 passes, first History, then Receivables
        'handles and buffers for PHF and RVF will be the same
        ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation

        Do While ilRet = BTRV_ERR_NONE
            gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slStr   'using entry date for filter
            llDate = gDateValue(slStr)                  'convert trans date to test if within requested limits

            'test for the valid transactions to include
            ilTransFound = False
            If (ilIncludeI) And (Left$(tmRvf.sTranType, 1) = "I") Then
                ilTransFound = True
            End If
            If (ilIncludeP) And (Left$(tmRvf.sTranType, 1) = "P") Then
                ilTransFound = True
            End If
            If (ilIncludeA) And (Left$(tmRvf.sTranType, 1) = "A") Then
                ilTransFound = True
            End If
            If (ilIncludeW) And (Left$(tmRvf.sTranType, 1) = "W") Then
                ilTransFound = True
            End If
            If (ilIncludeH) And (Left$(tmRvf.sTranType, 1) = "H") Then
                ilTransFound = True
            End If
            If tmRvf.sCashTrade = "C" And Not (imCash) Then
                ilTransFound = False
            ElseIf tmRvf.sCashTrade = "T" And Not (imTrade) Then
                ilTransFound = False
            ElseIf tmRvf.sCashTrade = "M" And Not (imMerchant) Then
                ilTransFound = False
            ElseIf tmRvf.sCashTrade = "P" And Not (imPromotion) Then
                ilTransFound = False
            End If


            If ilTransFound Then                        'so far, this is a valid trans, continue with the market selected
                If Not RptSelRP!ckcAll.Value = vbChecked Then        'all Markets selected?
                    ilTransFound = False
                    'filter the Markets selected
                    For ilLoop = 0 To RptSelRP!lbcSelection.ListCount - 1 Step 1
                        If RptSelRP!lbcSelection.Selected(ilLoop) Then
                            slNameCode = tgMktCode(ilLoop).sKey
                            ilRet = gParseItem(slNameCode, 1, "\", slName)
                            ilRet = gParseItem(slName, 3, "|", slName)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            'Determine which vehicle set to test
                            If tmRvf.iAirVefCode <> tmVef.iCode Then            'only read vef if not already in memory
                                tmVefSrchKey.iCode = tmRvf.iAirVefCode
                                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            End If
                            If tmVef.iMnfVehGp3Mkt = Val(slCode) Then
                                ilTransFound = True
                                Exit For
                            End If
                        End If
                    Next ilLoop
                End If
            End If


            If (ilTransFound) And (llDate >= llEarliestDate And llDate <= llLatestDate) Then
                LSet tmRvr = tmRvf
                tmRvr.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
                tmRvr.iGenDate(1) = igNowDate(1)
                gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                tmRvr.lGenTime = lgNowTime
                If ilLoopOnFile = 1 Then
                    tmRvr.sSource = "H"                    'let crystal know these records are histroy/receivables (vs contracts)
                Else
                    tmRvr.sSource = "R"
                End If
                ilRet = btrInsert(hmRvr, tmRvr, imRvrRecLen, INDEXKEY0)                                                '(for cash distribution reports)

            End If                                    'ilFoundTrans
            ilRet = btrGetNext(hmRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        ilRet = btrClose(hmRvf)

        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmRvf)
            btrDestroy hmRvf
            btrDestroy hmRvr
            Exit Sub
        End If
        imRvfRecLen = Len(tmRvf)
    Next ilLoopOnFile                                   '2 passes, first History, then Receivbles

    ilRet = btrClose(hmRvr)
    ilRet = btrClose(hmRvf)
    ilRet = btrClose(hmVef)
End Sub


