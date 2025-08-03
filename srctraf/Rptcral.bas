Attribute VB_Name = "RPTCRAL"
'*******************************************************************
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
'
' Description:
' This file contains the code for gathering the prepass data for the
' Alert Report (10-1-09)
'*******************************************************************
Option Explicit
Option Compare Text

Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GRF record length

Dim tmAuf As AUF
Dim hmAuf As Integer
Dim imAufRecLen As Integer        'AUF record length

Dim tmUrf As URF
Dim hmUrf As Integer
Dim imUrfRecLen As Integer        'URF record length

Dim tmUst As UST
Dim hmUst As Integer
Dim imUstRecLen As Integer        'UST record length
Dim tmSrchKey0 As INTKEY0

Dim tmChf As CHF
Dim hmCHF As Integer
Dim imCHFRecLen As Integer        'CHF record length
Dim tmLongSrchKey0 As LONGKEY0    'Key record image
'*********************************************************************
'       gCreateAlert()
'       Created 10/1/09
'       Create Alert report prepass file because of decryption of urfname
'
'*********************************************************************
Sub gCreateAlert()
Dim ilRet As Integer
Dim ilAlertContract As Integer
Dim ilAlertTraffic As Integer
Dim ilAlertAffiliate As Integer
Dim ilClearContract As Integer
Dim ilClearTraffic As Integer
Dim ilClearAffiliate As Integer
Dim llEffClearDate As Long
Dim slEffClearDate As String
Dim ilFound As Integer
Dim llAufClearDate As Long
Dim ilAlertPool As Integer
Dim ilClearPool As Integer

    hmAuf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAuf, "", sgDBPath & "Auf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAuf)
        btrDestroy hmAuf
        Exit Sub
    End If
    imAufRecLen = Len(tmAuf)

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmAuf)
        btrDestroy hmGrf
        btrDestroy hmAuf
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmUrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmAuf)
        btrDestroy hmUrf
        btrDestroy hmGrf
        btrDestroy hmAuf
        Exit Sub
    End If
    imUrfRecLen = Len(tmUrf)
        
    hmUst = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUst, "", sgDBPath & "Ust.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmUst)
        ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmAuf)
        btrDestroy hmUst
        btrDestroy hmUrf
        btrDestroy hmGrf
        btrDestroy hmAuf
        Exit Sub
    End If
    imUstRecLen = Len(tmUst)
    
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmUst)
        ilRet = btrClose(hmUrf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmAuf)
        btrDestroy hmCHF
        btrDestroy hmUst
        btrDestroy hmUrf
        btrDestroy hmGrf
        btrDestroy hmAuf
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)
        
'    slEffClearDate = RptSelAL!edcSelCFrom.Text   'Active From Date
    slEffClearDate = RptSelAL!CSI_CalFrom.Text      'active from date 9-9-19 change to use csi calendar vs edit box
    llEffClearDate = gDateValue(slEffClearDate)
    ilAlertContract = gSetCheck(RptSelAL!ckcAlert(0).Value)     'include contract alert
    ilAlertTraffic = gSetCheck(RptSelAL!ckcAlert(1).Value)      'include traffic alert
    ilAlertAffiliate = gSetCheck(RptSelAL!ckcAlert(2).Value)    'include affiliate alert
    ilAlertPool = gSetCheck(RptSelAL!ckcAlert(3).Value)         'include pool alert         Date: 10/15/2019    added pool alert
    
    ilClearContract = gSetCheck(RptSelAL!ckcClear(0).Value)     'include contract clear
    ilClearTraffic = gSetCheck(RptSelAL!ckcClear(1).Value)      'include traffic clear
    ilClearAffiliate = gSetCheck(RptSelAL!ckcClear(2).Value)    'include affiliate clear
    ilClearPool = gSetCheck(RptSelAL!ckcClear(3).Value)         'include pool clear         Date: 10/15/2019    added pool clear

    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime

    ilRet = btrGetFirst(hmAuf, tmAuf, imAufRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)    'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        'filter alert record
        'Include ready alerts for contract, traffic and/or affiliate export
        ilFound = False
        tmGrf.sGenDesc = ""
        If (tmAuf.sStatus = "R") Then
            'Date: 10/15/2019 added pool alert (type = "U"; subtype = "P")
            If ((ilAlertPool And tmAuf.sType = "U" And tmAuf.sSubType = "P") Or (ilAlertContract And tmAuf.sType = "C") Or (ilAlertTraffic And tmAuf.sType = "L") Or ((ilAlertAffiliate And tmAuf.sType = "R") Or (ilAlertAffiliate And tmAuf.sType = "F"))) Then
                ilFound = True
                If tmAuf.iCreateUrfCode > 0 Then
                    tmSrchKey0.iCode = tmAuf.iCreateUrfCode
                    ilRet = btrGetEqual(hmUrf, tmUrf, imUrfRecLen, tmSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        tmGrf.sGenDesc = gDecryptField(Trim$(tmUrf.sName))
                    End If
                Else
                    If tmAuf.iCreateUstCode > 0 Then
                        tmSrchKey0.iCode = tmAuf.iCreateUstCode
                        ilRet = btrGetEqual(hmUst, tmUst, imUstRecLen, tmSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            'tmGrf.sGenDesc = gDecryptField(Trim$(tmUst.sName))
                            tmGrf.sGenDesc = tmUst.sName
                        End If
                    End If
                End If
            End If
        End If
        
        'filter clear records; one of the clear types is selected
        'Date: 10/15/2019 added clear pool
        If ((ilClearContract) Or (ilClearTraffic) Or (ilClearAffiliate) Or (ilClearPool)) Then
                gUnpackDateLong tmAuf.iClearDate(0), tmAuf.iClearDate(1), llAufClearDate
                If llAufClearDate >= llEffClearDate Then
                    If (tmAuf.sStatus = "C") Then
                        'Date: 10/15/2019 added clear pool
                        If ((ilClearPool And tmAuf.sType = "U" And tmAuf.sSubType = "P") Or (ilClearContract And tmAuf.sType = "C") Or (ilClearTraffic And tmAuf.sType = "L") Or ((ilClearAffiliate And tmAuf.sType = "R") Or (ilAlertAffiliate And tmAuf.sType = "F"))) Then
                            ilFound = True
                            If tmAuf.iClearUrfCode > 0 Then
                                tmSrchKey0.iCode = tmAuf.iClearUrfCode
                                ilRet = btrGetEqual(hmUrf, tmUrf, imUrfRecLen, tmSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    tmGrf.sGenDesc = gDecryptField(Trim$(tmUrf.sName))
                                End If
                            Else
                                If tmAuf.iClearUstCode > 0 Then
                                    tmSrchKey0.iCode = tmAuf.iClearUstCode
                                    ilRet = btrGetEqual(hmUst, tmUst, imUstRecLen, tmSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet = BTRV_ERR_NONE Then
                                       ' tmGrf.sGenDesc = gDecryptField(Trim$(tmUst.sName))
                                       tmGrf.sGenDesc = tmUst.sName
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
        End If
        
        If ilFound Then
                'get the User for this entry
                tmGrf.lChfCode = 0          'contract #, not contract code
                If tmAuf.lChfCode > 0 Then
                    tmLongSrchKey0.lCode = tmAuf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmLongSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        tmGrf.lChfCode = tmChf.lCntrNo
                    End If
                End If
                tmGrf.lCode4 = tmAuf.lCode
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                End If
        End If
        ilRet = btrGetNext(hmAuf, tmAuf, imAufRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
    'Cleanup
    ilRet = btrClose(hmAuf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmUrf)
    ilRet = btrClose(hmUst)
    btrDestroy hmAuf
    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmUrf
    btrDestroy hmUst
   
End Sub
