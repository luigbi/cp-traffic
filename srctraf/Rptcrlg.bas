Attribute VB_Name = "RPTCRLG"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrlg.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmSsfSrchKey                                                                          *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptCRGet.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Report Get Data for Crystal screen code
Option Explicit
Option Compare Text
'Public igPdStartDate(0 To 1) As Integer
'Public sgPdType As String * 1
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
'Public igYear As Integer                'budget year used for filtering
'Public igODFGenDate(0 To 1) As Integer  '5-25-01
'Public igODFGenTime(0 To 1) As Integer  '5-25-01
'Public lgGenTime As Long                '10-01-01  Gen time for ODF
'Public lgNowTime As Long        'gen time creating prepass from ODF
Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey1 As SDFKEY1            'SDF record image (key 1)
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSdf As SDF
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVLF As Integer            'Vehicle Link file handle
Dim tmVlf As VLF                'VLF record image
Dim tmVlfSrchKey0 As VLFKEY0            'VLF by selling vehicle record image
Dim imVlfRecLen As Integer        'VLF record length
Dim imSVefCode() As Integer
Dim hmRcf As Integer            'Rate Card file handle
'Log Calendar
'Copy Report
Dim hmCpr As Integer            'Copy Report file handle
Dim tmCpr() As CPR                'CPR record image
Dim imCprRecLen As Integer        'CPR record length
'Copy inventory
Dim hmCif As Integer        'Copy inventory file handle
Dim tmCif As CIF            'CIF record image
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim imCifRecLen As Integer     'CIF record length
' Copy Combo Inventory File
'  Copy Product/Agency File
Dim hmCpf As Integer        'Copy Product/Agency file handle
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0 'CPF key record image
Dim imCpfRecLen As Integer     'CPF record length
' Time Zone Copy FIle
Dim hmTzf As Integer        'Time Zone Copy file handle
Dim tmTzf As TZF            'TZF record image
Dim tmTzfSrchKey As LONGKEY0 'TZF key record image
Dim imTzfRecLen As Integer     'TZF record length
'  Media code File
Dim hmMcf As Integer        'Media file handle
Dim tmMcf As MCF            'MCF record image
Dim tmMcfSrchKey As INTKEY0 'MCF key record image
Dim imMcfRecLen As Integer     'MCF record length

'12-27-04
Dim hmSsf As Integer            'Spot Summary file handle
Dim tmSsf As SSF                'SSF record image
Dim imSsfRecLen As Integer

Dim hmOdf As Integer
Dim tmOdfDay As ODF
Dim imOdfRecLen As Integer
Dim tmOdfSrchKey2 As ODFKEY2
Dim tmOdf0() As ODF

Dim hmCbf As Integer
Dim imCbfRecLen As Integer
Dim tmCbf As CBF

Dim hmMnf As Integer
Dim imMnfRecLen As Integer
Dim tmMnf As MNF
Dim tmMnfSrchKey As INTKEY0

Dim hmSbf As Integer
Dim imSbfRecLen As Integer
Dim tmSbf As SBF
Dim tmSbfSrchKey1 As LONGKEY0

Dim hmRdf As Integer
Dim imRdfRecLen As Integer
Dim tmRdf As RDF
Dim tmRdfSrchKey As INTKEY0

Dim hmCHF As Integer
Dim imCHFRecLen As Integer
Dim tmChf As CHF
Dim tmChfSrchKey As LONGKEY0

'Line record
Dim hmClf As Integer        'Line file
Dim tmClfSrchKey2 As LONGKEY0 'CLF key record image
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF



Dim tmCef As CEF
Type PLAYLISTLG
    sType As String * 1 'Vehicle type
    iVefCode As Integer 'Conventional or Selling Vehicle
    iLogCode As Integer 'Log Vehicle
    iAirCode As Integer 'Selected airing codes for selling vehicle
End Type
Type SKIPTIMES
    ianfCode As Integer
    lSkipStartTime As Long
End Type

'*******************************************************
'*                                                     *
'*      Procedure Name:gCRPlayListGen(ilDoVOFInfo)
'*      <input> ilDoVofInfo = true to obtain VOF header
'               and footer comments and place in CPR
'                           = false to ignore VOF file
'               (it may not be read in)
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Generate Play List Data         *
'*                     for Crystal report
'           6/21/00  MAI logs                          *
'*                                                     *
'*******************************************************
Sub gCRPlayListGenLg(ilDoVofInfo)
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilVehicle As Integer
    Dim ilVefCode As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilRec As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim slVehName As String
    ReDim tlPlayList(0 To 0) As PLAYLISTLG
    Dim tlVef As VEF
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmVLF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imVlfRecLen = Len(tmVlf)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)
    hmCpr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpr, "", sgDBPath & "Cpr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmCpr
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    ReDim tmCpr(0 To 0) As CPR
    imCprRecLen = Len(tmCpr(0))
    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmCif
        btrDestroy hmCpr
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imCifRecLen = Len(tmCif)
    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmCpf
        btrDestroy hmRcf
        btrDestroy hmCpr
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imCpfRecLen = Len(tmCpf)
    hmTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmRcf
        btrDestroy hmCpr
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imTzfRecLen = Len(tmTzf)
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmMcf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmRcf
        btrDestroy hmCpr
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imMcfRecLen = Len(tmMcf)

    '12-27-04 open SSF for airing vehicles to test valid airing day
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "SSf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmMcf)
        ilRet = btrClose(hmTzf)
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmCpr)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVLF)
        ilRet = btrClose(hmVef)
        btrDestroy hmSsf
        btrDestroy hmMcf
        btrDestroy hmTzf
        btrDestroy hmCpf
        btrDestroy hmRcf
        btrDestroy hmCpr
        btrDestroy hmSdf
        btrDestroy hmVLF
        btrDestroy hmVef
        Exit Sub
    End If
    imSsfRecLen = Len(tmSsf)

    slStartDate = RptSelLg!edcSelCFrom.Text
    slEndDate = Format$(gDateValue(slStartDate) + Val(RptSelLg!edcSelCFrom1.Text) - 1, "m/d/yy")
    tmVef.iCode = 0
    tmVefSrchKey.iCode = igcodes(0)         'new scheme only processes one vehicle at a time in this module
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        slVehName = tmVef.sName
       ' ilRet = btrClose(hmVef)
        'btrDestroy hmVef
    Else
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Sub
    End If

    'Loop to build list of vehicles to process:  required for Log vehicles to gather each conventional
    'vehicle for the Log
    If tmVef.sType = "L" Then
        ilRet = btrGetFirst(hmVef, tlVef, imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            If (tlVef.iVefCode = tmVef.iCode) Then
                ilFound = False
                For ilIndex = LBound(tlPlayList) To UBound(tlPlayList) - 1 Step 1
                    If tlPlayList(ilIndex).iVefCode = tlVef.iVefCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilIndex
                If Not ilFound Then
                    tlPlayList(UBound(tlPlayList)).sType = tlVef.sType
                    tlPlayList(UBound(tlPlayList)).iVefCode = tlVef.iCode
                    tlPlayList(UBound(tlPlayList)).iLogCode = tlVef.iVefCode
                    ReDim Preserve tlPlayList(0 To UBound(tlPlayList) + 1) As PLAYLISTLG
                End If
            End If
            ilRet = btrGetNext(hmVef, tlVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop

    ElseIf tmVef.sType = "A" Then
        gBuildLinkArray hmVLF, tmVef, slStartDate, imSVefCode()
        For ilLoop = LBound(imSVefCode) To UBound(imSVefCode) - 1 Step 1
            tlPlayList(UBound(tlPlayList)).sType = "S"
            tlPlayList(UBound(tlPlayList)).iVefCode = imSVefCode(ilLoop)
            tlPlayList(UBound(tlPlayList)).iLogCode = 0
            tlPlayList(UBound(tlPlayList)).iAirCode = tmVef.iCode
            ReDim Preserve tlPlayList(0 To UBound(tlPlayList) + 1) As PLAYLISTLG
        Next ilLoop
        'ilRet = btrGetFirst(hmVef, tlVef, imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        'Do While ilRet = BTRV_ERR_NONE
        '    If tlVef.sType = "S" Then
        '        ilVpfIndex = -1
        '        For ilLoop = 0 To UBound(tgVpf) Step 1
        '            If tlVef.iCode = tgVpf(ilLoop).iVefKCode Then
        '                ilVpfIndex = ilLoop
        '                Exit For
        '            End If
        '        Next ilLoop
        '        If ilVpfIndex >= 0 Then
        '            'For ilVehicle = 0 To UBound(ilCodes) - 1 Step 1
        '                ilVefCode = igCodes(0)
        '                For ilLoop = LBound(tgVpf(ilVpfIndex).iGLink) To UBound(tgVpf(ilVpfIndex).iGLink) Step 1
        '                    If tgVpf(ilVpfIndex).iGLink(ilLoop) > 0 Then
        '                        If tgVpf(ilVpfIndex).iGLink(ilLoop) = ilVefCode Then
        '                            ilFound = False
        '                            For ilIndex = LBound(tlPlayList) To UBound(tlPlayList) - 1 Step 1
        '                                If tlPlayList(ilIndex).iVefCode = tlVef.iVefCode Then
        '                                    For ilAirCode = 1 To 6 Step 1
        '                                        If tlPlayList(ilIndex).iAirCode(ilAirCode) = tgVpf(ilVpfIndex).iGLink(ilLoop) Then
        '                                            ilFound = True
        '                                            Exit For
        '                                        End If
        '                                    Next ilAirCode
        '                                    If Not ilFound Then
        '                                        For ilAirCode = 1 To 6 Step 1
        '                                            If tlPlayList(ilIndex).iAirCode(ilAirCode) = 0 Then
        '                                                tlPlayList(ilIndex).iAirCode(ilAirCode) = tgVpf(ilVpfIndex).iGLink(ilLoop)
        '
        '                                                Exit For
        '                                            End If
        '                                        Next ilAirCode
        '                                    End If
        '                                    ilFound = True
        '                                    Exit For
        '                                End If
        '                            Next ilIndex
        '                            If Not ilFound Then
        '                                tlPlayList(UBound(tlPlayList)).sType = tlVef.sType
        '                                tlPlayList(UBound(tlPlayList)).iVefCode = tlVef.iCode
        '                                tlPlayList(UBound(tlPlayList)).iLogCode = 0
        '                                tlPlayList(UBound(tlPlayList)).iAirCode(1) = tgVpf(ilVpfIndex).iGLink(ilLoop)
        '                                ReDim Preserve tlPlayList(0 To UBound(tlPlayList) + 1) As PLAYLISTLG
        '
        '                            End If
        '                            Exit For
        '                        End If
        '                    End If
        '                Next ilLoop
        '            'Next ilVehicle
        '        End If
        '    End If
        '    ilRet = btrGetNext(hmVef, tlVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        'Loop
    End If

    'ilRet = btrGetFirst(hmVef, tmVef, imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    'Do While ilRet = BTRV_ERR_NONE
    If tmVef.sType = "C" Or tmVef.sType = "G" Then
        'For ilVehicle = 0 To igCodes(0) - 1 Step 1
                ilVefCode = tmVef.iCode
                If (tmVef.iCode = ilVefCode) Or (tmVef.iVefCode = ilVefCode) Then
                    ilFound = False
                    For ilIndex = LBound(tlPlayList) To UBound(tlPlayList) - 1 Step 1
                        If tlPlayList(ilIndex).iVefCode = tmVef.iVefCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilIndex
                    If Not ilFound Then
                        tlPlayList(UBound(tlPlayList)).sType = tmVef.sType
                        tlPlayList(UBound(tlPlayList)).iVefCode = tmVef.iCode
                        tlPlayList(UBound(tlPlayList)).iLogCode = tmVef.iVefCode
                        ReDim Preserve tlPlayList(0 To UBound(tlPlayList) + 1) As PLAYLISTLG
                    End If
                    'Exit For
                End If
        'Next ilVehicle
    'ElseIf tmVef.sType = "S" Then
    '    ilVpfIndex = -1
    '    For ilLoop = 0 To UBound(tgVpf) Step 1
    '        If tmVef.iCode = tgVpf(ilLoop).iVefKCode Then
    '            ilVpfIndex = ilLoop
    '            Exit For
    '        End If
    '    Next ilLoop
    '    If ilVpfIndex >= 0 Then
    '        'For ilVehicle = 0 To igCodes(0) - 1 Step 1
    '                ilVefCode = igCodes(0)          'vehicle code passed from main log screen
    '                For ilLoop = LBound(tgVpf(ilVpfIndex).iGLink) To UBound(tgVpf(ilVpfIndex).iGLink) Step 1
    '                    If tgVpf(ilVpfIndex).iGLink(ilLoop) > 0 Then
    '                        If tgVpf(ilVpfIndex).iGLink(ilLoop) = ilVefCode Then
    '                            ilfound = False
    '                            For ilIndex = LBound(tlPlayList) To UBound(tlPlayList) - 1 Step 1
    '                                If tlPlayList(ilIndex).iVefCode = tmVef.iVefCode Then
    '                                    For ilAirCode = 1 To 6 Step 1
    '                                        If tlPlayList(ilIndex).iAirCode(ilAirCode) = tgVpf(ilVpfIndex).iGLink(ilLoop) Then
    '                                            ilfound = True
    '                                            Exit For
    '                                        End If
    '                                    Next ilAirCode
    '                                    If Not ilfound Then
    '                                        For ilAirCode = 1 To 6 Step 1
    '                                            If tlPlayList(ilIndex).iAirCode(ilAirCode) = 0 Then
    '                                                tlPlayList(ilIndex).iAirCode(ilAirCode) = tgVpf(ilVpfIndex).iGLink(ilLoop)
    '                                                Exit For
    '                                            End If
    '                                        Next ilAirCode
    '                                    End If
    '                                    ilfound = True
    '                                    Exit For
    '                                End If
    '                            Next ilIndex
    '                            If Not ilfound Then
    '                                tlPlayList(UBound(tlPlayList)).sType = tmVef.sType
    '                                tlPlayList(UBound(tlPlayList)).iVefCode = tmVef.iCode
    '                                tlPlayList(UBound(tlPlayList)).iLogCode = 0
    '                                tlPlayList(UBound(tlPlayList)).iAirCode(1) = tgVpf(ilVpfIndex).iGLink(ilLoop)
    '                                ReDim Preserve tlPlayList(0 To UBound(tlPlayList) + 1) As PLAYLISTLG
    '                            End If
    '                            Exit For
    '                        End If
    '                    End If
    '                Next ilLoop
    '        'Next ilVehicle
    '    End If
    End If
    'ilRet = btrGetNext(hmVef, tmVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    'Loop
    'Conventional Vehicles without log
    For ilVehicle = 0 To UBound(tlPlayList) - 1 Step 1
        If (tlPlayList(ilVehicle).sType = "C" Or tlPlayList(ilVehicle).sType = "G") And (tlPlayList(ilVehicle).iLogCode = 0) Then
            ilVefCode = tlPlayList(ilVehicle).iVefCode
            ReDim tmCpr(0 To 0) As CPR
            mGetPlayList ilVefCode, slStartDate, slEndDate, tlPlayList(ilVehicle)
            'Output records
            For ilRec = 0 To UBound(tmCpr) - 1 Step 1
                tmCpr(ilRec).lHd1CefCode = 0
                tmCpr(ilRec).lFt1CefCode = 0
                tmCpr(ilRec).lFt2CefCode = 0
                If ilDoVofInfo Then
                    tmCpr(ilRec).lHd1CefCode = tgVof.lHd1CefCode
                    tmCpr(ilRec).lFt1CefCode = tgVof.lFt1CefCode
                    tmCpr(ilRec).lFt2CefCode = tgVof.lFt2CefCode
                End If
                ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
            Next ilRec
        End If
    Next ilVehicle
    'Conventional Vehicle with Log
    ReDim tmCpr(0 To 0) As CPR
    For ilVehicle = 0 To UBound(tlPlayList) - 1 Step 1
        If (tlPlayList(ilVehicle).sType = "C") And (tlPlayList(ilVehicle).iLogCode <> 0) Then
            ilVefCode = tlPlayList(ilVehicle).iVefCode
            mGetPlayList ilVefCode, slStartDate, slEndDate, tlPlayList(ilVehicle)
        End If
    Next ilVehicle
    'Output records
    For ilRec = 0 To UBound(tmCpr) - 1 Step 1
        tmCpr(ilRec).lHd1CefCode = 0
        tmCpr(ilRec).lFt1CefCode = 0
        tmCpr(ilRec).lFt2CefCode = 0
        If ilDoVofInfo Then
            tmCpr(ilRec).lHd1CefCode = tgVof.lHd1CefCode
            tmCpr(ilRec).lFt1CefCode = tgVof.lFt1CefCode
            tmCpr(ilRec).lFt2CefCode = tgVof.lFt2CefCode
        End If
        ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
    Next ilRec
    'Selling Vehicles
    ReDim tmCpr(0 To 0) As CPR
    For ilVehicle = 0 To UBound(tlPlayList) - 1 Step 1
        If (tlPlayList(ilVehicle).sType = "S") Then
            ilVefCode = tlPlayList(ilVehicle).iVefCode
            mGetPlayList ilVefCode, slStartDate, slEndDate, tlPlayList(ilVehicle)
        End If
    Next ilVehicle
    'Output records
    For ilRec = 0 To UBound(tmCpr) - 1 Step 1
        tmCpr(ilRec).lHd1CefCode = 0
        tmCpr(ilRec).lFt1CefCode = 0
        tmCpr(ilRec).lFt2CefCode = 0
        If ilDoVofInfo Then
            tmCpr(ilRec).lHd1CefCode = tgVof.lHd1CefCode
            tmCpr(ilRec).lFt1CefCode = tgVof.lFt1CefCode
            tmCpr(ilRec).lFt2CefCode = tgVof.lFt2CefCode
        End If
        ilRet = btrInsert(hmCpr, tmCpr(ilRec), imCprRecLen, INDEXKEY0)
    Next ilRec
    Erase tmCpr
    Erase tlPlayList
    ilRet = btrClose(hmMcf)
    ilRet = btrClose(hmTzf)
    ilRet = btrClose(hmCpf)
    ilRet = btrClose(hmCif)
    ilRet = btrClose(hmCpr)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVLF)
    ilRet = btrClose(hmVef)
    btrDestroy hmMcf
    btrDestroy hmTzf
    btrDestroy hmCpf
    btrDestroy hmRcf
    btrDestroy hmCpr
    btrDestroy hmSdf
    btrDestroy hmVLF
    btrDestroy hmVef
    Exit Sub
End Sub
'
'
'                  Create a log that skips to a new page but
'                   prints the entire week (Monday-Friday)
'                   up thu the page skip on a page before
'                   skipping to new page
'
'                   gL14pageSkips ilDaysToDo - # days to process
'                   Created:  10/29/98
'
'                   1/29/99 - page skipping seq # corrected when
'                   there are multiple skips & named avails
'                   5/28/99 - fix page skipping (increments of seq. #s)
'                   when Monday doesnt exist.
'                   01-10-01 fix page skipping when avails are not the
'                   M-F: i.e.
'                       M/Tu avail page skip at 9A, first avail @ 5:30P
'                       W-Fi avail page skip at 9A, first avail @ 4:30P
'                       The 4:30P spots were not picking up that it needed skip
Sub gL14PageSkips(ilDaysToDo As Integer)
Dim ilRet As Integer
Dim llStartDate As Long             'user requested dates to genrate
Dim llEndDate As Long               'user requeted dates to generate
Dim llStartOfWk As Long
ReDim ilStartofWk(0 To 1) As Integer
Dim llStartTime As Long             'user requeted times to generate
Dim llEndTime As Long               'user requested times to generate
Dim ilVefCode As Integer            'vehicle to generate
Dim hlODF As Integer                'ODF handle
Dim ilSequence As Integer
Dim ilZoneLoop As Integer
Dim ilLoZone As Integer
Dim ilHiZone As Integer
Dim ilVpfIndex As Integer
Dim ilUseZone As Integer
Dim ilZones As Integer
Dim slZone As String
Dim llLoopDate As Long
Dim ilDayOfWeek As Integer
Dim ilFound As Integer
Dim ilLoop As Integer
Dim llOdfDate As Long           'converted time from Odf (from btrieve)
Dim llOdfTime As Long           'converted date from ODF (from btrieve)
Dim tlOdf As ODF
Dim tlOdfSrchKey As ODFKEY0            'ODF record image
Dim llNextStartTime As Long         'time of next skip time for named avail processing
Dim ilNext As Integer
Dim ilGotFirstDay As Integer
Dim ilFirstDay As Integer
    hlODF = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlODF, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlODF)
        btrDestroy hlODF
        Exit Sub
    End If
    llStartDate = gDateValue(sgLogStartDate)
    'llEndDate = llStartDate + Val(sgLogNoDays) - 1
    llEndDate = llStartDate + ilDaysToDo - 1
    llStartOfWk = llStartDate                       'get to Monday start week
    Do While (gWeekDayLong(llStartOfWk)) <> 0
        llStartOfWk = llStartOfWk - 1
    Loop
    'convert to btrieve format
    gPackDateLong llStartOfWk, ilStartofWk(0), ilStartofWk(1)
    llStartTime = CLng(gTimeToCurrency(sgLogStartTime, False))
    llEndTime = CLng(gTimeToCurrency(sgLogEndTime, True)) - 1
    ilVefCode = igcodes(0)                                  'passed for log function
    ilVpfIndex = -1
    ilVpfIndex = gVpfFind(RptSelLg, ilVefCode)              'determine vehicle options index

    ilUseZone = False                                       'assume not using zones until one is found in the vehicle options table
    ilZones = igZones                                       'save zones requested by user : 0=none, 1 =est, 2= cst, 3 =mst, 4 = pst
    ilLoZone = 1                                            'low loop factor to process zones
    ilHiZone = 4                                            'hi loop factor to process zones
    If ilZones <> 0 Then                                    'user has requested one zone in particular
        ilLoZone = ilZones
        ilHiZone = ilZones
    End If
    If ilVpfIndex >= 0 Then                                  'associated vehicle options record exists
        'If tgVpf(ilVpfIndex).sGZone(1) <> "   " Then
        If tgVpf(ilVpfIndex).sGZone(0) <> "   " Then
            ilUseZone = True
        Else
            'Zones not used, fake out flag to do 1 zone (EST)
            ilZones = 1
            ilLoZone = 1
            ilHiZone = 1
        End If
    Else
        'no vehicle options table
    End If

    ReDim tlSkip(0 To 0) As SKIPTIMES
    ilGotFirstDay = False
    For llLoopDate = llStartDate To llEndDate
        ilDayOfWeek = gWeekDayLong(llLoopDate)      '0=Mon
        'If ilDayofWeek < 5 Then                     'only process M-F  (5 = sa, 6 = su)
            For ilZoneLoop = ilLoZone To ilHiZone               'loop on all time zones (or just the selective one,  variety of zones are not allowed)
                Select Case ilZoneLoop
                    Case 1  'Eastern
                        If ilUseZone Then
                            slZone = "EST"
                        Else
                            slZone = "   "
                        End If
                    Case 2  'Central
                        slZone = "CST"
                    Case 3  'Mountain
                        slZone = "MST"
                    Case 4  'Pacific
                        slZone = "PST"
                End Select
            Next ilZoneLoop
        'End If                          'ildayofWeek

        'Access the ODF records and build table of the start time for the page skips for
        'the first day only.  As its building the start time for the page skips,
        'keep a sequential # to use it as amajor  sort key in Crystal.
        tlOdfSrchKey.iVefCode = ilVefCode
        gPackDateLong llLoopDate, tlOdfSrchKey.iAirDate(0), tlOdfSrchKey.iAirDate(1)
        gPackTimeLong llStartTime, tlOdfSrchKey.iLocalTime(0), tlOdfSrchKey.iLocalTime(1)
        tlOdfSrchKey.sZone = slZone
        tlOdfSrchKey.iSeqNo = 0
        ilRet = btrGetGreaterOrEqual(hlODF, tlOdf, Len(tlOdf), tlOdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)  '
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hlODF)
            btrDestroy hlODF
            Erase tlSkip
            Exit Sub
        Else
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hlODF)
                btrDestroy hlODF
                Erase tlSkip
                Exit Sub
            End If
        End If
        gUnpackDateLong tlOdf.iAirDate(0), tlOdf.iAirDate(1), llOdfDate
        gUnpackTimeLong tlOdf.iAirTime(0), tlOdf.iAirTime(1), False, llOdfTime
        Do While (ilRet = BTRV_ERR_NONE) And (ilVefCode = tlOdf.iVefCode) And (Trim$(tlOdf.sZone) = Trim$(slZone)) And (llOdfDate = llLoopDate) And (llOdfTime >= llStartTime And llOdfTime <= llEndTime)
            'If tlOdf.iEtfCode = 0 Then                  '01-10-01 get spots only
                'If tlOdf.sPageEjectFlag = "Y" Then     '01-10-01
                If tlOdf.iEtfCode = 10 Then                 '01-10-01 page skip
                    If ilDayOfWeek <> 0 Then
                        If Not ilGotFirstDay Then
                            ilGotFirstDay = True
                            ilFirstDay = ilDayOfWeek
                            tlSkip(UBound(tlSkip)).ianfCode = tlOdf.ianfCode
                            tlSkip(UBound(tlSkip)).lSkipStartTime = llOdfTime
                            ReDim Preserve tlSkip(0 To UBound(tlSkip) + 1) As SKIPTIMES
                        Else
                            If ilFirstDay = ilDayOfWeek Then
                                tlSkip(UBound(tlSkip)).ianfCode = tlOdf.ianfCode
                                tlSkip(UBound(tlSkip)).lSkipStartTime = llOdfTime
                                ReDim Preserve tlSkip(0 To UBound(tlSkip) + 1) As SKIPTIMES
                            End If
                        End If
                    Else
                'If tlOdf.sPageEjectFlag = "Y" And ilDayofWeek = 0 Then
                        ilGotFirstDay = True
                        tlSkip(UBound(tlSkip)).ianfCode = tlOdf.ianfCode
                        tlSkip(UBound(tlSkip)).lSkipStartTime = llOdfTime
                        ReDim Preserve tlSkip(0 To UBound(tlSkip) + 1) As SKIPTIMES
                    End If
                'End If             '01-10-01
                ElseIf tlOdf.iEtfCode = 0 Then             '1-10-01   spot
                'Determine Sequence # for this named avail
                ilSequence = 0
                ilFound = False
                'If ilDayofWeek = 0 Then
                    For ilLoop = LBound(tlSkip) To UBound(tlSkip)
                        If tlSkip(ilLoop).ianfCode = tlOdf.ianfCode Then
                            ilSequence = ilSequence + 1
                            'If ilLoop + 1 > UBound(tlSkip) - 1 Then
                                llNextStartTime = 86400     '12m end of day
                                If ilLoop + 1 <= UBound(tlSkip) Then
                                    For ilNext = ilLoop + 1 To UBound(tlSkip) - 1
                                        If tlSkip(ilNext).ianfCode = tlOdf.ianfCode Then
                                            llNextStartTime = tlSkip(ilNext).lSkipStartTime
                                            Exit For
                                        End If
                                    Next ilNext
                                    If llOdfTime >= tlSkip(ilLoop).lSkipStartTime And llOdfTime < llNextStartTime Then
                                        ilFound = True
                                        Exit For
                                    End If
                                End If
                                'ilFound = True
                                'Exit For
                            'Else
                            '    If llOdfTime >= tlSkip(ilLoop).lSkipStartTime And llOdfTime < tlSkip(ilLoop + 1).lSkipStartTime Then
                            '        ilFound = True
                            '        Exit For
                            '    Else
                            '        ilSequence = ilSequence + 1
                            '    End If
                            'End If
                        End If
                    Next ilLoop
                'Else
                '    ReDim llAltStartTimes(0 To 0) As Long
                '    llAltStartTimes(0) = llStartTime            'start time of day
                '    ReDim Preserve llAltStartTimes(0 To 1) As Long
                '    ilFound = False
                '    For ilLoop = LBound(tlSkip) To UBound(tlSkip)   'find all the other start times of page skips
                '        If tlSkip(ilLoop).iAnfCode = tlOdf.iAnfCode Then
                '            llAltStartTimes(UBound(llAltStartTimes)) = tlSkip(ilLoop).lSkipStartTime
                '            ReDim Preserve llAltStartTimes(0 To UBound(llAltStartTimes) + 1) As Long
                '        End If
                '    Next ilLoop
                '    llAltStartTimes(UBound(llAltStartTimes)) = llEndTime
                '    ReDim Preserve llAltStartTimes(0 To UBound(llAltStartTimes) + 1) As Long
                '    ilSequence = 0
                '    For ilLoop = 0 To UBound(llAltStartTimes) - 1
                '        If llOdfTime >= llAltStartTimes(ilLoop) And llOdfTime < llAltStartTimes(ilLoop + 1) Then
                '            ilFound = True
                '            ilSequence = ilLoop
                '            Exit For
                '        End If
                '    Next ilLoop
                'End If
                End If                      '01-10-01
                If Not ilFound Then             '
                    ilSequence = 0
                End If
                'Write record back to disk
                tlOdf.iSortSeq = ilSequence
                ilRet = btrUpdate(hlODF, tlOdf, Len(tlOdf))
                If ilRet <> BTRV_ERR_NONE Then
                    ilRet = btrClose(hlODF)
                    btrDestroy hlODF
                    Erase tlSkip
                    Exit Sub
                End If
            'End If                      '01-10-01 if tlodf.ietfcode = 0

            ilRet = btrGetNext(hlODF, tlOdf, Len(tlOdf), BTRV_LOCK_NONE, SETFORWRITE)
            gUnpackDateLong tlOdf.iAirDate(0), tlOdf.iAirDate(1), llOdfDate
            gUnpackTimeLong tlOdf.iAirTime(0), tlOdf.iAirTime(1), False, llOdfTime
        Loop                    'while ilret = btrv_err_none
    Next llLoopDate                     'for llLoopdate = llStartDate tollEndDate
    ilRet = btrClose(hlODF)
    btrDestroy hlODF
    Erase tlSkip
End Sub
'
'
'               L29 and C19 (C19 requires the DP times to be placed
'               back in ODF field tmOdf.sDPDesc. The DP times are
'               based on the vehicle options Interface table.  The
'               start time of day is assumed to be 6a)
'
'               1/6/00 Code added for C19 to place DP start/end times in ODF record
'               5-30-01 L71 log was sorting the DP sections by the ASCII string, and
'                   not by the true times (i.e. 10A-3P printed before 6A-10A)
'                   Insert the DP index into ODF for sorting in Crystal
'               5-20-04 assume start of day is 12M, not 6A for first DP found
'
Sub gSetSeqL29()
Dim hlODF As Integer
Dim hlCef As Integer
Dim ilRet As Integer
Dim slDate As String
Dim slLen As String
Dim ilUseZone As Integer            'true if at least one zone defined
Dim slZone As String * 3            'EST, CST, MST, PST or blank
Dim ilZoneLoop As Integer           'time zone to process in loop
Dim ilDay As Integer                'day of week to process
Dim llWeek As Long               'week to process
Dim ilDayOfWeek As Integer
Dim llStartOfWk As Long             'start date of week (temp)
ReDim ilStartofWk(0 To 1) As Integer  'start date of week btrieve format
Dim ilRec As Integer                'event within week to write to disk
Dim ilOdf As Integer                'event to process for day from ODF table
Dim ilVpfIndex As Integer           'vehicle options index
Dim ilZones As Integer              'zones requested by user
Dim ilLoZone As Integer             'lo limit loop factor to process zones.  If only 1 zone, this #
                                    'will be same as ilHizone.  if all Zones, this will be 1, and ilHIzones will be 4
Dim ilHiZone As Integer             'Hi limit loop factor to process zones.  If only 1 zone, this #
                                    'will be same as ilLozone.  If all zones, this will be a 4.
'ReDim llZoneEndTimes(1 To 6) As Long          'First field is always 0, remaining extracted for vpf (estendtime, cstendtime, mstendtime, & pstendtime)
Dim ilSeqNo As Integer              'running seq # to stored in each event (all avails get the same seq # if back-to back only)
Dim il1stAvailofBreak As Integer
Dim llPrevEndTime As Long
Dim llCurrAvailTime As Long
Dim slAvailSTime As String
Dim slAvailETime As String
Dim tlodfExtSrchKey As ODFKEY0
Dim tlOdf As ODF
Dim slComlID As String
Dim slCurrentID As String
'ReDim llZoneEndTimes(1 To 6) As Long          'First field is always 0, remaining extracted for vpf (estendtime, cstendtime, mstendtime, & pstendtime)
ReDim llZoneEndTimes(0 To 6) As Long          'First field is always 0, remaining extracted for vpf (estendtime, cstendtime, mstendtime, & pstendtime). Index zero ignored
Dim ilDPIndex As Integer                   'DP index used to separate the DP on crystal report (stored in svr)
Dim llOdfTime As Long
Dim llODFPosition() As Long
'**** parameters passed from Log program
Dim llStartTime As Long             'start time of log gen
Dim llEndTime As Long               'end time of log gen
Dim llStartDate As Long             'start date of log gen
Dim llEndDate As Long               'end date of log gen
Dim ilUrfCode As Integer           'user code requesting log
Dim ilVefCode As Integer            ' vehicle to process
Dim tlSrchKey As LONGKEY0    'Rcf key record image
Dim slXMid As String
'**** end parameters passed from Log program
    hlODF = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlODF, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlODF)
        btrDestroy hlODF
        Exit Sub
    End If
    hlCef = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlCef)
        btrDestroy hlCef
        Exit Sub
    End If

    llStartDate = gDateValue(sgLogStartDate)
    llEndDate = llStartDate + Val(sgLogNoDays) - 1
    llStartOfWk = llStartDate                       'get to Monday start week
    Do While (gWeekDayLong(llStartOfWk)) <> 0
        llStartOfWk = llStartOfWk - 1
    Loop
    'convert to btrieve format
    gPackDateLong llStartOfWk, ilStartofWk(0), ilStartofWk(1)
    llStartTime = CLng(gTimeToCurrency(sgLogStartTime, False))
    llEndTime = CLng(gTimeToCurrency(sgLogEndTime, True)) - 1
    ilUrfCode = Val(sgLogUserCode)                         'user code requesting log
    ilVefCode = igcodes(0)                                  'passed for log function
    ilVpfIndex = -1
    ilVpfIndex = gVpfFind(RptSelLg, ilVefCode)              'determine vehicle options index

    ilUseZone = False                                       'assume not using zones until one is found in the vehicle options table
    ilZones = igZones                                       'save zones requested by user : 0=none, 1 =est, 2= cst, 3 =mst, 4 = pst
    ilLoZone = 1                                            'low loop factor to process zones
    ilHiZone = 4                                            'hi loop factor to process zones
    If ilZones <> 0 Then                                    'user has requested one zone in particular
        ilLoZone = ilZones
        ilHiZone = ilZones
    End If
    If ilVpfIndex >= 0 Then                                  'associated vehicle options record exists
        'If tgVpf(ilVpfIndex).sGZone(1) <> "   " Then
        If tgVpf(ilVpfIndex).sGZone(0) <> "   " Then
            ilUseZone = True
        Else
            'Zones not used, fake out flag to do 1 zone (EST)
            ilZones = 1
            ilLoZone = 1
            ilHiZone = 1
        End If
    Else
        'no vehicle options table
    End If
    For ilZoneLoop = ilLoZone To ilHiZone               'loop on all time zones (or just the selective one,  variety of zones are not allowed)
        Select Case ilZoneLoop
            Case 1  'Eastern
                For ilRec = 1 To 5
                    llZoneEndTimes(ilRec + 1) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(ilRec - 1))
                Next ilRec
                If ilUseZone Then
                    slZone = "EST"
                Else
                    slZone = "   "
                End If
            Case 2  'Central
                For ilRec = 1 To 5
                    llZoneEndTimes(ilRec + 1) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(ilRec - 1))
                Next ilRec
                slZone = "CST"
            Case 3  'Mountain
                For ilRec = 1 To 5
                    llZoneEndTimes(ilRec + 1) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(ilRec - 1))
                Next ilRec
                slZone = "MST"
            Case 4  'Pacific
                For ilRec = 1 To 5
                    llZoneEndTimes(ilRec + 1) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(ilRec - 1))
                Next ilRec
                slZone = "PST"
        End Select

        For llWeek = llStartOfWk To llEndDate Step 7    '6-20-00 loop on weeks
        llStartDate = llWeek
        For ilDay = 0 To 6                                'loop on all days of the week
            slDate = Format$(llStartDate + ilDay, "m/d/yy")
            ilDayOfWeek = gWeekDayLong(llStartDate + ilDay)
            Select Case ilDay
                Case 0
                    mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
                Case 1
                    mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
                Case 2
                    mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
                Case 3
                    mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
                Case 4
                    mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
                Case 5
                    mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
                Case 6
                    mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
            End Select
            'Loop thru each event and build a unique record based on zone, vehicle, event type, time & position
            'for the entire week
            ilSeqNo = 0
            il1stAvailofBreak = True
            For ilOdf = LBound(tmOdf0) To UBound(tmOdf0) - 1 Step 1
                If tmOdf0(ilOdf).sZone = slZone Then        '6-9-10 need to filter out by time zone
                    If tmOdf0(ilOdf).iType = 4 Then         'spot, test for all spots that are back-to back.
                                                            'these spots will all maintain the same seq # (used for
                                                            'grouping in Crystal report
    
                        If il1stAvailofBreak Then
                            il1stAvailofBreak = False
                            ilSeqNo = ilSeqNo + 1
                            tmOdf0(ilOdf).iSortSeq = ilSeqNo
                            'gUnpackTimeLong tmOdf0(ilOdf).iAirTime(0), tmOdf0(ilOdf).iAirTime(1), False, llSTimeofBreak
                            gUnpackTime tmOdf0(ilOdf).iAirTime(0), tmOdf0(ilOdf).iAirTime(1), "A", "1", slAvailSTime
                            gUnpackLength tmOdf0(ilOdf).iLen(0), tmOdf0(ilOdf).iLen(1), "3", False, slLen
                            gAddTimeLength slAvailSTime, slLen, "A", "1", slAvailETime, slXMid
                            llPrevEndTime = gTimeToLong(slAvailETime, True)
                            slCurrentID = ""
                            slComlID = mComlID(hlCef, tmOdf0(ilOdf).lAvailcefCode, slCurrentID)
                            tmOdf0(ilOdf).sShortTitle = Trim$(slComlID)
                        Else
                            'not first spot of break
                            gUnpackTimeLong tmOdf0(ilOdf).iAirTime(0), tmOdf0(ilOdf).iAirTime(1), False, llCurrAvailTime
                            If llPrevEndTime < llCurrAvailTime Then    'new break, not back-to-back spots
                                'il1stAvailofBreak = True
                                ilSeqNo = ilSeqNo + 1
                                tmOdf0(ilOdf).iSortSeq = ilSeqNo
                                gUnpackTime tmOdf0(ilOdf).iAirTime(0), tmOdf0(ilOdf).iAirTime(1), "A", "1", slAvailSTime
                                gUnpackLength tmOdf0(ilOdf).iLen(0), tmOdf0(ilOdf).iLen(1), "3", False, slLen
                                gAddTimeLength slAvailSTime, slLen, "A", "1", slAvailETime, slXMid
                                llPrevEndTime = gTimeToLong(slAvailETime, True)
                                slCurrentID = ""
                                slComlID = mComlID(hlCef, tmOdf0(ilOdf).lAvailcefCode, slCurrentID)
                                tmOdf0(ilOdf).sShortTitle = Trim$(slComlID)
                            Else                                        'back-toback spots (same break)
                                tmOdf0(ilOdf).iSortSeq = ilSeqNo
                                'keep running time of events in the same break
                                'gUnpackTime tmOdf0(ilOdf).iAirTime(0), tmOdf0(ilOdf).iAirTime(1), "A", "1", slAvailSTime
                                slAvailSTime = gFormatTimeLong(llPrevEndTime, "A", "1")
                                gUnpackLength tmOdf0(ilOdf).iLen(0), tmOdf0(ilOdf).iLen(1), "3", False, slLen
                                gAddTimeLength slAvailSTime, slLen, "A", "1", slAvailETime, slXMid
                                llPrevEndTime = gTimeToLong(slAvailETime, True)
                                slCurrentID = slComlID
                                slComlID = mComlID(hlCef, tmOdf0(ilOdf).lAvailcefCode, slCurrentID)
                                tmOdf0(ilOdf).sShortTitle = Trim$(slComlID)
                            End If
                        End If
                        '9-11-00 If sgRnfRptName = "C19" Or sgRnfRptName = "L44" Then         'replace field containing DP description of R/c with DP times in vehicle options table
                        If sgRnfRptName = "C19" Or sgRnfRptName = "L71" Then         '9-11-00 replace field containing DP description of R/c with DP times in vehicle options table
                            'determine which daypart this spot belongs in based on Vehicle Options Interface table
                            gUnpackTimeLong tmOdf0(ilOdf).iAirTime(0), tmOdf0(ilOdf).iAirTime(1), False, llOdfTime   'dont convert 12m to end of day
                            For ilDPIndex = 1 To 5
                                If llOdfTime >= llZoneEndTimes(ilDPIndex) And llOdfTime < llZoneEndTimes(ilDPIndex + 1) Then
                                    Exit For
                                End If
                            Next ilDPIndex
                            'setup the DP text
                            tmOdf0(ilOdf).sDPDesc = ""
                            tmOdf0(ilOdf).iDPSort = 0      '5-30-01
                            If llZoneEndTimes(ilDPIndex) = 0 Then
                                tmOdf0(ilOdf).sDPDesc = "12M"       '5-20-04 chg start of day to 12m vs. "6A"
                                tmOdf0(ilOdf).iDPSort = ilDPIndex  '5-30-01
                            ElseIf llZoneEndTimes(ilDPIndex) = 86400 Then
                                tmOdf0(ilOdf).sDPDesc = ""       'found last entry whose end time is already 12m
                                tmOdf0(ilOdf).iDPSort = ilDPIndex  '5-30-1
                            Else
                                tmOdf0(ilOdf).sDPDesc = gFormatTimeLong(llZoneEndTimes(ilDPIndex), "A", "1")
                                tmOdf0(ilOdf).iDPSort = ilDPIndex  '5-30-01
                            End If
                            If ilDPIndex <= 5 Then          'valid entry to determine end time since there is only 6 entries in options table
                                tmOdf0(ilOdf).sDPDesc = Trim$(tmOdf0(ilOdf).sDPDesc) & "-" & gFormatTimeLong(llZoneEndTimes(ilDPIndex + 1), "A", "1")
                            End If
                        ElseIf sgRnfRptName = "L74" Then    '12-17-02 this log requires sorting by named avail comment description.  Crystal
                                                            'does not allow a memo field (avail comment) to be in a formula, and this field
                                                            'is one of the major sort fields.  Place 20 charac in the DPdescr field to use that as sort field
                            If tmOdf0(ilOdf).lAvailcefCode > 0 Then       'get the comment for this avail
                                tlSrchKey.lCode = tmOdf0(ilOdf).lAvailcefCode
                                ilRet = btrGetEqual(hlCef, tmCef, Len(tmCef), tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    tmOdf0(ilOdf).sDPDesc = Trim$(tmCef.sComment)
                                Else
                                    tmOdf0(ilOdf).sDPDesc = ""
                                End If
                            Else
                                tmOdf0(ilOdf).sDPDesc = ""   'if no avail comment, clear out the DP descr so it wont sort improperly
                            End If
                        End If
                    ElseIf tmOdf0(ilOdf).iType = 2 Then
                        il1stAvailofBreak = True
                        ilSeqNo = ilSeqNo + 1
                        tmOdf0(ilOdf).iSortSeq = ilSeqNo
                    Else
                        ilSeqNo = ilSeqNo + 1
                        tmOdf0(ilOdf).iSortSeq = ilSeqNo
                    End If
                End If
            Next ilOdf


            'go thru the odf records in memory and write back to disk withupdated seq #
            'which will be used as the major sort field in L29 log
            tlodfExtSrchKey.iVefCode = ilVefCode
            gPackDateLong llStartDate + ilDay, tlodfExtSrchKey.iAirDate(0), tlodfExtSrchKey.iAirDate(1)
            For ilOdf = LBound(tmOdf0) To UBound(tmOdf0) - 1 Step 1

                tlodfExtSrchKey.sZone = tmOdf0(ilOdf).sZone
                tlodfExtSrchKey.iSeqNo = tmOdf0(ilOdf).iSeqNo
                tlodfExtSrchKey.iLocalTime(0) = tmOdf0(ilOdf).iLocalTime(0)
                tlodfExtSrchKey.iLocalTime(1) = tmOdf0(ilOdf).iLocalTime(1)
                ilRet = btrGetEqual(hlODF, tlOdf, Len(tmOdf0(1)), tlodfExtSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    ilRet = btrClose(hlODF)
                    ilRet = btrClose(hlCef)
                    btrDestroy hlODF
                    btrDestroy hlCef
                    Erase tmOdf0
                    Exit Sub
                End If

                'ilRet = btrUpdate(hlODF, tmOdf0(ilOdf), Len(tmOdf0(1)))
                ilRet = btrUpdate(hlODF, tmOdf0(ilOdf), Len(tmOdf0(0)))
                If ilRet <> BTRV_ERR_NONE Then
                    ilRet = btrClose(hlODF)
                    ilRet = btrClose(hlCef)
                    btrDestroy hlODF
                    btrDestroy hlCef
                    Erase tmOdf0
                    Exit Sub
                End If
            Next ilOdf

            ReDim tmOdf0(0 To 0) As ODF
        Next ilDay
        Next llWeek
    Next ilZoneLoop                                         'for ilzone
    ilRet = btrClose(hlODF)
    ilRet = btrClose(hlCef)
    btrDestroy hlODF
    btrDestroy hlCef
    Erase tmOdf0
End Sub
'
'
'       Setup commercial ID for C14 Certificate of Performance
'       slComlID = mComlID ( hlCef as integer,slCurrentID as string)
'
'       7/11/99 D Hosaka
'       <input> slCurrentID - blank if first spot in avail (retrieve the
'                       comml ID from the CEF file, otherwise
'                            increment the previous Comml ID to next
'                            letter (ie. SN003A, next spot is SN003B, etc)
'               llCefCode - Avail Comments code
'               hlCef - Comments handle
'       <output> slComlID - ID to place in spot
'
Function mComlID(hlCef As Integer, llCefCode As Long, slCurrentID As String) As String
Dim ilRet As Integer
Dim ilLen As Integer
Dim ilLenTest As Integer
Dim ilAsc As Integer
Dim slLastChar As String * 1
Dim slComlID As String
Dim ilAllNumeric As Integer
Dim tlSrchKey As LONGKEY0    'Rcf key record image
    If sgRnfRptName = "C14" Then        'CP C14 only
        slComlID = slCurrentID
        If slComlID = "" Then          '1st time of avail or the avail doesnt have IDs
            If llCefCode > 0 Then       'get the comment for this avail
                tlSrchKey.lCode = llCefCode
                ilRet = btrGetEqual(hlCef, tmCef, Len(tmCef), tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    'slComlID = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
                    slComlID = gStripChr0(tmCef.sComment)
                End If
            End If
        Else                            'not 1st time of avail, keep running increments of coml id #
                                        'for Sheridan, its alphabetic
            ilLen = Len(slCurrentID)
            ilAllNumeric = True
            For ilLenTest = 1 To ilLen
                If Asc(Mid$(slCurrentID, ilLenTest, 1)) < 48 Or Asc(Mid$(slCurrentID, ilLenTest, 1)) > 57 Then
                    ilAllNumeric = False
                    Exit For
                End If
            Next ilLenTest
            If ilAllNumeric Then
                slComlID = str$(Val(slCurrentID) + 1)
            Else
                slLastChar = Mid$(slCurrentID, ilLen, 1)
                ilAsc = Asc(slLastChar)
                ilAsc = ilAsc + 1
                slLastChar = Chr$(ilAsc)
                Mid$(slComlID, ilLen, 1) = slLastChar
            End If
        End If
        mComlID = slComlID              'return the commercial ID to store in spot event
    End If
End Function
'
'
'               mRunAirTimes - make air times running end times rather than spots within
'               same break with same start times.  ie. 7:10a 30", 7:10a 30", 7:10a 30  --
'               change to 7:10a 30", 7:10:30a 30", 7:11a 30"
'
'               6/21/00 For logs that are monday thru fri/sun across
'
'               dh 10-24-00 Need to update the VOF header / comment pointers into ODF for C74/C75/C76/c77/C79
'               dh 12-01-00 When log run for more than 1 week, the vehicle vof information was not updated into ODF
Sub gFixAirTimes()
Dim hlODF As Integer
Dim hlCef As Integer
Dim ilRet As Integer
Dim slDate As String
Dim slLen As String
Dim ilUseZone As Integer            'true if at least one zone defined
Dim slZone As String * 3            'EST, CST, MST, PST or blank
Dim ilZoneLoop As Integer           'time zone to process in loop
Dim ilDay As Integer                'day of week to process
Dim ilDayOfWeek As Integer
Dim llStartOfWk As Long             'start date of week (temp)
ReDim ilStartofWk(0 To 1) As Integer  'start date of week btrieve format
Dim ilRec As Integer                'event within week to write to disk
Dim ilOdf As Integer                'event to process for day from ODF table
Dim ilVpfIndex As Integer           'vehicle options index
Dim ilZones As Integer              'zones requested by user
Dim ilLoZone As Integer             'lo limit loop factor to process zones.  If only 1 zone, this #
                                    'will be same as ilHizone.  if all Zones, this will be 1, and ilHIzones will be 4
Dim ilHiZone As Integer             'Hi limit loop factor to process zones.  If only 1 zone, this #
                                    'will be same as ilLozone.  If all zones, this will be a 4.
'ReDim llZoneEndTimes(1 To 6) As Long          'First field is always 0, remaining extracted for vpf (estendtime, cstendtime, mstendtime, & pstendtime)
Dim ilfirstTime As Integer
Dim llCurrAvailTime As Long
Dim llPrevAvailTime As Long
Dim llNextStartTime As Long
Dim slAvailSTime As String
Dim slAvailETime As String

Dim llCurrLocalTime As Long
Dim llPrevLocalTime As Long
Dim llNextLocalStartTime As Long
Dim slLocalSTime As String
Dim slLocalETime As String
 
Dim tlodfExtSrchKey As ODFKEY0
Dim tlOdf As ODF
'ReDim llZoneEndTimes(1 To 6) As Long          'First field is always 0, remaining extracted for vpf (estendtime, cstendtime, mstendtime, & pstendtime)
ReDim llZoneEndTimes(0 To 6) As Long          'First field is always 0, remaining extracted for vpf (estendtime, cstendtime, mstendtime, & pstendtime). Index zero ignored
Dim llODFPosition() As Long
'**** parameters passed from Log program
Dim llStartTime As Long             'start time of log gen
Dim llEndTime As Long               'end time of log gen
Dim llStartDate As Long             'start date of log gen
Dim llEndDate As Long               'end date of log gen
Dim ilUrfCode As Integer           'user code requesting log
Dim ilVefCode As Integer            ' vehicle to process
Dim llWeek As Long                  '12-01-00
Dim slXMid As String
'**** end parameters passed from Log program
    hlODF = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlODF, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlODF)
        btrDestroy hlODF
        Exit Sub
    End If
    hlCef = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlCef)
        btrDestroy hlCef
        Exit Sub
    End If

    llStartDate = gDateValue(sgLogStartDate)
    llEndDate = llStartDate + Val(sgLogNoDays) - 1
    llStartOfWk = llStartDate                       'get to Monday start week
    Do While (gWeekDayLong(llStartOfWk)) <> 0
        llStartOfWk = llStartOfWk - 1
    Loop
    'convert to btrieve format
    gPackDateLong llStartOfWk, ilStartofWk(0), ilStartofWk(1)
    llStartTime = CLng(gTimeToCurrency(sgLogStartTime, False))
    llEndTime = CLng(gTimeToCurrency(sgLogEndTime, True)) - 1
    ilUrfCode = Val(sgLogUserCode)                         'user code requesting log
    ilVefCode = igcodes(0)                                  'passed for log function
    ilVpfIndex = -1
    ilVpfIndex = gVpfFind(RptSelLg, ilVefCode)              'determine vehicle options index

    ilUseZone = False                                       'assume not using zones until one is found in the vehicle options table
    ilZones = igZones                                       'save zones requested by user : 0=none, 1 =est, 2= cst, 3 =mst, 4 = pst
    ilLoZone = 1                                            'low loop factor to process zones
    ilHiZone = 4                                            'hi loop factor to process zones
    If ilZones <> 0 Then                                    'user has requested one zone in particular
        ilLoZone = ilZones
        ilHiZone = ilZones
    End If
    If ilVpfIndex >= 0 Then                                  'associated vehicle options record exists
        'If tgVpf(ilVpfIndex).sGZone(1) <> "   " Then
        If tgVpf(ilVpfIndex).sGZone(0) <> "   " Then
            ilUseZone = True
        Else
            'Zones not used, fake out flag to do 1 zone (EST)
            ilZones = 1
            ilLoZone = 1
            ilHiZone = 1
        End If
    Else
        'no vehicle options table
    End If
    For ilZoneLoop = ilLoZone To ilHiZone               'loop on all time zones (or just the selective one,  variety of zones are not allowed)
        Select Case ilZoneLoop
            Case 1  'Eastern
                For ilRec = 1 To 5
                    llZoneEndTimes(ilRec + 1) = 60 * CLng(tgVpf(ilVpfIndex).iESTEndTime(ilRec - 1))
                Next ilRec
                If ilUseZone Then
                    slZone = "EST"
                Else
                    slZone = "   "
                End If
            Case 2  'Central
                For ilRec = 1 To 5
                    llZoneEndTimes(ilRec + 1) = 60 * CLng(tgVpf(ilVpfIndex).iCSTEndTime(ilRec - 1))
                Next ilRec
                slZone = "CST"
            Case 3  'Mountain
                For ilRec = 1 To 5
                    llZoneEndTimes(ilRec + 1) = 60 * CLng(tgVpf(ilVpfIndex).iMSTEndTime(ilRec - 1))
                Next ilRec
                slZone = "MST"
            Case 4  'Pacific
                For ilRec = 1 To 5
                    llZoneEndTimes(ilRec + 1) = 60 * CLng(tgVpf(ilVpfIndex).iPSTEndTime(ilRec - 1))
                Next ilRec
                slZone = "PST"
        End Select

        For llWeek = llStartOfWk To llEndDate Step 7        '12-1-00
            For ilDay = 0 To 6                                'loop on all days of the week
                'slDate = Format$(llStartDate + ilDay, "m/d/yy")  '12-1-00
                'ilDayofWeek = gWeekDayLong(llStartDate + ilDay) '12-1-00
                slDate = Format$(llWeek + ilDay, "m/d/yy")        '12-1-00
                ilDayOfWeek = gWeekDayLong(llWeek + ilDay)       '12-1-00
                Select Case ilDay
                    Case 0
                        mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
                    Case 1
                        mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
                    Case 2
                        mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
                    Case 3
                        mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
                    Case 4
                        mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
                    Case 5
                        mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
                    Case 6
                        mReadODF hlODF, ilUrfCode, ilVefCode, slZone, slDate, llStartTime, llEndTime, tmOdf0(), llODFPosition()
                End Select

                If sgRnfRptName <> "C74" And sgRnfRptName <> "C75" And sgRnfRptName <> "C76" And sgRnfRptName <> "C77" And sgRnfRptName <> "C79" Then
                    'Loop thru each event and build a unique record based on zone, vehicle, event type, time & position
                    'for the entire week
                    
                    '11-16-13 Currently, only odfAirTime is adjusted for running spot times.  odfLocalTime should be changed as well.  This may pose a problem for the log output, depending on its format.
                    'odfLocalTime comes into play with timezones.
                    
                    ilfirstTime = True
                    For ilOdf = LBound(tmOdf0) To UBound(tmOdf0) - 1 Step 1
                        If tmOdf0(ilOdf).iType = 4 Then         'spot, test for all spots that are back-to back.
                                                                'these spots will all maintain the same seq # (used for
                                                                'grouping in Crystal report

                            gUnpackTimeLong tmOdf0(ilOdf).iLocalTime(0), tmOdf0(ilOdf).iLocalTime(1), False, llCurrLocalTime
                            gUnpackTimeLong tmOdf0(ilOdf).iAirTime(0), tmOdf0(ilOdf).iAirTime(1), False, llCurrAvailTime
                            If ilfirstTime Then
                                ilfirstTime = False
                                llPrevAvailTime = llCurrAvailTime
                                gUnpackLength tmOdf0(ilOdf).iLen(0), tmOdf0(ilOdf).iLen(1), "3", False, slLen
                                gUnpackTime tmOdf0(ilOdf).iAirTime(0), tmOdf0(ilOdf).iAirTime(1), "A", "1", slAvailSTime
                                gAddTimeLength slAvailSTime, slLen, "A", "1", slAvailETime, slXMid
                                llNextStartTime = gTimeToLong(slAvailETime, True)
                                
'                                'adjust the zone local times
'                                llPrevLocalTime = llCurrLocalTime
'                                gUnpackLength tmOdf0(ilOdf).iLen(0), tmOdf0(ilOdf).iLen(1), "3", False, slLen
'                                gUnpackTime tmOdf0(ilOdf).iLocalTime(0), tmOdf0(ilOdf).iLocalTime(1), "A", "1", slLocalSTime
'                                gAddTimeLength slLocalSTime, slLen, "A", "1", slLocalETime, slXMid
'                                llNextLocalStartTime = gTimeToLong(slLocalETime, True)

                            Else
                                If llPrevAvailTime = llCurrAvailTime Then      'same break
                                    'same break, keep running time
                                    gPackTimeLong llNextStartTime, tmOdf0(ilOdf).iAirTime(0), tmOdf0(ilOdf).iAirTime(1)
                                    gUnpackLength tmOdf0(ilOdf).iLen(0), tmOdf0(ilOdf).iLen(1), "3", False, slLen
                                    gUnpackTime tmOdf0(ilOdf).iAirTime(0), tmOdf0(ilOdf).iAirTime(1), "A", "1", slAvailSTime
                                    gAddTimeLength slAvailSTime, slLen, "A", "1", slAvailETime, slXMid
                                    llNextStartTime = gTimeToLong(slAvailETime, True)
                                    gPackTime slAvailSTime, tmOdf0(ilOdf).iAirTime(0), tmOdf0(ilOdf).iAirTime(1)
                                Else                                        'different break
                                    llPrevAvailTime = llCurrAvailTime
                                    gUnpackLength tmOdf0(ilOdf).iLen(0), tmOdf0(ilOdf).iLen(1), "3", False, slLen
                                    gUnpackTime tmOdf0(ilOdf).iAirTime(0), tmOdf0(ilOdf).iAirTime(1), "A", "1", slAvailSTime
                                    gAddTimeLength slAvailSTime, slLen, "A", "1", slAvailETime, slXMid
                                    llNextStartTime = gTimeToLong(slAvailETime, True)
                                    gPackTime slAvailSTime, tmOdf0(ilOdf).iAirTime(0), tmOdf0(ilOdf).iAirTime(1)
                                End If

'                                'adjust Local zone times
'                                If llPrevLocalTime = llCurrLocalTime Then      'same break
'                                    'same break, keep running time
'                                    gPackTimeLong llNextLocalStartTime, tmOdf0(ilOdf).iLocalTime(0), tmOdf0(ilOdf).iLocalTime(1)
'                                    gUnpackLength tmOdf0(ilOdf).iLen(0), tmOdf0(ilOdf).iLen(1), "3", False, slLen
'                                    gUnpackTime tmOdf0(ilOdf).iLocalTime(0), tmOdf0(ilOdf).iLocalTime(1), "A", "1", slLocalSTime
'                                    gAddTimeLength slLocalSTime, slLen, "A", "1", slLocalETime, slXMid
'                                    llNextLocalStartTime = gTimeToLong(slLocalETime, True)
'                                    gPackTime slLocalSTime, tmOdf0(ilOdf).iLocalTime(0), tmOdf0(ilOdf).iLocalTime(1)
'                                Else                                        'different break
'                                    llPrevLocalTime = llCurrLocalTime
'                                    gUnpackLength tmOdf0(ilOdf).iLen(0), tmOdf0(ilOdf).iLen(1), "3", False, slLen
'                                    gUnpackTime tmOdf0(ilOdf).iLocalTime(0), tmOdf0(ilOdf).iLocalTime(1), "A", "1", slLocalSTime
'                                    gAddTimeLength slLocalSTime, slLen, "A", "1", slLocalETime, slXMid
'                                    llNextLocalStartTime = gTimeToLong(slLocalETime, True)
'                                    gPackTime slLocalSTime, tmOdf0(ilOdf).iLocalTime(0), tmOdf0(ilOdf).iLocalTime(1)
'                                End If
                            End If
                        End If
                    Next ilOdf
                End If


                'go thru the odf records in memory and write back to disk withupdated avail times
                'which will be used as one of the sort fields C70, C71 log .  Also, update the header and footer
                'comments from VOF
                tlodfExtSrchKey.iVefCode = ilVefCode
               '12-1-00 gPackDateLong llStartDate + ilDay, tlodfExtSrchKey.iAirDate(0), tlodfExtSrchKey.iAirDate(1)
                gPackDateLong llWeek + ilDay, tlodfExtSrchKey.iAirDate(0), tlodfExtSrchKey.iAirDate(1)  '12-1-00
                For ilOdf = LBound(tmOdf0) To UBound(tmOdf0) - 1 Step 1

                    'ilRet = btrGetDirect(hlODF, tlOdf, Len(tmOdf0(1)), llODFPosition(ilOdf), INDEXKEY0, BTRV_LOCK_NONE)
                    ilRet = btrGetDirect(hlODF, tlOdf, Len(tmOdf0(0)), llODFPosition(ilOdf), INDEXKEY0, BTRV_LOCK_NONE)

'                    tlodfExtSrchKey.sZone = tmOdf0(ilOdf).sZone
'                    tlodfExtSrchKey.iSeqNo = tmOdf0(ilOdf).iSeqNo
'                    tlodfExtSrchKey.iLocalTime(0) = tmOdf0(ilOdf).iLocalTime(0)
'                    tlodfExtSrchKey.iLocalTime(1) = tmOdf0(ilOdf).iLocalTime(1)
'                    ilRet = btrGetEqual(hlOdf, tlOdf, Len(tmOdf0(1)), tlodfExtSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        ilRet = btrClose(hlODF)
                        ilRet = btrClose(hlCef)
                        btrDestroy hlODF
                        btrDestroy hlCef
                        Erase tmOdf0
                        Exit Sub
                    End If

                    tmOdf0(ilOdf).lHd1CefCode = tgVof.lHd1CefCode
                    tmOdf0(ilOdf).lFt1CefCode = tgVof.lFt1CefCode
                    tmOdf0(ilOdf).lFt2CefCode = tgVof.lFt2CefCode

                    'ilRet = btrUpdate(hlODF, tmOdf0(ilOdf), Len(tmOdf0(1)))
                    ilRet = btrUpdate(hlODF, tmOdf0(ilOdf), Len(tmOdf0(0)))
                    If ilRet <> BTRV_ERR_NONE Then
                        ilRet = btrClose(hlODF)
                        ilRet = btrClose(hlCef)
                        btrDestroy hlODF
                        btrDestroy hlCef
                        Erase tmOdf0
                        Exit Sub
                    End If
                Next ilOdf

                ReDim tmOdf0(0 To 0) As ODF
            Next ilDay
        Next llWeek                             '12-1-00
    Next ilZoneLoop                                         'for ilzone
    ilRet = btrClose(hlODF)
    ilRet = btrClose(hlCef)
    btrDestroy hlODF
    btrDestroy hlCef
    Erase tmOdf0
End Sub
'************************************************************
'*                                                          *
'*      Procedure Name:mGetPlayList                         *
'*                                                          *
'*             Created:10/09/93      By:D. LeVine           *
'*            Modified:              By:                    *
'*                                                          *
'*            Comments:Obtain the Sdf records to be         *
'*                     reported                             *
'*                                                          *
'*      12/22/98 dh  if not using cart #s, place the        *
'*                  reel # field in the cpr cart #          *
'*                  field.  Global needs the cart           *
'*                  #s to print for Media America           *
'*                  vehicles                                *
'*      12/27/04    Test for valid airing day for an airing *
'                   vehicle.  If a airing vehicles library  *
'                   isnt M-F (i.e. Tuesday/Thursday not     *
'                   defined), the spots were still included *
'                   from the selling vehicles Tu/Th day     *
'*                                                          *
'************************************************************
Sub mGetPlayList(ilFdVefCode As Integer, slStartDate As String, slEndDate As String, tlPlayList As PLAYLISTLG)
'
'
'   Where
'
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilUpper As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slProduct As String
    Dim slISCI As String
    Dim slCreative As String
    Dim slCart As String
    Dim slZone As String
    Dim slDate As String
    Dim ilDay As Integer
    Dim slDay As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim blTestAirTimeUnits As Boolean       '10-12-16           Test honoring zero units feature
    Dim ilVff As Integer
    Dim llTemp As Long
    
    'ReDim ilVlfStartDate0(1 To 6) As Integer
    'ReDim ilVlfStartDate1(1 To 6) As Integer
    Dim ilVlfStartDate0 As Integer
    Dim ilVlfStartDate1 As Integer
    ReDim ilCurrDate(0 To 1) As Integer
    'ReDim ilVefCode(1 To 6) As Integer
    Dim ilVefCode As Integer
    Dim ilTerminated As Integer
    
    '10-13-16 Need to see if an airing vehicle, and to ignore avails defined as 0 units (HonorZeroUnits feature)
    blTestAirTimeUnits = False
    If tlPlayList.sType = "S" Then
        ilVff = gBinarySearchVff(tlPlayList.iAirCode)
        If ilVff <> -1 Then
            If tgVff(ilVff).sHonorZeroUnits = "Y" Then
                blTestAirTimeUnits = True
            End If
        End If
    End If
    
    btrExtClear hmSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmSdf   'Clear any previous extend operation
    tmSdfSrchKey1.iVefCode = ilFdVefCode
    gPackDate slStartDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
    tmSdfSrchKey1.iTime(0) = 0
    tmSdfSrchKey1.iTime(1) = 0
    tmSdfSrchKey1.sSchStatus = ""
    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
        tlIntTypeBuff.iType = ilFdVefCode
        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        If slStartDate <> "" Then
            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Sdf", "SdfDate")
            If slEndDate <> "" Then
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            End If
        End If
        If slEndDate <> "" Then
            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Sdf", "SdfDate")
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        End If
        ilRet = btrExtAddField(hmSdf, 0, ilExtLen)  'Extract Name
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilFound = False
                If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
                    'If vehicle is selling- obtain start date of Vlf
                    If tlPlayList.sType = "S" Then
                        If (ilCurrDate(0) <> tmSdf.iDate(0)) Or (ilCurrDate(1) <> tmSdf.iDate(1)) Or (blTestAirTimeUnits) Then
                            ilCurrDate(0) = tmSdf.iDate(0)
                            ilCurrDate(1) = tmSdf.iDate(1)
                            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                            ilDay = gWeekDayStr(slDate)
                            If (ilDay >= 0) And (ilDay <= 4) Then
                                slDay = "0"
                            ElseIf ilDay = 5 Then
                                slDay = "6"
                            Else
                                slDay = "7"
                            End If
                            'For ilVef = 1 To 6 Step 1
                                If tlPlayList.iAirCode > 0 Then
                                    'if airing vehicle, see if zero units should be excluded
                                    tmVlfSrchKey0.iSellTime(0) = 0
                                    tmVlfSrchKey0.iSellTime(1) = 6144    '24*256
                                    If blTestAirTimeUnits Then          '10-13-16 use time of selling vehicle to search the links if honoring zero units, need to test the avail
                                        tmVlfSrchKey0.iSellTime(0) = tmSdf.iTime(0)
                                        tmVlfSrchKey0.iSellTime(1) = tmSdf.iTime(1)
                                    End If
                                    ilVlfStartDate0 = 0
                                    ilVlfStartDate1 = 0
                                    tmVlfSrchKey0.iSellCode = ilFdVefCode
                                    tmVlfSrchKey0.iSellDay = Val(slDay)
                                    tmVlfSrchKey0.iEffDate(0) = tmSdf.iDate(0)
                                    tmVlfSrchKey0.iEffDate(1) = tmSdf.iDate(1)
'                                    tmVlfSrchKey0.iSellTime(0) = 0
'                                    tmVlfSrchKey0.iSellTime(1) = 6144    '24*256
                                    tmVlfSrchKey0.iSellPosNo = 32000
                                    tmVlfSrchKey0.iSellSeq = 32000
                                    ilRet = btrGetLessOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = ilFdVefCode) And (tmVlf.iSellDay = Val(slDay))
                                        ilTerminated = False
                                        If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                            If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                                ilTerminated = True
                                            End If
                                        End If
                                        If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                            '12-27-04   determine if this day is a valid air vehicle air date
                                            If tmVlf.iAirCode = tlPlayList.iAirCode Then
                                                If blTestAirTimeUnits Then
                                                    If tmSdf.iTime(0) <> tmVlf.iSellTime(0) Or tmSdf.iTime(1) <> tmVlf.iSellTime(1) Then
                                                        ilRet = False
                                                    Else
                                                        ilRet = gTestAirVefValidDay(hmSsf, slDate, tlPlayList.iAirCode, tmVlf, blTestAirTimeUnits)
                                                    End If
                                                Else
                                                    ilRet = gTestAirVefValidDay(hmSsf, slDate, tlPlayList.iAirCode, tmVlf, blTestAirTimeUnits)
                                                End If
                                                    
                                                If ilRet Then       'found valid airing day
                                                    ilVlfStartDate0 = tmVlf.iEffDate(0)
                                                    ilVlfStartDate1 = tmVlf.iEffDate(1)
                                                    Exit Do
                                                Else
                                                    ilVlfStartDate0 = 0
                                                    ilVlfStartDate1 = 0
                                                End If
                                            End If
                                        End If
                                        ilRet = btrGetPrevious(hmVLF, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
                                Else
                                    ilVlfStartDate0 = 0
                                    ilVlfStartDate1 = 0
                                End If
                            'Next ilVef
                        End If
                        'For ilVef = 1 To 6 Step 1
                        '    ilVefCode(ilVef) = 0
                        'Next ilVef
                        'For ilVef = 1 To 6 Step 1
                            If (tlPlayList.iAirCode > 0) And (ilVlfStartDate0 <> 0 Or ilVlfStartDate1 <> 0) Then        '12-27-04 test the dates
                                ilVefCode = 0
                                tmVlfSrchKey0.iSellCode = ilFdVefCode
                                tmVlfSrchKey0.iSellDay = Val(slDay)
                                tmVlfSrchKey0.iEffDate(0) = ilVlfStartDate0
                                tmVlfSrchKey0.iEffDate(1) = ilVlfStartDate1
                                tmVlfSrchKey0.iSellTime(0) = tmSdf.iTime(0)
                                tmVlfSrchKey0.iSellTime(1) = tmSdf.iTime(1)    '24*256
                                tmVlfSrchKey0.iSellPosNo = 0
                                tmVlfSrchKey0.iSellSeq = 0
                                ilRet = btrGetGreaterOrEqual(hmVLF, tmVlf, imVlfRecLen, tmVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                                Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = ilFdVefCode) And (tmVlf.iSellDay = Val(slDay)) And (tmVlf.iSellTime(0) = tmSdf.iTime(0)) And (tmVlf.iSellTime(1) = tmSdf.iTime(1))
                                    ilTerminated = False
                                    If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                        If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                            ilTerminated = True
                                        End If
                                    End If
                                    If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                        If tmVlf.iAirCode = tlPlayList.iAirCode Then
                                            ilVefCode = tlPlayList.iAirCode
                                            Exit Do
                                        End If
                                    End If
                                    ilRet = btrGetNext(hmVLF, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                            Else
                                ilVefCode = 0
                            End If
                        'Next ilVef
                    Else
                        'For ilVef = 1 To 6 Step 1
                        '    ilVefCode(ilVef) = 0
                        'Next ilVef
                        If tlPlayList.iLogCode > 0 Then
                            ilVefCode = tlPlayList.iLogCode
                        Else
                            ilVefCode = tmSdf.iVefCode
                        End If
                    End If
                    'For ilVef = 1 To 6 Step 1
                        If ilVefCode > 0 Then
                            If tmSdf.sPtType = "1" Then  '  Single Copy
                                ' Read CIF using lCopyCode from SDF
                                tmCifSrchKey.lCode = tmSdf.lCopyCode
                                ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    slZone = ""
                                    If tmCif.lcpfCode > 0 Then
                                        tmCpfSrchKey.lCode = tmCif.lcpfCode
                                        ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            tmCpf.sISCI = ""
                                            tmCpf.sName = ""
                                            tmCpf.sCreative = ""
                                        End If
                                        slISCI = Trim$(tmCpf.sISCI)
                                        slProduct = Trim$(tmCpf.sName)
                                        slCreative = Trim$(tmCpf.sCreative)
                                    End If
                                    If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
                                        If tmCif.iMcfCode <> tmMcf.iCode Then
                                            tmMcfSrchKey.iCode = tmCif.iMcfCode
                                            ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet <> BTRV_ERR_NONE Then
                                                tmMcf.sName = ""
                                            End If
                                            slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                        Else
                                            slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                        End If
                                    Else
                                        'slCart = ""
                                        slCart = Trim$(tmCif.sReel)
                                    End If
                                    ilFound = False
                                    For ilLoop = LBound(tmCpr) To UBound(tmCpr) - 1 Step 1
                                        If (tmCpr(ilLoop).iVefCode = ilVefCode) And (tmCpr(ilLoop).iAdfCode = tmSdf.iAdfCode) And (tmCpr(ilLoop).iLen = tmSdf.iLen) Then
                                            If (Trim$(tmCpr(ilLoop).sProduct) = slProduct) And (Trim$(tmCpr(ilLoop).sZone) = slZone) And (Trim$(tmCpr(ilLoop).sCartNo) = slCart) And (Trim$(tmCpr(ilLoop).sISCI) = slISCI) And (Trim$(tmCpr(ilLoop).sCreative) = slCreative) Then
                                                tmCpr(ilLoop).iLineNo = tmCpr(ilLoop).iLineNo + 1
                                                ilFound = True
                                                Exit For
                                            End If
                                        End If
                                    Next ilLoop
                                    If Not ilFound Then
                                        ilUpper = UBound(tmCpr)
                                        tmCpr(ilUpper).iGenDate(0) = igNowDate(0)
                                        tmCpr(ilUpper).iGenDate(1) = igNowDate(1)
                                        '10-10-01
                                        tmCpr(ilUpper).lGenTime = lgNowTime
                                        'tmCpr(ilUpper).iGenTime(0) = igNowTime(0)
                                        'tmCpr(ilUpper).iGenTime(1) = igNowTime(1)
                                        tmCpr(ilUpper).iVefCode = ilVefCode
                                        tmCpr(ilUpper).iAdfCode = tmSdf.iAdfCode
                                        tmCpr(ilUpper).iLen = tmSdf.iLen
                                        tmCpr(ilUpper).sProduct = slProduct
                                        tmCpr(ilUpper).sZone = slZone
                                        tmCpr(ilUpper).sCartNo = slCart
                                        tmCpr(ilUpper).sISCI = slISCI
                                        tmCpr(ilUpper).sCreative = slCreative
                                        tmCpr(ilUpper).iLineNo = 1
                                        ReDim Preserve tmCpr(0 To ilUpper + 1) As CPR
                                    End If
                                End If
                            ElseIf tmSdf.sPtType = "2" Then  '  Combo Copy
                            ElseIf tmSdf.sPtType = "3" Then  '  Time Zone Copy
                                ' Read TZF using lCopyCode from SDF
                                tmTzfSrchKey.lCode = tmSdf.lCopyCode
                                ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    ' Look for the first positive lZone value
                                    For ilIndex = 1 To 6 Step 1
                                        If (tmTzf.lCifZone(ilIndex - 1) > 0) Then ' Process just the first positive Zone
                                            ' Read CIF using lCopyCode from SDF
                                            tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                                            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet = BTRV_ERR_NONE Then
                                                slZone = Trim$(tmTzf.sZone(ilIndex - 1))
                                                If tmCif.lcpfCode > 0 Then
                                                    tmCpfSrchKey.lCode = tmCif.lcpfCode
                                                    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                    If ilRet <> BTRV_ERR_NONE Then
                                                        tmCpf.sISCI = ""
                                                        tmCpf.sName = ""
                                                        tmCpf.sCreative = ""
                                                    End If
                                                    slISCI = Trim$(tmCpf.sISCI)
                                                    slProduct = Trim$(tmCpf.sName)
                                                    slCreative = Trim$(tmCpf.sCreative)
                                                End If
                                                If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
                                                    If tmCif.iMcfCode <> tmMcf.iCode Then
                                                        tmMcfSrchKey.iCode = tmCif.iMcfCode
                                                        ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        If ilRet <> BTRV_ERR_NONE Then
                                                            tmMcf.sName = ""
                                                        End If
                                                        slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                                    Else
                                                        slCart = Trim$(tmMcf.sName) & Trim$(tmCif.sName) & Trim$(tmCif.sCut)
                                                    End If
                                                Else
                                                    'slCart = ""
                                                    slCart = Trim$(tmCif.sReel)
                                                End If
                                                ilFound = False
                                                For ilLoop = LBound(tmCpr) To UBound(tmCpr) - 1 Step 1
                                                    If (tmCpr(ilLoop).iVefCode = ilVefCode) And (tmCpr(ilLoop).iAdfCode = tmSdf.iAdfCode) And (tmCpr(ilLoop).iLen = tmSdf.iLen) Then
                                                        If (Trim$(tmCpr(ilLoop).sProduct) = slProduct) And (Trim$(tmCpr(ilLoop).sZone) = slZone) And (Trim$(tmCpr(ilLoop).sCartNo) = slCart) And (Trim$(tmCpr(ilLoop).sISCI) = slISCI) And (Trim$(tmCpr(ilLoop).sCreative) = slCreative) Then
                                                            tmCpr(ilLoop).iLineNo = tmCpr(ilLoop).iLineNo + 1
                                                            ilFound = True
                                                            Exit For
                                                        End If
                                                    End If
                                                Next ilLoop
                                                If Not ilFound Then
                                                    ilUpper = UBound(tmCpr)
                                                    tmCpr(ilUpper).iGenDate(0) = igNowDate(0)
                                                    tmCpr(ilUpper).iGenDate(1) = igNowDate(1)
                                                    '10-10-01
                                                    tmCpr(ilUpper).lGenTime = lgNowTime
                                                    'tmCpr(ilUpper).iGenTime(0) = igNowTime(0)
                                                    'tmCpr(ilUpper).iGenTime(1) = igNowTime(1)
                                                    tmCpr(ilUpper).iVefCode = ilVefCode
                                                    tmCpr(ilUpper).iAdfCode = tmSdf.iAdfCode
                                                    tmCpr(ilUpper).iLen = tmSdf.iLen
                                                    tmCpr(ilUpper).sProduct = slProduct
                                                    tmCpr(ilUpper).sZone = slZone
                                                    tmCpr(ilUpper).sCartNo = slCart
                                                    tmCpr(ilUpper).sISCI = slISCI
                                                    tmCpr(ilUpper).sCreative = slCreative
                                                    tmCpr(ilUpper).iLineNo = 1
                                                    ReDim Preserve tmCpr(0 To ilUpper + 1) As CPR
                                                End If
                                            End If
                                        End If
                                    Next ilIndex
                                End If
                            End If
                        End If
                    'Next ilVef
                End If
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainExtOdf                   *
'*          Extended read on ODF file for given vehicle
'*          and date/time
'*                                                     *
'*             Created:7/9/99       By:D. Hosaka       *
'*            Modified:              By:               *
'*                                                     *
'*                                                     *
''  5-25-01 only retrieve the ODF records just created
'           by using the ODFGenDate & Time sent via
'           command statement from Logs
'  6-12-02 change extracting odftime from time field to long
'*******************************************************
Sub mReadODF(hlODF As Integer, ilUrfCode As Integer, ilVefCode As Integer, slZone As String, slDate As String, llStartTime As Long, llEndTime As Long, tlOdfExt() As ODF, llODFPosition() As Long)
'
'   mObtainOdf
'   Where:
'
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim slTime As String
    Dim ilUpper As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE          '6-12-02 long
    Dim tlStrTypeBuff As POPSTRINGTYPE
    Dim tlodfExtSrchKey As ODFKEY0
    Dim ilAirDate0 As Integer
    Dim ilAirDate1 As Integer
    Dim tlOdf As ODF
    'ReDim tlOdfExt(1 To 1) As ODF
    ReDim tlOdfExt(0 To 0) As ODF
    'ReDim llODFPosition(1 To 1) As Long '7-10-08 save the record position for duplicate keys when updating
    ReDim llODFPosition(0 To 0) As Long '7-10-08 save the record position for duplicate keys when updating
    'ilRecLen = Len(tlOdfExt(1)) 'btrRecordLength(hlAdf)  'Get and save record length
    ilRecLen = Len(tlOdfExt(0)) 'btrRecordLength(hlAdf)  'Get and save record length
    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hlODF   'Clear any previous extend operation
    'tlOdfExtSrchKey.iUrfCode = ilUrfCode
    tlodfExtSrchKey.iVefCode = ilVefCode
    gPackDate slDate, tlodfExtSrchKey.iAirDate(0), tlodfExtSrchKey.iAirDate(1)
    gPackDate slDate, ilAirDate0, ilAirDate1
    slTime = gCurrencyToTime(CCur(llStartTime))
    gPackTime slTime, tlodfExtSrchKey.iLocalTime(0), tlodfExtSrchKey.iLocalTime(1)
    tlodfExtSrchKey.sZone = "" 'slZone
    tlodfExtSrchKey.iSeqNo = 0
    ilRet = btrGetGreaterOrEqual(hlODF, tlOdf, Len(tlOdf), tlodfExtSrchKey, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        Exit Sub
    Else
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
    End If
    Call btrExtSetBounds(hlODF, llNoRec, -1, "UC", "ODF", "")  'Set extract limits (all records)

    'tlIntTypeBuff.iType = ilUrfCode
    'ilOffset = gFieldOffsetExtra("ODF", "OdfUrfCode")
    'ilRet = btrExtAddLogicConst(hlOdf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
    
    
    tlIntTypeBuff.iType = ilVefCode
    ilOffSet = gFieldOffsetExtra("ODF", "OdfVefCode")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)

    tlDateTypeBuff.iDate0 = igODFGenDate(0)     '5-25-01
    tlDateTypeBuff.iDate1 = igODFGenDate(1)
    ilOffSet = gFieldOffsetExtra("ODF", "ODFGenDate")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlDateTypeBuff, 4)
    'tlDateTypeBuff.iDate0 = igODFGenTime(0)     '5-25-01
    'tlDateTypeBuff.iDate1 = igODFGenTime(1)
    tlLongTypeBuff.lCode = lgGenTime            '6-12-02
    ilOffSet = gFieldOffsetExtra("ODF", "ODFGenTime")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)

    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffsetExtra("ODF", "OdfAirDate")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlDateTypeBuff, 4)
    slTime = gCurrencyToTime(CCur(llEndTime))
    gPackTime slTime, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffsetExtra("ODF", "OdfLocalTime")
    ilRet = btrExtAddLogicConst(hlODF, BTRV_KT_TIME, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

    ilUpper = UBound(tlOdfExt)
    ilRet = btrExtAddField(hlODF, 0, ilRecLen)  'Extract the whole record

    ilRet = btrExtGetNext(hlODF, tlOdf, ilRecLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            Exit Sub
        End If
    End If
    'ilRet = btrExtGetFirst(hlIcf, tlIcf, ilRecLen, llRecPos)
    Do While ilRet = BTRV_ERR_REJECT_COUNT
        ilRet = btrExtGetNext(hlODF, tlOdf, ilRecLen, llRecPos)
    Loop
    Do While ilRet = BTRV_ERR_NONE
        'If tlOdfExt(ilUpper).iMnfSubFeed = 0 Then   'Or ((Trim$(slZone) = "") Or (Trim$(slZone) = tlOdfExt(ilUpper).sZone)) Then      'bypass records with subfeed
        If (tlOdf.iMnfSubFeed = 0) And ((Trim$(slZone) = "") Or (Trim$(slZone) = Trim$(tlOdf.sZone))) Then       'bypass records with subfeed
           tlOdfExt(ilUpper) = tlOdf
            llODFPosition(ilUpper) = llRecPos           '7-10-08
            'ReDim Preserve tlOdfExt(1 To ilUpper + 1) As ODF
            ReDim Preserve tlOdfExt(0 To ilUpper + 1) As ODF
            'ReDim Preserve llODFPosition(1 To ilUpper + 1) As Long
            ReDim Preserve llODFPosition(0 To ilUpper + 1) As Long
            ilUpper = ilUpper + 1
        End If
        ilRet = btrExtGetNext(hlODF, tlOdf, ilRecLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlODF, tlOdf, ilRecLen, llRecPos)
        Loop
    Loop
    Exit Sub
End Sub
'
'           All air time and NTR records have been created for L87
'           Since there is a combination of both types of records, Crystal
'           cannot point to 2 different files, using one or the other from
'           the base file.  Therefore, create another temporary file built
'           of all the fields required to handle 2 different record types
'           requiring 2 different internal pointers to supporting filese
'           Using the created log ODF records, create 1 record in Cbf for
'           each air time (type 4) and NTR record (type 5).
'
'           The NTR records have been created in the ODF creation since its know
'           what vehicles to process at that time, depending if log vehicle, conv vehicle
'
'       cbfGenDate - generation time to filter records for crystal
'       cbfgenTime - generation time to filter records for crystal
'       cbfChfCode - contract header internal code
'       chfvefCode - Log vehicle or same as the alternate vehicle code
'       cbfDysTms - daypart name with overrides
'       cbfLength - spot length (air time)
'       cbfvehsort - alternate vehicle code (when using log vehicles, this is the conventional code of the vehicles merged)
'       cbfSortField1 - DP name or NTR type (this field is compared to cbfDysTms.  If same, show this field only.  If different, show cbfDysTms field in another column
'       cbfSortField2 - ISCI for air time only
'       cbfSurvey - Creative title for air time only
'       cbfIntComment - vehicle footer comment
'       cbfOtherComment - vehicle footer comment #2
'       cbfLineComment - vehicle header comment; 8/9/16 changed to cbfChgRComment
'       cbfLineComment = line comment (clf)
'       cbfResort - zone
'       cbfType - A = airtime, N = ntr
'       cbfAudioType - from clflivecopy
'       cbfSTartDate - rotation start date, stored in CIF
Public Sub gGenL87Master()

Dim ilRet As Integer

        hmOdf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmOdf, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmOdf)
            btrDestroy hmOdf
            Exit Sub
        End If
        imOdfRecLen = Len(tmOdfDay)
        
        hmCbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmCbf)
            ilRet = btrClose(hmOdf)
            btrDestroy hmOdf
            btrDestroy hmCbf
            Exit Sub
        End If
        imCbfRecLen = Len(tmCbf)
        
        hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmCbf)
            ilRet = btrClose(hmOdf)
            ilRet = btrClose(hmSbf)
            btrDestroy hmOdf
            btrDestroy hmCbf
            btrDestroy hmSbf
            Exit Sub
        End If
        imSbfRecLen = Len(tmSbf)

        hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmCbf)
            ilRet = btrClose(hmOdf)
            ilRet = btrClose(hmSbf)
            ilRet = btrClose(hmRdf)
            btrDestroy hmOdf
            btrDestroy hmCbf
            btrDestroy hmSbf
            btrDestroy hmRdf
            Exit Sub
        End If
        imRdfRecLen = Len(tmRdf)
        
        hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmCbf)
            ilRet = btrClose(hmOdf)
            ilRet = btrClose(hmSbf)
            ilRet = btrClose(hmRdf)
            btrDestroy hmCHF
            btrDestroy hmOdf
            btrDestroy hmCbf
            btrDestroy hmSbf
            btrDestroy hmRdf
            Exit Sub
        End If
        imCHFRecLen = Len(tmChf)
        
        hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmCbf)
            ilRet = btrClose(hmOdf)
            ilRet = btrClose(hmSbf)
            ilRet = btrClose(hmRdf)
            btrDestroy hmClf
            btrDestroy hmCHF
            btrDestroy hmOdf
            btrDestroy hmCbf
            btrDestroy hmSbf
            btrDestroy hmRdf
            Exit Sub
        End If
        imClfRecLen = Len(tmClf)
       
        hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmMnf)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmCbf)
            ilRet = btrClose(hmOdf)
            ilRet = btrClose(hmSbf)
            ilRet = btrClose(hmRdf)
            btrDestroy hmMnf
            btrDestroy hmClf
            btrDestroy hmCHF
            btrDestroy hmOdf
            btrDestroy hmCbf
            btrDestroy hmSbf
            btrDestroy hmRdf
            Exit Sub
        End If
        imMnfRecLen = Len(tmMnf)
        
        hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmCif)
            ilRet = btrClose(hmMnf)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmCbf)
            ilRet = btrClose(hmOdf)
            ilRet = btrClose(hmSbf)
            ilRet = btrClose(hmRdf)
            btrDestroy hmCif
            btrDestroy hmMnf
            btrDestroy hmClf
            btrDestroy hmCHF
            btrDestroy hmOdf
            btrDestroy hmCbf
            btrDestroy hmSbf
            btrDestroy hmRdf
            Exit Sub
        End If
        imCifRecLen = Len(tmCif)
        
        hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmCpf)
            ilRet = btrClose(hmCif)
            ilRet = btrClose(hmMnf)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmCbf)
            ilRet = btrClose(hmOdf)
            ilRet = btrClose(hmSbf)
            ilRet = btrClose(hmRdf)
            btrDestroy hmCpf
            btrDestroy hmCif
            btrDestroy hmMnf
            btrDestroy hmClf
            btrDestroy hmCHF
            btrDestroy hmOdf
            btrDestroy hmCbf
            btrDestroy hmSbf
            btrDestroy hmRdf
            Exit Sub
        End If
        imCpfRecLen = Len(tmCpf)

        
        tmOdfSrchKey2.iGenDate(0) = igNowDate(0)
        tmOdfSrchKey2.iGenDate(1) = igNowDate(1)
        '10-10-01
        tmOdfSrchKey2.lGenTime = lgGenTime
        ilRet = btrGetGreaterOrEqual(hmOdf, tmOdfDay, imOdfRecLen, tmOdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmOdfDay.iGenDate(0) = igNowDate(0)) And (tmOdfDay.lGenTime = lgGenTime)
            tmCbf.iGenDate(0) = tmOdfDay.iGenDate(0)
            tmCbf.iGenDate(1) = tmOdfDay.iGenDate(1)
            tmCbf.lGenTime = lgGenTime
            tmCbf.iUrfCode = tmOdfDay.iUrfCode
            tmCbf.lIntComment = tmOdfDay.lFt1CefCode            'general log footer comment from vehicle
            tmCbf.lOtherComment = tmOdfDay.lFt2CefCode          'general log footer comment from vehicle
'            tmCbf.lLineComment = tmOdfDay.lHd1CefCode           'general Log header comment from vehicle
            tmCbf.lChgRComment = tmOdfDay.lHd1CefCode           '8-5-16 chg general Log header comment from vehicle
            tmCbf.sSortField2 = ""                          'init isci field
            tmCbf.sSurvey = ""                              'init creative title field
            tmCbf.iStartDate(0) = 0                         'rotation start date
            tmCbf.iStartDate(1) = 0
           'get the contract header
            If tmOdfDay.iType = 4 Then          'air time
                tmClfSrchKey2.lCode = tmOdfDay.lClfCode
                ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                gLogBtrError ilRet, "gGenL87Master: btrGetEqualClf"
                tmCbf.lLineComment = tmClf.lCxfCode         '8-5-16 line comments to show on l87
                'get the Daypart name
                tmRdfSrchKey.iCode = tmClf.iRdfCode
                ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                gLogBtrError ilRet, "gGenL87Master: btrGetEqualRdf"
                tmCbf.sSortField1 = Trim$(tmRdf.sName)          'DP name to be compared with theoverrirde field in crystal
                tmCbf.sDysTms = Trim$(tmOdfDay.sDPDesc)         'dp name, if override applies, it has override info
        
                tmCifSrchKey.lCode = tmOdfDay.lCifCode
                ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    tmCbf.iStartDate(0) = tmCif.iRotStartDate(0)
                    tmCbf.iStartDate(1) = tmCif.iRotStartDate(1)
                    tmCpfSrchKey.lCode = tmCif.lcpfCode
                    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        tmCbf.sSortField2 = tmCpf.sISCI
                        tmCbf.sSurvey = tmCpf.sCreative
                    End If
                End If
                
                tmCbf.lChfCode = tmClf.lChfCode
                tmCbf.iVefCode = tmOdfDay.iVefCode      'log vehicle, else same as the alternate vehicle
                tmCbf.iVehSort = tmOdfDay.iAlternateVefCode     'conventional vehicle
                tmCbf.iLen = tmClf.iLen                 'spot length
                tmCbf.sType = "A"                       'flag air time
                tmCbf.sResort = tmOdfDay.sZone             'time zone
                tmCbf.sAudioType = tmClf.sLiveCopy      'audio type
                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                gLogBtrError ilRet, "gGenL87Master: btrInsertCBf"
                
            ElseIf tmOdfDay.iType = 5 Then      'NTR
                tmSbfSrchKey1.lCode = tmOdfDay.lClfCode         'internal code for the NTR
                ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                gLogBtrError ilRet, "gGenL87Master: btrGetEqualSbf"
                
                'get the SBFType
                tmMnfSrchKey.iCode = tmSbf.iMnfItem
                ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                gLogBtrError ilRet, "gGenL87Master: btrGetEqualMnf"
                
                tmCbf.sSortField1 = Trim$(tmMnf.sName)      'show the ntr item in only Placement column (DP column).  two fields will be compared
                                                            'in crystal and if same, only show once.  One field is for override info for air time
                tmCbf.sDysTms = Trim$(tmMnf.sName)
                tmCbf.sSortField2 = ""                      'NTR doesnt have isci
                tmCbf.sSurvey = ""                          'NTR doesnt have creative title
                tmCbf.lChfCode = tmSbf.lChfCode
                tmCbf.iVefCode = tmOdfDay.iVefCode
                tmCbf.iVehSort = tmOdfDay.iAlternateVefCode
                tmCbf.iLen = 0
                tmCbf.sType = "N"
                tmCbf.sAudioType = ""
                tmCbf.sResort = tmOdfDay.sZone             'time zone
                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                gLogBtrError ilRet, "gGenL87Master: btrInsertCBf"
            End If
            ilRet = btrGetNext(hmOdf, tmOdfDay, imOdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        
        ilRet = btrClose(hmCpf)
        ilRet = btrClose(hmCif)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmOdf)
        ilRet = btrClose(hmSbf)
        ilRet = btrClose(hmRdf)
        btrDestroy hmCpf
        btrDestroy hmCif
        btrDestroy hmMnf
        btrDestroy hmClf
        btrDestroy hmCHF
        btrDestroy hmOdf
        btrDestroy hmCbf
        btrDestroy hmSbf
        btrDestroy hmRdf
        Exit Sub

End Sub
