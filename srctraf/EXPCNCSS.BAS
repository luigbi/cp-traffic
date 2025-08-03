Attribute VB_Name = "ExpCNCSSSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of EXPCNCSS.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpCNCSS.BAS
'
' Release: 1.0
'
' Description:
'  This file contains the ExpCncSS subs and functions
Option Explicit
Option Compare Text
'Type SPOTTIMES
'    iVefCode As Integer
'    sCodeStn As String * 5
'    lAirDate As Long   'Air Date
'    lAvailTime As Long  'Avail Time
'    lNextSpotTime As Long   'Next spot time
'End Type
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainCifCode                  *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Copy                    *
'*                                                     *
'*******************************************************
Function mObtainCifCodeSS(tlSdf As SDF, slZone As String, hlTzf As Integer, ilOther As Integer) As Long
'
'   mObtainCifCode
'       Where:
'           tlSdf(I)- Spot record
'           slZone(I)-Zone
'           hlTzf(I)- TZF Handle
'           ilOther(O)- Output for other
'           mObtainCifCode(O)- Cif Code or Zero if not found
'
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim tlTzf As TZF
    Dim ilTzfRecLen As Integer
    Dim tlTzfSrchKey As LONGKEY0
    ilOther = True
    If tlSdf.sPtType = "1" Then  '  Single Copy
        ' Read CIF using lCopyCode from SDF
        mObtainCifCodeSS = tlSdf.lCopyCode
        Exit Function
    ElseIf tlSdf.sPtType = "2" Then  '  Combo Copy
    ElseIf tlSdf.sPtType = "3" Then  '  Time Zone Copy
        ' Read TZF using lCopyCode from SDF
        ilTzfRecLen = Len(tlTzf)
        tlTzfSrchKey.lCode = tlSdf.lCopyCode
        ilRet = btrGetEqual(hlTzf, tlTzf, ilTzfRecLen, tlTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ' Look for the first positive lZone value
        For ilIndex = 1 To 6 Step 1
            If tlTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                If StrComp(tlTzf.sZone(ilIndex - 1), slZone, 1) = 0 Then
                    ' Read CIF using lCopyCode from SDF
                    mObtainCifCodeSS = tlTzf.lCifZone(ilIndex - 1)
                    If StrComp(slZone, "Oth", 1) <> 0 Then
                        ilOther = False
                    End If
                    Exit Function
                End If
            End If
        Next ilIndex
        For ilIndex = 1 To 6 Step 1
            If tlTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                If StrComp(tlTzf.sZone(ilIndex - 1), "Oth", 1) = 0 Then
                    ' Read CIF using lCopyCode from SDF
                    mObtainCifCodeSS = tlTzf.lCifZone(ilIndex - 1)
                    Exit Function
                End If
            End If
        Next ilIndex
    End If
    mObtainCifCodeSS = 0
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainCrfCsfCodeSS             *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Rotation Comment Code   *
'*                                                     *
'*******************************************************
Function mObtainCrfCsfCodeSS(tlSdf As SDF, slZone As String, hlCrf As Integer, hlTzf As Integer, ilCrfVefCode As Integer, ilPkgVefCode As Integer, ilLnVefCode As Integer, tlCrf As CRF) As Long
'
'   mObtainCrfCsfCodeSS
'       Where:
'           tlSdf(I)- Spot record
'           hlCrf(I)- CRF Handle
'           hlTzf(I)- TZF Handle
'           mObtainCrfCsfCodeSS(O)- Crf CsfCode or Zero if not found
'
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim ilVefCode As Integer
    'Time zone
    Dim tlTzf As TZF
    Dim ilTzfRecLen As Integer
    Dim tlTzfSrchKey As LONGKEY0
    'Copy rotation record information
    Dim tlCrfSrchKey1 As CRFKEY1 'CRF key record image
    Dim ilCrfRecLen As Integer  'CRF record length
    Dim ilRotNo As Integer
    Dim slType As String
    mObtainCrfCsfCodeSS = 0
    ilRotNo = -1
    tlCrf.lCode = 0
    If tlSdf.sPtType = "3" Then  '  Time zone
        ilTzfRecLen = Len(tlTzf)
        tlTzfSrchKey.lCode = tlSdf.lCopyCode
        ilRet = btrGetEqual(hlTzf, tlTzf, ilTzfRecLen, tlTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ' Look for the first positive lZone value
        For ilIndex = 1 To 6 Step 1
            If tlTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                If StrComp(tlTzf.sZone(ilIndex - 1), slZone, 1) = 0 Then
                    ilRotNo = tlTzf.iRotNo(ilIndex - 1)
                    Exit For
                End If
            End If
        Next ilIndex
        If ilRotNo = -1 Then
            Exit Function
        End If
    Else
        If (tlSdf.sPtType <> "1") And (tlSdf.sPtType <> "2") Then
            Exit Function
        End If
        ilRotNo = tlSdf.iRotNo
    End If
    ilCrfRecLen = Len(tlCrf)
    ilVefCode = ilCrfVefCode
    Do
        slType = "A"
        tlCrfSrchKey1.sRotType = slType
        tlCrfSrchKey1.iEtfCode = 0
        tlCrfSrchKey1.iEnfCode = 0
        tlCrfSrchKey1.iAdfCode = tlSdf.iAdfCode
        tlCrfSrchKey1.lChfCode = tlSdf.lChfCode
        tlCrfSrchKey1.lFsfCode = 0
        tlCrfSrchKey1.iVefCode = ilVefCode  'ilCrfVefCode   'tlSdf.iVefCode
        tlCrfSrchKey1.iRotNo = ilRotNo  '32000
        'ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
        'Do While (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tlSdf.iAdfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (tlCrf.iVefCode = ilVefCode) 'And (tlCrf.iVefCode = ilCrfVefCode)    'tlSdf.iVefCode)
        '    If ilRotNo = tlCrf.iRotNo Then
        '        mObtainCrfCsfCodeSS = tlCrf.lCsfCode
        '        Exit Function
        '    End If
        '    ilRet = btrGetNext(hlCrf, tlCrf, ilCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        'Loop
        ilRet = btrGetEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            mObtainCrfCsfCodeSS = tlCrf.lCsfCode
            Exit Function
        End If
        If ilPkgVefCode > 0 Then
            ilVefCode = ilPkgVefCode
        Else
            If (ilCrfVefCode = ilLnVefCode) Or (ilLnVefCode = 0) Then
                Exit Do
            End If
            ilVefCode = ilLnVefCode
            ilLnVefCode = 0
        End If
        ilPkgVefCode = 0
    Loop While ilVefCode > 0
    tlCrf.lCode = 0
    Exit Function
End Function
