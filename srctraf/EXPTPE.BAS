Attribute VB_Name = "EXPTPESubs"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Extract.BAS
'
' Release: 1.0
'
' Description:
'  This file contains the Extract subs and functions
Option Explicit
Option Compare Text
Type BASEREC
    sChar(1 To 20000) As Byte   'string * 1 'Record
End Type
Type PKAGELNCHECK
    iPkLineNo As Integer
    iInclude As Integer
End Type
'Public tgCxf As CXF
Public tgPkageLnCheck() As PKAGELNCHECK
Public igPECode As Integer      'Participant Extraction Code
Public igPECSRNVefCode() As Integer 'Participant Extraction Conventional/Selling Vehicles
Public igPEAVefCode() As Integer 'Participant Extraction Airing Vehicles
Public igPEPVefCode() As Integer 'Participant Extraction Package Vehicles
Public igPELVefCode() As Integer 'Participant Extraction Log Vehicles
Public tgPEMnf() As MNF         'Participant MNF records
Public lgChfCode() As Long      'contract codes
Public lgLCFAirStartDate() As Long
Public lgSSFAirStartDate() As Long
Public igPELtfCode() As Integer     'Library Title records extracted
Public lgPELvfCode() As Long        'Library version records extracted
Public igAddPkVefCode() As Integer  'Package vehicles required to be added
Public igPECntrAdfCode() As Integer
Public igPECntrAgfCode() As Integer
Public igPECntrSlfCode() As Integer

Type COPYUSED
    lCifCode As Long
    iMcfCode As Integer
    sUsed As String * 1
    sPurged As String * 1
End Type
Public tgCopyUsed() As COPYUSED


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
Function mObtainCifCode(tlSdf As SDF, slZone As String, hlTzf As Integer, ilOther As Integer) As Long
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
        mObtainCifCode = tlSdf.lCopyCode
        Exit Function
    ElseIf tlSdf.sPtType = "2" Then  '  Combo Copy
    ElseIf tlSdf.sPtType = "3" Then  '  Time Zone Copy
        ' Read TZF using lCopyCode from SDF
        ilTzfRecLen = Len(tlTzf)
        tlTzfSrchKey.lCode = tlSdf.lCopyCode
        ilRet = btrGetEqual(hlTzf, tlTzf, ilTzfRecLen, tlTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ' Look for the first positive lZone value
        For ilIndex = 1 To 6 Step 1
            If tlTzf.lCifZone(ilIndex) > 0 Then ' Process just the first positive Zone
                If StrComp(tlTzf.sZone(ilIndex), slZone, 1) = 0 Then
                    ' Read CIF using lCopyCode from SDF
                    mObtainCifCode = tlTzf.lCifZone(ilIndex)
                    If StrComp(slZone, "Oth", 1) <> 0 Then
                        ilOther = False
                    End If
                    Exit Function
                End If
            End If
        Next ilIndex
        For ilIndex = 1 To 6 Step 1
            If tlTzf.lCifZone(ilIndex) > 0 Then ' Process just the first positive Zone
                If StrComp(tlTzf.sZone(ilIndex), "Oth", 1) = 0 Then
                    ' Read CIF using lCopyCode from SDF
                    mObtainCifCode = tlTzf.lCifZone(ilIndex)
                    Exit Function
                End If
            End If
        Next ilIndex
    End If
    mObtainCifCode = 0
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainCrfCsfCode               *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Rotation Comment Code   *
'*                                                     *
'*******************************************************
Function mObtainCrfCsfCode(tlSdf As SDF, slZone As String, hlCrf As Integer, hlTzf As Integer, ilCrfVefCode As Integer, tlCrf As CRF) As Long
'
'   mObtainCrfCsfCode
'       Where:
'           tlSdf(I)- Spot record
'           hlCrf(I)- CRF Handle
'           hlTzf(I)- TZF Handle
'           mObtainCrfCsfCode(O)- Crf CsfCode or Zero if not found
'
    Dim ilRet As Integer
    Dim ilIndex As Integer
    'Time zone
    Dim tlTzf As TZF
    Dim ilTzfRecLen As Integer
    Dim tlTzfSrchKey As LONGKEY0
    'Copy rotation record information
    Dim tlCrfSrchKey1 As CRFKEY1 'CRF key record image
    Dim ilCrfRecLen As Integer  'CRF record length
    Dim ilRotNo As Integer
    Dim slType As String
    mObtainCrfCsfCode = 0
    ilRotNo = -1
    tlCrf.lCode = 0
    If tlSdf.sPtType = "3" Then  '  Time zone
        ilTzfRecLen = Len(tlTzf)
        tlTzfSrchKey.lCode = tlSdf.lCopyCode
        ilRet = btrGetEqual(hlTzf, tlTzf, ilTzfRecLen, tlTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ' Look for the first positive lZone value
        For ilIndex = 1 To 6 Step 1
            If tlTzf.lCifZone(ilIndex) > 0 Then ' Process just the first positive Zone
                If StrComp(tlTzf.sZone(ilIndex), slZone, 1) = 0 Then
                    ilRotNo = tlTzf.iRotNo(ilIndex)
                    Exit For
                End If
            End If
        Next ilIndex
        If ilRotNo = -1 Then
            Exit Function
        End If
    Else
        ilRotNo = tlSdf.iRotNo
    End If
    ilCrfRecLen = Len(tlCrf)
    slType = "A"
    tlCrfSrchKey1.sRotType = slType
    tlCrfSrchKey1.iEtfCode = 0
    tlCrfSrchKey1.iEnfCode = 0
    tlCrfSrchKey1.iAdfCode = tlSdf.iAdfCode
    tlCrfSrchKey1.lChfCode = tlSdf.lChfCode
    tlCrfSrchKey1.iVefCode = ilCrfVefCode   'tlSdf.iVefCode
    tlCrfSrchKey1.iRotNo = 32000
    ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tlSdf.iAdfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (tlCrf.iVefCode = ilCrfVefCode)    'tlSdf.iVefCode)
        If ilRotNo = tlCrf.iRotNo Then
            mObtainCrfCsfCode = tlCrf.lCsfCode
            Exit Function
        End If
        ilRet = btrGetNext(hlCrf, tlCrf, ilCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    tlCrf.lCode = 0
    Exit Function
End Function
