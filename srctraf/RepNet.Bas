Attribute VB_Name = "RepNet"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RepNet.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions for the Rep-Net Link interface
Option Explicit
Public igNrfIndex As Integer
'Contract header
Public tgRNChfSrchKey0 As LONGKEY0
Public tgRNChfSrchKey1 As CHFKEY1    'chf key record image
Public hgRNChf As Integer    'Contract header file handle
Public igRNChfRecLen As Integer        'Chf record length
Public tgRNChf As CHF

'Contract Lines
Public tgRNClf() As CLF        'Clf record image
Public tgRNClfSrchKey0 As CLFKEY0
Public tgRNClfSrchKey1 As CLFKEY1
Public tgRNClfSrchKey2 As LONGKEY0    'Clf key record image
Public hgRNClf As Integer    'Contract Line file handle
Public igRNClfRecLen As Integer        'CLF record length

'Contract Lines flights
Public tgRNCff() As CFF        'Cff record image
Public tgRNCffSrchKey0 As CFFKEY0
Public hgRNCff As Integer    'Contract Line file handle
Public igRNCffRecLen As Integer        'CfF record length

'Contract Lines vehicles
Public tgRNVsf() As VSF        'Clf record image
Public tgRNVsfSrchKey0 As LONGKEY0
Public hgRNVsf As Integer    'Contract Line file handle
Public igRNVsfRecLen As Integer        'CLF record length

'Salesperson
Public tgRNSlf As SLF        'Slf record image
Public tgRNSlfSrchKey0 As INTKEY0
Public hgRNSlf As Integer    'Salesperson file handle
Public igRNSlfRecLen As Integer        'SLF record length

'Sales Office
Public tgRNSof As SOF        'Sof record image
Public tgRNSofSrchKey0 As INTKEY0
Public hgRNSof As Integer    'Sales office file handle
Public igRNSofRecLen As Integer        'SOF record length

'Vehicle
Public tgRNVef As VEF        'Clf record image
Public tgRNVefSrchKey0 As INTKEY0
Public hgRNVef As Integer    'Contract Line file handle
Public igRNVefRecLen As Integer        'CLF record length

'Gather Info file (Used to communicate to stand alone program that contract should be sent)
Public tgRNGif As GIF            'GIF record image
Public tgRNGifSrchKey0 As LONGKEY0 'GIF key record image (agency)
Public tgRNGifSrchKey1 As GIFKEY1
Public hgRNGif As Integer        'GIF Handle
Public igRNGifRecLen As Integer     'GIF record length

Public tgRNQif As QIF            'QIF record image
Public tgRNQifSrchKey0 As LONGKEY0 'QIF key record image (agency)
Public tgRNQifSrchKey1 As QIFKEY1
Public hgRNQif As Integer        'QIF Handle
Public igRNQifRecLen As Integer     'QIF record length

Public hgRNNrf As Integer 'Avail name file handle
Public tgRNNrfSrchKey As INTKEY0    'ANF key record image
Public igRNNrfRecLen As Integer        'ANF record length

'Advertiser
Public hgRNAdf As Integer
Public tgRNAdf As ADF            'ADF record image
Public tgRNAdfSrchKey As INTKEY0 'ADF key record image
Public igRNAdfRecLen As Integer     'ADF record length

'Agency
Public hgRNAgf As Integer
Public tgRNAgf As AGF            'AGF record image
Public tgRNAgfSrchKey As INTKEY0 'AGF key record image
Public igRNAgfRecLen As Integer     'AGF record length

'Multi-Name
Public hgRNMnf As Integer
Public tgRNMnf As MNF            'MNF record image
Public tgRNMnfSrchKey As INTKEY0 'MNF key record image
Public tgRNMnfSrchKey3 As MNFKEY3
Public igRNMnfRecLen As Integer     'MNF record length

'Avail Name
Public hgRNAnf As Integer
Public tgRNAnf As ANF            'ANF record image
Public tgRNAnfSrchKey As INTKEY0 'ANF key record image
Public igRNAnfRecLen As Integer     'ANF record length

'Comment record-Header/Line
Public hgRNCxf As Integer
Public tgRNCxf As CXF            'CXF record image
Public tgRNCxfSrchKey As LONGKEY0  'CXF key record image
Public igRNCxfRecLen As Integer      'CXF record length

'Map Info
Public hgRNMif As Integer
Public tgRNMif As MIF            'MNF record image
Public tgRNMifSrchKey As LONGKEY0 'MNF key record image
Public tgRNMifSrchKey1 As MIFKEY1
Public tgRNMifSrchKey2 As MIFKEY2
Public igRNMifRecLen As Integer     'MNF record length

'Rate Card
Public hgRNRcf As Integer
Public tgRNRcf As RCF            'RCF record image
Public tgRNRcfSrchKey As INTKEY0 'RCF key record image
Public igRNRcfRecLen As Integer     'RCF record length

'Daypart
Public hgRNRdf As Integer
Public tgRNRdf As RDF            'RDF record image
Public tgRNRdfSrchKey As INTKEY0 'RDF key record image
Public igRNRdfRecLen As Integer     'RDF record length

'Comment for Message
Public hgRNCef As Integer
Public tgRNCef As CEF            'CEF record image
Public tgRNCefSrchKey As LONGKEY0 'CEF key record image
Public igRNCefRecLen As Integer     'CEF record length

'Transmit Message
Public hgRNSmf As Integer
Public tgRNSmf As SMF            'SMF record image
Public tgRNSmfSrchKey0 As SMFKEY0 'SMF key record image
Public tgRNSmfSrchKey1 As LONGKEY0
Public tgRNSmfSrchKey2 As LONGKEY0
Public igRNSmfRecLen As Integer     'SMF record length

'Alert Message
Public hgRNAuf As Integer
Public tgRNAuf As AUF            'AUF record image
Public tgRNAufSrchKey As LONGKEY0 'AUF key record image
Public igRNAufRecLen As Integer     'AUF record length

'Detail Spots
Public tgRNSdfSrchKey3 As LONGKEY0 'SDF key record image
Public tgRNSdfSrchKey4 As SDFKEY4    'sdf key record image
Public hgRNSdf As Integer    'Spot Detail file handle
Public igRNSdfRecLen As Integer        'Sdf record length
Public tgRNSdf As SDF

'Detail Spots
Public tgRNSbfSrchKey0 As SBFKEY0    'sbf key record image
Public hgRNSbf As Integer    'Special bill file handle
Public igRNSbfRecLen As Integer        'Sbf record length
Public tgRNSbf As SBF

'Copy Inventory
Public tgRNCifSrchKey0 As LONGKEY0    'cif key record image
Public hgRNCif As Integer    'Copy Inventory file handle
Public igRNCifRecLen As Integer        'Cif record length
Public tgRNCif As CIF

'Copy Product/ISCI
Public tgRNCpfSrchKey0 As LONGKEY0    'cpf key record image
Public tgRNCpfSrchKey1 As CPFKEY1
Public hgRNCpf As Integer    'Copy Product/ISCI file handle
Public igRNCpfRecLen As Integer        'Cpf record length
Public tgRNCpf As CPF

'Avail Summary
Public tgRNAsfSrchKey0 As INTKEY0    'asf key record image
Public hgRNAsf As Integer    'Avail Summary file handle
Public igRNAsfRecLen As Integer        'Asf record length
Public tgRNAsf As ASF

Public smHdFieldNames(0 To 70) As String
Public smLnFieldNames(0 To 75) As String
Private smFields(0 To 74) As String

Private smRepDBID As String
Private smNetDBID As String
Private smCntrNo As String
Private smAdvertiserName As String
Private smVehicleName As String
Private smProdProt1 As String
Private smProdProt2 As String
'Contract header special fields
Private smCommentInt As String
Private smShowOnInt As String
Private smCommentChgR As String
Private smShowOnChgR As String
Private smCommentCanc As String
Private smShowOnCanc As String
Private smCommentOther As String
Private smShowOnOther As String
'Contract line special fields
Private smLnComment() As String * 5000
Private smShowOnLnComment() As String
Private imProdProt1Req As Integer 'Product protection 1 required
Private imProdProt2Req As Integer 'Product protection 2 required

Private tmSbf() As SBF
Private tmClf As CLF
Private lmInvStdEndDate As Long
Private lmInvCalEndDate As Long

Type NETCHFSORT
    lChfCode As Long
End Type

Public Function gCreateContractHeaderRec(slRepDBID As String, slNetDBID As String, llGenDate As Long, llGenTime As Long, llSeqNo As Long, tlChf As CHF, slSlspFirstName As String, slSlspLastName As String) As String
    Dim ilRet As Integer
    Dim slRecord As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slShowNo As String
    
    gCreateContractHeaderRec = ""
    slRecord = ""
    
    slRecord = mAddQuotes(slRepDBID) & ","
    
    slRecord = slRecord & mAddQuotes(slNetDBID) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llGenDate))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llGenTime))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llSeqNo))) & ","
    slRecord = slRecord & mAddQuotes(tlChf.sType) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlChf.lCntrNo))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlChf.iExtRevNo))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlChf.iCntRevNo))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlChf.iPropVer))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlChf.iAdfCode))) & ","
    tgRNAdfSrchKey.iCode = tlChf.iAdfCode
    ilRet = btrGetEqual(hgRNAdf, tgRNAdf, igRNAdfRecLen, tgRNAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Advertiser(adfCode = " & Trim$(Str$(tlChf.iAdfCode)) & ", Pervasive Error # " & ilRet, True
        Exit Function
    End If
    slRecord = slRecord & mAddQuotes(tgRNAdf.sName) & ","
    slRecord = slRecord & mAddQuotes(tgRNAdf.sAbbr) & ","
    slRecord = slRecord & mAddQuotes(tgRNAdf.sBillAgyDir) & ","
    slRecord = slRecord & mAddQuotes(tgRNAdf.sAddrID) & ","
    slRecord = slRecord & mAddQuotes(tgRNAdf.sCntrAddr(0)) & ","
    slRecord = slRecord & mAddQuotes(tgRNAdf.sCntrAddr(1)) & ","
    slRecord = slRecord & mAddQuotes(tgRNAdf.sCntrAddr(2)) & ","
    slRecord = slRecord & mAddQuotes(tgRNAdf.sBillAddr(0)) & ","
    slRecord = slRecord & mAddQuotes(tgRNAdf.sBillAddr(1)) & ","
    slRecord = slRecord & mAddQuotes(tgRNAdf.sBillAddr(2)) & ","
    slRecord = slRecord & mAddQuotes(tlChf.sProduct) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlChf.iAgfCode))) & ","
    
    If tgSpf.sInsertAddr = "P" Then
        If tlChf.iAgfCode > 0 Then
            tgRNAgfSrchKey.iCode = tlChf.iAgfCode
            ilRet = btrGetEqual(hgRNAgf, tgRNAgf, igRNAgfRecLen, tgRNAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Agency(agfCode = " & Trim$(Str$(tlChf.iAgfCode)) & ", Pervasive Error # " & ilRet, True
                Exit Function
            End If
            slRecord = slRecord & mAddQuotes(tgRNAgf.sName) & ","
            slRecord = slRecord & mAddQuotes(tgRNAgf.sAbbr) & ","
            slRecord = slRecord & mAddQuotes(tgRNAgf.sCityID) & ","
            slRecord = slRecord & mAddQuotes(tgRNAgf.sCntrAddr(0)) & ","
            slRecord = slRecord & mAddQuotes(tgRNAgf.sCntrAddr(1)) & ","
            slRecord = slRecord & mAddQuotes(tgRNAgf.sCntrAddr(2)) & ","
            slRecord = slRecord & mAddQuotes(tgRNAgf.sBillAddr(0)) & ","
            slRecord = slRecord & mAddQuotes(tgRNAgf.sBillAddr(1)) & ","
            slRecord = slRecord & mAddQuotes(tgRNAgf.sBillAddr(2)) & ","
            slRecord = slRecord & mAddQuotes(Trim$(Str$(tgRNAgf.iComm))) & ","
        Else
            'D.S. 3/27/08
           ' slRecord = slRecord & "0" & ","
            slRecord = slRecord & """" & """" & ","
            slRecord = slRecord & """" & """" & ","
            slRecord = slRecord & """" & """" & ","
            slRecord = slRecord & """" & """" & ","
            slRecord = slRecord & """" & """" & ","
            slRecord = slRecord & """" & """" & ","
            slRecord = slRecord & """" & """" & ","
            slRecord = slRecord & """" & """" & ","
            slRecord = slRecord & """" & """" & ","
            slRecord = slRecord & """" & """" & ","
        End If
    Else
        tgRNSlfSrchKey0.iCode = tlChf.iSlfCode(0)
        ilRet = btrGetEqual(hgRNSlf, tgRNSlf, igRNSlfRecLen, tgRNSlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Salesperson(slfCode = " & Trim$(Str$(tlChf.iSlfCode(0))) & ", Pervasive Error # " & ilRet, True
            Exit Function
        End If
        tgRNSofSrchKey0.iCode = tgRNSlf.iSofCode
        ilRet = btrGetEqual(hgRNSof, tgRNSof, igRNSofRecLen, tgRNSofSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Sales Office(sofCode = " & Trim$(Str$(tgRNSlf.iSofCode)) & ", Pervasive Error # " & ilRet, True
            Exit Function
        End If
        slRecord = slRecord & mAddQuotes(tgSpf.sBPayName) & ","             'Name
        slRecord = slRecord & mAddQuotes(Left$(tgSpf.sBPayName, 5)) & ","   'Abbreviation
        '7/14/10:  Replace SOFName with PayName so that all agency $ reports will be to the same agency on the Network.
        'slRecord = slRecord & mAddQuotes(Left$(tgRNSof.sName, 5)) & ","     'City ID
        slRecord = slRecord & mAddQuotes(Left$(tgSpf.sBPayName, 5)) & ","   'City ID replacement
        slRecord = slRecord & mAddQuotes(tgSpf.sBPayAddr(0)) & ","
        slRecord = slRecord & mAddQuotes(tgSpf.sBPayAddr(1)) & ","
        slRecord = slRecord & mAddQuotes(tgSpf.sBPayAddr(2)) & ","
        slRecord = slRecord & """" & """" & ","
        slRecord = slRecord & """" & """" & ","
        slRecord = slRecord & """" & """" & ","
        slRecord = slRecord & mAddQuotes(Trim$(Str$(0))) & ","
    End If
    For ilLoop = 0 To 1 Step 1
        If tlChf.iMnfComp(ilLoop) > 0 Then
            tgRNMnfSrchKey.iCode = tlChf.iMnfComp(ilLoop)
            ilRet = btrGetEqual(hgRNMnf, tgRNMnf, igRNMnfRecLen, tgRNMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Product Protection(mnfCode = " & Trim$(Str$(tlChf.iMnfComp(ilLoop))) & ", Pervasive Error # " & ilRet, True
                Exit Function
            End If
            slRecord = slRecord & mAddQuotes(tgRNMnf.sName) & ","
        Else
            slRecord = slRecord & """" & """" & ","
        End If
    Next ilLoop
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlChf.iPctTrade))) & ","
    For ilLoop = 0 To 3 Step 1
        If tlChf.iMnfDemo(ilLoop) > 0 Then
            tgRNMnfSrchKey.iCode = tlChf.iMnfDemo(ilLoop)
            ilRet = btrGetEqual(hgRNMnf, tgRNMnf, igRNMnfRecLen, tgRNMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Demo Name(mnfCode = " & Trim$(Str$(tlChf.iMnfDemo(ilLoop))) & ", Pervasive Error # " & ilRet, True
                Exit Function
            End If
            slRecord = slRecord & mAddQuotes(tgRNMnf.sName) & ","
        Else
            slRecord = slRecord & """" & """" & ","
        End If
    Next ilLoop
    slRecord = slRecord & mAddQuotes(slSlspFirstName) & ","
    slRecord = slRecord & mAddQuotes(slSlspLastName) & ","
    slRecord = slRecord & mAddQuotes(tlChf.sBuyer) & ","
    slRecord = slRecord & mAddQuotes(Trim$(tlChf.sPhone)) & ","
    For ilLoop = 0 To 3 Step 1
        Select Case ilLoop
            Case 0
                slRecord = slRecord & mAddQuotes(gGetComment(tlChf.lCxfInt, slShowNo)) & ","
                slRecord = slRecord & mAddQuotes(slShowNo) & ","
            Case 1
                slRecord = slRecord & mAddQuotes(gGetComment(tlChf.lCxfChgR, slShowNo)) & ","
                slRecord = slRecord & mAddQuotes(slShowNo) & ","
            Case 2
                slRecord = slRecord & mAddQuotes(gGetComment(tlChf.lCxfCanc, slShowNo)) & ","
                slRecord = slRecord & mAddQuotes(slShowNo) & ","
            Case 3
                slRecord = slRecord & mAddQuotes(gGetComment(tlChf.lCxfCode, slShowNo)) & ","
                slRecord = slRecord & mAddQuotes(slShowNo)
        End Select
    Next ilLoop
    gCreateContractHeaderRec = slRecord
End Function


Public Function gCreateContractLineRec(slRepDBID As String, slNetDBID As String, llGenDate As Long, llGenTime As Long, tlChf As CHF, tlClf As CLF) As String
    Dim ilRet As Integer
    Dim slRecord As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilDay As Integer
    Dim ilPreDay As Integer
    Dim ilRdf As Integer
    Dim slRdfStartTime As String
    Dim slRdfEndTime As String
    Dim ilCount As Integer
    
    gCreateContractLineRec = ""
    slRecord = ""
    
    slRecord = mAddQuotes(slRepDBID) & ","
    slRecord = slRecord & mAddQuotes(slNetDBID) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llGenDate))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llGenTime))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlChf.lCntrNo))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlClf.iCntRevNo))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlClf.iPropVer))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlClf.iLine))) & ","
    tgRNVefSrchKey0.iCode = tlClf.iVefCode
    ilRet = btrGetEqual(hgRNVef, tgRNVef, igRNVefRecLen, tgRNVefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Line Vehicle(vefCode = " & Trim$(Str$(tlClf.iVefCode)) & ", Pervasive Error # " & ilRet, True
        Exit Function
    End If
    slRecord = slRecord & mAddQuotes(tgRNVef.sName) & ","
    'slStr = ""
    'tgRNMifSrchKey2.sType = "V"
    'tgRNMifSrchKey2.sRepDBID = ""
    'tgRNMifSrchKey2.lRepDBCode = tlClf.iVefCode
    'ilRet = btrGetGreaterOrEqual(hgRNMif, tgRNMif, igRNMifRecLen, tgRNMifSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    'Do While (ilRet = BTRV_ERR_NONE) And (tgRNMif.sType = "V")
    '    If StrComp(Trim$(tgRNMif.sNetDBID), Trim$(slNetDBID), vbTextCompare) = 0 Then
    '        If tgRNMif.lRepDBCode = tlClf.iVefCode Then
    '            slStr = Trim$(Str$(tgRNMif.lNetDBCode))
    '            Exit Do
    '        End If
    '    End If
    '    ilRet = btrGetNext(hgRNMif, tgRNMif, igRNMifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    'Loop
    'slRecord = slRecord & mAddQuotes(Trim$(slStr)) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlClf.iVefCode))) & ","
    
    tgRNRdfSrchKey.iCode = tlClf.iRdfCode
    ilRet = btrGetEqual(hgRNRdf, tgRNRdf, igRNRdfRecLen, tgRNRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Daypart Name(rdfCode = " & Trim$(Str$(tlClf.iRdfCode)) & ", Pervasive Error # " & ilRet, True
        Exit Function
    End If
    slRecord = slRecord & mAddQuotes(tgRNRdf.sName) & ","
    'slStr = ""
    'tgRNMifSrchKey2.sType = "D"
    'tgRNMifSrchKey2.sRepDBID = ""
    'tgRNMifSrchKey2.lRepDBCode = tlClf.iRdfcode
    'ilRet = btrGetGreaterOrEqual(hgRNMif, tgRNMif, igRNMifRecLen, tgRNMifSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    'Do While (ilRet = BTRV_ERR_NONE) And (tgRNMif.sType = "D")
    '    If StrComp(Trim$(tgRNMif.sNetDBID), Trim$(slNetDBID), vbTextCompare) = 0 Then
    '        If tgRNMif.lRepDBCode = tlClf.iRdfcode Then
    '            slStr = Trim$(Str$(tgRNMif.lNetDBCode))
    '            Exit Do
    '        End If
    '    End If
    '    ilRet = btrGetNext(hgRNMif, tgRNMif, igRNMifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    'Loop
    'slRecord = slRecord & mAddQuotes(Trim$(slStr)) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlClf.iRdfCode))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(tgRNRdf.sInOut)) & ","
    If tgRNRdf.ianfCode > 0 Then
        tgRNAnfSrchKey.iCode = tgRNRdf.ianfCode
        ilRet = btrGetEqual(hgRNAnf, tgRNAnf, igRNAnfRecLen, tgRNAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Avail Name(anfCode = " & Trim$(Str$(tgRNRdf.ianfCode)) & " failed, Pervasive Error # " & ilRet, True
            Exit Function
        End If
        slRecord = slRecord & mAddQuotes(tgRNAnf.sName) & ","
    Else
        slRecord = slRecord & "" & ","
    End If
    ilCount = 0
    For ilRdf = LBound(tgRNRdf.iStartTime, 2) To UBound(tgRNRdf.iStartTime, 2) Step 1
        If (tgRNRdf.iStartTime(0, ilRdf) <> 1) Or (tgRNRdf.iStartTime(1, ilRdf) <> 0) Then
            slStr = ""
            For ilDay = LBound(tgRNRdf.sWkDays, 1) To UBound(tgRNRdf.sWkDays, 1) Step 1
                slStr = slStr & tgRNRdf.sWkDays(ilRdf, ilDay)
            Next ilDay
            slRecord = slRecord & slStr & ","
            gUnpackTime tgRNRdf.iStartTime(0, ilRdf), tgRNRdf.iStartTime(1, ilRdf), "A", "1", slRdfStartTime
            gUnpackTime tgRNRdf.iEndTime(0, ilRdf), tgRNRdf.iEndTime(1, ilRdf), "A", "1", slRdfEndTime
            slRecord = slRecord & mAddQuotes(Trim$(slRdfStartTime)) & ","
            slRecord = slRecord & mAddQuotes(Trim$(slRdfEndTime)) & ","
            slRecord = slRecord & Trim$(Str$(tgRNRdf.iSpotPct(ilRdf))) & ","
            ilCount = ilCount + 1
            If ilCount = 2 Then
                Exit For
            End If
        End If
    Next ilRdf
    If ilCount = 0 Then
        slRecord = slRecord & "" & ","
        slRecord = slRecord & "" & ","
        slRecord = slRecord & "" & ","
        slRecord = slRecord & "" & ","
        ilCount = 1
    End If
    If ilCount = 1 Then
        slRecord = slRecord & "" & ","
        slRecord = slRecord & "" & ","
        slRecord = slRecord & "" & ","
        slRecord = slRecord & "" & ","
    End If
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlClf.iLen))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlClf.iBBOpenLen))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlClf.iBBCloseLen))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(tlClf.sSoloAvail)) & ","
    slRecord = slRecord & mAddQuotes(Trim$(tlClf.sOV2DefinedBits)) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlClf.iPosition))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlClf.iPctAllocation))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlClf.lAcquisitionCost))) & ","
    gUnpackDate tlClf.iStartDate(0), tlClf.iStartDate(1), slStr
    slStr = gAdjYear(slStr)
    slRecord = slRecord & mAddQuotes(Trim$(slStr)) & ","
    gUnpackDate tlClf.iEndDate(0), tlClf.iEndDate(1), slStr
    slStr = gAdjYear(slStr)
    slRecord = slRecord & mAddQuotes(Trim$(slStr)) & ","
    If ((tlClf.iStartTime(0) <> 1) Or (tlClf.iStartTime(1) <> 0)) And ((tlClf.iEndTime(0) <> 1) Or (tlClf.iEndTime(1) <> 0)) Then
        gUnpackTime tlClf.iStartTime(0), tlClf.iStartTime(1), "A", "1", slStr
        slRecord = slRecord & mAddQuotes(Trim$(slStr)) & ","
        gUnpackTime tlClf.iEndTime(0), tlClf.iEndTime(1), "A", "1", slStr
        slRecord = slRecord & mAddQuotes(Trim$(slStr)) & ","
    Else
        slRecord = slRecord & """" & """" & ","
        slRecord = slRecord & """" & """" & ","
    End If
    gUnpackTime tlClf.iPrefStartTime(0), tlClf.iPrefStartTime(1), "A", "1", slStr
    slRecord = slRecord & mAddQuotes(Trim$(slStr)) & ","
    gUnpackTime tlClf.iPrefEndTime(0), tlClf.iPrefEndTime(1), "A", "1", slStr
    slRecord = slRecord & mAddQuotes(Trim$(slStr)) & ","
    slStr = ""
    ilPreDay = False
    For ilDay = 0 To 6 Step 1
        If (Trim$(tlClf.sPrefDays(ilDay)) <> "N") And (Trim$(tlClf.sPrefDays(ilDay)) <> "") Then
            ilPreDay = True
            Exit For
        End If
    Next ilDay
    If ilPreDay Then
        For ilDay = 0 To 6 Step 1
            If tlClf.sPrefDays(ilDay) = "Y" Then
                slStr = slStr & "Y"
            Else
                slStr = slStr & "N"
            End If
        Next ilDay
    Else
        slStr = ""
    End If
    slRecord = slRecord & mAddQuotes(Trim$(slStr)) & ","
    slRecord = slRecord & mAddQuotes(Trim$(tlClf.sPreempt)) & ","
    slRecord = slRecord & mAddQuotes(Trim$(tlClf.sLiveCopy)) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlClf.iAdvtSepFlag))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlClf.iPriority))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(tlClf.sHideCBS)) & ","
    gCreateContractLineRec = slRecord
End Function


Public Function gCreateContractFlightRec(tlCff As CFF) As String
    Dim slRecord As String
    Dim slStr As String
    Dim ilLoop As Integer
    
    gCreateContractFlightRec = ""
    slRecord = ""
    gUnpackDate tlCff.iStartDate(0), tlCff.iStartDate(1), slStr
    slStr = gAdjYear(slStr)
    slRecord = slRecord & mAddQuotes(Trim$(slStr)) & ","
    gUnpackDate tlCff.iEndDate(0), tlCff.iEndDate(1), slStr
    slStr = gAdjYear(slStr)
    slRecord = slRecord & mAddQuotes(Trim$(slStr)) & ","
    slRecord = slRecord & mAddQuotes(Trim$(tlCff.sDyWk)) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlCff.iSpotsWk))) & ","
    For ilLoop = 0 To 6 Step 1
        slRecord = slRecord & mAddQuotes(Trim$(Str$(tlCff.iDay(ilLoop)))) & ","
    Next ilLoop
    slRecord = slRecord & mAddQuotes(Trim$(tlCff.sPriceType)) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlCff.lActPrice))) & ","
    gCreateContractFlightRec = slRecord
End Function
Public Function gGetComment(llCxfCode As Long, slShowOn As String) As String
    Dim ilRet As Integer
    Dim slComment As String
    Dim slStr As String
    Dim slStr1 As String
    
    slComment = ""
    slStr = ""
    tgRNCxfSrchKey.lCode = llCxfCode
    If tgRNCxfSrchKey.lCode > 0 Then
        igRNCxfRecLen = Len(tgRNCxf) '5027
        ilRet = btrGetEqual(hgRNCxf, tgRNCxf, igRNCxfRecLen, tgRNCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        slStr1 = gStripChr0(tgRNCxf.sComment)
        'If tgRNCxf.iStrLen > 0 Then
        If slStr1 <> "" Then
            ''slComment = mAddQuotes(Left$(tgRNCxf.sComment, tgRNCxf.iStrLen))
            'slComment = mAddQuotes(slStr1)
            slComment = slStr1
        End If
        If tgRNCxf.sShProp = "Y" Then
            slStr = "Y"
        Else
            slStr = "N"
        End If
        If tgRNCxf.sShOrder = "Y" Then
            slStr = slStr & "Y"
        Else
            slStr = slStr & "N"
        End If
        If tgRNCxf.sShSpot = "Y" Then
            slStr = slStr & "Y"
        Else
            slStr = slStr & "N"
        End If
        If tgRNCxf.sShInv = "Y" Then
            slStr = slStr & "Y"
        Else
            slStr = slStr & "N"
        End If
    End If
    slShowOn = slStr
    gGetComment = slComment
End Function
Public Function gPutComment(slInComment As String, slShowOn As String, llCxfCode As Long) As Integer
    Dim slComment As String
    Dim ilRet As Integer
    
    llCxfCode = 0
    slComment = Trim$(slInComment)
    If Trim$(slComment) <> "" Then
        tgRNCxf.sShProp = Mid$(slShowOn, 1, 1)
        tgRNCxf.sShOrder = Mid$(slShowOn, 2, 1)
        tgRNCxf.sShSpot = Mid$(slShowOn, 3, 1)
        tgRNCxf.sShInv = Mid$(slShowOn, 4, 1)
        'For now set to N
        tgRNCxf.sShInsertion = "N"
        'tgRNCxf.iStrLen = Len(slComment)
        tgRNCxf.sComment = Trim$(slComment) & Chr$(0) '& Chr$(0) 'sgTB
        igRNCxfRecLen = Len(tgRNCxf) '- Len(tgRNCxf.sComment) + Len(Trim$(tgRNCxf.sComment)) ' + 2    '25 = fixed record length; 2=Length value which is part of the variable record
        tgRNCxf.lCode = 0 'Autoincrement
        tgRNCxf.iRemoteID = 0
        tgRNCxf.lAutoCode = 0
        tgRNCxf.iSourceID = 0
        ilRet = btrInsert(hgRNCxf, tgRNCxf, igRNCxfRecLen, INDEXKEY0)
        If ilRet = BTRV_ERR_NONE Then
            tgRNCxf.iRemoteID = 0
            tgRNCxf.lAutoCode = tgRNCxf.lCode
            tgRNCxf.iSourceID = 0
            gPackDate Format$(gNow(), "m/d/yy"), tgRNCxf.iSyncDate(0), tgRNCxf.iSyncDate(1)
            gPackTime Format$(gNow(), "h:mm:ssAM/PM"), tgRNCxf.iSyncTime(0), tgRNCxf.iSyncTime(1)
            igRNCxfRecLen = Len(tgRNCxf) '- Len(tgRNCxf.sComment) + Len(Trim$(tgRNCxf.sComment))
            ilRet = btrUpdate(hgRNCxf, tgRNCxf, igRNCxfRecLen)
            If ilRet <> BTRV_ERR_NONE Then
                gPutComment = False
                Exit Function
            End If
            llCxfCode = tgRNCxf.lCode
        Else
            gPutComment = False
            Exit Function
        End If
    Else
        llCxfCode = 0
    End If
    gPutComment = True
End Function


Public Function gTextToContract(tlQif As QIF) As Integer
    Dim hlTextChf As Integer
    Dim hlTextClf As Integer
    Dim ilRet As Integer
    Dim slHeader As String
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilContractHeaderFound As Integer
    Dim ilContractLineFound As Integer
    Dim ilIndex As Integer
    Dim ilLineNo As Integer
    Dim ilLineOrFlightInError As Integer
    Dim ilContractInError As Integer
    
    ilRet = mOpenTextForRead(hlTextChf, tlQif.sFileName1)
    If Not ilRet Then
        gTextToContract = False
        Exit Function
    End If
    ilRet = mOpenTextForRead(hlTextClf, tlQif.sFileName2)
    If Not ilRet Then
        gTextToContract = False
        Exit Function
    End If
    ilRet = gParseItem(tlQif.sFileName1, 1, "_", smNetDBID)
    ilRet = gParseItem(tlQif.sFileName1, 2, "_", smRepDBID)
    smCntrNo = ""
    smAdvertiserName = ""
    smVehicleName = ""
    ilContractHeaderFound = False
    ilContractInError = False
    ilContractLineFound = False
    'Read and Build Contract header
    Do
        ilRet = 0
        err.Clear
        'On Error GoTo gTestToContractErr:
        Line Input #hlTextChf, slHeader
        On Error GoTo 0
        ilRet = err.Number
        If ilRet <> 0 Then
            Exit Do
        End If
        If Len(slHeader) > 0 Then
            If (Asc(slHeader) = 26) Then     'Ctrl Z
                Exit Do
            Else
                If StrComp(Left$(slHeader, 3), "cht", vbTextCompare) = 0 Then
                    gParseCDFields slHeader, False, smHdFieldNames()
                    ilContractHeaderFound = True
                ElseIf ilContractHeaderFound Then
                    gParseCDFields slHeader, False, smFields()
                    ilIndex = mFindFieldName(smHdFieldNames(), "chtNetDBID")
                    If ilIndex <> -1 Then
                        smNetDBID = Trim$(smFields(ilIndex))
                        If StrComp(Trim$(smFields(ilIndex)), Trim$(tgNrf.sDBID), vbTextCompare) = 0 Then
                            ilIndex = mFindFieldName(smHdFieldNames(), "chtRepDBID")
                            If ilIndex <> -1 Then
                                smRepDBID = Trim$(smFields(ilIndex))
                            Else
                                smRepDBID = ""
                            End If
                            'Create contract header
                            ilRet = mCreateContract()
                            ilContractInError = Not ilRet
                            If ilRet Then
                                ReDim tgRNClf(0 To 0) As CLF
                                ReDim smLnComment(0 To 0) As String * 5000
                                ReDim smShowOnLnComment(0 To 0) As String
                                ReDim tgRNCff(0 To 0) As CFF
                                ilContractLineFound = False
                                ilLineOrFlightInError = False
                                'Create Lines and Header for the contract
                                Do
                                    ilRet = 0
                                    err.Clear
                                    'On Error GoTo gTestToContractErr:
                                    Line Input #hlTextClf, slLine
                                    On Error GoTo 0
                                    ilRet = err.Number
                                    If ilRet <> 0 Then
                                        Exit Do
                                    End If
                                    If Len(slLine) > 0 Then
                                        If (Asc(slLine) = 26) Then     'Ctrl Z
                                            Exit Do
                                        Else
                                            If (StrComp(Left$(slLine, 3), "clt", vbTextCompare) = 0) And (Not ilContractLineFound) Then
                                                gParseCDFields slLine, False, smLnFieldNames()
                                                ilContractLineFound = True
                                                ilLineNo = -1
                                            ElseIf ilContractLineFound Then
                                                gParseCDFields slLine, False, smFields()
                                                ilIndex = mFindFieldName(smLnFieldNames(), "cltNetDBID")
                                                If ilIndex <> -1 Then
                                                    If StrComp(Trim$(smFields(ilIndex)), Trim$(tgNrf.sDBID), vbTextCompare) = 0 Then
                                                        'Create contract header
                                                        ilIndex = mFindFieldName(smLnFieldNames(), "cltLineNo")
                                                        If ilIndex <> -1 Then
                                                            If Val(smFields(ilIndex)) <> ilLineNo Then
                                                                ilRet = mCreateLine()
                                                                If Not ilRet Then
                                                                    ilLineOrFlightInError = True
                                                                    Exit Do
                                                                End If
                                                                ilLineNo = Val(smFields(ilIndex))
                                                            Else
                                                                ilRet = True
                                                            End If
                                                            If ilRet Then
                                                                ilRet = mCreateFlight()
                                                                If Not ilRet Then
                                                                    ilLineOrFlightInError = True
                                                                    Exit Do
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Loop Until ilEof
                                If Not ilLineOrFlightInError Then
                                    If UBound(tgRNClf) > 0 Then
                                        ilRet = btrBeginTrans(hgRNChf, 1000)
                                        ilRet = mClearContracts()
                                        If Not ilRet Then
                                            ilRet = btrAbortTrans(hgRNChf)
                                        Else
                                            tgRNChf.lAirTimeGross = mSetChfTotalDollars()
                                            tgRNChf.lInputGross = tgRNChf.lAirTimeGross
                                            tgRNChf.lVefCode = mSetHdVef()
                                            If tgRNChf.lVefCode <> 0 Then
                                                'Save Contract
                                                ilRet = mAddContract()
                                                If Not ilRet Then
                                                    ilRet = btrAbortTrans(hgRNChf)
                                                Else
                                                    ilRet = btrEndTrans(hgRNChf)
                                                End If
                                            Else
                                                ilRet = btrAbortTrans(hgRNChf)
                                            End If
                                        End If
                                    Else
                                        gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "||" & "Contract Not Added as No Lines found.", True
                                    End If
                                End If
                            Else
                                gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "||" & "Contract Not Added, see Previous message.", True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Loop Until ilEof
    If ilContractHeaderFound = False Then
        gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & "|" & "||" & "Contract Not Added as Header not found in Text file " & Trim$(tlQif.sFileName1), True
    Else
        If ilContractInError Then
            gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "||" & "Contract Not Added.  See previous message. " & Trim$(tlQif.sFileName1), True
        Else
            If ilContractLineFound = False Then
                gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "||" & "Contract Not Added as Header or Lines not found in Text file " & Trim$(tlQif.sFileName2), True
            End If
        End If
    End If
    'Read and Build Contract lines and Flights
    Close hlTextChf
    Close hlTextClf
    gTextToContract = True
    Exit Function
'gTestToContractErr:
'    ilRet = Err.Number
'    Resume Next
End Function

Private Function mOpenTextForRead(hlText As Integer, slFileName As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    'On Error GoTo mOpenTextForReadErr:
    slToFile = sgImportPath & slFileName
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        'hlText = FreeFile
        'Open slToFile For Input Access Read As hlText
        ilRet = gFileOpen(slToFile, "Input Access Read", hlText)
        If ilRet <> 0 Then
            gCreateMessage "O|E|" & Trim$(tgNrf.sDBID) & "|" & "|" & "||" & "Open Text file " & slFileName & " for Read failed, error #: " & ilRet, True
            mOpenTextForRead = False
            Exit Function
        End If
    Else
        gCreateMessage "O|E|" & Trim$(tgNrf.sDBID) & "|" & "|" & "||" & "Open Text file " & slFileName & " for Read failed, error #: " & ilRet, True
        Exit Function
    End If
    On Error GoTo 0
    mOpenTextForRead = True
    Exit Function
'mOpenTextForReadErr:
'    ilRet = Err.Number
'    Resume Next
End Function


Private Function mFindFieldName(slFieldNames() As String, slFindName As String) As Integer
    Dim ilLoop As Integer
    
    For ilLoop = LBound(slFieldNames) To UBound(slFieldNames) Step 1
        If StrComp(slFieldNames(ilLoop), Trim$(slFindName), vbTextCompare) = 0 Then
            mFindFieldName = ilLoop
            Exit Function
        End If
    Next ilLoop
    gCreateMessage "F|E|" & Trim$(tgNrf.sDBID) & "|" & "|" & "||" & "Unable to find Field name " & slFindName, True
    mFindFieldName = -1
End Function

Private Function mCreateContract() As Integer
    Dim ilIndex As Integer
    Dim ilRepAdfCode As Integer
    Dim ilRepAgfCode As Integer
    Dim slDemoName(0 To 3) As String
    Dim slSlspFirstName As String
    Dim slSlspLastName As String
    Dim tlAdf As ADF
    Dim tlagf As AGF
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilOlfCntrFd As Integer
    Dim tlChf As CHF
    
    mCreateContract = False
    ilOlfCntrFd = False
    mInitChf
    If tgRNChf.iRcfCode <= 0 Then
        Exit Function
    End If
    mInitAdf
    mInitAgf
    tgRNChf.sRepDBID = smRepDBID
    ilIndex = mFindFieldName(smHdFieldNames(), "chtContractType")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNChf.sType = smFields(ilIndex)
    Select Case tgRNChf.sType
        Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
            tgRNChf.sSchStatus = "M"
            tgRNChf.sStatus = "O"
        Case Else   'Standard or Reservation
            'tgRNChf.sStatus = "N" 'N is type O but needs scheduling
            'tgRNChf.sSchStatus = "A"  '"N"
            'Jim: Change to Complete
            tgRNChf.sStatus = "C" 'Complete
            tgRNChf.sSchStatus = "P"  '"N"
    End Select
    ilIndex = mFindFieldName(smHdFieldNames(), "chtContractNo")
    If ilIndex = -1 Then
        Exit Function
    End If
    smCntrNo = smFields(ilIndex)
    tgRNChf.lCntrNo = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "chtExtRevNo")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNChf.iExtRevNo = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "chtIntVerNo")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNChf.iCntRevNo = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "chtProposalVersion")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNChf.iPropVer = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "chtAdvertiserCode")
    If ilIndex = -1 Then
        Exit Function
    End If
    ilRepAdfCode = Val(smFields(ilIndex))
    If ilRepAdfCode > 0 Then
        tgRNMifSrchKey2.sType = "A"
        tgRNMifSrchKey2.sRepDBID = smRepDBID
        tgRNMifSrchKey2.lRepDBCode = ilRepAdfCode
        ilRet = btrGetEqual(hgRNMif, tgRNMif, igRNMifRecLen, tgRNMifSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            tgRNAdfSrchKey.iCode = tgRNMif.lNetDBCode
            ilRet = btrGetEqual(hgRNAdf, tgRNAdf, igRNAdfRecLen, tgRNAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                tgRNChf.iAdfCode = tgRNAdf.iCode
            End If
        End If
    End If
    ilIndex = mFindFieldName(smHdFieldNames(), "chtAdvertiserName")
    If ilIndex = -1 Then
        Exit Function
    End If
    smAdvertiserName = smFields(ilIndex)
    tgRNAdf.sName = smFields(ilIndex)
    ilIndex = mFindFieldName(smHdFieldNames(), "chtAdvertiserAbbr")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNAdf.sAbbr = smFields(ilIndex)
    ilIndex = mFindFieldName(smHdFieldNames(), "chtAdvertiserType")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNAdf.sBillAgyDir = smFields(ilIndex)
    If Trim$(tgRNAdf.sBillAgyDir) = "D" Then
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAdvertiserAddrID")
        If ilIndex = -1 Then
            Exit Function
        End If
        smAdvertiserName = smAdvertiserName & "," & smFields(ilIndex)
        tgRNAdf.sAddrID = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAdvertiserCntrAddr_1")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAdf.sCntrAddr(0) = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAdvertiserCntrAddr_2")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAdf.sCntrAddr(1) = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAdvertiserCntrAddr_3")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAdf.sCntrAddr(2) = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAdvertiserBillAddr_1")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAdf.sBillAddr(0) = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAdvertiserBillAddr_2")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAdf.sBillAddr(1) = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAdvertiserBillAddr_3")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAdf.sBillAddr(2) = smFields(ilIndex)
    Else
        tgRNAdf.sAddrID = ""
        tgRNAdf.sCntrAddr(0) = ""
        tgRNAdf.sCntrAddr(1) = ""
        tgRNAdf.sCntrAddr(2) = ""
        tgRNAdf.sBillAddr(0) = ""
        tgRNAdf.sBillAddr(1) = ""
        tgRNAdf.sBillAddr(2) = ""
    End If
    If tgRNChf.iAdfCode > 0 Then
        ilRet = btrUpdate(hgRNAdf, tgRNAdf, igRNAdfRecLen)
    Else
        'Search for matching name
        ilRet = btrGetFirst(hgRNAdf, tlAdf, igRNAdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While ilRet = BTRV_ERR_NONE
            If StrComp(Trim$(tgRNAdf.sName), Trim$(tlAdf.sName), vbTextCompare) = 0 Then
                If StrComp(Trim$(tgRNAdf.sAddrID), Trim$(tlAdf.sAddrID), vbTextCompare) = 0 Then
                    tgRNChf.iAdfCode = tlAdf.iCode
                    Exit Do
                End If
            End If
            ilRet = btrGetNext(hgRNAdf, tlAdf, igRNAdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    If tgRNChf.iAdfCode <= 0 Then
        'Add advertiser
        ilRet = btrInsert(hgRNAdf, tgRNAdf, igRNAdfRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            If Trim$(tlAdf.sAddrID) = "" Then
                gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "||" & "Contract Not Added as unable to add Advertiser, Pervasive Error # " & ilRet, True
            Else
                gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "||" & "Contract Not Added as unable to add Advertiser, Pervasive Error # " & ilRet, True
            End If
            Exit Function
        End If
        'ilRet = btrUpdate(hgRNAdf, tgRNAdf, igRNAdfRecLen)
        tgRNChf.iAdfCode = tgRNAdf.iCode
        tgRNMif.lCode = 0
        tgRNMif.sNetDBID = smNetDBID
        tgRNMif.sRepDBID = smRepDBID
        tgRNMif.lNetDBCode = tgRNAdf.iCode
        tgRNMif.lRepDBCode = ilRepAdfCode
        tgRNMif.sType = "A"
        tgRNMif.sUnused = ""
        ilRet = btrInsert(hgRNMif, tgRNMif, igRNMifRecLen, INDEXKEY0)
    End If
    ilIndex = mFindFieldName(smHdFieldNames(), "chtProductName")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNChf.sProduct = smFields(ilIndex)
    ilIndex = mFindFieldName(smHdFieldNames(), "chtAgencyCode")
    If ilIndex = -1 Then
        Exit Function
    End If
    ilRepAgfCode = Val(smFields(ilIndex))
    If ilRepAgfCode > 0 Then
        tgRNMifSrchKey2.sType = "G"
        tgRNMifSrchKey2.sRepDBID = smRepDBID
        tgRNMifSrchKey2.lRepDBCode = ilRepAgfCode
        ilRet = btrGetEqual(hgRNMif, tgRNMif, igRNMifRecLen, tgRNMifSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            tgRNAgfSrchKey.iCode = tgRNMif.lNetDBCode
            ilRet = btrGetEqual(hgRNAgf, tgRNAgf, igRNAgfRecLen, tgRNAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                tgRNChf.iAgfCode = tgRNAgf.iCode
            End If
        End If
    End If
    If ilRepAgfCode > 0 Then
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAgencyName")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAgf.sName = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAgencyAbbr")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAgf.sAbbr = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAgencyCityID")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAgf.sCityID = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAgencyCntrAddr_1")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAgf.sCntrAddr(0) = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAgencyCntrAddr_2")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAgf.sCntrAddr(1) = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAgencyCntrAddr_3")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAgf.sCntrAddr(2) = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAgencyBillAddr_1")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAgf.sBillAddr(0) = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAgencyBillAddr_2")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAgf.sBillAddr(1) = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAgencyBillAddr_3")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAgf.sBillAddr(2) = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtAgencyPctComm")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNAgf.iComm = Val(smFields(ilIndex))
        If tgRNChf.iAgfCode > 0 Then
            ilRet = btrUpdate(hgRNAgf, tgRNAgf, igRNAgfRecLen)
        Else
            'Search for matching name
            ilRet = btrGetFirst(hgRNAgf, tlagf, igRNAgfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            Do While ilRet = BTRV_ERR_NONE
                If StrComp(Trim$(tgRNAgf.sName), Trim$(tlagf.sName), vbTextCompare) = 0 Then
                    If StrComp(Trim$(tgRNAgf.sCityID), Trim$(tlagf.sCityID), vbTextCompare) = 0 Then
                        tgRNChf.iAgfCode = tlagf.iCode
                        Exit Do
                    End If
                End If
                ilRet = btrGetNext(hgRNAgf, tlagf, igRNAgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
        If tgRNChf.iAgfCode <= 0 Then
            'Add advertiser
            ilRet = btrInsert(hgRNAgf, tgRNAgf, igRNAgfRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                If ilRet >= 30000 Then
                    ilRet = csiHandleValue(0, 7)
                End If
                If Trim$(tlagf.sCityID) = "" Then
                    gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "||" & "Contract Not Added as unable to add Agency " & Trim$(tgRNAgf.sName) & ", Pervasive Error # " & ilRet, True
                Else
                    gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "||" & "Contract Not Added as unable to add Agency " & Trim$(tgRNAgf.sName) & "," & Trim$(tlagf.sCityID) & ", Pervasive Error # " & ilRet, True
                End If
                Exit Function
            End If
            tgRNChf.iAgfCode = tgRNAgf.iCode
            '10148
           ' tgRNAgf.iAutoCode = tgRNAgf.iCode
            ilRet = btrUpdate(hgRNAgf, tgRNAgf, igRNAgfRecLen)
            tgRNMif.lCode = 0
            tgRNMif.sNetDBID = smNetDBID
            tgRNMif.sRepDBID = smRepDBID
            tgRNMif.lNetDBCode = tgRNAgf.iCode
            tgRNMif.lRepDBCode = ilRepAgfCode
            tgRNMif.sType = "G"
            tgRNMif.sUnused = ""
            ilRet = btrInsert(hgRNMif, tgRNMif, igRNMifRecLen, INDEXKEY0)
        End If
    End If
    ilIndex = mFindFieldName(smHdFieldNames(), "chtProductProtection_1")
    If ilIndex = -1 Then
        Exit Function
    End If
    smProdProt1 = Trim$(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "chtProductProtection_2")
    If ilIndex = -1 Then
        Exit Function
    End If
    smProdProt2 = Trim$(smFields(ilIndex))
    imProdProt1Req = False
    imProdProt2Req = False
    'If (smProdProt1 <> "") Or (smProdProt2 <> "") Then
        'Get Product from previous contract
        tgRNChfSrchKey1.lCntrNo = tgRNChf.lCntrNo
        tgRNChfSrchKey1.iCntRevNo = 32000
        tgRNChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hgRNChf, tlChf, igRNChfRecLen, tgRNChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tgRNChf.lCntrNo = tlChf.lCntrNo)
            If StrComp(Trim$(tgRNChf.sRepDBID), Trim$(tlChf.sRepDBID), vbTextCompare) = 0 Then
                ilOlfCntrFd = True
                tgRNChf.iMnfComp(0) = tlChf.iMnfComp(0)
                tgRNChf.iMnfComp(1) = tlChf.iMnfComp(1)
                For ilLoop = 0 To 4 Step 1
                    tgRNChf.iMnfRevSet(ilLoop) = tlChf.iMnfRevSet(ilLoop)
                Next ilLoop
                tgRNChf.iMnfBus = tlChf.iMnfBus
                Exit Do
            End If
            ilRet = btrGetNext(hgRNChf, tlChf, igRNChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If ((tgRNChf.iMnfComp(0) <= 0) And (smProdProt1 <> "")) Or ((tgRNChf.iMnfComp(1) <= 0) And (smProdProt2 <> "")) Then
            tgRNMnfSrchKey3.sType = "C"
            ilRet = btrGetEqual(hgRNMnf, tgRNMnf, igRNMnfRecLen, tgRNMnfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tgRNMnf.sType = "C")
                If (tgRNChf.iMnfComp(0) <= 0) And (smProdProt1 <> "") Then
                    If StrComp(Trim$(tgRNMnf.sName), smProdProt1, vbTextCompare) = 0 Then
                        tgRNChf.iMnfComp(0) = tgRNMnf.iCode
                    End If
                End If
                If (tgRNChf.iMnfComp(1) <= 0) And (smProdProt2 <> "") Then
                    If StrComp(Trim$(tgRNMnf.sName), smProdProt2, vbTextCompare) = 0 Then
                        tgRNChf.iMnfComp(1) = tgRNMnf.iCode
                    End If
                End If
                If (tgRNChf.iMnfComp(0) > 0) And (tgRNChf.iMnfComp(1) > 0) Then
                    Exit Do
                End If
                If (tgRNChf.iMnfComp(0) > 0) And (smProdProt2 = "") Then
                    Exit Do
                End If
                If (smProdProt1 = "") And (tgRNChf.iMnfComp(1) > 0) Then
                    Exit Do
                End If
                ilRet = btrGetNext(hgRNMnf, tgRNMnf, igRNMnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
        'If ((tgRNChf.iMnfComp(0) <= 0) And (smProdProt1 <> "")) Or ((tgRNChf.iMnfComp(1) <= 0) And (smProdProt2 <> "")) Then
        If (tgRNChf.iMnfComp(0) <= 0) Or (tgRNChf.iMnfComp(1) <= 0) Then
            tgRNAdfSrchKey.iCode = tgRNChf.iAdfCode
            ilRet = btrGetEqual(hgRNAdf, tgRNAdf, igRNAdfRecLen, tgRNAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                If StrComp(Trim$(tgRNAdf.sProduct), Trim$(tgRNChf.sProduct), vbTextCompare) = 0 Then
                    'If (tgRNChf.iMnfComp(0) <= 0) And (smProdProt1 <> "") Then
                    If (tgRNChf.iMnfComp(0) <= 0) Then
                        tgRNChf.iMnfComp(0) = tgRNAdf.iMnfComp(0)
                    End If
                    'If (tgRNChf.iMnfComp(1) <= 0) And (smProdProt2 <> "") Then
                    If (tgRNChf.iMnfComp(1) <= 0) Then
                        tgRNChf.iMnfComp(1) = tgRNAdf.iMnfComp(1)
                    End If
                End If
            End If
        End If
        If (tgRNChf.iMnfComp(0) <= 0) And (smProdProt1 <> "") Then
            imProdProt1Req = True
            gCreateAlert "M||" & "|" & tgRNChf.lCntrNo & "|" & smAdvertiserName & "||" & "Product Protection reference " & smProdProt1 & " removed from Contract", True
        Else
            If (tgRNChf.iMnfComp(0) <= 0) Then
                imProdProt1Req = True
                gCreateAlert "M||" & "|" & tgRNChf.lCntrNo & "|" & smAdvertiserName & "||" & "Product Protection missing from contract", True
            End If
        End If
        If (tgRNChf.iMnfComp(1) <= 0) And (smProdProt2 <> "") Then
            imProdProt2Req = True
            gCreateAlert "M||" & "|" & tgRNChf.lCntrNo & "|" & smAdvertiserName & "||" & "Product Protection reference " & smProdProt2 & " removed from Contract", True
        End If
    'End If
    ilIndex = mFindFieldName(smHdFieldNames(), "chtPctTrade")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNChf.iPctTrade = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "chtDemoGraphicName_1")
    If ilIndex = -1 Then
        Exit Function
    End If
    slDemoName(0) = Trim$(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "chtDemoGraphicName_2")
    If ilIndex = -1 Then
        Exit Function
    End If
    slDemoName(1) = Trim$(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "chtDemoGraphicName_3")
    If ilIndex = -1 Then
        Exit Function
    End If
    slDemoName(2) = Trim$(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "chtDemoGraphicName_4")
    If ilIndex = -1 Then
        Exit Function
    End If
    slDemoName(3) = Trim$(smFields(ilIndex))
    For ilLoop = 0 To 3 Step 1
        If (slDemoName(ilLoop) <> "") Then
            tgRNMnfSrchKey3.sType = "D"
            ilRet = btrGetEqual(hgRNMnf, tgRNMnf, igRNMnfRecLen, tgRNMnfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tgRNMnf.sType = "D")
                If StrComp(Trim$(tgRNMnf.sName), slDemoName(ilLoop), vbTextCompare) = 0 Then
                    tgRNChf.iMnfDemo(ilLoop) = tgRNMnf.iCode
                    Exit Do
                End If
                If ilLoop = 0 Then
                    If slDemoName(1) <> "" Then
                        If StrComp(Trim$(tgRNMnf.sName), slDemoName(1), vbTextCompare) = 0 Then
                            tgRNChf.iMnfDemo(1) = tgRNMnf.iCode
                            slDemoName(1) = ""
                        End If
                    End If
                    If slDemoName(2) <> "" Then
                        If StrComp(Trim$(tgRNMnf.sName), slDemoName(2), vbTextCompare) = 0 Then
                            tgRNChf.iMnfDemo(2) = tgRNMnf.iCode
                            slDemoName(2) = ""
                        End If
                    End If
                    If slDemoName(3) <> "" Then
                        If StrComp(Trim$(tgRNMnf.sName), slDemoName(3), vbTextCompare) = 0 Then
                            tgRNChf.iMnfDemo(3) = tgRNMnf.iCode
                            slDemoName(3) = ""
                        End If
                    End If
                ElseIf ilLoop = 1 Then
                    If slDemoName(2) <> "" Then
                        If StrComp(Trim$(tgRNMnf.sName), slDemoName(2), vbTextCompare) = 0 Then
                            tgRNChf.iMnfDemo(2) = tgRNMnf.iCode
                            slDemoName(2) = ""
                        End If
                    End If
                    If slDemoName(3) <> "" Then
                        If StrComp(Trim$(tgRNMnf.sName), slDemoName(3), vbTextCompare) = 0 Then
                            tgRNChf.iMnfDemo(3) = tgRNMnf.iCode
                            slDemoName(3) = ""
                        End If
                    End If
                ElseIf ilLoop = 2 Then
                    If StrComp(Trim$(tgRNMnf.sName), slDemoName(3), vbTextCompare) = 0 Then
                        tgRNChf.iMnfDemo(3) = tgRNMnf.iCode
                        slDemoName(3) = ""
                    End If
                End If
                ilRet = btrGetNext(hgRNMnf, tgRNMnf, igRNMnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
    Next ilLoop
    ilIndex = mFindFieldName(smHdFieldNames(), "chtSalespersonFirstName")
    If ilIndex = -1 Then
        Exit Function
    End If
    slSlspFirstName = Trim$(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "chtSalespersonLastName")
    If ilIndex = -1 Then
        Exit Function
    End If
    slSlspLastName = Trim$(smFields(ilIndex))
    ilRet = btrGetFirst(hgRNSlf, tgRNSlf, igRNSlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If StrComp(Trim$(tgRNSlf.sFirstName), slSlspFirstName, vbTextCompare) = 0 Then
            If StrComp(Trim$(tgRNSlf.sLastName), slSlspLastName, vbTextCompare) = 0 Then
                tgRNChf.iSlfCode(0) = tgRNSlf.iCode
                'tgRNChf.iSlspCommPct(0) = 10000
                'tgRNChf.lComm(0) = 0
                'tgRNChf.iMnfSubCmpy(0) = 0
                Exit Do
            End If
        End If
        ilRet = btrGetNext(hgRNSlf, tgRNSlf, igRNSlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If tgRNChf.iSlfCode(0) <= 0 Then
        gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "||" & "Contract Not Added as unable to find Salesperson " & Trim$(slSlspFirstName) & " " & slSlspLastName, True
        Exit Function
    End If
    If tgSpf.sCommByCntr = "Y" Then
        tgRNChf.iSlspCommPct(0) = gStrDecToInt("100", 2)
        tgRNChf.lComm(0) = gStrDecToLong("100", 4)
        tgRNChf.iMnfSubCmpy(0) = 0
    ElseIf tgSpf.sSubCompany = "Y" Then
        tgRNChf.iSlspCommPct(0) = gStrDecToInt("100", 2)
        tgRNChf.lComm(0) = gStrDecToLong("", 4)
        tgRNChf.iMnfSubCmpy(0) = 0
    Else
        tgRNChf.iSlspCommPct(0) = gStrDecToInt("100", 2)
        tgRNChf.lComm(0) = gStrDecToLong("100", 4)
        tgRNChf.iMnfSubCmpy(0) = 0
    End If
    ilIndex = mFindFieldName(smHdFieldNames(), "chtBuyerName")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNChf.sBuyer = smFields(ilIndex)
    ilIndex = mFindFieldName(smHdFieldNames(), "chtBuyerPhoneNo")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNChf.sPhone = smFields(ilIndex)
    If ilOlfCntrFd Then
        For ilLoop = 0 To 3 Step 1
            Select Case ilLoop
                Case 0
                    smCommentInt = gGetComment(tlChf.lCxfInt, smShowOnInt)
                Case 1
                    smCommentChgR = gGetComment(tlChf.lCxfChgR, smShowOnChgR)
                Case 2
                    smCommentCanc = gGetComment(tlChf.lCxfCanc, smShowOnCanc)
                Case 3
                    smCommentOther = gGetComment(tlChf.lCxfCode, smShowOnOther)
            End Select
        Next ilLoop
    Else
        ilIndex = mFindFieldName(smHdFieldNames(), "chtInternalComment")
        If ilIndex = -1 Then
            Exit Function
        End If
        smCommentInt = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtShowInternalOn")
        If ilIndex = -1 Then
            Exit Function
        End If
        smShowOnInt = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtChangeReasonComment")
        If ilIndex = -1 Then
            Exit Function
        End If
        smCommentChgR = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtShowChangeOn")
        If ilIndex = -1 Then
            Exit Function
        End If
        smShowOnChgR = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtCancelComment")
        If ilIndex = -1 Then
            Exit Function
        End If
        smCommentCanc = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtShowCancelOn")
        If ilIndex = -1 Then
            Exit Function
        End If
        smShowOnCanc = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtOtherComment")
        If ilIndex = -1 Then
            Exit Function
        End If
        smCommentOther = smFields(ilIndex)
        ilIndex = mFindFieldName(smHdFieldNames(), "chtShowOtherOn")
        If ilIndex = -1 Then
            Exit Function
        End If
        smShowOnOther = smFields(ilIndex)
    End If
    If Not mOkToAddContract() Then
        mCreateContract = False
        Exit Function
    End If
    mCreateContract = True
End Function



Private Function mCreateLine() As Integer
    Dim ilIndex As Integer
    Dim slCntrNo As String
    Dim slVehicleName As String
    Dim ilRepVefCode As Integer
    Dim slDayPartName As String
    Dim ilRepRdfCode As Integer
    Dim ilUpper As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilLoop As Integer
    Dim slDays As String
    Dim ilRet As Integer
    'Dim slDPInOut As String
    'Dim slDPAvailName As String
    Dim ilDay As Integer
    'Dim slDPDay_1 As String
    Dim llDPStartTime_1 As Long
    Dim llDPEndTime_1 As Long
    'Dim ilDPSpotPct_1 As Integer
    'Dim slDPDay_2 As String
    'Dim llDPStartTime_2 As Long
    'Dim llDPEndTime_2 As Long
    'Dim ilDPSpotPct_2 As Integer
    Dim ilRdf As Integer
    Dim llRdfStartTime As Long
    Dim llRdfEndTime As Long
    'Dim ilDP1Found As Integer
    'Dim ilDP2Found As Integer
    Dim ilCount As Integer
    Dim slStr As String
    'Dim ilExtraDayCount As Integer
    'Dim ilBFRdfCode As Integer
    'Dim ilBFExtraDayCount As Integer
    Dim slDPInOut As String
    Dim slDPAvailName As String
    
    ilUpper = UBound(tgRNClf)
    mCreateLine = False
    ilIndex = mFindFieldName(smLnFieldNames(), "cltContractNo")
    If ilIndex = -1 Then
        Exit Function
    End If
    slCntrNo = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltIntVerNo")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).iCntRevNo = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltProposalVersion")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).iPropVer = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltLineNo")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).iLine = Val(smFields(ilIndex))
    tgRNClf(ilUpper).iVefCode = 0
    ilIndex = mFindFieldName(smLnFieldNames(), "cltVehicleName")
    If ilIndex = -1 Then
        Exit Function
    End If
    smVehicleName = Trim$(smFields(ilIndex))
    slVehicleName = Trim$(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltNetVehicleCode")
    If ilIndex = -1 Then
        Exit Function
    End If
    ilRepVefCode = Val(smFields(ilIndex))
    tgRNClf(ilUpper).iVefCode = 0
    tgRNMifSrchKey2.sType = "V"
    tgRNMifSrchKey2.sRepDBID = smRepDBID
    tgRNMifSrchKey2.lRepDBCode = ilRepVefCode
    ilRet = btrGetEqual(hgRNMif, tgRNMif, igRNMifRecLen, tgRNMifSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If (ilRet = BTRV_ERR_NONE) Then
        tgRNClf(ilUpper).iVefCode = tgRNMif.lNetDBCode
    End If
    If tgRNClf(ilUpper).iVefCode <= 0 Then
        'Search for matching name
        ilRet = btrGetFirst(hgRNVef, tgRNVef, igRNVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While ilRet = BTRV_ERR_NONE
            If StrComp(Trim$(tgRNVef.sName), slVehicleName, vbTextCompare) = 0 Then
                tgRNClf(ilUpper).iVefCode = tgRNVef.iCode
                tgRNMif.lCode = 0
                tgRNMif.sNetDBID = smNetDBID
                tgRNMif.sRepDBID = smRepDBID
                tgRNMif.lNetDBCode = tgRNVef.iCode
                tgRNMif.lRepDBCode = ilRepVefCode
                tgRNMif.sType = "V"
                tgRNMif.sUnused = ""
                ilRet = btrInsert(hgRNMif, tgRNMif, igRNMifRecLen, INDEXKEY0)
                Exit Do
            End If
            ilRet = btrGetNext(hgRNVef, tgRNVef, igRNVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    If tgRNClf(ilUpper).iVefCode <= 0 Then
        gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "|" & smVehicleName & "|" & "Contract Not Added as unable to find Vehicle", True
        Exit Function
    End If
    ilIndex = mFindFieldName(smLnFieldNames(), "cltDaypartName")
    If ilIndex = -1 Then
        Exit Function
    End If
    slDayPartName = Trim$(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltNetDaypartCode")
    If ilIndex = -1 Then
        Exit Function
    End If
    ilRepRdfCode = Val(smFields(ilIndex))
    tgRNClf(ilUpper).iRdfCode = 0
    'Remove speedup as content could change but code not change
    'tgRNMifSrchKey2.sType = "D"
    'tgRNMifSrchKey2.sRepDBID = smRepDBID
    'tgRNMifSrchKey2.lRepDBCode = ilRepRdfCode
    'ilRet = btrGetEqual(hgRNMif, tgRNMif, igRNMifRecLen, tgRNMifSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    'If (ilRet = BTRV_ERR_NONE) Then
    '    tgRNClf(ilUpper).iRdfcode = tgRNMif.lNetDBCode
    'End If
    'If tgRNClf(ilUpper).iRdfcode <= 0 Then
        tgRNClf(ilUpper).iRdfCode = mSearchForDaypart(smLnFieldNames(), "clt", False, slDPInOut, slDPAvailName)
        If tgRNClf(ilUpper).iRdfCode > 0 Then
            'tgRNMif.lCode = 0
            'tgRNMif.sNetDBID = smNetDBID
            'tgRNMif.sRepDBID = smRepDBID
            'tgRNMif.lNetDBCode = tgRNClf(ilUpper).iRdfcode
            'tgRNMif.lRepDBCode = ilRepRdfCode
            'tgRNMif.sType = "D"
            'tgRNMif.sUnused = ""
            'ilRet = btrInsert(hgRNMif, tgRNMif, igRNMifRecLen, INDEXKEY0)
        ElseIf (slDPInOut <> "I") And (slDPInOut <> "O") Then
            'Look for 12m-12m Mo-Su
            ilRet = btrGetFirst(hgRNRdf, tgRNRdf, igRNRdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            Do While ilRet = BTRV_ERR_NONE
                If tgRNRdf.sState = "A" Then
                    For ilRdf = LBound(tgRNRdf.iStartTime, 2) To UBound(tgRNRdf.iStartTime, 2) Step 1
                        If (tgRNRdf.iStartTime(0, ilRdf) <> 1) Or (tgRNRdf.iStartTime(1, ilRdf) <> 0) Then
                            slStr = ""
                            For ilDay = LBound(tgRNRdf.sWkDays, 1) To UBound(tgRNRdf.sWkDays, 1) Step 1
                                slStr = slStr & tgRNRdf.sWkDays(ilRdf, ilDay)
                            Next ilDay
                            gUnpackTimeLong tgRNRdf.iStartTime(0, ilRdf), tgRNRdf.iStartTime(1, ilRdf), False, llRdfStartTime
                            gUnpackTimeLong tgRNRdf.iEndTime(0, ilRdf), tgRNRdf.iEndTime(1, ilRdf), False, llRdfEndTime
                            If StrComp(slStr, "YYYYYYY", vbTextCompare) = 0 Then
                                If (llRdfStartTime = 0) And (llRdfEndTime = 0) Then
                                    tgRNClf(ilUpper).iRdfCode = tgRNRdf.iCode
                                    Exit Do
                                End If
                            End If
                        End If
                    Next ilRdf
                End If
                ilRet = btrGetNext(hgRNRdf, tgRNRdf, igRNRdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If tgRNClf(ilUpper).iRdfCode <= 0 Then
                'Add Mo-Su 12m-12m
                mInitRdf
                gPackTime "12M", tgRNRdf.iStartTime(0, 7), tgRNRdf.iStartTime(1, 7)
                gPackTime "12M", tgRNRdf.iEndTime(0, 7), tgRNRdf.iEndTime(1, 7)
                tgRNRdf.sWkDays(7, 1) = "Y"
                tgRNRdf.sWkDays(7, 2) = "Y"
                tgRNRdf.sWkDays(7, 3) = "Y"
                tgRNRdf.sWkDays(7, 4) = "Y"
                tgRNRdf.sWkDays(7, 5) = "Y"
                tgRNRdf.sWkDays(7, 6) = "Y"
                tgRNRdf.sWkDays(7, 7) = "Y"
                tgRNRdf.iSpotPct(7) = 10000
                ilRet = btrInsert(hgRNRdf, tgRNRdf, igRNRdfRecLen, INDEXKEY0)
                If ilRet = BTRV_ERR_NONE Then
                    tgRNRdf.iRemoteID = 0
                    tgRNRdf.iAutoCode = tgRNRdf.iCode
                    tgRNRdf.iSourceID = 0
                    gPackDate Format$(gNow(), "m/d/yy"), tgRNRdf.iSyncDate(0), tgRNRdf.iSyncDate(1)
                    gPackTime Format$(gNow(), "h:mm:ssAM/PM"), tgRNRdf.iSyncTime(0), tgRNRdf.iSyncTime(1)
                    ilRet = btrUpdate(hgRNRdf, tgRNRdf, igRNRdfRecLen)
                End If
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "|" & smVehicleName & "|" & "Contract Not Added as unable to add Daypart Mo-Su 12M-12M, Pervasive Error # " & ilRet, True
                    Exit Function
                End If
                tgRNClf(ilUpper).iRdfCode = tgRNRdf.iCode
            End If
            'Create override times as dayparts don't match
            ilIndex = mFindFieldName(smLnFieldNames(), "cltOverrideStartTime")
            If ilIndex = -1 Then
                Exit Function
            End If
            If Trim$(smFields(ilIndex)) = "" Then
                smFields(ilIndex) = gFormatTimeLong(llDPStartTime_1, "A", "1")
            End If
            ilIndex = mFindFieldName(smLnFieldNames(), "cltOverrideEndTime")
            If ilIndex = -1 Then
                Exit Function
            End If
            If Trim$(smFields(ilIndex)) = "" Then
                smFields(ilIndex) = gFormatTimeLong(llDPEndTime_1, "A", "1")
            End If
        End If
    ''Else
    ''    tgRNClf(ilUpper).iRdfcode = ilNetRdfCode
    'End If
    If tgRNClf(ilUpper).iRdfCode <= 0 Then
        gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "|" & smVehicleName & "|" & "Contract Not Added as unable to find daypart " & slDayPartName, True
        Exit Function
    End If
    ilIndex = mFindFieldName(smLnFieldNames(), "cltSpotLength")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).iLen = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltBBOpenLength")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).iBBOpenLen = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltBBCloseLength")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).iBBCloseLen = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltSoloAvail")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).sSoloAvail = smFields(ilIndex)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltOv2DefinedBits")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).sOV2DefinedBits = smFields(ilIndex)
    ilIndex = mFindFieldName(smLnFieldNames(), "clt1stPosition")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).iPosition = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltPctAllocation")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).iPctAllocation = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltAcquisitionCost")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).lAcquisitionCost = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltLineStartDate")
    If ilIndex = -1 Then
        Exit Function
    End If
    slStartDate = smFields(ilIndex)
    gPackDate slStartDate, tgRNClf(ilUpper).iStartDate(0), tgRNClf(ilUpper).iStartDate(1)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltLineEndDate")
    If ilIndex = -1 Then
        Exit Function
    End If
    slEndDate = smFields(ilIndex)
    gPackDate slEndDate, tgRNClf(ilUpper).iEndDate(0), tgRNClf(ilUpper).iEndDate(1)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltOverrideStartTime")
    If ilIndex = -1 Then
        Exit Function
    End If
    slStartTime = smFields(ilIndex)
    gPackTime slStartTime, tgRNClf(ilUpper).iStartTime(0), tgRNClf(ilUpper).iStartTime(1)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltOverrideEndTime")
    If ilIndex = -1 Then
        Exit Function
    End If
    slEndTime = smFields(ilIndex)
    gPackTime slEndTime, tgRNClf(ilUpper).iEndTime(0), tgRNClf(ilUpper).iEndTime(1)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltPreferredStartTime")
    If ilIndex = -1 Then
        Exit Function
    End If
    slStartTime = smFields(ilIndex)
    gPackTime slStartTime, tgRNClf(ilUpper).iPrefStartTime(0), tgRNClf(ilUpper).iPrefStartTime(1)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltPreferredEndTime")
    If ilIndex = -1 Then
        Exit Function
    End If
    slEndTime = smFields(ilIndex)
    gPackTime slEndTime, tgRNClf(ilUpper).iPrefEndTime(0), tgRNClf(ilUpper).iPrefEndTime(1)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltPreferredDays")
    If ilIndex = -1 Then
        Exit Function
    End If
    slDays = smFields(ilIndex)
    For ilLoop = 0 To 6 Step 1
        tgRNClf(ilUpper).sPrefDays(ilLoop) = Mid$(slDays, ilLoop + 1, 1)
    Next ilLoop
    ilIndex = mFindFieldName(smLnFieldNames(), "cltPreemptible")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).sPreempt = smFields(ilIndex)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltLiveCopy")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).sLiveCopy = smFields(ilIndex)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltAdvtSepFlag")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).iAdvtSepFlag = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltPriority")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).iPriority = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltHideCBS")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNClf(ilUpper).sHideCBS = smFields(ilIndex)
    
    ilIndex = mFindFieldName(smLnFieldNames(), "cltComment")
    If ilIndex = -1 Then
        Exit Function
    End If
    smLnComment(ilUpper) = smFields(ilIndex)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltShowCommentOn")
    If ilIndex = -1 Then
        Exit Function
    End If
    smShowOnLnComment(ilUpper) = smFields(ilIndex)
    
    ReDim Preserve tgRNClf(0 To ilUpper + 1) As CLF
    ReDim Preserve smLnComment(0 To ilUpper + 1) As String * 5000
    ReDim Preserve smShowOnLnComment(0 To ilUpper + 1) As String
    mCreateLine = True
End Function

Private Function mCreateFlight() As Integer
    Dim ilIndex As Integer
    Dim ilDay As Integer
    Dim ilUpper As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    
    mCreateFlight = False
    ilUpper = UBound(tgRNCff)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltIntVerNo")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNCff(ilUpper).iCntRevNo = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltProposalVersion")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNCff(ilUpper).iPropVer = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltLineNo")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNCff(ilUpper).iClfLine = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltFlightStartDate")
    If ilIndex = -1 Then
        Exit Function
    End If
    slStartDate = smFields(ilIndex)
    gPackDate slStartDate, tgRNCff(ilUpper).iStartDate(0), tgRNCff(ilUpper).iStartDate(1)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltFlightEndDate")
    If ilIndex = -1 Then
        Exit Function
    End If
    slEndDate = smFields(ilIndex)
    gPackDate slEndDate, tgRNCff(ilUpper).iEndDate(0), tgRNCff(ilUpper).iEndDate(1)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltDailyOrWeeklyBuy")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNCff(ilUpper).sDyWk = smFields(ilIndex)
    If tgRNCff(ilUpper).sDyWk = "W" Then
        ilIndex = mFindFieldName(smLnFieldNames(), "cltSpotsPerWeek")
        If ilIndex = -1 Then
            Exit Function
        End If
        tgRNCff(ilUpper).iSpotsWk = Val(smFields(ilIndex))
    Else
        tgRNCff(ilUpper).iSpotsWk = 0
    End If
    ilIndex = mFindFieldName(smLnFieldNames(), "cltSpots_1_Mo")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNCff(ilUpper).iDay(0) = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltSpots_2_Tu")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNCff(ilUpper).iDay(1) = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltSpots_3_We")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNCff(ilUpper).iDay(2) = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltSpots_4_Th")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNCff(ilUpper).iDay(3) = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltSpots_5_Fr")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNCff(ilUpper).iDay(4) = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltSpots_6_Sa")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNCff(ilUpper).iDay(5) = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltSpots_7_Su")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNCff(ilUpper).iDay(6) = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smLnFieldNames(), "cltPriceType")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNCff(ilUpper).sPriceType = smFields(ilIndex)
    ilIndex = mFindFieldName(smLnFieldNames(), "cltActualPrice")
    If ilIndex = -1 Then
        Exit Function
    End If
    tgRNCff(ilUpper).lActPrice = Val(smFields(ilIndex))
    tgRNCff(ilUpper).iXSpotsWk = 0
    tgRNCff(ilUpper).lAdjPrice = 0
    tgRNCff(ilUpper).lBBPrice = 0
    tgRNCff(ilUpper).lPropPrice = 0
    tgRNCff(ilUpper).sDelete = "N"
    tgRNCff(ilUpper).sXDay(0) = ""
    tgRNCff(ilUpper).sXDay(1) = ""
    tgRNCff(ilUpper).sXDay(2) = ""
    tgRNCff(ilUpper).sXDay(3) = ""
    tgRNCff(ilUpper).sXDay(4) = ""
    tgRNCff(ilUpper).sXDay(5) = ""
    tgRNCff(ilUpper).sXDay(6) = ""
    tgRNCff(ilUpper).sUnused = ""
    ReDim Preserve tgRNCff(0 To ilUpper + 1) As CFF
    mCreateFlight = True
End Function

Private Sub mInitAdf()
    Dim ilLoop As Integer
    Dim slDate As String
    Dim slStr As String
    
    tgRNAdf.iCode = 0    'Internal code number for advertiser
    tgRNAdf.sName = ""  'Name
    tgRNAdf.sAbbr = ""
    tgRNAdf.sProduct = "" 'Product name
    tgRNAdf.iSlfCode = 0 'Salesperson code number
    tgRNAdf.iAgfCode = 0 'Agency code number
    tgRNAdf.sBuyer = ""   'Buyers name
    tgRNAdf.sCodeRep = "" 'Rep advertiser Code
    tgRNAdf.sCodeAgy = ""
    tgRNAdf.sCodeStn = "" 'Station advertiser Code
    tgRNAdf.iMnfComp(0) = 0 'Competitive code
    tgRNAdf.iMnfComp(1) = 0 'Competitive code
    tgRNAdf.iMnfExcl(0) = 0 'Program Exclusions code
    tgRNAdf.iMnfExcl(1) = 0 'Program Exclusions code
    tgRNAdf.sCppCpm = "N"    'P=CPP; M=CPM; N=N/A
    For ilLoop = 0 To 3 Step 1
        tgRNAdf.sDemo(ilLoop) = ""    'First-four Demo target
        'slStr = ""
        'gStrToPDN slStr, 2, 4, tgRNAdf.sTarget(ilLoop)
        tgRNAdf.iMnfDemo(ilLoop) = 0
        tgRNAdf.lTarget(ilLoop) = 0
    Next ilLoop
    tgRNAdf.sCreditRestr = "N"
    'tgRNAdf.sCreditLimit = "0"
    'slStr = ""
    'gStrToPDN slStr, 2, 5, tgRNAdf.sCreditLimit
    tgRNAdf.lCreditLimit = 0
    tgRNAdf.sPaymRating = "1"
    tgRNAdf.sShowISCI = "N"
    tgRNAdf.iMnfSort = 0
    tgRNAdf.sBillAgyDir = "A"
    For ilLoop = 0 To 2 Step 1
        tgRNAdf.sCntrAddr(ilLoop) = ""
    Next ilLoop
    For ilLoop = 0 To 2 Step 1
        tgRNAdf.sBillAddr(ilLoop) = ""
    Next ilLoop
    tgRNAdf.iArfLkCode = 0
    'Phone number (123) 456-789A Ext(BCDE)
    'Stored as 123456789ABCDE
    tgRNAdf.sPhone = "______________"
    tgRNAdf.sFax = "__________"
    tgRNAdf.iArfCntrCode = 0
    tgRNAdf.iArfInvCode = 0
    tgRNAdf.sCntrPrtSz = "N"
    '12/17/06-Change to tax by agency or vehicle
    'tgRNAdf.sSlsTax(0) = "N"
    'tgRNAdf.sSlsTax(1) = "N"
    tgRNAdf.iTrfCode = 0
    tgRNAdf.sCrdApp = "A" '"R" changed 6/30/00 via jim request
    tgRNAdf.sCrdRtg = ""
    tgRNAdf.iPnfBuyer = 0
    tgRNAdf.iPnfPay = 0
    'slStr = ""
    'gStrToPDN slStr, 0, 2, tgRNAdf.sPct90
    tgRNAdf.iPct90 = 0
    slStr = ""
    gStrToPDN slStr, 2, 6, tgRNAdf.sCurrAR
    slStr = ""
    gStrToPDN slStr, 2, 6, tgRNAdf.sUnbilled
    slStr = ""
    gStrToPDN slStr, 2, 6, tgRNAdf.sHiCredit
    slStr = ""
    gStrToPDN slStr, 2, 6, tgRNAdf.sTotalGross
    slDate = Format$(gNow(), "m/d/yy")
    gPackDate slDate, tgRNAdf.iDateEntrd(0), tgRNAdf.iDateEntrd(1)
    tgRNAdf.iNSFChks = 0
    tgRNAdf.iDateLstInv(0) = 0  'No date
    tgRNAdf.iDateLstInv(1) = 0
    tgRNAdf.iDateLstPaym(0) = 0  'No date
    tgRNAdf.iDateLstPaym(1) = 0
    tgRNAdf.iAvgToPay = 0
    tgRNAdf.iLstToPay = 0
    tgRNAdf.iNoInvPd = 0
    tgRNAdf.sNewBus = "N"
    tgRNAdf.iEndDate(0) = 0
    tgRNAdf.iEndDate(1) = 0
    tgRNAdf.iMerge = 0
    tgRNAdf.iUrfCode = 2
    tgRNAdf.sState = "A"
    tgRNAdf.iCrdAppDate(0) = 0
    tgRNAdf.iCrdAppDate(1) = 0
    tgRNAdf.iCrdAppTime(0) = 0
    tgRNAdf.iCrdAppTime(1) = 0
    tgRNAdf.sPkInvShow = "T"
    tgRNAdf.sRateOnInv = "Y"
    tgRNAdf.iMnfBus = 0
    tgRNAdf.lGuar = 0
    tgRNAdf.sAllowRepMG = "N"
    tgRNAdf.sBonusOnInv = "Y"
    tgRNAdf.sRepInvGen = "I"
    tgRNAdf.iMnfInvTerms = 0
    tgRNAdf.sPolitical = "N"
    tgRNAdf.sAddrID = ""
    tgRNAdf.iTrfCode = 0
    
    'ilRet = btrUpdate(hmAdf, tmAdf, imAdfRecLen)

End Sub

Private Sub mInitAgf()
    Dim ilLoop As Integer
    Dim slDate As String
    Dim slStr As String
    
    tgRNAgf.iCode = 0
    tgRNAgf.sName = ""
    tgRNAgf.sAbbr = ""
    tgRNAgf.sCityID = ""
    slStr = "15.00"
    'gStrToPDN slStr, 2, 3, tgRNAgf.sComm
    tgRNAgf.iComm = gStrDecToInt(slStr, 2)
    tgRNAgf.iSlfCode = 0
    tgRNAgf.sBuyer = ""
    tgRNAgf.sCodeRep = ""
    tgRNAgf.sCodeStn = ""
    tgRNAgf.sCreditRestr = "N"
    'tgRNAgf.sCreditLimit = "0"
    'slStr = ""
    'gStrToPDN slStr, 2, 5, tgRNAgf.sCreditLimit
    tgRNAgf.lCreditLimit = 0
    tgRNAgf.sPaymRating = "1"
    tgRNAgf.sShowISCI = "N"
    tgRNAgf.iMnfSort = 0
    tgRNAgf.sCntrAddr(0) = ""
    tgRNAgf.sCntrAddr(1) = ""
    tgRNAgf.sCntrAddr(2) = ""
    tgRNAgf.sBillAddr(0) = ""
    tgRNAgf.sBillAddr(1) = ""
    tgRNAgf.sBillAddr(2) = ""
    tgRNAgf.iArfLkCode = 0
    'Phone number (123) 456-789A Ext(BCDE)
    'Stored as 123456789ABCDE
    tgRNAgf.sPhone = "______________"
    tgRNAgf.sFax = "__________"
    tgRNAgf.iArfCntrCode = 0
    tgRNAgf.iArfInvCode = 0
    tgRNAgf.sCntrPrtSz = "N"
    '12/17/06-Change to tax by agency or vehicle
    'tgRNAgf.sSlsTax(0) = "N"
    'tgRNAgf.sSlsTax(1) = "N"
    tgRNAgf.iTrfCode = 0
    tgRNAgf.sCrdApp = "A"    '"R"
    tgRNAgf.sCrdRtg = ""
    tgRNAgf.iPnfBuyer = 0
    tgRNAgf.iPnfPay = 0
    'slStr = ""
    'gStrToPDN slStr, 0, 2, tgRNAgf.sPct90
    tgRNAgf.iPct90 = 0
    slStr = ""
    gStrToPDN slStr, 2, 6, tgRNAgf.sCurrAR
    slStr = ""
    gStrToPDN slStr, 2, 6, tgRNAgf.sUnbilled
    slStr = ""
    gStrToPDN slStr, 2, 6, tgRNAgf.sHiCredit
    slStr = ""
    gStrToPDN slStr, 2, 6, tgRNAgf.sTotalGross
    slDate = Format$(gNow(), "m/d/yy")
    gPackDate slDate, tgRNAgf.iDateEntrd(0), tgRNAgf.iDateEntrd(1)
    tgRNAgf.iNSFChks = 0
    tgRNAgf.iDateLstInv(0) = 0  'No date
    tgRNAgf.iDateLstInv(1) = 0
    tgRNAgf.iDateLstPaym(0) = 0  'No date
    tgRNAgf.iDateLstPaym(1) = 0
    tgRNAgf.iAvgToPay = 0
    tgRNAgf.iLstToPay = 0
    tgRNAgf.iNoInvPd = 0
    tgRNAgf.iMerge = 0
    tgRNAgf.iUrfCode = 2
    tgRNAgf.sState = "A"
    tgRNAgf.iCrdAppDate(0) = 0
    tgRNAgf.iCrdAppDate(1) = 0
    tgRNAgf.iCrdAppTime(0) = 0
    tgRNAgf.iCrdAppTime(1) = 0
    tgRNAgf.sPkInvShow = "T"
    tgRNAgf.iTrfCode = 0
        '10148
'    tgRNAgf.iRemoteID = 0
'    tgRNAgf.iAutoCode = 0
'    tgRNAgf.iSourceID = 0
'    gPackDate Format$(gNow(), "m/d/yy"), tgRNAgf.iSyncDate(0), tgRNAgf.iSyncDate(1)
'    gPackTime Format$(gNow(), "h:mm:ssAM/PM"), tgRNAgf.iSyncTime(0), tgRNAgf.iSyncTime(1)
    'ilRet = btrUpdate(hmAgf, tmAgf, imAgfRecLen)

End Sub

Private Sub mInitChf()
    Dim ilLoop As Integer
    Dim slStr As String
    
    tgRNChf.lCode = 0    'Internal code number for Contract
    tgRNChf.lCntrNo = 0  'Contract number
    tgRNChf.sAgyEstNo = ""  'Agenct estimate number
    tgRNChf.iExtRevNo = 0  'Master Contract number
    'tgRNChf.iPkageGenMeth = 0 'weekly dollar generation
    tgRNChf.iXMLSentUrfCode = 0
    gPackDate Format$(gNow(), "m/d/yy"), tgRNChf.iOHDDate(0), tgRNChf.iOHDDate(1)
    gPackTime Format$(gNow(), "h:mm:ssAM/PM"), tgRNChf.iOHDTime(0), tgRNChf.iOHDTime(1)
    tgRNChf.iCntRevNo = 0   'Revision number
    tgRNChf.iPropVer = 0
    tgRNChf.iPropDate(0) = 0
    tgRNChf.iPropDate(1) = 0
    tgRNChf.iPropTime(0) = 0
    tgRNChf.iPropTime(1) = 0
    tgRNChf.sType = ""  'Type of contract: P=Proposal; H=Hold; V=Reservation; J=Rejection;
                        'E=Order; C=Contract; B=Allocation contract; D=Deferred Contract;
                        'T=Remnant; R=Direct Response; Q=Per inQuiry; S=PSA; M=Promo
    tgRNChf.iAgfCode = 0
    tgRNChf.iAdfCode = 0
    tgRNChf.sProduct = ""   'Product (default to advertiser product)
    tgRNChf.lSifCode = 0
    For ilLoop = 0 To 9 Step 1
        tgRNChf.iSlfCode(ilLoop) = 0
        tgRNChf.lComm(ilLoop) = 0
        tgRNChf.iMnfSubCmpy(ilLoop) = 0
        tgRNChf.iSlspCommPct(ilLoop) = 0
    Next ilLoop
    tgRNChf.iMnfComp(0) = 0
    tgRNChf.iMnfComp(1) = 0   'mGetCompCode(slName) 'Competitive code number
    tgRNChf.iMnfExcl(0) = 0 'Program exclusions
    tgRNChf.iMnfExcl(1) = 0 'Program exclusions
    tgRNChf.sBuyer = ""   'Buyers name (default name from advertiser)
    tgRNChf.sPhone = "______________"   'Buyers Phone number and extension(if direct: default to advt otherwisw default to agency)
    'tgRNChf.sPriceCntr = "Y"  'Y=Show prices on contract; N=Don't show prices
    'tgRNChf.sPriceInv = "Y" 'Y=Show prices on invoices; N=Don't show prices
    'If tgNrf.sTransmitPrice = "N" Then
    '    tgRNChf.sPriceCntr = "N"  'Y=Show prices on contract; N=Don't show prices
    '    tgRNChf.sPriceInv = "N" 'Y=Show prices on invoices; N=Don't show prices
    'End If
    tgRNChf.iEDSSentUrfCode = 0
    tgRNChf.iPctBudget = 0
    For ilLoop = 0 To 4 Step 1
        tgRNChf.iMnfRevSet(ilLoop) = 0 '
    Next ilLoop
    tgRNChf.sCppCpm = "N"    'P=CPP; M=CPM; N=N/A
    tgRNChf.iMerchPct = 0
    tgRNChf.iPromoPct = 0
    For ilLoop = 0 To 9 Step 1
        tgRNChf.iSlspCommPct(ilLoop) = 0
    Next ilLoop
    For ilLoop = 0 To 3 Step 1
        tgRNChf.iMnfDemo(ilLoop) = 0    'First-four Demo target
        tgRNChf.lTarget(ilLoop) = 0
    Next ilLoop
    slStr = "0"
    'gStrToPDN slStr, 0, 2, tgRNChf.sPctTrade
    tgRNChf.iPctTrade = gStrDecToInt(slStr, 0)
    tgRNChf.iRcfCode = mObtainLatestRcf() 'Rate Card code
    'gStrToPDN slStr, 2, 6, tgRNChf.sInputGross
    tgRNChf.lInputGross = gStrDecToLong("0", 2)
    'slStr = smFieldValues(26)
    'gStrToPDN slStr, 2, 6, tmIcf.sGross
    tgRNChf.sBillCycle = "S" 'S=Standard; C=Calendar; W=Weekly
    tgRNChf.sInvGp = "A"       'Invoicing grouping flag: A= all Spots; P= per Product; T= per Tag
    'slStr = ""
    'gStrToPDN slStr, 2, 3, tgRNChf.sPctTag
    'tgRNChf.iPctTag = 0
    tgRNChf.lCxfCode = 0        'Comment code number
    tgRNChf.lCxfChgR = 0        'Comment code number
    tgRNChf.lCxfInt = 0        'Comment code number
    tgRNChf.lCxfMerch = 0        'Comment code number
    tgRNChf.lCxfProm = 0        'Comment code number
    tgRNChf.lCxfCanc = 0        'Comment code number
    tgRNChf.iPropVer = 0
    If tgSpf.sInvAirOrder = "O" Then
        tgRNChf.sMGMiss = "M"
    Else
        tgRNChf.sMGMiss = "G"
    End If
    'If imPorO = 0 Then
    '    tgRNChf.sSchStatus = "P"
    '    tgRNChf.sStatus = "W"
    'Else
    '    Select Case tgRNChf.sType
    '        Case "T", "R", "Q", "S", "M" 'Remnant; Direct response; per Inquires; PSA; Promo
    '            tgRNChf.sSchStatus = "M"
    '            tgRNChf.sStatus = "O"
    '        Case Else   'Standard or Reservation
    '            tgRNChf.sStatus = "N" 'N is type O but needs scheduling
    '            tgRNChf.sSchStatus = "A"  '"N"
    '    End Select
    'End If
    tgRNChf.sStatus = "N"
    tgRNChf.sSchStatus = "A"
    tgRNChf.sTitle = ""
    tgRNChf.iDtNeed(0) = 0
    tgRNChf.iDtNeed(1) = 0
    tgRNChf.iMnfBus = 0
    tgRNChf.lGuar = 0
    tgRNChf.iEDSSentExtRevNo = 0
    For ilLoop = 0 To 6 Step 1
        tgRNChf.iMnfCmpy(ilLoop) = 0
        tgRNChf.iCmpyPct(ilLoop) = 0
    Next ilLoop
    tgRNChf.sResvNew = "N"
    tgRNChf.lChfCode = 0
    tgRNChf.iMnfPotnType = 0
    tgRNChf.sPrint = "N"    'Contract print state: N=New and not printed; C=Changed and not printed; P=already printed (don't ask revision number until printed)
    tgRNChf.sDiscrep = "N"
    tgRNChf.sNewBus = "N"   'New business state: Y= advertiser new business flag is Y-> set it to N; N= not new business
    tgRNChf.sAgyCTrade = "N"  'Trades Commissionable
    tgRNChf.iStartDate(0) = 0
    tgRNChf.iStartDate(1) = 0
    tgRNChf.iEndDate(0) = 0
    tgRNChf.iEndDate(1) = 0
    tgRNChf.lVefCode = 0 'Vehicle Code (- => combo. note: Vehicle of all lines.  If one does not exist for all lines it will be created)
    tgRNChf.iUrfCode = 2     'Last user who modified contract
    tgRNChf.sCBSOrder = "N"
    tgRNChf.iHdChg = 0   'Bit field indicating header changes
                        ' 0     Advertiser changed
                        ' 1     Competitive # 1 changed
                        ' 2     Competitive # 2 changed
                        ' 3     Exclusion # 1 changed
                        ' 4     Exclusion # 2 changed
                        ' 5     Status P (Prevent) changed
                        ' 6     Status M (Manual) changed
    'tgRNChf.iVerifyDate(0) = 0
    'tgRNChf.iVerifyDate(1) = 0
    'tgRNChf.iVerifyTime(0) = 0
    'tgRNChf.iVerifyTime(1) = 0
    'tgRNChf.sVerifyFlag = "Y"
    gPackDate "", tgRNChf.iEDSSentDate(0), tgRNChf.iEDSSentDate(1)
    gPackTime "12AM", tgRNChf.iEDSSentTime(0), tgRNChf.iEDSSentTime(1)
    'Dan M per Dick 9/3/15
    'tgRNChf.iSlfLock = 0
    'tgRNChf.iLockDate(0) = 0
    'tgRNChf.iLockDate(1) = 0
    tgRNChf.iXMLSentExtRevNo = 0
    gPackDate "", tgRNChf.iXMLSentDate(0), tgRNChf.iXMLSentDate(1)
    gPackTime "12AM", tgRNChf.iXMLSentTime(0), tgRNChf.iXMLSentTime(1)
    tgRNChf.sDelete = "N"
    'tgRNChf.lAudMGAdj = 0
    tgRNChf.lSpotChfCode = 0
    tgRNChf.sNoAssigned = "Y"
    tgRNChf.lEffCode = 0
End Sub

Private Function mObtainLatestRcf() As Integer
    Dim ilRet As Integer
    Dim ilYear As Integer
    Dim llDate As Long
    Dim llTDate As Long
    Dim ilRcfCode As Integer
    Dim tlRcf As RCF
    
    ilYear = 0
    llDate = 0
    ilRcfCode = 0
    ilRet = btrGetFirst(hgRNRcf, tlRcf, igRNRcfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If Left$(tlRcf.sName, 1) <> "~" Then
            If tlRcf.iYear > ilYear Then
                ilRcfCode = tlRcf.iCode
                ilYear = tlRcf.iYear
                gUnpackDateLong tlRcf.iStartDate(0), tlRcf.iStartDate(1), llDate
            ElseIf tlRcf.iYear = ilYear Then
                gUnpackDateLong tlRcf.iStartDate(0), tlRcf.iStartDate(1), llTDate
                If llTDate > llDate Then
                    ilRcfCode = tlRcf.iCode
                    ilYear = tlRcf.iYear
                    llDate = llTDate
                End If
            End If
        End If
        ilRet = btrGetNext(hgRNRcf, tlRcf, igRNRcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mObtainLatestRcf = ilRcfCode
End Function

Private Function mSetHdVef() As Long
    Dim ilNoVsf As Integer
    Dim ilCount As Integer
    Dim ilIndex As Integer
    Dim ilVsf As Integer
    Dim ilClf As Integer
    Dim ilFound As Integer
    Dim llLkVsfCode As Long
    Dim ilLoop As Integer
    Dim ilRet As Integer
    ReDim tgRNVsf(0 To 1) As VSF
    ilNoVsf = 0
    ilCount = 0
    ilIndex = 0
    For ilClf = LBound(tgRNClf) To UBound(tgRNClf) - 1 Step 1
        ilFound = False
        For ilVsf = LBound(tgRNVsf) To UBound(tgRNVsf) - 1 Step 1
            For ilLoop = LBound(tgRNVsf(0).iFSCode) To UBound(tgRNVsf(0).iFSCode) Step 1
                If tgRNVsf(ilVsf).iFSCode(ilLoop) > 0 Then
                    If tgRNVsf(ilVsf).iFSCode(ilLoop) = tgRNClf(ilClf).iVefCode Then
                        ilFound = True
                    End If
                Else
                    Exit For
                End If
            Next ilLoop
        Next ilVsf
        If Not ilFound Then
            If ilNoVsf > UBound(tgRNVsf(0).iFSCode) Then
                ilNoVsf = 0
                ilIndex = ilIndex + 1
                ReDim Preserve tgRNVsf(0 To UBound(tgRNVsf) + 1) As VSF
            End If
            tgRNVsf(ilIndex).iFSCode(ilNoVsf) = tgRNClf(ilClf).iVefCode
            ilNoVsf = ilNoVsf + 1
            ilCount = ilCount + 1
        End If
    Next ilClf
    If ilCount = 0 Then
        gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "|" & "|" & "Contract Not Added as not vehicles defined with Lines", True
        mSetHdVef = 0
        Exit Function
    ElseIf ilCount = 1 Then
        mSetHdVef = tgRNVsf(0).iFSCode(0)
        Exit Function
    Else
        llLkVsfCode = 0
        For ilVsf = UBound(tgRNVsf) - 1 To 0 Step 1
            tgRNVsf(ilVsf).lCode = 0
            tgRNVsf(ilVsf).sType = "F"
            tgRNVsf(ilVsf).sName = ""
            tgRNVsf(ilVsf).lLkVsfCode = llLkVsfCode
            For ilLoop = LBound(tgRNVsf(0).lFSComm) To UBound(tgRNVsf(0).lFSComm) Step 1
                tgRNVsf(ilVsf).iNoSpots(ilLoop) = 0
            Next ilLoop
            tgRNVsf(ilVsf).sSource = "S"
            tgRNVsf(ilVsf).iMerge = 0
            ilRet = btrInsert(hgRNVsf, tgRNVsf(ilVsf), igRNVsfRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                mSetHdVef = 0
                Exit Function
            End If
            llLkVsfCode = tgRNVsf(ilVsf).lCode    '.lLkVsfCode chged to .lcode
        Next ilVsf
        mSetHdVef = -tgRNVsf(0).lCode
        Exit Function
    End If
End Function

Private Function mOkToAddContract() As Integer
    'Test if contract exist or a newer one exist
    Dim tlChf As CHF
    Dim ilRet As Integer
    
    tgRNChfSrchKey1.lCntrNo = tgRNChf.lCntrNo
    tgRNChfSrchKey1.iCntRevNo = 32000
    tgRNChfSrchKey1.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hgRNChf, tlChf, igRNChfRecLen, tgRNChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tgRNChf.lCntrNo = tlChf.lCntrNo)
        If StrComp(Trim$(tgRNChf.sRepDBID), Trim$(tlChf.sRepDBID), vbTextCompare) <> 0 Then
            If Trim$(tlChf.sRepDBID) <> "" Then
                gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "|" & "|" & "Contract Not Added as another Rep using number, Rep ID " & Trim$(tlChf.sRepDBID), True
            Else
                gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "|" & "|" & "Contract Not Added as Network using number", True
            End If
            mOkToAddContract = False
            Exit Function
        Else
            If tgRNChf.iCntRevNo <= tlChf.iCntRevNo Then
                If (tlChf.sSchStatus <> "F") Then
                    gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "|" & "|" & "Contract Not Added as a newer one exist " & tgRNChf.iCntRevNo & " vs " & tlChf.iCntRevNo, True
                    mOkToAddContract = False
                    Exit Function
                Else
                    If ((imProdProt1Req) And (tlChf.iMnfComp(0) > 0)) Or ((imProdProt2Req) And (tlChf.iMnfComp(1) > 0)) Then
                        'Assume that version changed because of product protection changed
                    Else
                        gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "|" & "|" & "Contract Not Added as a newer one exist " & tgRNChf.iCntRevNo & " vs " & tlChf.iCntRevNo, True
                        mOkToAddContract = False
                        Exit Function
                    End If
                End If
            End If
            'If (imProdProt1Req) And (tlChf.iMnfComp(0) > 0) Then
            '    tgRNChf.iMnfComp(0) = tlChf.iMnfComp(0)
            'End If
            'If (imProdProt2Req) And (tlChf.iMnfComp(1) > 0) Then
            '    tgRNChf.iMnfComp(1) = tlChf.iMnfComp(1)
            'End If
        End If
        ilRet = btrGetNext(hgRNChf, tlChf, igRNChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'If imProdProt1Req Then
    '    gCreateAlert "M||" & "|" & tgRNChf.lCntrNo & "|" & smAdvertiserName & "||" & "Product Protection reference " & smProdProt1 & " removed from Contract", True
    'End If
    'If imProdProt2Req Then
    '    gCreateAlert "M||" & "|" & tgRNChf.lCntrNo & "|" & smAdvertiserName & "||" & "Product Protection reference " & smProdProt2 & " removed from Contract", True
    'End If
    mOkToAddContract = True
End Function

Private Function mClearContracts() As Integer
'
'   Clear any contract not scheduled or that have a larger internal revision number
'   Larger internal revision number is ignored in Test if Ok to add contract if Product protection is required
'
    Dim tlChf As CHF
    Dim tlClf As CLF
    Dim tlCff As CFF
    Dim tlVsf As VSF
    Dim llLkVsfCode As Long
    Dim ilRet As Integer
    
    tgRNChfSrchKey1.lCntrNo = tgRNChf.lCntrNo
    tgRNChfSrchKey1.iCntRevNo = 32000
    tgRNChfSrchKey1.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hgRNChf, tlChf, igRNChfRecLen, tgRNChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tgRNChf.lCntrNo = tlChf.lCntrNo)
        If ((tgRNChf.iCntRevNo <= tlChf.iCntRevNo) And (tlChf.sSchStatus = "F")) Then
            gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "|" & "|" & "Contract Not Added as a newer one exist " & tgRNChf.iCntRevNo & " vs " & tlChf.iCntRevNo, True
            mClearContracts = False
            Exit Function
        End If
        If (tlChf.sSchStatus <> "F") Then
            tgRNCffSrchKey0.lChfCode = tlChf.lCode
            tgRNCffSrchKey0.iClfLine = 0
            tgRNCffSrchKey0.iCntRevNo = 0
            tgRNCffSrchKey0.iPropVer = 0
            tgRNCffSrchKey0.iStartDate(0) = 0
            tgRNCffSrchKey0.iStartDate(1) = 0
            ilRet = btrGetGreaterOrEqual(hgRNCff, tlCff, igRNCffRecLen, tgRNCffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tlCff.lChfCode = tlChf.lCode)
                ilRet = btrDelete(hgRNCff)
                tgRNCffSrchKey0.lChfCode = tlChf.lCode
                tgRNCffSrchKey0.iClfLine = 0
                tgRNCffSrchKey0.iCntRevNo = 0
                tgRNCffSrchKey0.iPropVer = 0
                tgRNCffSrchKey0.iStartDate(0) = 0
                tgRNCffSrchKey0.iStartDate(1) = 0
                ilRet = btrGetGreaterOrEqual(hgRNCff, tlCff, igRNCffRecLen, tgRNCffSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
            Loop
            tgRNClfSrchKey1.lChfCode = tlChf.lCode
            tgRNClfSrchKey1.iVefCode = 0
            ilRet = btrGetGreaterOrEqual(hgRNClf, tlClf, igRNClfRecLen, tgRNClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlChf.lCode)
                ilRet = btrDelete(hgRNClf)
                tgRNClfSrchKey1.lChfCode = tlChf.lCode
                tgRNClfSrchKey1.iVefCode = 0
                ilRet = btrGetGreaterOrEqual(hgRNClf, tlClf, igRNClfRecLen, tgRNClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Loop
            If tlChf.lVefCode < 0 Then
                llLkVsfCode = -tlChf.lVefCode
                Do While llLkVsfCode > 0
                    tgRNVsfSrchKey0.lCode = llLkVsfCode
                    ilRet = btrGetEqual(hgRNVsf, tlVsf, igRNVsfRecLen, tgRNVsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        ilRet = btrDelete(hgRNVsf)
                    End If
                    llLkVsfCode = tlVsf.lLkVsfCode
                Loop
            
            End If
            ilRet = btrDelete(hgRNChf)
        End If
        ilRet = btrGetNext(hgRNChf, tlChf, igRNChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mClearContracts = True
End Function

Private Function mSetChfTotalDollars() As Long
    Dim ilCff As Integer
    Dim ilDay As Integer
    Dim llTotal As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slStartDate As String
    Dim llDate As Long
    Dim llMoStartDate As Long
    
    llTotal = 0
    If tgNrf.sTransmitPrice <> "N" Then
        For ilCff = 0 To UBound(tgRNCff) - 1 Step 1
            gUnpackDateLong tgRNCff(ilCff).iStartDate(0), tgRNCff(ilCff).iStartDate(1), llStartDate
            gUnpackDateLong tgRNCff(ilCff).iEndDate(0), tgRNCff(ilCff).iEndDate(1), llEndDate
            If llStartDate <= llEndDate Then
                If tgRNCff(ilCff).sPriceType = "T" Then
                    gUnpackDate tgRNCff(ilCff).iStartDate(0), tgRNCff(ilCff).iStartDate(1), slStartDate
                    slStartDate = gObtainPrevMonday(slStartDate)
                    llMoStartDate = gDateValue(slStartDate)
                    For llDate = llMoStartDate To llEndDate Step 7
                        If (tgRNCff(ilCff).sDyWk = "D") Then
                            For ilDay = 0 To 6 Step 1
                                If (ilDay + llDate >= llStartDate) And (ilDay + llDate <= llEndDate) Then
                                    llTotal = llTotal + tgRNCff(ilCff).iDay(ilDay) * tgRNCff(ilCff).lActPrice
                                End If
                            Next ilDay
                        Else
                            llTotal = llTotal + tgRNCff(ilCff).iSpotsWk * tgRNCff(ilCff).lActPrice
                        End If
                    Next llDate
                End If
            End If
        Next ilCff
    End If
    mSetChfTotalDollars = llTotal
End Function

Private Function mAddContract() As Integer
    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim ilCff As Integer
        
    mAddContract = False
    ilRet = gPutComment(smCommentInt, smShowOnInt, tgRNChf.lCxfInt)
    ilRet = gPutComment(smCommentChgR, smShowOnChgR, tgRNChf.lCxfChgR)
    ilRet = gPutComment(smCommentCanc, smShowOnCanc, tgRNChf.lCxfCanc)
    ilRet = gPutComment(smCommentOther, smShowOnOther, tgRNChf.lCxfCode)
    ilRet = btrInsert(hgRNChf, tgRNChf, igRNChfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "|" & "|" & "Contract Not Added, Pervasive Error # " & ilRet, True
        Exit Function
    End If
    For ilClf = 0 To UBound(tgRNClf) - 1 Step 1
        ilRet = gPutComment(smLnComment(ilClf), smShowOnLnComment(ilClf), tgRNClf(ilClf).lCxfCode)
        tgRNClf(ilClf).lChfCode = tgRNChf.lCode
        ilRet = btrInsert(hgRNClf, tgRNClf(ilClf), igRNClfRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "|" & "|" & "Contract Not Added as unable to add Line # " & tgRNClf(ilClf).iLine & ", Pervasive Error # " & ilRet, True
            Exit Function
        End If
    Next ilClf
    For ilCff = 0 To UBound(tgRNCff) - 1 Step 1
        tgRNCff(ilCff).lChfCode = tgRNChf.lCode
        ilRet = btrInsert(hgRNCff, tgRNCff(ilCff), igRNCffRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            gCreateMessage "C|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "|" & "|" & "Contract Not Added as unable to add flight for Line # " & tgRNClf(ilClf).iLine & ", Pervasive Error # " & ilRet, True
            Exit Function
        End If
    Next ilCff
    gCreateMessage "C|O|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "|" & "|" & "Contract Received successfully on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), True
    ilRet = gAlertAdd("C", "C", tgRNChf.lCode, 0, "")
    mAddContract = True
End Function

Public Sub gCreateMessage(slMessage As String, ilLogMsg As Integer)
    Dim ilRet As Integer
    
    'tgRNCef.iStrLen = Len(Trim$(slMessage))
    tgRNCef.sComment = Trim$(slMessage) & Chr$(0) '& Chr$(0) 'sgTB
    igRNCefRecLen = Len(tgRNCef)    '5 + Len(Trim$(tgRNCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
    tgRNCef.lCode = 0 'Autoincrement
    ilRet = btrInsert(hgRNCef, tgRNCef, igRNCefRecLen, INDEXKEY0)
    tgRNGif.lCode = 0
    tgRNGif.sNRProcessed = "N"
    tgRNGif.sType = "M"
    tgRNGif.lGenDate = gDateValue(Format$(gNow(), "m/d/yy"))
    tgRNGif.lGenTime = gTimeToLong(Format$(gNow(), "h:mm:ssAM/PM"), False)
    Select Case Left$(gGetLocalTZName(), 1)
        Case "E"
        Case "C"
            tgRNGif.lGenTime = tgRNGif.lGenTime + 3600
        Case "M"
            tgRNGif.lGenTime = tgRNGif.lGenTime + 7200
        Case "P"
            tgRNGif.lGenTime = tgRNGif.lGenTime + 10800
    End Select
    tgRNGif.lCntrNo = 0
    tgRNGif.iVefCode = 0
    tgRNGif.iStartDate(0) = 0
    tgRNGif.iStartDate(1) = 0
    tgRNGif.iEndDate(0) = 0
    tgRNGif.iEndDate(1) = 0
    tgRNGif.sMessageRepDBID = smRepDBID
    tgRNGif.sMessageNetDBID = smNetDBID
    tgRNGif.lMessageCefCode = tgRNCef.lCode
    gPackDate "", tgRNGif.iProcessedDate(0), tgRNGif.iProcessedDate(1)
    gPackTime "", tgRNGif.iProcessedTime(0), tgRNGif.iProcessedTime(1)
    ilRet = btrInsert(hgRNGif, tgRNGif, igRNGifRecLen, INDEXKEY0)
    Sleep (1500)
    If ilLogMsg Then
        gLogMsg slMessage, "RepNetLink.txt", False
    End If
End Sub



Public Sub gCreateAlert(slMessage As String, ilLogMsg)
    Dim ilRet As Integer
    
    'tgRNCef.iStrLen = Len(Trim$(slMessage))
    tgRNCef.sComment = Trim$(slMessage) & Chr$(0) '& Chr$(0) 'sgTB
    igRNCefRecLen = Len(tgRNCef)    '5 + Len(Trim$(tgRNCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
    tgRNCef.lCode = 0 'Autoincrement
    ilRet = btrInsert(hgRNCef, tgRNCef, igRNCefRecLen, INDEXKEY0)
    If ilRet = BTRV_ERR_NONE Then
        ilRet = gAlertAdd("M", "", 0, 0, Trim$(Str$(tgRNCef.lCode)))
    End If
    If ilLogMsg Then
        gLogMsg slMessage, "RepNetLink.txt", False
    End If
End Sub

Public Function gCreateMessageRec(slRepDBID As String, slNetDBID As String, llGenDate As Long, llGenTime As Long, llSeqNo As Long, tlCef As CEF) As String
    Dim slRecord As String
    Dim slStr As String
    
    gCreateMessageRec = ""
    'slStr = Left$(tlCef.sComment, tlCef.iStrLen)
    slStr = gStripChr0(tlCef.sComment)
    slRecord = ""
    
    slRecord = mAddQuotes(slRepDBID) & ","
    slRecord = slRecord & mAddQuotes(slNetDBID) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llGenDate))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llGenTime))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llSeqNo))) & ","
    If InStr(1, slStr, "Contract", vbTextCompare) > 0 Then
        slRecord = slRecord & mAddQuotes("C") & ","
        If InStr(1, slStr, "Error", vbTextCompare) > 0 Then
            slRecord = slRecord & mAddQuotes("E") & ","
        Else
            slRecord = slRecord & mAddQuotes("I") & ","
        End If
    ElseIf InStr(1, slStr, "Post Spots", vbTextCompare) > 0 Then
        slRecord = slRecord & mAddQuotes("S") & ","
        If InStr(1, slStr, "Error", vbTextCompare) > 0 Then
            slRecord = slRecord & mAddQuotes("E") & ","
        Else
            slRecord = slRecord & mAddQuotes("I") & ","
        End If
    Else
        slRecord = slRecord & mAddQuotes("M") & ","
        slRecord = slRecord & mAddQuotes("I") & ","
    End If
    slRecord = slRecord & mAddQuotes(slStr)
    gCreateMessageRec = slRecord
End Function

Public Function gTextToMessage(tlQif As QIF) As Integer
    Dim hlTextMsg As Integer
    Dim ilRet As Integer
    Dim slMessage As String
    Dim ilEof As Integer
    Dim ilMessageHeaderFound As Integer
    Dim ilIndex As Integer
    
    ilRet = mOpenTextForRead(hlTextMsg, tlQif.sFileName1)
    If Not ilRet Then
        gTextToMessage = False
        Exit Function
    End If
    If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
        ilRet = gParseItem(tlQif.sFileName1, 1, "_", smRepDBID)
        ilRet = gParseItem(tlQif.sFileName1, 2, "_", smNetDBID)
    ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
        ilRet = gParseItem(tlQif.sFileName1, 1, "_", smNetDBID)
        ilRet = gParseItem(tlQif.sFileName1, 2, "_", smRepDBID)
    End If
    smRepDBID = Trim$(smRepDBID)
    smNetDBID = Trim$(smNetDBID)
    ilMessageHeaderFound = False
    'Read and Build Contract header
    Do
        ilRet = 0
        err.Clear
        'On Error GoTo gTextToMessageErr:
        Line Input #hlTextMsg, slMessage
        On Error GoTo 0
        ilRet = err.Number
        If ilRet <> 0 Then
            Exit Do
        End If
        If Len(slMessage) > 0 Then
            If (Asc(slMessage) = 26) Then     'Ctrl Z
                Exit Do
            Else
                If StrComp(Left$(slMessage, 3), "mit", vbTextCompare) = 0 Then
                    gParseCDFields slMessage, False, smHdFieldNames()
                    ilMessageHeaderFound = True
                ElseIf ilMessageHeaderFound Then
                    gParseCDFields slMessage, False, smFields()
                    If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                        ilIndex = mFindFieldName(smHdFieldNames(), "mitRepDBID")
                    ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
                        ilIndex = mFindFieldName(smHdFieldNames(), "mitNetDBID")
                    End If
                    If ilIndex <> -1 Then
                        If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                            smRepDBID = Trim$(smFields(ilIndex))
                        ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
                            smNetDBID = Trim$(smFields(ilIndex))
                        End If
                        If StrComp(Trim$(smFields(ilIndex)), Trim$(tgNrf.sDBID), vbTextCompare) = 0 Then
                            If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                                ilIndex = mFindFieldName(smHdFieldNames(), "mitNetDBID")
                            ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
                                ilIndex = mFindFieldName(smHdFieldNames(), "mitRepDBID")
                            End If
                            If ilIndex <> -1 Then
                                If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                                    smNetDBID = Trim$(smFields(ilIndex))
                                ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
                                    smRepDBID = Trim$(smFields(ilIndex))
                                End If
                            Else
                                If ((Asc(tgSpf.sAutoType2) And RN_REP) = RN_REP) Then
                                    smNetDBID = ""
                                ElseIf ((Asc(tgSpf.sAutoType2) And RN_NET) = RN_NET) Then
                                    smRepDBID = ""
                                End If
                            End If
                            'Create message
                            ilIndex = mFindFieldName(smHdFieldNames(), "mitMessage")
                            If ilIndex = -1 Then
                                Exit Function
                            End If
                            slMessage = smFields(ilIndex)
                            gCreateAlert slMessage, False
                        End If
                    End If
                End If
            End If
        End If
    Loop Until ilEof
    If ilMessageHeaderFound = False Then
        gCreateMessage "M|E|" & "|" & "|" & "||" & "Message Not Added as Header not found in Text file " & Trim$(tlQif.sFileName1), True
    End If
    'Read and Build Contract lines and Flights
    Close hlTextMsg
    gTextToMessage = True
    Exit Function
'gTextToMessageErr:
'    ilRet = Err.Number
'    Resume Next
End Function


Private Sub mInitRdf()
    Dim ilLoop As Integer
    
    tgRNRdf.iCode = 0   'Internal code number for Rate Card program
    tgRNRdf.sName = "Mo-Su 12M-12M"   'Program/Time Name
    tgRNRdf.sBase = "N" 'Base daypart (Y=Yes, N=No)
    tgRNRdf.sReport = "N" 'Include as Daypart on reports (Y=Yes, N=No)
    tgRNRdf.sInOut = "N"   'N=N/A; I=Book into avail specified; O=Exclude avail specified
    tgRNRdf.ianfCode = 0   'Avail name code if sInOut = I or O
    tgRNRdf.sTimeOver = "Y" 'Y=OK to override times specified; N=Times can't to override
    tgRNRdf.sState = "A"    'A=Active; D=Drmant
    'Start Time, End Time , Week day and % spot only valid if NO library specified
    'The values are stored in a shifted format- high number indicies filled first
    'i.e. If two times defined, the first time will be stored in index 6 and the second time will
    '     be stored as index 7.  The first five indicies will be set to zero and blank.
    'For ilLoop = 1 To 7 Step 1
    For ilLoop = 0 To 6 Step 1
        tgRNRdf.iStartTime(0, ilLoop) = 1   'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                        'If Hund Sec = 1 and all other times = 0, then time not defined
                        'In Basic the left most dimension expands first
        tgRNRdf.iStartTime(1, ilLoop) = 0
        tgRNRdf.iEndTime(0, ilLoop) = 0   'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
        tgRNRdf.iEndTime(1, ilLoop) = 0
        tgRNRdf.sWkDays(ilLoop, 1) = "N" 'Dimension 1: Time; Dimension 2:Days;
                                            'Dimension 2: Index 1= Monday, 2= tuesday,...
                                            'Weekday flag: Y=day allowed; N=Day disallowed
        tgRNRdf.sWkDays(ilLoop, 2) = "N" 'Dimension 1: Time; Dimension 2:Days;
        tgRNRdf.sWkDays(ilLoop, 3) = "N" 'Dimension 1: Time; Dimension 2:Days;
        tgRNRdf.sWkDays(ilLoop, 4) = "N" 'Dimension 1: Time; Dimension 2:Days;
        tgRNRdf.sWkDays(ilLoop, 5) = "N" 'Dimension 1: Time; Dimension 2:Days;
        tgRNRdf.sWkDays(ilLoop, 6) = "N" 'Dimension 1: Time; Dimension 2:Days;
        tgRNRdf.sWkDays(ilLoop, 7) = "N" 'Dimension 1: Time; Dimension 2:Days;
        tgRNRdf.iSpotPct(ilLoop) = 0 'Percentage of spots (xxx.xx)
    Next ilLoop
    tgRNRdf.sContrRef = "Y"   'Y=Referenced by contract; N=Not Referenced
    tgRNRdf.iLtfCode(0) = 0    'Library code (set is automatically, 0=manual)
    tgRNRdf.iLtfCode(1) = 0
    tgRNRdf.iLtfCode(2) = 0
    tgRNRdf.iSortCode = 0   'Sort code number
    tgRNRdf.iRemoteID = 0       'Remote ID (Note: Unique ID=Remote ID + AutoCode)
    tgRNRdf.iAutoCode = 0       'Auto Incr Code (Note: Unique ID=Remote ID + AutoCode)
    tgRNRdf.iSyncDate(0) = 0 'Sync Date (from Master Server or Remote System)
    tgRNRdf.iSyncDate(1) = 0 'Sync Date (from Master Server or Remote System)
    tgRNRdf.iSyncTime(0) = 0  'Sync Time (from Master Server or Remote System)
    tgRNRdf.iSyncTime(1) = 0  'Sync Time (from Master Server or Remote System)
    tgRNRdf.iSourceID = 0       'Remote User ID (avoid resending to sender)
    tgRNRdf.sUnused = ""
End Sub



Public Function gCreatePostedSpotRec(slRepDBID As String, slNetDBID As String, llGenDate As Long, llGenTime As Long, llSeqNo As Long, tlSdf As SDF, tlChf As CHF) As String
    Dim ilAdf As Integer
    Dim slBonusOnInv As String
    Dim slBonusSpotType As String
    Dim slStr As String
    Dim ilVefCode As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slRecord As String
    Dim ilRet As Integer
    
    gCreatePostedSpotRec = ""
    slRecord = ""
    
    slRecord = mAddQuotes(slRepDBID) & ","
    slRecord = slRecord & mAddQuotes(slNetDBID) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llGenDate))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llGenTime))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llSeqNo))) & ","
    slRecord = slRecord & tlChf.lCntrNo & ","
    slRecord = slRecord & tlSdf.iLineNo & ","
    slRecord = slRecord & tlSdf.lCode & ","
    'If tlSdf.sSpotType = "X" Then
    '    slBonusOnInv = "Y"
    '    tgRNAdfSrchKey.iCode = tlChf.iAdfCode
    '    ilRet = btrGetEqual(hgRNAdf, tgRNAdf, igRNAdfRecLen, tgRNAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    '    If ilRet = BTRV_ERR_NONE Then
    '        slBonusOnInv = tgRNAdf.sBonusOnInv
    '    End If
    '    If tlSdf.sPriceType = "+" Then
    '        slBonusSpotType = "+"
    '    ElseIf tlSdf.sPriceType = "-" Then
    '        slBonusSpotType = "-"
    '    Else
    '        If slBonusOnInv <> "N" Then
    '            slBonusSpotType = "+"
    '        Else
    '            slBonusSpotType = "-"
    '        End If
    '    End If
    '    slRecord = slRecord & slBonusSpotType & ","
    'ElseIf tlSdf.sSpotType = "O" Then
    '    slRecord = slRecord & "B" & ","       'Begin BB: Open
    'ElseIf tlSdf.sSpotType = "C" Then
    '    slRecord = slRecord & "E" & ","       'End BB: Close
    'Else
    '    slRecord = slRecord & tlSdf.sSchStatus & ","  'M=issed;S=Scheduled;O=Outside;G=MG;H=Hidden;C=Cancel
    'End If
    slRecord = slRecord & tlSdf.sSpotType & ","
    slRecord = slRecord & tlSdf.sPriceType & ","
    slRecord = slRecord & tlSdf.sSchStatus & ","
    If (tlSdf.sSchStatus = "O") Or (tlSdf.sSchStatus = "G") Then
        tgRNSmfSrchKey2.lCode = tlSdf.lCode
        ilRet = btrGetEqual(hgRNSmf, tgRNSmf, igRNSmfRecLen, tgRNSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            'gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Line Vehicle(vefCode = " & Trim$(Str$(tlSdf.iVefCode)) & ", Pervasive Error # " & ilRet, True
            Exit Function
        End If
        ilVefCode = tgRNSmf.iOrigSchVef
    Else
        ilVefCode = tlSdf.iVefCode
    End If
    tgRNVefSrchKey0.iCode = ilVefCode
    ilRet = btrGetEqual(hgRNVef, tgRNVef, igRNVefRecLen, tgRNVefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        'gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Line Vehicle(vefCode = " & Trim$(Str$(ilVefCode)) & ", Pervasive Error # " & ilRet, True
        Exit Function
    End If
    slRecord = slRecord & mAddQuotes(tgRNVef.sName) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(ilVefCode))) & ","
    If (tlSdf.sSchStatus = "O") Or (tlSdf.sSchStatus = "G") Then
        gUnpackDate tgRNSmf.iMissedDate(0), tgRNSmf.iMissedDate(1), slDate
        gUnpackTime tgRNSmf.iMissedTime(0), tgRNSmf.iMissedTime(1), "A", "1", slTime
    Else
        gUnpackDate tlSdf.iDate(0), tlSdf.iDate(1), slDate
        gUnpackTime tlSdf.iTime(0), tlSdf.iTime(1), "A", "1", slTime
    End If
    slDate = gAdjYear(slDate)
    slRecord = slRecord & mAddQuotes(Trim$(slDate)) & ","
    slRecord = slRecord & mAddQuotes(Trim$(slTime)) & ","
    
    ilVefCode = tlSdf.iVefCode
    tgRNVefSrchKey0.iCode = ilVefCode
    ilRet = btrGetEqual(hgRNVef, tgRNVef, igRNVefRecLen, tgRNVefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        'gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Line Vehicle(vefCode = " & Trim$(Str$(ilVefCode)) & ", Pervasive Error # " & ilRet, True
        Exit Function
    End If
    slRecord = slRecord & mAddQuotes(tgRNVef.sName) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(ilVefCode))) & ","
    gUnpackDate tlSdf.iDate(0), tlSdf.iDate(1), slDate
    gUnpackTime tlSdf.iTime(0), tlSdf.iTime(1), "A", "1", slTime
    slDate = gAdjYear(slDate)
    slRecord = slRecord & mAddQuotes(Trim$(slDate)) & ","
    slRecord = slRecord & mAddQuotes(Trim$(slTime)) & ","
    slRecord = slRecord & tlSdf.sXCrossMidnight & ","
    If tlSdf.sPtType = "1" Then
        tgRNCifSrchKey0.lCode = tlSdf.lCopyCode
        ilRet = btrGetEqual(hgRNCif, tgRNCif, igRNCifRecLen, tgRNCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            tgRNCpfSrchKey0.lCode = tgRNCif.lcpfCode
            ilRet = btrGetEqual(hgRNCpf, tgRNCpf, igRNCpfRecLen, tgRNCpfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                slRecord = slRecord & mAddQuotes(Trim$(tgRNCpf.sISCI))
            Else
                slRecord = slRecord
            End If
        Else
            slRecord = slRecord
        End If
    Else
        slRecord = slRecord
    End If
    gCreatePostedSpotRec = slRecord
End Function

Public Function gCreateAvailSumRec(slRepDBID As String, slNetDBID As String, llGenDate As Long, llGenTime As Long, llSeqNo As Long, tlAsf As ASF) As String
    Dim slRecord As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim ilLoop As Integer
    Dim slRdfStartTime As String
    Dim slRdfEndTime As String
    Dim ilCount As Integer
    Dim slStr As String
    Dim ilDay As Integer
    Dim ilRdf As Integer
    
    gCreateAvailSumRec = ""
    slRecord = ""
    
    slRecord = mAddQuotes(slRepDBID) & ","
    slRecord = slRecord & mAddQuotes(slNetDBID) & ","
    tgRNVefSrchKey0.iCode = tlAsf.iVefCode
    ilRet = btrGetEqual(hgRNVef, tgRNVef, igRNVefRecLen, tgRNVefSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        'gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Line Vehicle(vefCode = " & Trim$(Str$(ilVefCode)) & ", Pervasive Error # " & ilRet, True
        Exit Function
    End If
    slRecord = slRecord & mAddQuotes(tgRNVef.sName) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlAsf.iVefCode))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llGenDate))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llGenTime))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(llSeqNo))) & ","
    tgRNRdfSrchKey.iCode = tlAsf.iRdfCode
    ilRet = btrGetEqual(hgRNRdf, tgRNRdf, igRNRdfRecLen, tgRNRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        'gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Daypart Name(rdfCode = " & Trim$(Str$(tlClf.iRdfcode)) & ", Pervasive Error # " & ilRet, True
        Exit Function
    End If
    slRecord = slRecord & mAddQuotes(tgRNRdf.sName) & ","
    slRecord = slRecord & mAddQuotes(Trim$(Str$(tlAsf.iRdfCode))) & ","
    slRecord = slRecord & mAddQuotes(Trim$(tgRNRdf.sInOut)) & ","
    If tgRNRdf.ianfCode > 0 Then
        tgRNAnfSrchKey.iCode = tgRNRdf.ianfCode
        ilRet = btrGetEqual(hgRNAnf, tgRNAnf, igRNAnfRecLen, tgRNAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            'gCreateAlert "C|E|" & "|" & Trim$(Str$(tlChf.lCntrNo)) & "|" & "||" & "Accessing Avail Name(anfCode = " & Trim$(Str$(tgRNRdf.ianfCode)) & " failed, Pervasive Error # " & ilRet, True
            Exit Function
        End If
        slRecord = slRecord & mAddQuotes(tgRNAnf.sName) & ","
    Else
        slRecord = slRecord & "" & ","
    End If
    ilCount = 0
    For ilRdf = LBound(tgRNRdf.iStartTime, 2) To UBound(tgRNRdf.iStartTime, 2) Step 1
        If (tgRNRdf.iStartTime(0, ilRdf) <> 1) Or (tgRNRdf.iStartTime(1, ilRdf) <> 0) Then
            slStr = ""
            For ilDay = LBound(tgRNRdf.sWkDays, 1) To UBound(tgRNRdf.sWkDays, 1) Step 1
                slStr = slStr & tgRNRdf.sWkDays(ilRdf, ilDay)
            Next ilDay
            slRecord = slRecord & slStr & ","
            gUnpackTime tgRNRdf.iStartTime(0, ilRdf), tgRNRdf.iStartTime(1, ilRdf), "A", "1", slRdfStartTime
            gUnpackTime tgRNRdf.iEndTime(0, ilRdf), tgRNRdf.iEndTime(1, ilRdf), "A", "1", slRdfEndTime
            slRecord = slRecord & mAddQuotes(Trim$(slRdfStartTime)) & ","
            slRecord = slRecord & mAddQuotes(Trim$(slRdfEndTime)) & ","
            slRecord = slRecord & Trim$(Str$(tgRNRdf.iSpotPct(ilRdf))) & ","
            ilCount = ilCount + 1
            If ilCount = 2 Then
                Exit For
            End If
        End If
    Next ilRdf
    If ilCount = 0 Then
        slRecord = slRecord & "" & ","
        slRecord = slRecord & "" & ","
        slRecord = slRecord & "" & ","
        slRecord = slRecord & "" & ","
        ilCount = 1
    End If
    If ilCount = 1 Then
        slRecord = slRecord & "" & ","
        slRecord = slRecord & "" & ","
        slRecord = slRecord & "" & ","
        slRecord = slRecord & "" & ","
    End If
    gUnpackDate tlAsf.iStartWkDate(0), tlAsf.iStartWkDate(1), slDate
    slDate = gAdjYear(slDate)
    slRecord = slRecord & mAddQuotes(Trim$(slDate)) & ","
    For ilLoop = LBound(tlAsf.sAvail60) To UBound(tlAsf.sAvail60) Step 2
        slRecord = slRecord & mAddQuotes(Trim$(Str$(Asc(tlAsf.sAvail60(ilLoop)) * 256 + Asc(tlAsf.sAvail60(ilLoop + 1))))) & ","
    Next ilLoop
    For ilLoop = LBound(tlAsf.sAvail30) To UBound(tlAsf.sAvail30) Step 2
        slRecord = slRecord & mAddQuotes(Trim$(Str$(Asc(tlAsf.sAvail30(ilLoop)) * 256 + Asc(tlAsf.sAvail30(ilLoop + 1))))) & ","
    Next ilLoop
    For ilLoop = LBound(tlAsf.sAvail10) To UBound(tlAsf.sAvail10) Step 2
        slRecord = slRecord & mAddQuotes(Trim$(Str$(Asc(tlAsf.sAvail10(ilLoop)) * 256 + Asc(tlAsf.sAvail10(ilLoop + 1))))) & ","
    Next ilLoop
    gUnpackDate tlAsf.iCreatedDate(0), tlAsf.iCreatedDate(1), slDate
    slDate = gAdjYear(slDate)
    slRecord = slRecord & mAddQuotes(Trim$(slDate)) & ","
    gUnpackTime tlAsf.iCreatedTime(0), tlAsf.iCreatedTime(1), "A", "1", slStr
    slRecord = slRecord & mAddQuotes(Trim$(slStr))
    gCreateAvailSumRec = slRecord
End Function
Public Function gTextToPostedSpots(tlQif As QIF) As Integer
    Dim hlTextSdf As Integer
    Dim ilRet As Integer
    Dim ilPSHeaderFound As Integer
    Dim ilEof As Integer
    Dim slPostedSpot As String
    Dim ilIndex As Integer
    
    ilRet = gObtainVef()
    If Not ilRet Then
        gTextToPostedSpots = False
        Exit Function
    End If
    ilRet = mOpenTextForRead(hlTextSdf, tlQif.sFileName1)
    If Not ilRet Then
        gTextToPostedSpots = False
        Exit Function
    End If
    
    ReDim tmSbf(0 To 0) As SBF
    lmInvStdEndDate = tlQif.lCntrNo 'Used to pass Invoice end date
    If lmInvStdEndDate > 0 Then
        lmInvCalEndDate = gDateValue(gObtainEndCal(Format$(lmInvStdEndDate, "m/d/yy")))
    Else
        lmInvCalEndDate = 0
    End If
    ilRet = gParseItem(tlQif.sFileName1, 1, "_", smRepDBID)
    ilRet = gParseItem(tlQif.sFileName1, 2, "_", smNetDBID)
    
    Do
        ilRet = 0
        err.Clear
        'On Error GoTo gTextToPostedSpotsErr:
        Line Input #hlTextSdf, slPostedSpot
        On Error GoTo 0
        ilRet = err.Number
        If ilRet <> 0 Then
            Exit Do
        End If
        If Len(slPostedSpot) > 0 Then
            If (Asc(slPostedSpot) = 26) Then     'Ctrl Z
                Exit Do
            Else
                If StrComp(Left$(slPostedSpot, 3), "pst", vbTextCompare) = 0 Then
                    gParseCDFields slPostedSpot, False, smHdFieldNames()
                    ilPSHeaderFound = True
                ElseIf ilPSHeaderFound Then
                    gParseCDFields slPostedSpot, False, smFields()
                    ilIndex = mFindFieldName(smHdFieldNames(), "pstRepDBID")
                    If ilIndex <> -1 Then
                        smRepDBID = Trim$(smFields(ilIndex))
                        If StrComp(Trim$(smFields(ilIndex)), Trim$(tgNrf.sDBID), vbTextCompare) = 0 Then
                            ilIndex = mFindFieldName(smHdFieldNames(), "pstNetDBID")
                            If ilIndex <> -1 Then
                                smNetDBID = Trim$(smFields(ilIndex))
                            Else
                                smNetDBID = ""
                            End If
                            ilRet = mCreatePostedSpot()
                            If Not ilRet Then
                                'Message not required as mCreatePostedSpots have generated the error messages
                                'gCreateMessage "P|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "||" & "Contract Not Added, see Previous message.", True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Loop Until ilEof
    If ilPSHeaderFound = False Then
        gCreateMessage "P|E|" & Trim$(tgNrf.sDBID) & "|" & "|" & "||" & "Posted Spots Not Added as Header not found in Text file " & Trim$(tlQif.sFileName1), True
    End If
    Close hlTextSdf
    ilRet = mUpdateSbf()
    gTextToPostedSpots = True
    Exit Function
'gTextToPostedSpotsErr:
'    ilRet = Err.Number
'    Resume Next

End Function
Public Function gTextToAvailSum(tlQif As QIF) As Integer
    Dim hlTextAsf As Integer
    Dim ilRet As Integer
    Dim ilASHeaderFound As Integer
    Dim ilEof As Integer
    Dim slAvailSum As String
    Dim ilIndex As Integer
    Dim ilfirstTime As Integer
    
    ilRet = gObtainVef()
    If Not ilRet Then
        gTextToAvailSum = False
        Exit Function
    End If
    ilRet = mOpenTextForRead(hlTextAsf, tlQif.sFileName1)
    If Not ilRet Then
        gTextToAvailSum = False
        Exit Function
    End If
    ilRet = gParseItem(tlQif.sFileName1, 1, "_", smRepDBID)
    ilRet = gParseItem(tlQif.sFileName1, 2, "_", smNetDBID)
    ilfirstTime = True
    Do
        ilRet = 0
        err.Clear
        'On Error GoTo gTextToAvailSumErr:
        Line Input #hlTextAsf, slAvailSum
        On Error GoTo 0
        ilRet = err.Number
        If ilRet <> 0 Then
            Exit Do
        End If
        If Len(slAvailSum) > 0 Then
            If (Asc(slAvailSum) = 26) Then     'Ctrl Z
                Exit Do
            Else
                If StrComp(Left$(slAvailSum, 3), "amt", vbTextCompare) = 0 Then
                    gParseCDFields slAvailSum, False, smHdFieldNames()
                    ilASHeaderFound = True
                ElseIf ilASHeaderFound Then
                    gParseCDFields slAvailSum, False, smFields()
                    ilIndex = mFindFieldName(smHdFieldNames(), "amtRepDBID")
                    If ilIndex <> -1 Then
                        smRepDBID = Trim$(smFields(ilIndex))
                        If StrComp(Trim$(smFields(ilIndex)), Trim$(tgNrf.sDBID), vbTextCompare) = 0 Then
                            ilRet = mCreateAvailSum(ilfirstTime)
                            ilfirstTime = False
                            If Not ilRet Then
                                'Message not required as mCreateAvailSum have generated the error messages
                                'gCreateMessage "A|E|" & Trim$(tgNrf.sDBID) & "|" & smCntrNo & "|" & smAdvertiserName & "||" & "Contract Not Added, see Previous message.", True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Loop Until ilEof
    If ilASHeaderFound = False Then
        'gCreateMessage "M|E|" & "|" & "|" & "||" & "Message Not Added as Header not found in Text file " & Trim$(tlQif.sFileName1), True
        gCreateMessage "A|E|" & Trim$(tgNrf.sDBID) & "|" & "|" & "||" & "Avail Summary Not Added as Header not found in Text file " & Trim$(tlQif.sFileName1), True
    End If
    
    Close hlTextAsf
    gTextToAvailSum = True
    Exit Function
'gTextToAvailSumErr:
'    ilRet = Err.Number
'    Resume Next
End Function


Private Function mCreatePostedSpot() As Integer
    'slTitle = "pstRepDBID,pstNetDBID,pstGenDate,pstGenTime,pstSeqNo,pstContractNo,pstLineNo,pstSpotID,pstSpotType,pstPriceType,pstSchStatus,pstSchdVehicleName,pstSchdVehicleCode,pstSchdDate,pstSchdTime,pstAirVehicleName,pstAirVehicleCode,pstAirDate,pstAirTime,pstXmidnight,pstISCI"
    Dim ilIndex As Integer
    Dim llCntrNo As Long
    Dim ilLineNo As Integer
    Dim llSpotID As Long
    Dim slSpotType As String
    Dim slPriceType As String
    Dim slSchStatus As String
    Dim slSchdVehicleName As String
    Dim slSchdDate As String
    Dim slSchdTime As String
    Dim slAiredVehicleName As String
    Dim slAiredDate As String
    Dim slAiredTime As String
    Dim slXMidnight As String
    Dim slISCI As String
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilAirVehicleCode As Integer
    Dim tlSdf As SDF
    Dim slDate As String
    Dim llPrice As Long
    Dim slPrice As String
    Dim slPaidOrBonus As String
    
    mCreatePostedSpot = False
    ilIndex = mFindFieldName(smHdFieldNames(), "pstContractNo")
    If ilIndex = -1 Then
        Exit Function
    End If
    llCntrNo = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "pstLineNo")
    If ilIndex = -1 Then
        Exit Function
    End If
    ilLineNo = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "pstSpotID")
    If ilIndex = -1 Then
        Exit Function
    End If
    llSpotID = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "pstSpotType")
    If ilIndex = -1 Then
        Exit Function
    End If
    slSpotType = smFields(ilIndex)
    ilIndex = mFindFieldName(smHdFieldNames(), "pstPriceType")
    If ilIndex = -1 Then
        Exit Function
    End If
    slPriceType = smFields(ilIndex)
    ilIndex = mFindFieldName(smHdFieldNames(), "pstSchStatus")
    If ilIndex = -1 Then
        Exit Function
    End If
    slSchStatus = smFields(ilIndex)
    ilIndex = mFindFieldName(smHdFieldNames(), "pstSchdVehicleName")
    If ilIndex = -1 Then
        Exit Function
    End If
    slSchdVehicleName = smFields(ilIndex)
    ilIndex = mFindFieldName(smHdFieldNames(), "pstSchdDate")
    If ilIndex = -1 Then
        Exit Function
    End If
    slSchdDate = smFields(ilIndex)
    ilIndex = mFindFieldName(smHdFieldNames(), "pstSchdTime")
    If ilIndex = -1 Then
        Exit Function
    End If
    slSchdTime = smFields(ilIndex)
    
    ilIndex = mFindFieldName(smHdFieldNames(), "pstAirVehicleName")
    If ilIndex = -1 Then
        Exit Function
    End If
    slAiredVehicleName = Trim$(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "pstAirVehicleCode")
    If ilIndex = -1 Then
        Exit Function
    End If
    ilAirVehicleCode = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(smHdFieldNames(), "pstAirDate")
    If ilIndex = -1 Then
        Exit Function
    End If
    slAiredDate = smFields(ilIndex)
    ilIndex = mFindFieldName(smHdFieldNames(), "pstAirTime")
    If ilIndex = -1 Then
        Exit Function
    End If
    slAiredTime = smFields(ilIndex)
    
    ilIndex = mFindFieldName(smHdFieldNames(), "pstXmidnight")
    If ilIndex = -1 Then
        Exit Function
    End If
    slXMidnight = smFields(ilIndex)
    ilIndex = mFindFieldName(smHdFieldNames(), "pstISCI")
    If ilIndex = -1 Then
        Exit Function
    End If
    slISCI = smFields(ilIndex)
    
    tgRNChfSrchKey1.lCntrNo = llCntrNo
    tgRNChfSrchKey1.iCntRevNo = 32000
    tgRNChfSrchKey1.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hgRNChf, tgRNChf, igRNChfRecLen, tgRNChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tgRNChf.lCntrNo = llCntrNo)
        If ((tgRNChf.sSchStatus = "F") Or (tgRNChf.sSchStatus = "M")) And (tgRNChf.sDelete <> "Y") Then
            Exit Do
        End If
        ilRet = btrGetNext(hgRNChf, tgRNChf, igRNChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (tgRNChf.lCntrNo <> llCntrNo) Or (tgRNChf.sDelete = "Y") Then
        'Add error
        'Send message to Rep
        gCreateAlert "P|E|" & smNetDBID & "|" & llCntrNo & "|" & "||" & "Contract not found.", True
        'Send message back to Net
        gCreateMessage "P|E|" & smRepDBID & "|" & llCntrNo & "|" & "||" & "Contract not found.", True
        Exit Function
    End If
    
    tgRNClfSrchKey0.lChfCode = tgRNChf.lCode
    tgRNClfSrchKey0.iLine = ilLineNo
    tgRNClfSrchKey0.iCntRevNo = tgRNChf.iCntRevNo
    tgRNClfSrchKey0.iPropVer = tgRNChf.iPropVer
    ilRet = btrGetGreaterOrEqual(hgRNClf, tmClf, igRNClfRecLen, tgRNClfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (ilRet <> BTRV_ERR_NONE) Or (tmClf.lChfCode <> tgRNChf.lCode) Or (tmClf.iLine <> ilLineNo) Then
        'Add error
        'Send message to Rep
        gCreateAlert "P|E|" & smNetDBID & "|" & llCntrNo & "|" & "||" & "Contract line " & ilLineNo & " not found.", True
        'Send message back to Net
        gCreateMessage "P|E|" & smRepDBID & "|" & llCntrNo & "|" & "||" & "Contract line " & ilLineNo & " not found.", True
        Exit Function
    End If
    
    tgRNSdf.lCode = 0
    tgRNMifSrchKey1.sType = "P"
    tgRNMifSrchKey1.sNetDBID = smNetDBID
    tgRNMifSrchKey1.lNetDBCode = llSpotID
    ilRet = btrGetEqual(hgRNMif, tgRNMif, igRNMifRecLen, tgRNMifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If (ilRet = BTRV_ERR_NONE) Then
        tgRNSdf.lCode = tgRNMif.lRepDBCode
    End If
    
    tgRNSdf.iVefCode = 0
    tgRNMifSrchKey1.sType = "V"
    tgRNMifSrchKey1.sNetDBID = smNetDBID
    tgRNMifSrchKey1.lNetDBCode = ilAirVehicleCode
    ilRet = btrGetEqual(hgRNMif, tgRNMif, igRNMifRecLen, tgRNMifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If (ilRet = BTRV_ERR_NONE) Then
        tgRNSdf.iVefCode = tgRNMif.lRepDBCode
    End If
    If tgRNSdf.iVefCode = 0 Then
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If StrComp(Trim$(tgMVef(ilVef).sName), slAiredVehicleName, vbTextCompare) = 0 Then
                tgRNSdf.iVefCode = tgMVef(ilVef).iCode
                tgRNMif.lCode = 0
                tgRNMif.sNetDBID = smNetDBID
                tgRNMif.sRepDBID = smRepDBID
                tgRNMif.lNetDBCode = ilAirVehicleCode
                tgRNMif.lRepDBCode = tgMVef(ilVef).iCode
                tgRNMif.sType = "V"
                tgRNMif.sUnused = ""
                ilRet = btrInsert(hgRNMif, tgRNMif, igRNMifRecLen, INDEXKEY0)
                Exit For
            End If
        Next ilVef
    End If
    If tgRNSdf.iVefCode = 0 Then
        'Add error message
        'Send message to Rep
        gCreateAlert "P|E|" & smNetDBID & "|" & llCntrNo & "|" & "||" & "Vehicle " & slAiredVehicleName & " not found.", True
        'Send message back to Net
        gCreateMessage "P|E|" & smRepDBID & "|" & llCntrNo & "|" & "||" & "Vehicle " & slAiredVehicleName & " not found.", True
        Exit Function
    End If
    tgRNSdf.lChfCode = tgRNChf.lCode
    tgRNSdf.iLineNo = ilLineNo
    tgRNSdf.iAdfCode = tgRNChf.iAdfCode
    gPackDate slAiredDate, tgRNSdf.iDate(0), tgRNSdf.iDate(1)
    gPackTime slAiredTime, tgRNSdf.iTime(0), tgRNSdf.iTime(1)
    tgRNSdf.sPriceType = slPriceType
    tgRNSdf.sSpotType = slSpotType
    tgRNSdf.sSchStatus = slSchStatus
    If slSpotType = "O" Then
        tgRNSdf.iLen = tmClf.iBBOpenLen
    ElseIf slSpotType = "C" Then
        tgRNSdf.iLen = tmClf.iBBCloseLen
    Else
        tgRNSdf.iLen = tmClf.iLen
    End If
    
    tgRNSdf.iMnfMissed = 0
    tgRNSdf.sTracer = ""
    tgRNSdf.sAffChg = ""
    tgRNSdf.sPtType = "0"
    tgRNSdf.lCopyCode = mAddCopy(slISCI, tmClf.iLen)
    If tgRNSdf.lCopyCode > 0 Then
        tgRNSdf.sPtType = "1"
    End If
    tgRNSdf.sBill = "N"
    tgRNSdf.iRotNo = 0
    tgRNSdf.lSmfCode = 0
    tgRNSdf.iUrfCode = tgUrf(0).iCode
    tgRNSdf.lFsfCode = 0
    tgRNSdf.iGameNo = 0
    tgRNSdf.sXCrossMidnight = slXMidnight
    tgRNSdf.sUnused = ""
    If tgRNSdf.lCode = 0 Then
        ilRet = btrInsert(hgRNSdf, tgRNSdf, igRNSdfRecLen, INDEXKEY0)
        tgRNMif.lCode = 0
        tgRNMif.sNetDBID = smNetDBID
        tgRNMif.sRepDBID = smRepDBID
        tgRNMif.lNetDBCode = llSpotID
        tgRNMif.lRepDBCode = tgRNSdf.lCode
        tgRNMif.sType = "P"
        tgRNMif.sUnused = ""
        ilRet = btrInsert(hgRNMif, tgRNMif, igRNMifRecLen, INDEXKEY0)
    Else
        Do
            tgRNSdfSrchKey3.lCode = tgRNSdf.lCode
            ilRet = btrGetEqual(hgRNSdf, tlSdf, igRNSdfRecLen, tgRNSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrUpdate(hgRNSdf, tgRNSdf, igRNSdfRecLen)
            Else
                tgRNSdf.lCode = 0
                ilRet = btrInsert(hgRNSdf, tgRNSdf, igRNSdfRecLen, INDEXKEY0)
                Exit Do
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
    End If
    If (slSchStatus = "O") Or (slSchStatus = "G") Then
        'Create Outside/MG Specification (smf)
        ilRet = mCreateSMF(slSpotType, slSchdDate, slSchdTime, slAiredDate, slAiredTime, tmClf.iVefCode)
        Do
            tgRNSdfSrchKey3.lCode = tgRNSdf.lCode
            ilRet = btrGetEqual(hgRNSdf, tlSdf, igRNSdfRecLen, tgRNSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                tgRNSdf.lSmfCode = tgRNSmf.lCode
                ilRet = btrUpdate(hgRNSdf, tgRNSdf, igRNSdfRecLen)
            Else
                Exit Function
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
    End If
    ilVef = gBinarySearchVef(tmClf.iVefCode)
    If ilVef <> -1 Then
        If tgMVef(ilVef).sType = "R" Then
            If tgRNSdf.sSpotType = "X" Then
                slPaidOrBonus = "B"
                llPrice = 0
            Else
                slPaidOrBonus = "P"
                ilRet = gGetFlightPrice(tgRNSdf, tmClf, hgRNCff, hgRNSmf, slPrice)
                If ilRet Then
                    If InStr(1, slPrice, ".", vbTextCompare) > 0 Then
                        llPrice = gStrDecToLong(slPrice, 2)
                    Else
                        llPrice = 0
                    End If
                Else
                    llPrice = 0
                End If
            End If
            If lmInvStdEndDate <= 0 Then
                lmInvStdEndDate = gDateValue(gObtainEndStd(slAiredDate))
                lmInvCalEndDate = gDateValue(gObtainEndCal(Format$(lmInvStdEndDate, "m/d/yy")))
            End If
            mInitSbf llPrice, slAiredDate, slPaidOrBonus
        End If
    End If
    mCreatePostedSpot = True
End Function

Private Function mAddCopy(slISCI As String, ilLen As Integer) As Long
    Dim ilRet As Integer
    Dim slDate As String
    
    If Trim$(slISCI) = "" Then
        mAddCopy = 0
        Exit Function
    End If
    tgRNCpfSrchKey1.sISCI = slISCI
    ilRet = btrGetEqual(hgRNCpf, tgRNCpf, igRNCpfRecLen, tgRNCpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        'Add Product/ISCI
        tgRNCpf.lCode = 0
        tgRNCpf.sName = ""
        tgRNCpf.sISCI = slISCI
        tgRNCpf.sCreative = ""
        tgRNCpf.iRotEndDate(0) = 0
        tgRNCpf.iRotEndDate(1) = 0
        tgRNCpf.lSifCode = 0
        ilRet = btrInsert(hgRNCpf, tgRNCpf, igRNCpfRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            mAddCopy = 0
            Exit Function
        End If
    End If
    'Add copy inventory
    tgRNCif.lCode = 0  'Autoincrement
    tgRNCif.iMcfCode = 0
    tgRNCif.sName = ""
    tgRNCif.sCut = ""
    tgRNCif.sReel = ""
    tgRNCif.iLen = ilLen
    tgRNCif.iEtfCode = 0
    tgRNCif.iEnfCode = 0
    tgRNCif.iAdfCode = tgChfRep.iAdfCode
    tgRNCif.lcpfCode = tgRNCpf.lCode
    tgRNCif.iMnfComp(0) = 0
    tgRNCif.iMnfComp(1) = 0
    tgRNCif.iMnfAnn = 0
    tgRNCif.sHouse = "Y"
    tgRNCif.sCleared = "Y"
    tgRNCif.lCsfCode = 0
    tgRNCif.iNoTimesAir = 0
    tgRNCif.sCartDisp = "N"
    tgRNCif.sTapeDisp = "N"
    tgRNCif.sPurged = "A"
    tgRNCif.iPurgeDate(0) = 0
    tgRNCif.iPurgeDate(1) = 0
    slDate = Format$(gNow(), "m/d/yy")
    gPackDate slDate, tgRNCif.iDateEntrd(0), tgRNCif.iDateEntrd(1)
    tgRNCif.iUsedDate(0) = 0
    tgRNCif.iUsedDate(1) = 0
    tgRNCif.iRotStartDate(0) = 0
    tgRNCif.iRotStartDate(1) = 0
    tgRNCif.iRotEndDate(0) = 0
    tgRNCif.iRotEndDate(1) = 0
    tgRNCif.iUrfCode = tgUrf(0).iCode
    tgRNCif.sPrint = "N"
    tgRNCif.iLangMnfCode = 0
    tgRNCif.sUnused = ""
    ilRet = btrInsert(hgRNCif, tgRNCif, igRNCifRecLen, INDEXKEY0)
    If ilRet = BTRV_ERR_NONE Then
        mAddCopy = tgRNCif.lCode
    Else
        mAddCopy = 0
    End If
End Function

Private Function mCreateAvailSum(ilfirstTime As Integer) As Integer
    'slTitle = "amtRepDBID,amtNetDBID,amtVehicleName,amtNetVehicleCode,amtGenDate,amtGenTime,amtSeqNo,"
    'slTitle = slTitle & "amtDaypartName,amtNetDaypartCode,amtDPInOut,amtDPAvailName,amtDPDay_1,amtDPStartTime_1,amtDPEndTime_1,amtDPSpotPct_1,amtDPDay_2,amtDPStartTime_2,amtDPEndTime_2,amtDPSpotPct_2,amtStartWeekDate,"
    'slTitle = slTitle & "amtAvail_60_Wk01,amtAvail_60_Wk03,amtAvail_60_Wk05,amtAvail_60_Wk07,amtAvail_60_Wk09,amtAvail_60_Wk11,amtAvail_60_Wk13,amtAvail_60_Wk15,amtAvail_60_Wk17,amtAvail_60_Wk19,amtAvail_60_Wk21,amtAvail_60_Wk23,amtAvail_60_Wk25,"
    'slTitle = slTitle & "amtAvail_30_Wk01,amtAvail_30_Wk03,amtAvail_30_Wk05,amtAvail_30_Wk07,amtAvail_30_Wk09,amtAvail_30_Wk11,amtAvail_30_Wk13,amtAvail_30_Wk15,amtAvail_30_Wk17,amtAvail_30_Wk19,amtAvail_30_Wk21,amtAvail_30_Wk23,amtAvail_30_Wk25,"
    'slTitle = slTitle & "amtAvail_10_Wk01,amtAvail_10_Wk03,amtAvail_10_Wk05,amtAvail_10_Wk07,amtAvail_10_Wk09,amtAvail_10_Wk11,amtAvail_10_Wk13,amtAvail_10_Wk15,amtAvail_10_Wk17,amtAvail_10_Wk19,amtAvail_10_Wk21,amtAvail_10_Wk23,amtAvail_10_Wk25,"
    'slTitle = slTitle & "amtCreatedDate,amtCreatedTime"
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim slRepDBID As String
    Dim slVehicleName As String
    Dim ilNetVehicleCode As Integer
    Dim slDayPartName As String
    Dim ilNetDaypartCode As Integer
    Dim slStartWeekDate As String
    Dim slDate As String
    Dim slTime As String
    Dim ilValue As Integer
    Dim ilVef As Integer
    Dim ilRet As Integer
    Dim tlRdf As RDF
    Dim tlAsf As ASF
    Dim slDPInOut As String
    Dim slDPAvailName As String
    
    mCreateAvailSum = False
    ilIndex = mFindFieldName(smHdFieldNames(), "amtVehicleName")
    If ilIndex = -1 Then
        Exit Function
    End If
    slVehicleName = smFields(ilIndex)
    ilIndex = mFindFieldName(smHdFieldNames(), "amtNetVehicleCode")
    If ilIndex = -1 Then
        Exit Function
    End If
    ilNetVehicleCode = Val(smFields(ilIndex))
    
    tgRNAsf.iVefCode = 0
    tgRNMifSrchKey1.sType = "V"
    tgRNMifSrchKey1.sNetDBID = smNetDBID
    tgRNMifSrchKey1.lNetDBCode = ilNetVehicleCode
    ilRet = btrGetEqual(hgRNMif, tgRNMif, igRNMifRecLen, tgRNMifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If (ilRet = BTRV_ERR_NONE) Then
        tgRNAsf.iVefCode = tgRNMif.lRepDBCode
    End If
    If tgRNAsf.iVefCode = 0 Then
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If StrComp(Trim$(tgMVef(ilVef).sName), slVehicleName, vbTextCompare) = 0 Then
                tgRNAsf.iVefCode = tgMVef(ilVef).iCode
                tgRNMif.lCode = 0
                tgRNMif.sNetDBID = smNetDBID
                tgRNMif.sRepDBID = smRepDBID
                tgRNMif.lNetDBCode = ilNetVehicleCode
                tgRNMif.lRepDBCode = tgMVef(ilVef).iCode
                tgRNMif.sType = "V"
                tgRNMif.sUnused = ""
                ilRet = btrInsert(hgRNMif, tgRNMif, igRNMifRecLen, INDEXKEY0)
                Exit For
            End If
        Next ilVef
    End If
    If tgRNAsf.iVefCode = 0 Then
        'Add error message
        'Send message to Rep
        gCreateAlert "A|E|" & smNetDBID & "|" & "|" & "||" & "Vehicle " & slVehicleName & " not found.", True
        'Send message back to Net
        gCreateMessage "A|E|" & smRepDBID & "|" & "|" & "||" & "Vehicle " & slVehicleName & " not found.", True
        Exit Function
    End If


    ilIndex = mFindFieldName(smHdFieldNames(), "amtDaypartName")
    If ilIndex = -1 Then
        Exit Function
    End If
    slDayPartName = smFields(ilIndex)
    ilIndex = mFindFieldName(smHdFieldNames(), "amtNetDaypartCode")
    If ilIndex = -1 Then
        Exit Function
    End If
    ilNetDaypartCode = Val(smFields(ilIndex))
    
    tgRNAsf.iRdfCode = 0
    'tgRNMifSrchKey1.sType = "D"
    'tgRNMifSrchKey1.sNetDBID = smNetDBID
    'tgRNMifSrchKey1.lNetDBCode = ilNetDaypartCode
    'ilRet = btrGetEqual(hgRNMif, tgRNMif, igRNMifRecLen, tgRNMifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    'If (ilRet = BTRV_ERR_NONE) Then
    '    tgRNAsf.iRdfcode = tgRNMif.lRepDBCode
    'End If
    'If tgRNAsf.iRdfcode = 0 Then
        tgRNAsf.iRdfCode = mSearchForDaypart(smHdFieldNames(), "amt", True, slDPInOut, slDPAvailName)
        'If tgRNAsf.iRdfcode > 0 Then
        '    tgRNMif.lCode = 0
        '    tgRNMif.sNetDBID = smNetDBID
        '    tgRNMif.sRepDBID = smRepDBID
        '    tgRNMif.lNetDBCode = ilNetDaypartCode
        '    tgRNMif.lRepDBCode = tgRNAsf.iRdfcode
        '    tgRNMif.sType = "D"
        '    tgRNMif.sUnused = ""
        '    ilRet = btrInsert(hgRNMif, tgRNMif, igRNMifRecLen, INDEXKEY0)
        'End If
    'End If
    If tgRNAsf.iRdfCode = 0 Then
        'Add error message
        'Send message to Rep
        gCreateAlert "A|E|" & smNetDBID & "|" & "|" & "||" & "Daypart " & slDayPartName & " not found.", True
        'Send message back to Net
        gCreateMessage "A|E|" & smRepDBID & "|" & "|" & "||" & "Daypart " & slDayPartName & " not found.", True
        Exit Function
    End If

    
    ilIndex = mFindFieldName(smHdFieldNames(), "amtStartWeekDate")
    If ilIndex = -1 Then
        Exit Function
    End If
    slStartWeekDate = smFields(ilIndex)
    gPackDate slStartWeekDate, tgRNAsf.iStartWkDate(0), tgRNAsf.iStartWkDate(1)


    ilIndex = mFindFieldName(smHdFieldNames(), "amtAvail_60_Wk01")
    If ilIndex = -1 Then
        Exit Function
    End If
    For ilLoop = LBound(tgRNAsf.sAvail60) To UBound(tgRNAsf.sAvail60) Step 2
        ilValue = Val(smFields(ilIndex))
        tgRNAsf.sAvail60(ilLoop) = Chr(ilValue / 256)
        tgRNAsf.sAvail60(ilLoop + 1) = Chr(ilValue And &HFF)
        ilIndex = ilIndex + 1
    Next ilLoop
    ilIndex = mFindFieldName(smHdFieldNames(), "amtAvail_30_Wk01")
    If ilIndex = -1 Then
        Exit Function
    End If
    For ilLoop = LBound(tgRNAsf.sAvail30) To UBound(tgRNAsf.sAvail30) Step 2
        ilValue = Val(smFields(ilIndex))
        tgRNAsf.sAvail30(ilLoop) = Chr(ilValue / 256)
        tgRNAsf.sAvail30(ilLoop + 1) = Chr(ilValue And &HFF)
        ilIndex = ilIndex + 1
    Next ilLoop
    ilIndex = mFindFieldName(smHdFieldNames(), "amtAvail_10_Wk01")
    If ilIndex = -1 Then
        Exit Function
    End If
    For ilLoop = LBound(tgRNAsf.sAvail10) To UBound(tgRNAsf.sAvail10) Step 2
        ilValue = Val(smFields(ilIndex))
        tgRNAsf.sAvail10(ilLoop) = Chr(ilValue / 256)
        tgRNAsf.sAvail10(ilLoop + 1) = Chr(ilValue And &HFF)
        ilIndex = ilIndex + 1
    Next ilLoop
    ilIndex = mFindFieldName(smHdFieldNames(), "amtCreatedDate")
    If ilIndex = -1 Then
        Exit Function
    End If
    slDate = smFields(ilIndex)
    gPackDate slDate, tgRNAsf.iCreatedDate(0), tgRNAsf.iCreatedDate(1)
    ilIndex = mFindFieldName(smHdFieldNames(), "amtCreatedTime")
    If ilIndex = -1 Then
        Exit Function
    End If
    slTime = smFields(ilIndex)
    gPackTime slTime, tgRNAsf.iCreatedTime(0), tgRNAsf.iCreatedTime(1)

    'Remove record by vehicle and add this one
    If ilfirstTime Then
        Do
            tgRNAsfSrchKey0.iCode = tgRNAsf.iVefCode
            ilRet = btrGetEqual(hgRNAsf, tlAsf, igRNAsfRecLen, tgRNAsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hgRNAsf)
            Else
                Exit Do
            End If
        Loop
    End If
    tgRNAsf.lCode = 0
    ilRet = btrInsert(hgRNAsf, tgRNAsf, igRNAsfRecLen, INDEXKEY1)
    If ilRet = BTRV_ERR_NONE Then
        mCreateAvailSum = True
    Else
        mCreateAvailSum = False
    End If
End Function

Private Function mCreateSMF(slSchStatus As String, slMissedDate As String, slMissedTime As String, slSchDate As String, slSchTime As String, ilOrigVefCode As Integer) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    'Create SMF record
    igRNSmfRecLen = Len(tgRNSmf)  'Get and save SMF record length
    tgRNSmf.lCode = 0
    tgRNSmf.lSdfCode = tgRNSdf.lCode
    tgRNSmf.lChfCode = tgRNSdf.lChfCode        'Contract code
    tgRNSmf.iLineNo = tgRNSdf.iLineNo      'Line number
    tgRNSmf.lFsfCode = tgRNSdf.lFsfCode
    gPackDate slMissedDate, tgRNSmf.iMissedDate(0), tgRNSmf.iMissedDate(1)
    gPackTime slMissedTime, tgRNSmf.iMissedTime(0), tgRNSmf.iMissedTime(1)
    gPackDate slSchDate, tgRNSmf.iStartDate(0), tgRNSmf.iStartDate(1)
    gPackDate slSchDate, tgRNSmf.iEndDate(0), tgRNSmf.iEndDate(1)
    gPackDate slSchDate, tgRNSmf.iActualDate(0), tgRNSmf.iActualDate(1)
    gPackTime slSchTime, tgRNSmf.iStartTime(0), tgRNSmf.iStartTime(1)
    gPackTime slSchTime, tgRNSmf.iEndTime(0), tgRNSmf.iEndTime(1)
    gPackTime slSchTime, tgRNSmf.iActualTime(0), tgRNSmf.iActualTime(1)
    For ilLoop = 1 To 7
        tgRNSmf.sWkDays(ilLoop - 1) = "N" 'Weekday flag: Y=day allowed; N=Day disallowed
    Next ilLoop
    tgRNSmf.sWkDays(gWeekDayStr(slSchDate) + 1) = "Y" 'Weekday flag: Y=day allowed; N=Day disallowed
    If tgRNSdf.iVefCode <> ilOrigVefCode Then
        tgRNSmf.sMGSource = "A"     'A=Create by mouse move to a different vehicle
    Else
        tgRNSmf.sMGSource = "M"     'S=Created by MG Spec, M=Created by mouse move outside contract
    End If
    tgRNSmf.sSchStatus = slSchStatus    'G=Makegood; O=Outside contract limits
    tgRNSmf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
    tgRNSmf.iOrigSchVef = ilOrigVefCode
    'Retain copy assignment for as Order billing
    tgRNSmf.sPtType = "0"
    tgRNSmf.lCopyCode = 0
    tgRNSmf.iRotNo = 0
    tgRNSmf.lMtfCode = 0
    tgRNSmf.iGameNo = 0
    ilRet = btrInsert(hgRNSmf, tgRNSmf, igRNSmfRecLen, INDEXKEY0)
    
End Function

Private Function mSearchForDaypart(slFieldNames() As String, slFileName As String, ilFindMatchOnly As Integer, slDPInOut As String, slDPAvailName As String) As Integer
    Dim ilIndex As Integer
    Dim ilCount As Integer
    Dim ilRdf As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilDay As Integer
    Dim llRdfStartTime As Long
    Dim llRdfEndTime As Long
    Dim ilRdfCode As Integer
    Dim ilBFRdfCode As Integer
    Dim ilBFExtraDayCount As Integer
    Dim ilExtraDayCount As Integer
    'Dim slDPInOut As String
    'Dim slDPAvailName As String
    Dim slDPDay_1 As String
    Dim llDPStartTime_1 As Long
    Dim llDPEndTime_1 As Long
    Dim ilDPSpotPct_1 As Integer
    Dim slDPDay_2 As String
    Dim llDPStartTime_2 As Long
    Dim llDPEndTime_2 As Long
    Dim ilDPSpotPct_2 As Integer
    Dim ilInOutMatch As Integer
    
    ilRdfCode = 0
    ilIndex = mFindFieldName(slFieldNames(), slFileName & "DPInOut")
    If ilIndex = -1 Then
        Exit Function
    End If
    slDPInOut = Trim$(smFields(ilIndex))
    ilIndex = mFindFieldName(slFieldNames(), slFileName & "DPAvailName")
    If ilIndex = -1 Then
        Exit Function
    End If
    slDPAvailName = Trim$(smFields(ilIndex))
    ilIndex = mFindFieldName(slFieldNames(), slFileName & "DPDay_1")
    If ilIndex = -1 Then
        Exit Function
    End If
    slDPDay_1 = Trim$(smFields(ilIndex))
    ilIndex = mFindFieldName(slFieldNames(), slFileName & "DPStartTime_1")
    If ilIndex = -1 Then
        Exit Function
    End If
    llDPStartTime_1 = gTimeToLong(Trim$(smFields(ilIndex)), False)
    ilIndex = mFindFieldName(slFieldNames(), slFileName & "DPEndTime_1")
    If ilIndex = -1 Then
        Exit Function
    End If
    llDPEndTime_1 = gTimeToLong(Trim$(smFields(ilIndex)), True)
    ilIndex = mFindFieldName(slFieldNames(), slFileName & "DPSpotPct_1")
    If ilIndex = -1 Then
        Exit Function
    End If
    ilDPSpotPct_1 = Val(smFields(ilIndex))
    ilIndex = mFindFieldName(slFieldNames(), slFileName & "DPDay_2")
    If ilIndex = -1 Then
        Exit Function
    End If
    slDPDay_2 = Trim$(smFields(ilIndex))
    ilIndex = mFindFieldName(slFieldNames(), slFileName & "DPStartTime_2")
    If ilIndex = -1 Then
        Exit Function
    End If
    llDPStartTime_2 = gTimeToLong(Trim$(smFields(ilIndex)), False)
    ilIndex = mFindFieldName(slFieldNames(), slFileName & "DPEndTime_2")
    If ilIndex = -1 Then
        Exit Function
    End If
    llDPEndTime_2 = gTimeToLong(Trim$(smFields(ilIndex)), True)
    ilIndex = mFindFieldName(slFieldNames(), slFileName & "DPSpotPct_2")
    If ilIndex = -1 Then
        Exit Function
    End If
    ilDPSpotPct_2 = Val(smFields(ilIndex))
    
    'Search for matching times and days
    ilRet = btrGetFirst(hgRNRdf, tgRNRdf, igRNRdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If tgRNRdf.sState = "A" Then
            ilInOutMatch = False
            If (slDPInOut = "I") Or (slDPInOut = "O") Then
                If Trim$(tgRNRdf.sInOut) = slDPInOut Then
                    tgRNAnfSrchKey.iCode = tgRNRdf.ianfCode
                    ilRet = btrGetEqual(hgRNAnf, tgRNAnf, igRNAnfRecLen, tgRNAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If StrComp(Trim$(tgRNAnf.sName), slDPAvailName, vbTextCompare) = 0 Then
                        ilInOutMatch = True
                    End If
                End If
            Else
                If (Trim$(tgRNRdf.sInOut) <> "I") And (Trim$(tgRNRdf.sInOut) <> "O") Then
                    ilInOutMatch = True
                End If
            End If
            If ilInOutMatch Then
                ilCount = 0
                For ilRdf = LBound(tgRNRdf.iStartTime, 2) To UBound(tgRNRdf.iStartTime, 2) Step 1
                    If (tgRNRdf.iStartTime(0, ilRdf) <> 1) Or (tgRNRdf.iStartTime(1, ilRdf) <> 0) Then
                        slStr = ""
                        For ilDay = LBound(tgRNRdf.sWkDays, 1) To UBound(tgRNRdf.sWkDays, 1) Step 1
                            slStr = slStr & tgRNRdf.sWkDays(ilRdf, ilDay)
                        Next ilDay
                        gUnpackTimeLong tgRNRdf.iStartTime(0, ilRdf), tgRNRdf.iStartTime(1, ilRdf), False, llRdfStartTime
                        gUnpackTimeLong tgRNRdf.iEndTime(0, ilRdf), tgRNRdf.iEndTime(1, ilRdf), True, llRdfEndTime
                        If ilCount = 0 Then
                            If StrComp(slDPDay_1, slStr, vbTextCompare) = 0 Then
                                If (llDPStartTime_1 = llRdfStartTime) And (llDPEndTime_1 = llRdfEndTime) Then
                                    If slDPDay_2 = "" Then
                                        ilRdfCode = tgRNRdf.iCode
                                        Exit Do
                                    End If
                                    ilCount = 1
                                End If
                            End If
                        Else
                            If StrComp(slDPDay_2, slStr, vbTextCompare) = 0 Then
                                If (llDPStartTime_2 = llRdfStartTime) And (llDPEndTime_2 = llRdfEndTime) Then
                                    ilRdfCode = tgRNRdf.iCode
                                    Exit Do
                                End If
                            End If
                        End If
                    End If
                Next ilRdf
            End If
        End If
        ilRet = btrGetNext(hgRNRdf, tgRNRdf, igRNRdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRdfCode <= 0) And (slDPDay_2 = "") And (Not ilFindMatchOnly) Then
        ilBFRdfCode = -1
        ilBFExtraDayCount = -1
        'Look for matching times and best fit days
        ilRet = btrGetFirst(hgRNRdf, tgRNRdf, igRNRdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While ilRet = BTRV_ERR_NONE
            If tgRNRdf.sState = "A" Then
                ilInOutMatch = False
                If (slDPInOut = "I") Or (slDPInOut = "O") Then
                    If Trim$(tgRNRdf.sInOut) = slDPInOut Then
                        tgRNAnfSrchKey.iCode = tgRNRdf.ianfCode
                        ilRet = btrGetEqual(hgRNAnf, tgRNAnf, igRNAnfRecLen, tgRNAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If StrComp(Trim$(tgRNAnf.sName), slDPAvailName, vbTextCompare) = 0 Then
                            ilInOutMatch = True
                        End If
                    End If
                Else
                    If (Trim$(tgRNRdf.sInOut) <> "I") And (Trim$(tgRNRdf.sInOut) <> "O") Then
                        ilInOutMatch = True
                    End If
                End If
                If ilInOutMatch Then
                    ilCount = 0
                    For ilRdf = LBound(tgRNRdf.iStartTime, 2) To UBound(tgRNRdf.iStartTime, 2) Step 1
                        If (tgRNRdf.iStartTime(0, ilRdf) <> 1) Or (tgRNRdf.iStartTime(1, ilRdf) <> 0) Then
                            slStr = ""
                            For ilDay = LBound(tgRNRdf.sWkDays, 1) To UBound(tgRNRdf.sWkDays, 1) Step 1
                                slStr = slStr & tgRNRdf.sWkDays(ilRdf, ilDay)
                            Next ilDay
                            gUnpackTimeLong tgRNRdf.iStartTime(0, ilRdf), tgRNRdf.iStartTime(1, ilRdf), False, llRdfStartTime
                            gUnpackTimeLong tgRNRdf.iEndTime(0, ilRdf), tgRNRdf.iEndTime(1, ilRdf), True, llRdfEndTime
                            If ilCount = 0 Then
                                If (llDPStartTime_1 = llRdfStartTime) And (llDPEndTime_1 = llRdfEndTime) Then
                                    'Count matching days
                                    ilExtraDayCount = 0
                                    For ilDay = 1 To 7 Step 1
                                        If Mid(slStr, ilDay, 1) = "Y" Then
                                            If Mid(slDPDay_1, ilDay, 1) = "N" Then
                                                ilExtraDayCount = ilExtraDayCount + 1
                                            End If
                                        Else
                                            If Mid(slDPDay_1, ilDay, 1) = "Y" Then
                                                ilExtraDayCount = -1
                                                Exit For
                                            End If
                                        End If
                                    Next ilDay
                                    If ilExtraDayCount = -1 Then
                                        Exit For
                                    End If
                                Else
                                    ilExtraDayCount = -1
                                    Exit For
                                End If
                            Else
                                ilExtraDayCount = -1
                                Exit For
                            End If
                        End If
                    Next ilRdf
                    If ilExtraDayCount > 0 Then
                        If ilBFRdfCode = -1 Then
                            ilBFRdfCode = tgRNRdf.iCode
                            ilBFExtraDayCount = ilExtraDayCount
                        Else
                            If ilExtraDayCount < ilBFExtraDayCount Then
                                ilBFRdfCode = tgRNRdf.iCode
                                ilBFExtraDayCount = ilExtraDayCount
                            End If
                        End If
                    End If
                End If
            End If
            ilRet = btrGetNext(hgRNRdf, tgRNRdf, igRNRdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If ilBFRdfCode <> -1 Then
            ilRdfCode = ilBFRdfCode
        End If
    End If
    mSearchForDaypart = ilRdfCode
End Function
Private Sub mInitSbf(llPrice As Long, slDate As String, slPaidOrBonus As String)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim llDate As Long
    
    llDate = gDateValue(slDate)
    For ilLoop = 0 To UBound(tmSbf) - 1 Step 1
        If (tmSbf(ilLoop).lChfCode = tgRNChf.lCode) And (tmSbf(ilLoop).iLineNo = tmClf.iLine) Then
            If tmSbf(ilLoop).lSpotPrice = llPrice Then
                If tgRNChf.sBillCycle = "C" Then
                    If slPaidOrBonus = "B" Then
                        tmSbf(ilLoop).iBonusNoSpots = tmSbf(ilLoop).iBonusNoSpots + 1
                    Else
                        tmSbf(ilLoop).iAirNoSpots = tmSbf(ilLoop).iAirNoSpots + 1
                    End If
                Else
                    If llDate > lmInvStdEndDate Then
                        If slPaidOrBonus = "B" Then
                            tmSbf(ilLoop).iCalCarryBonus = tmSbf(ilLoop).iCalCarryBonus + 1
                        Else
                            tmSbf(ilLoop).iCalCarryOver = tmSbf(ilLoop).iCalCarryOver + 1
                        End If
                    Else
                        If slPaidOrBonus = "B" Then
                            tmSbf(ilLoop).iBonusNoSpots = tmSbf(ilLoop).iBonusNoSpots + 1
                        Else
                            tmSbf(ilLoop).iAirNoSpots = tmSbf(ilLoop).iAirNoSpots + 1
                        End If
                    End If
                End If
                Exit Sub
            End If
        End If
    Next ilLoop
    ilUpper = UBound(tmSbf)
    tmSbf(ilUpper).lCode = 0
    tmSbf(ilUpper).lChfCode = tgRNChf.lCode
    If tgRNChf.sBillCycle = "C" Then
        gPackDateLong lmInvCalEndDate, tmSbf(ilUpper).iDate(0), tmSbf(ilUpper).iDate(1)
    Else
        gPackDateLong lmInvStdEndDate, tmSbf(ilUpper).iDate(0), tmSbf(ilUpper).iDate(1)
    End If
    tmSbf(ilUpper).sTranType = "T"
    tmSbf(ilUpper).iBillVefCode = tmClf.iVefCode
    tmSbf(ilUpper).iMnfItem = 0
    tmSbf(ilUpper).iNoItems = 0
    tmSbf(ilUpper).lGross = 0
    tmSbf(ilUpper).sUnused1 = ""
    gPackDate "", tmSbf(ilUpper).iPrintInvDate(0), tmSbf(ilUpper).iPrintInvDate(1)
    tmSbf(ilUpper).sDescr = ""
    tmSbf(ilUpper).sAgyComm = "Y"
    tmSbf(ilUpper).sSlsComm = "Y"
    tmSbf(ilUpper).sUnused2 = ""
    tmSbf(ilUpper).sBilled = "N"
    If tgRNChf.iPctTrade = 0 Then
        tmSbf(ilUpper).sCashTrade = "C"
    ElseIf tgRNChf.iPctTrade = 100 Then
        tmSbf(ilUpper).sCashTrade = "T"
    Else
        tmSbf(ilUpper).sCashTrade = "B"
    End If
    tmSbf(ilUpper).iAirVefCode = tgRNSdf.iVefCode   'tmSbf(ilUpper).iBillVefCode
    If tgRNChf.sBillCycle = "C" Then
        If slPaidOrBonus = "B" Then
            tmSbf(ilUpper).iBonusNoSpots = 1
        Else
            tmSbf(ilUpper).iAirNoSpots = 1
        End If
    Else
        If llDate > lmInvStdEndDate Then
            If slPaidOrBonus = "B" Then
                tmSbf(ilUpper).iCalCarryBonus = 1
            Else
                tmSbf(ilUpper).iCalCarryOver = 1
            End If
        Else
            If slPaidOrBonus = "B" Then
                tmSbf(ilUpper).iBonusNoSpots = 1
            Else
                tmSbf(ilUpper).iAirNoSpots = 1
            End If
        End If
    End If
    tmSbf(ilUpper).iBonusNoSpots = 0
    tmSbf(ilUpper).iTrfCode = 0
    tmSbf(ilUpper).sUnused3 = ""
    gPackDate "", tmSbf(ilUpper).iImportDate(0), tmSbf(ilUpper).iImportDate(1)
    gPackDate "", tmSbf(ilUpper).iExportDate(0), tmSbf(ilUpper).iExportDate(1)
    tmSbf(ilUpper).lRefInvNo = 0
    tmSbf(ilUpper).iCombineID = 0
    tmSbf(ilUpper).iUrfCode = tgUrf(0).iCode
    tmSbf(ilUpper).lOGross = 0
    tmSbf(ilUpper).iCommPct = 0
    gPackDate "", tmSbf(ilUpper).iRecDate(0), tmSbf(ilUpper).iRecDate(1)
    gPackDate Format$(gNow(), "m/d/yy"), tmSbf(ilUpper).iPostDate(0), tmSbf(ilUpper).iPostDate(1)
    tmSbf(ilUpper).sPostStatus = "P"
    tmSbf(ilUpper).iCalCarryOver = 0
    tmSbf(ilUpper).lSpotPrice = llPrice
    tmSbf(ilUpper).iMissCarryOver = 0
    tmSbf(ilUpper).sInserted = "N"
    tmSbf(ilUpper).sBarterPaid = "N"
    tmSbf(ilUpper).lAcquisitionCost = 0
    tmSbf(ilUpper).iSpotLen = tmClf.iLen
    tmSbf(ilUpper).iIhfCode = 0
    tmSbf(ilUpper).iLineNo = tmClf.iLine
    ReDim Preserve tmSbf(0 To ilUpper + 1) As SBF
End Sub

Private Function mUpdateSbf() As Integer
    Dim ilRet As Integer
    Dim ilSbf As Integer
    Dim ilDeleted As Integer
    Dim ilPass As Integer   'Pass 0:  Save the billed flag; Pass = 1:  Delete all records for line
    Dim tlSbf As SBF
    
    For ilPass = 0 To 1 Step 1
        For ilSbf = 0 To UBound(tmSbf) - 1 Step 1
            'Find and remove previously created record
            tgRNChfSrchKey1.lCntrNo = tmSbf(ilSbf).lChfCode
            tgRNChfSrchKey1.iCntRevNo = 32000
            tgRNChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hgRNChf, tgRNChf, igRNChfRecLen, tgRNChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tgRNChf.lCntrNo = tmSbf(ilSbf).lChfCode)
                If ((tgRNChf.sSchStatus = "F") Or (tgRNChf.sSchStatus = "M")) And (tgRNChf.sDelete <> "Y") Then
                    Exit Do
                End If
                ilRet = btrGetNext(hgRNChf, tgRNChf, igRNChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            tgRNSbfSrchKey0.lChfCode = tmSbf(ilSbf).lChfCode
            If tgRNChf.sBillCycle = "C" Then
                gPackDateLong lmInvCalEndDate, tgRNSbfSrchKey0.iDate(0), tgRNSbfSrchKey0.iDate(1)
            Else
                gPackDateLong lmInvStdEndDate, tgRNSbfSrchKey0.iDate(0), tgRNSbfSrchKey0.iDate(1)
            End If
            tgRNSbfSrchKey0.sTranType = "T"
            ilRet = btrGetEqual(hgRNSbf, tlSbf, igRNSbfRecLen, tgRNSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)  'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tlSbf.lChfCode = tmSbf(ilSbf).lChfCode)
                ilDeleted = False
                If (tlSbf.iLineNo = tmSbf(ilSbf).iLineNo) And (tlSbf.sTranType = "T") Then
                    If (tlSbf.iDate(0) = tmSbf(ilSbf).iDate(0)) And (tlSbf.iDate(1) = tmSbf(ilSbf).iDate(1)) Then
                        If ilPass = 0 Then
                            If tlSbf.lSpotPrice = tmSbf(ilSbf).lSpotPrice Then
                                tmSbf(ilSbf).sBilled = tlSbf.sBilled
                                Exit Do
                            End If
                        Else
                            ilRet = btrDelete(hgRNSbf)
                            ilDeleted = True
                        End If
                    End If
                End If
                If ilDeleted Then
                    tgRNSbfSrchKey0.lChfCode = tmSbf(ilSbf).lChfCode
                    If tgRNChf.sBillCycle = "C" Then
                        gPackDateLong lmInvCalEndDate, tgRNSbfSrchKey0.iDate(0), tgRNSbfSrchKey0.iDate(1)
                    Else
                        gPackDateLong lmInvStdEndDate, tgRNSbfSrchKey0.iDate(0), tgRNSbfSrchKey0.iDate(1)
                    End If
                    tgRNSbfSrchKey0.sTranType = "T"
                    ilRet = btrGetEqual(hgRNSbf, tlSbf, igRNSbfRecLen, tgRNSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)  'Get first record as starting point of extend operation
                Else
                    ilRet = btrGetNext(hgRNSbf, tlSbf, igRNSbfRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                End If
            Loop
        Next ilSbf
    Next ilPass
    'Save record
    For ilSbf = 0 To UBound(tmSbf) - 1 Step 1
        tmSbf(ilSbf).lGross = tmSbf(ilSbf).lSpotPrice * tmSbf(ilSbf).iAirNoSpots
        ilRet = btrInsert(hgRNSbf, tmSbf(ilSbf), igRNSbfRecLen, INDEXKEY1)
    Next ilSbf

End Function

Private Function mAddQuotes(slStr As String) As String
    mAddQuotes = gAddQuotesRemoveChars(slStr, True)
End Function
