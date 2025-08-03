Attribute VB_Name = "CONTRACTXML"
'DAN NOTE
' If programmatic <> 1  IS my code.
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ContractXML.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the extra XML routine definitions for contracts
Option Explicit
' 8364 to move mExportProposal to this module
Public Type ExportProposalFacts
    hmMnf As Integer
    tmMnf As MNF
    imMnfRecLen As Integer
    tmMnfSrchKey As INTKEY0
    imCppCpm As Integer
    imVefCode As Integer
    imVpfIndex As Integer
    imTerminate As Integer
    imChgMode As Integer
    hmDrf As Integer
    hmDpf As Integer
    hmDef As Integer
    hmRaf As Integer
    hmCHF As Integer
    hmClf As Integer
    tmRegionCode() As SORTCODE
    imStep1Index As Integer
    lmCPriDemoGRP As Long
    lmCPriDemoGrImp As Long
    lmCPriDemoCPP As Long
    lmCPriDemoCPM As Long
    imCPriDemoRating As Integer
    smAgencyXMLPath As String
    smStationXMLWOPath As String
    smStationXMLBasePath As String
    imLB1Or2 As Integer
    imDiffMode As Integer
    smLnSave() As String
End Type
Public tgExportFacts As ExportProposalFacts
Public tgExportAgf As AGF
Dim lmCPriDemoGRP As Long       'Primary Demo contract total GRP, set in mSetPriDemoResearch
Dim lmCPriDemoGrImp As Long     'Primary Demo contract total GrImp, set in mSetPriDemoResearch
Dim lmCPriDemoCPP As Long       'Primary Demo contract total CPP, set in mSetPriDemoResearch
Dim lmCPriDemoCPM As Long       'Primary Demo contract total CPM, set in mSetPriDemoResearch
Dim imCPriDemoRating As Integer 'Primary Demo contract total Avg rating, set in mSetPriDemoResearch

'Dan M 4/27/2011 proposal xml
Public bmXmlOutput As Boolean
Public smClfInfo() As String
Public smClfDates As String
Dim smFields(55) As String
Dim smValues(55) As String
'for array immediately below:
'DEMOMAX is not count, but array ubound
Private Const XMLDEMOMAX As Integer = 3 '4
Private Const XMLDEMOGROUP As Integer = 0
Private Const XMLDEMOFROM As Integer = 1
Private Const XMLDEMOTO As Integer = 2
'not using?
'Private Const XMLDEMORATING As Integer = 3
'Private Const XMLDEMOAVGAUD As Integer = 4
Private Const XMLIMPRESSIONS As Integer = 5
' Dan 1/28/13 5908 proposal xml
'store each demo info: name,age from, age to
Dim smDemoHeaders(XMLDEMOMAX, XMLDEMOTO) As String
Dim rsDemoValues As ADODB.Recordset
'Private Type xmlDemos
'    iIndex As Integer
'    slRating As String
'    slAvgAud As String
'End Type
''values to write out in periods for each demo
'Dim smDemoValues() As xmlDemos
'ttp 6143
Private smDemoComment As String
'8915
Dim myRst As ADODB.Recordset
Private Const XMLUNKNOWN As String = "unknown"
'8032-moved to help compile
'Public igStationXmlChoice As Integer
'parse xml line
Const XMLCLFLINE As Integer = 0
'these are also used with smfields/smvalues
Const XMLMON As Integer = 1
Const XMLTUE As Integer = 2
Const XMLWED As Integer = 3
Const XMLTHU As Integer = 4
Const XMLFRI As Integer = 5
Const XMLSAT As Integer = 6
Const XMLSUN As Integer = 7
'end smfields/smvalues
Const XMLDAILY As Integer = 8
Const XMLVEHICLE As Integer = 9
Const XMLTIMES As Integer = 10
Const XMLSPOTLENGTH As Integer = 11
Const XMLRATE As Integer = 12
Const XMLAVGAUD As Integer = 13
Const XMLRATING As Integer = 15
Const XMLGRP As Integer = 16
Const XMLCPP As Integer = 17
Const XMLCPM As Integer = 18
Const XMLGRIMP As Integer = 19
Const XMLPOPULATION As Integer = 20
Const XMLPACKAGE As Integer = 21 '20
Const XMLRATINGSOURCE As Integer = 22 '21
Const XMLTOTALSPOTS As Integer = 23 '22
'tag names
Private Const PROPOSALID As Integer = 45 '0
Private Const PVERSION As Integer = 46 '1
Private Const PSENDDATETIME As Integer = 47 '2
Private Const PWEEKSTARTDAY As Integer = 48 '3
Private Const PSTARTDATE As Integer = 49 '4
Private Const PENDDATE As Integer = 50 '5
Private Const SCOMPANYNAME As Integer = 51 '6
Private Const SOFFICENAME As Integer = 52 '7
Private Const SALESNAME As Integer = 8
Private Const BCOMPANYNAME As Integer = 9
Private Const BOFFICENAME As Integer = 10
Private Const BNAME As Integer = 11
Private Const ADVNAME As Integer = 12
Private Const ProductName As Integer = 13
Private Const PROPOSALNAME As Integer = 14
Private Const SELLERREFERENCE As Integer = 15
Private Const RADIOCALLLETTERS As Integer = 16
Private Const RADIOBAND As Integer = 17
Private Const RADIOOUTLETID As Integer = 18
Private Const AVAILID As Integer = 19
Private Const AVAILSTARTDATE As Integer = 20
Private Const AVAILENDDATE As Integer = 21
Private Const AVAILISPACKAGE As Integer = 22
Private Const AVAILNAME As Integer = 23
Private Const OUTLETREF As Integer = 24
Private Const OUTLETID As Integer = 25
Private Const COMMENTPROPOSAL As Integer = 26
Private Const DEMOID As Integer = 27
Private Const DEMOREF As Integer = 28
Private Const DEMOTYPE As Integer = 29
Private Const DEMOGROUP As Integer = 30
Private Const DEMOAGEFROM As Integer = 31
Private Const DEMOAGETO As Integer = 32
Private Const DAYTIMESTARTTIME As Integer = 33
Private Const DAYTIMEENDTIME As Integer = 34
Private Const DAYNAME As Integer = 35
Private Const DAYAVAILNAME As Integer = 36
Private Const DAYSPOTLENGTH As Integer = 37
Private Const COMMENTLINE As Integer = 38
Private Const PERIODSTARTDATE As Integer = 39
Private Const PERIODENDDATE As Integer = 40
Private Const PERIODRATE As Integer = 41
Private Const PERIODSPOTSPERWEEK As Integer = 42
Private Const PERIODDEMOVALUE As Integer = 43
Private Const PERIODDEMOREF As Integer = 44
Private Const AVAILOUTLET As Integer = 0 '46
Private Const XMLDATE As String = "yyyy-mm-dd"
Private Const XMLTIME As String = "hh:mm"
Private Const XMLSPOTTIME As String = "hh:mm:ss"
'ttp 5230
Private Const PERIODDEMOAUD As Integer = 53
Private Const DEMOAVGAUDID As Integer = 54
'8032
Public ogEOrder As CEOExport
Public ogStationXmlErrors As CLogger
'which station xml turned on by client? set this at minit and use when setting each order
Public bgStationXmlSiteAllowed(0 To 1) As Boolean
'order using which station xml
Public bgStationXmlOrderAllowed(0 To 1) As Boolean
'client wants this SERVICE to go to web, or just generate file?
Public bgStationXmlWebServiceAllowed(0 To 1) As Boolean
Public Enum StationXmlIndexes
    WO = 0
    MARKETRONEO = 1
End Enum
Public Const STATIONXMLMARKETRONNAME As String = "Marketron-Electronic Orders"
Public Const STATIONXMLWIDEORBITNAME As String = "WideOrbit-Electronic Contracts"
'7/11/16 moved
''8032 be careful!  These are the indexes to radio buttons on vehOpt!
'Public Const STATIONXMLWIDEORBIT As Integer = 5
'Public Const STATIONXMLMARKETRON As Integer = 7
'Public Const STATIONXMLNONE As Integer = 6
'Public Const AGENCYXML As Integer = 0

' XML Web Service Functions-- Dan proposal xml
Type CSIRspGetXMLStatus
    sStatus As String * 2048
End Type

'6807 6865
'Declare Function csiXMLStart Lib "csi_utils.dll" (ByVal slIni$, ByVal slSection$, ByVal slType$, ByVal slFileName$, ByVal slLineEndChar$) As Integer
Declare Function csiXMLStart Lib "csi_utils.dll" (ByVal slIni$, ByVal slSection$, ByVal slType$, ByVal slFileName$, ByVal slLineEndChar$, ByVal slXmlErrorResponseFile$) As Integer
Declare Function csiXMLSetMethod Lib "csi_utils.dll" (ByVal slMethodName$, ByVal slTableName$, ByVal slTransmissionID$, ByVal slSchemaName$) As Integer
Declare Function csiXMLData Lib "csi_utils.dll" (ByVal slType$, ByVal slTag$, ByVal slvalue$) As Integer
Declare Function csiXMLDataEx Lib "csi_utils.dll" (ByVal slType$) As Integer
Declare Function csiXMLWrite Lib "csi_utils.dll" (ByVal FlushQueue As Long) As Integer
Declare Function csiXMLEnd Lib "csi_utils.dll" () As Integer
Declare Function csiXMLStatus Lib "csi_utils.dll" (ByRef sStatus As CSIRspGetXMLStatus) As Integer

Public Function mTestForProposalXml(slExportTypeFromEmail As String) As Boolean
    'If tmAgf.sCntrExptForm = "X" Then
    If slExportTypeFromEmail = "X" Then
        mTestForProposalXml = True
    Else
        mTestForProposalXml = False
    End If
End Function

Public Function mXmlStationInfo(slAgencyBand As String, slRepVehicle As String) As Boolean
    
On Error GoTo ERRBOX
    '8032
    If igStationXmlChoice = AGENCYXML Then
        If Trim$(slAgencyBand) <> "" Then
            smValues(RADIOBAND) = Trim$(slAgencyBand)
        Else
            '8287 added DV SM N
            'Dan 11/16/11 most of our users will have 'X'...but this isn't allowed in schema.  'Dummy' 'Am' instead
            If Trim(tgSpf.sEDIBand) = "FM" Or Trim(tgSpf.sEDIBand) = "DV" Or Trim(tgSpf.sEDIBand) = "SM" Or Trim(tgSpf.sEDIBand) = "N" Then
                smValues(RADIOBAND) = tgSpf.sEDIBand
            Else
                smValues(RADIOBAND) = "AM"
            End If
        End If
        smValues(RADIOCALLLETTERS) = tgSpf.sEDICallLetter
        If Trim$(tgExportAgf.sXMLCallLetters) <> "" Then
            smValues(RADIOCALLLETTERS) = Trim$(tgExportAgf.sXMLCallLetters)
        End If
        '8032 moved to validation
    ''    If (Len(smValues(RADIOCALLLETTERS)) = 4 Or Len(smValues(RADIOCALLLETTERS)) = 3) And (smValues(RADIOBAND) = "FM" Or smValues(RADIOBAND) = "AM") Then
    '    If (Len(smValues(RADIOCALLLETTERS)) = 4 Or Len(smValues(RADIOCALLLETTERS)) = 3) And (Len(Trim(smValues(RADIOBAND))) > 0) Then
    '        mXmlStationInfo = True
    '    Else
    '        mXmlStationInfo = False
    '    End If
    Else
        '8060 adjust outlet to the vehicle
        gAdjustStationForAgency slRepVehicle
    End If
    mXmlStationInfo = True
    Exit Function
ERRBOX:
    mXmlStationInfo = False
End Function
Public Sub gAdjustStationForAgency(slVehicle As String)
    Dim ilPos As Integer
    Dim slCall As String
    Dim slBand As String
    
    slCall = ""
    slBand = ""
    ilPos = InStr(slVehicle, "-")
    If ilPos > 0 Then
        slCall = Mid(slVehicle, 1, ilPos - 1)
        slBand = Mid(slVehicle, ilPos + 1)
    End If
    smValues(RADIOBAND) = slBand
    smValues(RADIOCALLLETTERS) = slCall
End Sub
Public Function mXmlSetBasicAndValidate(slVehicle As String, llMoStartDate As Long, llEndDate As Long, slDelimiter As String, slErrorMessage As String, tlagf As AGF, ilStepIndex As Integer, blRan As Boolean) As Boolean
    Dim blRet As Boolean
    Dim llWeek As Long
    Dim c As Integer
    
    slErrorMessage = ""
    blRet = True
    If Not blRan Then
        '8032 from mXmlSetFields
        'first time, everyone needs this
        For c = 0 To UBound(smValues)
            smValues(c) = ""
        Next c
    End If
    '8063 move this because we need to do this each time for stationxml
    Select Case igStationXmlChoice
        '0 is agency xml
        Case STATIONXMLWIDEORBIT, AGENCYXML
            mXmlSetFields
        Case STATIONXMLMARKETRON
            mXmlSetFieldsAgencyMarketron
        Case STATIONXMLNONE
    End Select
    '7673 circumvent doing this more than once
'    If blRan Then
'        mXmlSetBasicAndValidate = blRet
'        Exit Function
'    End If
    'mXmlSetFields
    '8032
    If Not blRan Then
        blRan = True
         For llWeek = llMoStartDate To llEndDate Step 7
             smClfDates = smClfDates & Format(llWeek, XMLDATE) & slDelimiter
         Next llWeek
         If Not mXmlProposalInfo(ilStepIndex) Then
             slErrorMessage = "There was a problem creating the xml file: mXmlProposalInfo"
             blRet = False
         ElseIf Not mXmlSellerBuyerInfo(tlagf) Then
             slErrorMessage = "There was a problem creating the xml file: mXmlSellerBuyerInfo"
             blRet = False
         ElseIf Not mXmlStationInfo(tlagf.sXMPProposalBand, slVehicle) Then
             'slErrorMessage = "Cannot export to xml unless Edi Call Letters and Edi Band are set in site options; Call letters can only be 3 or 4 letters long and band cannot be blank."
             slErrorMessage = "There was a problem creating the xml file:mXmlStationInfo"
             blRet = False
         'note: this sub uses values that are set above!
         ElseIf Not mXmlAvailListInfo(llMoStartDate, llEndDate) Then
            slErrorMessage = "There was a problem creating the xml file: mXmlAvailListInfo"
            blRet = False
         '5908 move xXmlDemoInfo before gather demo info.
        ' Else
        '    mXmlDemoInfo
         End If
         If blRet Then
            blRet = mXmlValidateBasic(slErrorMessage)
        End If
    'need to run this for each vehicle in stationXml. Agency only ever comes in here once, so don't worry about
    Else
        'this is in mxmlstationinfo
        gAdjustStationForAgency (slVehicle)
        blRet = mXmlValidateBasic(slErrorMessage)
    End If
    mXmlSetBasicAndValidate = blRet
End Function
Private Function mXmlValidateBasic(slErrorMessage As String) As Boolean

'Private Const PROPOSALID As Integer = 45 '0        R   CONSTANT
'Private Const PVERSION As Integer = 46 '1          R
'Private Const PSENDDATETIME As Integer = 47 '2     R
'Private Const PWEEKSTARTDAY As Integer = 48 '3     R   CONSTANT
'Private Const PSTARTDATE As Integer = 49 '4        R
'Private Const PENDDATE As Integer = 50 '5          R
'Seller not required, but if it's there:
'Private Const SCOMPANYNAME As Integer = 51 '6      R   don't test
'Private Const SOFFICENAME As Integer = 52 '7           don't test
'Private Const SALESNAME As Integer = 8             R   don't test
'Buyer not required but if it's there:
'Private Const BCOMPANYNAME As Integer  = 9         R   don't test
'Private Const BOFFICENAME As Integer = 10              don't test
'Private Const BNAME As Integer = 11                    don't test
'Advertiser not required, but if it's there;
'Private Const ADVNAME As Integer = 12              R ? don't test
'Private Const ProductName As Integer = 13              don't test
'Private Const PROPOSALNAME As Integer = 14         R   don't test
'Private Const SELLERREFERENCE As Integer = 15          don't test
'Private Const RADIOCALLLETTERS As Integer = 16     R 3 to 4 letters
'Private Const RADIOBAND As Integer = 17            R Am FM
'Private Const RADIOOUTLETID As Integer = 18        R   CONSTANT
'Private Const AVAILID As Integer = 19              R   CONSTANT
'Private Const AVAILSTARTDATE As Integer = 20       R
'Private Const AVAILENDDATE As Integer = 21         R
'Private Const AVAILISPACKAGE As Integer = 22       R  Y N CONSTANT
'Private Const AVAILNAME As Integer = 23            ?
'Private Const OUTLETREF As Integer = 24            R   CONSTANT
'Private Const OUTLETID As Integer = 25             R   CONSTANT
    Dim ilCount As Integer
    Dim blRet As Boolean
    'to do test radiocallletters and radioband
    'to do find what failed and write it out
    
    blRet = True
    'test all
    If Len(smValues(PVERSION)) = 0 Then
        blRet = False
        slErrorMessage = slErrorMessage & " Order Version cannot be blank."
    End If
    If Len(smValues(PSTARTDATE)) = 0 Or Len(smValues(PENDDATE)) = 0 Then
        blRet = False
        slErrorMessage = slErrorMessage & " Order start and end dates cannot be blank."
    End If
    Select Case igStationXmlChoice
        Case STATIONXMLWIDEORBIT, AGENCYXML
            For ilCount = AVAILSTARTDATE To AVAILENDDATE
                If Len(smValues(ilCount)) = 0 Then
                    blRet = False
                    slErrorMessage = slErrorMessage & " Avail dates cannot be blank."
                End If
            Next ilCount
            If Len(smValues(AVAILISPACKAGE)) = 0 Then
                blRet = False
                slErrorMessage = slErrorMessage & " Avail 'package' cannot be blank."
            End If
            If Len(smValues(AVAILNAME)) = 0 Then
                blRet = False
                slErrorMessage = slErrorMessage & " Avail name cannot be blank."
            End If
'            For ilCount = PVERSION To PENDDATE
'                If Len(smValues(ilCount)) = 0 Then
'                    blRet = False
'                End If
'            Next ilCount
            If Len(smValues(PSENDDATETIME)) = 0 Then
                blRet = False
                slErrorMessage = slErrorMessage & " SendDateTime cannot be blank."
            End If
            'call letters
            If Len(smValues(RADIOCALLLETTERS)) > 4 Or Len(smValues(RADIOCALLLETTERS)) < 3 Then
                blRet = False
                slErrorMessage = " Outlet's 'call letters' must be 3 or 4 characters. " & slErrorMessage
            End If
''            'band
            If Not (InStr(smValues(RADIOBAND), "AM") > 0 Or InStr(smValues(RADIOBAND), "FM") > 0 Or InStr(smValues(RADIOBAND), "DV") > 0 Or InStr(smValues(RADIOBAND), "SM") > 0 Or InStr(smValues(RADIOBAND), "N") > 0) Then
                blRet = False
                slErrorMessage = slErrorMessage & " Outlet's 'band' not an allowed value."
            End If
        Case STATIONXMLMARKETRON
            'buy co.
            If Len(smValues(BCOMPANYNAME)) = 0 Then
                blRet = False
                slErrorMessage = slErrorMessage & " Buying company name may not be blank."
            End If
            'buy co office
            If Len(smValues(BOFFICENAME)) = 0 Then
                blRet = False
                slErrorMessage = slErrorMessage & " Buying company's office may not be blank."
            End If
            'sales co
            If Len(Trim$(smValues(SCOMPANYNAME))) = 0 Then
                blRet = False
                slErrorMessage = slErrorMessage & " Selling company name may not be blank."
            End If
            'sales co office
            If Len(Trim$(smValues(SOFFICENAME))) = 0 Then
                blRet = False
                slErrorMessage = slErrorMessage & " Selling company's office may not be blank."
            End If
            'call letters
            If Len(smValues(RADIOCALLLETTERS)) > 9 Or Len(smValues(RADIOCALLLETTERS)) < 3 Then
                blRet = False
                slErrorMessage = slErrorMessage & " Vehicle's 'call letters' must be between 3 and 9 characters."
            End If
'            'band
            If Not (InStr(smValues(RADIOBAND), "AM") > 0 Or InStr(smValues(RADIOBAND), "FM") > 0 Or InStr(smValues(RADIOBAND), "HD1") > 0 Or InStr(smValues(RADIOBAND), "HD2") > 0 Or InStr(smValues(RADIOBAND), "HD3") > 0 Or InStr(smValues(RADIOBAND), "Stream") > 0) Then
                blRet = False
                slErrorMessage = slErrorMessage & " Vehicle's 'band' not an allowed value."
            End If
    End Select
    mXmlValidateBasic = blRet
End Function
Private Function mXmlAvailListInfo(llMoStartDate As Long, llEndDate As Long) As Boolean
    Dim llWeek As Long
    
On Error GoTo ERRORBOX
    smValues(AVAILNAME) = smValues(SELLERREFERENCE) & "-" & smValues(ADVNAME) & "-" & smValues(ProductName)
    smValues(AVAILSTARTDATE) = Format(llMoStartDate, XMLDATE)
    smValues(AVAILENDDATE) = Format(llEndDate, XMLDATE)
    mXmlAvailListInfo = True
    Exit Function
ERRORBOX:
     mXmlAvailListInfo = False
End Function
Private Function mXmlProposalInfo(ilStepIndex As Integer) As Boolean
On Error GoTo ERRORBOX
    Dim ilVersionOrRevision As Integer
    Dim ilClf As Integer
    Dim ilClfIndex As Integer
    Dim llCff As Long
    Dim llEarliestStartDate As Long
    Dim llLatestEndDate As Long
    Dim ilDay As Integer
    Dim llDate As Long
    
    smValues(PROPOSALID) = tgChfCntr.lCntrNo
    smValues(PSENDDATETIME) = Format(Now, XMLDATE) & "T" & Format(Now, XMLSPOTTIME)
    '9/10/18
    If tgExportAgf.sXMLDates <> "A" Then
        smValues(PSTARTDATE) = mConvertDate(tgChfCntr.iStartDate)
        smValues(PENDDATE) = mConvertDate(tgChfCntr.iEndDate)
    Else
        llEarliestStartDate = -1
        llLatestEndDate = -1
        For ilClf = 0 To UBound(smClfInfo) - 1
            ilClfIndex = mGetClfLineNumber(smClfInfo(ilClf))
            llCff = tgClfCntr(ilClfIndex).lFirstCff
            Do While llCff <> -1
                If tgCffCntr(llCff).CffRec.iSpotsWk > 0 Or tgCffCntr(llCff).CffRec.sDyWk = "D" Then
                    If (llEarliestStartDate = -1) Or (tgCffCntr(llCff).lStartDate <= llEarliestStartDate) Then
                        llDate = tgCffCntr(llCff).lStartDate
                        For ilDay = gWeekDayLong(tgCffCntr(llCff).lStartDate) To 6 Step 1
                            If tgCffCntr(llCff).CffRec.iDay(ilDay) > 0 Then
                                Exit For
                            End If
                            llDate = llDate + 1
                        Next ilDay
                        If (llEarliestStartDate = -1) Then
                            llEarliestStartDate = llDate
                        ElseIf llDate < llEarliestStartDate Then
                            llEarliestStartDate = llDate
                        End If
                        
                    End If
                    If (llLatestEndDate = -1) Or (tgCffCntr(llCff).lEndDate >= llLatestEndDate) Then
                        llDate = tgCffCntr(llCff).lEndDate
                        For ilDay = gWeekDayLong(tgCffCntr(llCff).lEndDate) To 0 Step -1
                            If tgCffCntr(llCff).CffRec.iDay(ilDay) > 0 Then
                                Exit For
                            End If
                            llDate = llDate - 1
                        Next ilDay
                        If (llLatestEndDate = -1) Then
                            llLatestEndDate = llDate
                        ElseIf llDate > llLatestEndDate Then
                            llLatestEndDate = llDate
                        End If
                    End If
                End If
                llCff = tgCffCntr(llCff).lNextCff
            Loop
        Next ilClf
        smValues(PSTARTDATE) = Format(llEarliestStartDate, XMLDATE)
        smValues(PENDDATE) = Format(llLatestEndDate, XMLDATE)
    End If
    smValues(ProductName) = tgChfCntr.sProduct
    '7580
    'smValues(PROPOSALNAME) = Trim$(str$(tgChfCntr.lCntrNo)) & "-" & Trim$(str$(tgChfCntr.iPropVer)) & " " & Trim$(tgSpf.sGClient)
    smValues(PROPOSALNAME) = mXmlContractAndVersionOrRevision(ilStepIndex, ilVersionOrRevision) & " " & Trim$(tgSpf.sGClient)
    smValues(PVERSION) = ilVersionOrRevision
    mXmlProposalInfo = True
    Exit Function
ERRORBOX:
    mXmlProposalInfo = False
End Function
Public Function mXmlContractAndVersionOrRevision(ilStepIndex As Integer, Optional ilVersionOrRevisionNumber As Integer = 0) As String
    'returns "XX-Y#"  where XX is the contract#, Y is status P,O,H and # is version(if status not O or H) or revision + 1.
    'O- ilVersionOrRevisionNumber. For 'Version' on xml.  Because this CANNOT be 0, it is possible this number will not match the # in the string.
    
    Dim slType As String
    Dim ilNumber As Integer
    Dim blRet As String
    
    blRet = ""
    slType = "P"
    If tgChfCntr.sStatus = "H" Or tgChfCntr.sStatus = "G" Then
        ilNumber = tgChfCntr.iExtRevNo
        slType = "H"
    'Dan 7580 they changed to 'revise hold or order' so check cntrrevno or ilstepindex(tgChfCntr.sStatus = "W" And tgChfCntr.iCntRevNo > 0) works after save
    ElseIf tgChfCntr.sStatus = "O" Or tgChfCntr.sStatus = "N" Or (tgChfCntr.sStatus = "W" And (tgChfCntr.iCntRevNo > 0 Or ilStepIndex = 4 Or ilStepIndex = 5)) Then
        ilNumber = tgChfCntr.iExtRevNo
        slType = "O"
    Else
        ilNumber = tgChfCntr.iPropVer
    End If
    blRet = Trim$(str$(tgChfCntr.lCntrNo)) & "-" & slType & ilNumber
    If ilNumber = 0 Then
        ilNumber = ilNumber + 1
    End If
    ilVersionOrRevisionNumber = ilNumber
    mXmlContractAndVersionOrRevision = blRet
End Function
Public Function mXmlDemoInfo(hlDrf As Integer, hlMnf As Integer, hlDpf As Integer, hlDef As Integer, hlRaf As Integer, tlRegionCode() As SORTCODE, imStep1Index As Integer) As Integer
    ' -1 bad custom 0 no demo 1 ok
    '6339
    Dim ilRet As Integer
    Dim c As Integer
    Dim J As Integer
    
    ilRet = 0
    smDemoComment = ""
    '6339
    ilRet = mXmlDemoHeader()
    If ilRet = 1 Then
        mXmlSetDemoValues hlDrf, hlMnf, hlDpf, hlDef, hlRaf, tlRegionCode(), imStep1Index
    End If
    mXmlDemoInfo = ilRet
End Function
Private Function mXmlDemoHeader() As Integer
    ' -1 bad custom 0 no demo 1 ok
    '5908
    Dim c As Integer
    Dim J As Integer
    '6339
    Dim blIsSafeDemo As Boolean
    Dim blIsDemo As Boolean
    Dim ilRet As Integer
    
    ilRet = 0
    blIsDemo = False
    blIsSafeDemo = False
    For c = 0 To XMLDEMOMAX
        smDemoHeaders(c, XMLDEMOGROUP) = ""
        For J = 1 To 2
            smDemoHeaders(c, J) = 0
        Next J
#If programmatic <> 1 Then
        'need to get count of how many demos there actually are
        If Contract.lbcDemo.Count > c Then
            If Contract.lbcDemo(c).ListIndex >= 1 Then
                blIsDemo = True
                'stuffing into smvalues doesn't make as much sense now.  Should just pass back locally
                smValues(DEMOGROUP) = ""
                smValues(DEMOAGEFROM) = ""
                smValues(DEMOAGETO) = ""
                If mParseDemoName(Contract.lbcDemo(c).List(Contract.lbcDemo(c).ListIndex), c + 1) Then
                    blIsSafeDemo = True
                    smDemoHeaders(c, XMLDEMOGROUP) = smValues(DEMOGROUP)
                    smDemoHeaders(c, XMLDEMOFROM) = smValues(DEMOAGEFROM)
                    smDemoHeaders(c, XMLDEMOTO) = smValues(DEMOAGETO)
                End If
            End If
        End If
#Else
    Dim ilDemoCount As Integer
    
      '8915
    For J = 0 To UBound(tgChfCntr.iMnfDemo)
        If tgChfCntr.iMnfDemo(J) > 0 Then
            ilDemoCount = ilDemoCount + 1
        End If
    Next J
        If ilDemoCount >= c Then
            If tgChfCntr.iMnfDemo(c) > 0 Then
                blIsDemo = True
                'stuffing into smvalues doesn't make as much sense now.  Should just pass back locally
                smValues(DEMOGROUP) = ""
                smValues(DEMOAGEFROM) = ""
                smValues(DEMOAGETO) = ""
                '8915
                Set myRst = gSQLSelectCall("Select mnfName from MNF_Multi_Names where mnfCode = " & tgChfCntr.iMnfDemo(c))
                If Not myRst.EOF Then
                    If mParseDemoName(Trim(myRst!mnfName), c + 1) Then
                        blIsSafeDemo = True
                        smDemoHeaders(c, XMLDEMOGROUP) = smValues(DEMOGROUP)
                        smDemoHeaders(c, XMLDEMOFROM) = smValues(DEMOAGEFROM)
                        smDemoHeaders(c, XMLDEMOTO) = smValues(DEMOAGETO)
                    End If
                End If
            End If
        End If
#End If
    Next c
    If blIsDemo And blIsSafeDemo Then
        ilRet = 1
    ElseIf Not blIsDemo Then
        ilRet = 0
    Else
        ilRet = -1
    End If
    mXmlDemoHeader = ilRet
End Function
Private Sub mXmlSetDemoValues(hlDrf As Integer, hlMnf As Integer, hlDpf As Integer, hlDef As Integer, hlRaf As Integer, tlRegionCode() As SORTCODE, imStep1Index As Integer)

    Set rsDemoValues = mPrepRecordset()
    Dim c As Integer
    Dim ilAvgAud As Integer
    Dim ilRating As Integer
    
    '5890
    'now do primary demo also, but only do if NOT custom
   ' For c = 0 To XMLDEMOMAX
    '8459
    For c = XMLDEMOMAX To 0 Step -1
        If Len(smDemoHeaders(c, XMLDEMOGROUP)) > 0 And Len(smDemoHeaders(c, XMLDEMOFROM)) > 0 And Len(smDemoHeaders(c, XMLDEMOTO)) > 0 Then
            'D.S. 5/30/18 needs fix
            #If programmatic <> 1 Then
                mSetPriDemoResearch hlDrf, hlMnf, hlDpf, hlDef, hlRaf, tlRegionCode(), imStep1Index, lmCPriDemoGRP, lmCPriDemoGrImp, lmCPriDemoCPP, lmCPriDemoCPM, imCPriDemoRating, c, rsDemoValues
            #Else
                mSetXMLResearchNo
            #End If
        End If
    Next c

    'If Len(smValues(DEMOGROUP)) > 0 And Len(smValues(DEMOAGEFROM)) > 0 And Len(smValues(DEMOAGETO)) > 0 Then
    'Dan +1?  not doing primary demo!
'    For c = lbcDemo.LBound + 1 To lbcDemo.UBound '- 1
'        mSetPriDemoResearch hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imStep1Index, lmCPriDemoGRP, lmCPriDemoGrImp, lmCPriDemoCPP, lmCPriDemoCPM, imCPriDemoRating, c, rsDemoValues
'    Next c
End Sub
Private Function mPrepRecordset() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    Set myRs = New ADODB.Recordset
        '5890 fixed Audience
        With myRs.Fields
            .Append "Rating", adChar, 10
            .Append "Audience", adInteger
            .Append "DemoIndex", adInteger
            .Append "LineIndex", adInteger
        End With
    myRs.Open
    Set mPrepRecordset = myRs

End Function

Public Function mWriteXmlDemoHeader() As Boolean
'5980
    Dim c As Integer
    Dim ilFirstValid As Integer
    
    mWriteXmlDemoHeader = False
    ilFirstValid = -1
    '7209
    If tgChfCntr.sHideDemo <> "Y" Then
'        mCSIXMLData "OT", "DemoCategories", ""
        For c = 0 To XMLDEMOMAX
            ' don't really need to stick into smValues, but code already written that way.
            smValues(DEMOGROUP) = smDemoHeaders(c, XMLDEMOGROUP)
            smValues(DEMOAGEFROM) = smDemoHeaders(c, XMLDEMOFROM)
            smValues(DEMOAGETO) = smDemoHeaders(c, XMLDEMOTO)
            If Len(smValues(DEMOGROUP)) > 0 And Len(smValues(DEMOAGEFROM)) > 0 And Len(smValues(DEMOAGETO)) > 0 Then
                If ilFirstValid = -1 Then
                    ilFirstValid = c
                    'Dan moved here 10/22/14..so won't write if empty  7214
                    mCSIXMLData "OT", "DemoCategories", ""
               End If
                mCSIXMLData "OT", "DemoCategory", smFields(DEMOID) & "='DM" & CStr(c) & "'" ' mXmlAttribute(DEMOID)
                mCSIXMLData "CD", "tvb:DemoType", "Rating"
                mCSIXMLData "CD", smFields(DEMOGROUP), smValues(DEMOGROUP)
                mCSIXMLData "CD", smFields(DEMOAGEFROM), smValues(DEMOAGEFROM)
                mCSIXMLData "CD", smFields(DEMOAGETO), smValues(DEMOAGETO)
                mCSIXMLData "CT", "DemoCategory", ""
                'ttp5230
                mCSIXMLData "OT", "DemoCategory", smFields(DEMOID) & "='DM" & CStr(c + XMLIMPRESSIONS) & "'" 'mXmlAttribute(DEMOAVGAUDID)
                mCSIXMLData "CD", "tvb:DemoType", "Impression"
                mCSIXMLData "CD", smFields(DEMOGROUP), smValues(DEMOGROUP)
                mCSIXMLData "CD", smFields(DEMOAGEFROM), smValues(DEMOAGEFROM)
                mCSIXMLData "CD", smFields(DEMOAGETO), smValues(DEMOAGETO)
                mCSIXMLData "CT", "DemoCategory", ""
                mWriteXmlDemoHeader = True
            End If
        Next c
'        mCSIXMLData "CT", "DemoCategories", ""
    End If
    If ilFirstValid <> -1 Then
        'dan moved here 10/22/14 so won't write if empty 7214
        mCSIXMLData "CT", "DemoCategories", ""
        mCSIXMLData "CA", "TargetDemo", "demoRef=""" & "DM" & ilFirstValid & """"  'smValues(DEMOID)
    End If
End Function
Public Function mParseDemoName(slName As String, ilIndex As Integer) As Boolean
    Dim slLetter As String
    Dim slCompare As String
    Dim ilPos As Integer
    Dim slNumber As String
    Dim slMin As String
    Dim slMax As String
    Dim c As Integer
    Dim slCompareEthnic As String
    Dim ilCharToTest As Integer
    'block custom names currently allows custom names where in format of m5-99
    mParseDemoName = False
    'Dan 8/23/13 cumulus doesn't capitalize all custom demos
    slName = UCase(slName)
    '6143 allow certain customs:  S panish H ispanic B lack U rban  ex: HA15-25
    slCompareEthnic = "SHBU"
    slCompare = "AMWTCPH"
    If InStr(1, slCompareEthnic, Mid(slName, 1, 1)) > 0 And InStr(1, slCompare, Mid(slName, 2, 1)) > 0 Then
        ilCharToTest = 2
    Else
        ilCharToTest = 1
    End If
On Error GoTo ERRORBOX
    slLetter = Mid(slName, ilCharToTest, 1)
    ilPos = InStr(1, slCompare, slLetter, vbTextCompare)
    If ilPos > 0 Then
        Select Case ilPos
            Case 1
                smValues(DEMOGROUP) = "Adults"
            Case 2
                smValues(DEMOGROUP) = "Men"
            Case 3
                smValues(DEMOGROUP) = "Women"
            Case 4
                smValues(DEMOGROUP) = "Teens"
            Case 5
                smValues(DEMOGROUP) = "Children"
            Case 6
                 smValues(DEMOGROUP) = "Persons"
           Case 7
                smValues(DEMOGROUP) = "Homes"
        End Select
        ilPos = Len(slName)
        'm15-25..get 15 then get 25
        For c = ilCharToTest + 1 To ilPos
            slNumber = Mid(slName, c, 1)
            If IsNumeric(slNumber) Then
                slMin = slMin & slNumber
            Else
                ilPos = c
                Exit For
            End If
        Next c
        If ilPos <= Len(slName) Then
            slNumber = Mid(slName, ilPos + 1)
            If IsNumeric(slNumber) Then
                slMax = slNumber
            ElseIf slNumber = "+" Then
                slMax = "99"
            ElseIf InStr(1, slName, "+") > 0 Then
                slMax = "99"
            End If
        End If
        'if not acceptable number, set demoagefrom and to are blank so won't be written
        If IsNumeric(slMin) And IsNumeric(slMax) Then
            slCompare = ",2,6,12,18,21,25,35,45,50,55,65,"
            If InStr(1, slCompare, "," & slMin & ",", vbBinaryCompare) = 0 Then
                Exit Function
            End If
            slCompare = ",5,11,17,20,24,34,44,49,54,64,99,"
            If InStr(1, slCompare, "," & slMax & ",", vbBinaryCompare) = 0 Then
                Exit Function
            End If
            mParseDemoName = True
            smValues(DEMOAGEFROM) = slMin
            smValues(DEMOAGETO) = slMax
            '6143. add after all tests ok.
            If ilCharToTest = 2 Then
                Select Case Mid(slName, 1, 1)
                    Case "S", "s"
                        smDemoComment = smDemoComment & "Demo #" & ilIndex & ": Spanish "
                    Case "H", "h"
                        smDemoComment = smDemoComment & "Demo #" & ilIndex & ": Hispanic "
                    Case "B", "b"
                        smDemoComment = smDemoComment & "Demo #" & ilIndex & ": Black "
                    Case "U", "u"
                        smDemoComment = smDemoComment & "Demo #" & ilIndex & ": Urban "
                    Case Else
                        ilCharToTest = 1
                End Select

            End If
        End If
    End If
    Exit Function
ERRORBOX:
    mParseDemoName = False
End Function
Private Function mXmlSellerBuyerInfo(tlagf As AGF) As Boolean
    Dim tlSlfSrchKey As INTKEY0
    Dim hlSlf As Integer
    Dim ilRet As Integer
    Dim ilSlfRecLen As Integer
    Dim tlSlf As SLF
    Dim slStr As String
    Dim hlSof As Integer
    Dim ilSofRecLen As Integer
    Dim tlSofSrchKey As INTKEY0
    Dim tlSof As SOF
    Dim ilLoop As Integer
    Dim ilPos As Integer
    
On Error GoTo ERRORBOX
    smValues(SELLERREFERENCE) = tgChfCntr.lCntrNo
    smValues(BCOMPANYNAME) = tlagf.sName
    smValues(BNAME) = tgChfCntr.sBuyer
    For ilLoop = 2 To 0 Step -1
        ilPos = InStr(1, tlagf.sCntrAddr(ilLoop), ",", vbTextCompare)
        If ilPos > 0 Then
            smValues(BOFFICENAME) = Left$(tlagf.sCntrAddr(ilLoop), ilPos - 1)
            Exit For
        End If
    Next ilLoop
'8915
#If programmatic <> 1 Then
    If Contract.lbcAdvt.ListIndex >= 0 Then
        smValues(ADVNAME) = Contract.lbcAdvt.List(Contract.lbcAdvt.ListIndex)
    End If
#Else
    Set myRst = gSQLSelectCall("Select  adfname from ADF_Advertisers where adfCode = " & tgChfCntr.iAdfCode)
    If Not myRst.EOF Then
        smValues(ADVNAME) = Trim$(myRst!adfName)
    End If
#End If
    smValues(SCOMPANYNAME) = tgSpf.sGClient
    tlSlfSrchKey.iCode = tgChfCntr.iSlfCode(0)
    hlSlf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlSlf, "", sgDBPath & "Slf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    #If programmatic <> 1 Then
        gBtrvErrorMsg ilRet, "mXmlSellerBuyerInfo (btrOpen: Slf.Btr)", Contract
    #Else
        gBtrvErrorMsg ilRet, "mXmlSellerBuyerInfo (btrOpen: Slf.Btr)", ProgrammaticBuy
    #End If
    If ilRet = BTRV_ERR_NONE Then
        ilSlfRecLen = Len(tlSlf)
        ilRet = btrGetEqual(hlSlf, tlSlf, ilSlfRecLen, tlSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smValues(SALESNAME) = Trim$(tlSlf.sFirstName) & " " & Trim$(tlSlf.sLastName)
            If tlSlf.iSofCode > 0 Then
                hlSof = CBtrvTable(ONEHANDLE)
                ilRet = btrOpen(hlSof, "", sgDBPath & "Sof.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                #If programmatic <> 1 Then
                    gBtrvErrorMsg ilRet, "mXmlSellerBuyerInfo (btrOpen: Sof.Btr)", Contract
                #Else
                    gBtrvErrorMsg ilRet, "mXmlSellerBuyerInfo (btrOpen: Sof.Btr)", ProgrammaticBuy
                #End If
                If ilRet = BTRV_ERR_NONE Then
                    ilSofRecLen = Len(tlSof)
                    tlSofSrchKey.iCode = tlSlf.iSofCode
                    ilRet = btrGetEqual(hlSof, tlSof, ilSofRecLen, tlSofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                            '2-18-05 reference short client name if exists vs long client name
                           smValues(SOFFICENAME) = tlSof.sName
                    End If
                End If
                ilRet = btrClose(hlSof)
                btrDestroy hlSof
            End If
        End If
    End If
    mXmlSellerBuyerInfo = True
CLEANUP:
    ilRet = btrClose(hlSlf)
    btrDestroy hlSlf
    Exit Function
ERRORBOX:
    mXmlSellerBuyerInfo = False
    GoTo CLEANUP
End Function
Private Function mConvertDate(llPacked() As Integer) As String
    Dim slDate As String
    gUnpackDate llPacked(0), llPacked(1), slDate
    mConvertDate = Format(slDate, XMLDATE)
End Function
Public Function gRepLineName(ilRepCounter As Integer) As String
    '7687
    Dim slVehicle As String
    Dim slLine() As String
    Dim slDelimiter As String
    
    slDelimiter = Chr(9)
    slVehicle = ""
    If ilRepCounter > -1 And ilRepCounter < UBound(smClfInfo) Then
        slLine = Split(smClfInfo(ilRepCounter), slDelimiter)
        slVehicle = slLine(XMLVEHICLE)
        If slVehicle = "" Then
            slVehicle = mReturnPreviousName(ilRepCounter)
        End If
    End If
    Erase slLine
    gRepLineName = slVehicle
    
End Function
Public Function gXMLNameFilter(slInName As String) As String
    Dim slName As String
    Dim ilPos As Integer
    Dim ilStartPos As Integer
    Dim ilFound As Integer
    
    slName = slInName
    'Remove " and '
    ilStartPos = 1
    Do
        ilFound = False
        ilPos = InStr(ilStartPos, slName, "&", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & "&amp;" & Mid$(slName, ilPos + 1)
            ilStartPos = ilPos + Len("&amp;")
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, "<", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & "&lt;" & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, ">", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & "&gt;" & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, "'", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & "&apos;" & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    Do
        ilFound = False
        ilPos = InStr(1, slName, """", 1)
        If ilPos > 0 Then
            slName = Left$(slName, ilPos - 1) & "&quot;" & Mid$(slName, ilPos + 1)
            ilFound = True
        End If
    Loop While ilFound
    '9158
    slName = gRemoveIllegalChars(slName)
    gXMLNameFilter = Trim$(slName)
End Function

Private Function mReturnPreviousName(ilRepCounter As Integer) As String
    '7673 11/11/15  handle 'sub-lines' where the rate has changed mid line  note no vehicle name.  in previous line
    '2   1   1   1   1   1   0   0   False       6AM-12AM        55.00                                               0   0   0   0   7   0
    Dim J As Integer
    Dim slRet As String
    Dim slLine() As String
    Dim slDelimiter As String
    
    slDelimiter = Chr(9)
    slRet = ""
    If ilRepCounter > 0 Then
        'go backwards to find the vehicle
        For J = ilRepCounter - 1 To 0 Step -1
            slLine = Split(smClfInfo(J), slDelimiter)
            slRet = slLine(XMLVEHICLE)
            If slRet <> "" Then
                Exit For
            End If
        Next J
    End If
    Erase slLine
    mReturnPreviousName = slRet
End Function
Public Function gAllowRepLine(ilIndex As Integer, ilRepCounter As Integer, slRepVehicle As String) As Boolean
    '7687
    Dim blRet As Boolean
    Dim slLine() As String
    Dim slDelimiter As String
    Dim c As Integer
    Dim blIsFirst As Boolean
    '7673
    Dim slVehicle As String
    
    blRet = True
    If ilRepCounter > -1 Then
        blIsFirst = True
        slDelimiter = Chr(9)
        'don't allow vehicle if already sent
        For c = 0 To ilRepCounter - 1 Step 1
            slLine = Split(smClfInfo(c), slDelimiter)
            If slRepVehicle = slLine(XMLVEHICLE) Then
                blIsFirst = False
                blRet = False
                Exit For
            End If
        Next c
        'ok to write this out...and any that follow
        If blIsFirst Then
            'a match means this line is good to go.  Let's test other lines to see if they go out too
            If ilRepCounter <> ilIndex Then
                blRet = False
                'allow this one if name matches.
                slLine = Split(smClfInfo(ilIndex), slDelimiter)
                slVehicle = slLine(XMLVEHICLE)
                If slVehicle = "" Then
                    slVehicle = mReturnPreviousName(ilIndex)
                End If
                If slRepVehicle = slVehicle Then
                    blRet = True
                End If
            End If
        End If
    '-1 counts as true
    Else
        blRet = True
    End If
    Erase slLine
    gAllowRepLine = blRet
End Function
Public Function mWriteXml(slXmlErrorMessage As String, ilRepCounter As Integer, blSupressRate As Boolean) As Boolean
    'I: ilRepCounter = -1 for agencyXml
    Dim c As Integer
    Dim slLine As String
    Dim ilFlightAsLine As Integer
    Dim blDaily As Boolean
    Dim blWriteDemo As Boolean
    'this line doesn't want to write the AvailList info, just the availLine info
    'Dim blSeparateAvailLine As Boolean
    'this line is availList info, but don't close when done
   ' Dim blWillContainMultipleAvailLines As Boolean
    Dim slDemoForline As String
    Dim slDemoHtml() As String
    Dim ilLines As Integer
    '7687
    Dim slRepVehicle As String
    Dim slHiddenLines() As String
    Dim ilLineNumber As Integer
    
    '8032 added 'if'.  WideOrbit and Agency go out as original proposalxml
    If ilRepCounter = -1 Or igStationXmlChoice = STATIONXMLWIDEORBIT Then
       ilLines = UBound(smClfInfo)
       ReDim slDemoHtml(ilLines) As String
    On Error GoTo ERRORBOX
       mWriteXmlHeader
       mWriteXmlSellerBuyer
       For c = PROPOSALNAME To SELLERREFERENCE
           mCSIXMLData "CD", smFields(c), smValues(c)
       Next c
       mWriteXmlStations
       mWriteXmlListStart
       blWriteDemo = mWriteXmlDemoHeader()
       '7687 expand the allowed lines to the same vehicle and don't allow if vehicle in previous line
       slRepVehicle = gRepLineName(ilRepCounter)
       'proposal xml writes out each line into one file; insertion orders has to go through hoops to group only similar lines together
       For c = 0 To UBound(smClfInfo) - 1
           'ttp 5311.  Found that I never use blWillContainMultipleAvailLines
         '  blWillContainMultipleAvailLines = mMultipleAvailLines(c)
         '  If mParseForXmlLine(c, blWillContainMultipleAvailLines, blDaily, slDemoForline, slXmlErrorMessage) Then
         '7673 added 'if' wrapper  if ilRepCounter is -1, do each line; otherwise, only do the matching line
           'If ilRepCounter = -1 Or c = ilRepCounter Then
           If gAllowRepLine(c, ilRepCounter, slRepVehicle) Then
             '6472 false now means no spots in the line
             '8940 added blSupressRate
               If mParseForXmlLine(c, smClfInfo, blDaily, slDemoForline, blSupressRate) Then
                   mWriteXmlAvailLine blDaily, blWriteDemo, c
                   slDemoHtml(c) = slDemoForline
       '        Else
       '            mWriteXml = False
       '            GoTo CLEANUP
               End If
           End If
       Next c
       '7209
#If programmatic <> 1 Then
       If blWriteDemo Then
            '8940 adds blsupressrate
           mWriteXmlDemoExtra slDemoHtml, blSupressRate
       Else
           mCSIXMLData "CT", "AvailList", ""
       End If
#Else
        mCSIXMLData "CT", "AvailList", ""
#End If
       mCloseXmlHeader
       mWriteXml = csiXMLWrite(1)
    'marketron or other StationXML besides wideOrbit
    Else
        Select Case igStationXmlChoice
        Case STATIONXMLMARKETRON
            '8278
            If Len(ogEOrder.ErrorMessage) = 0 Then
               ilLines = UBound(smClfInfo)
               ReDim slDemoHtml(ilLines) As String
            On Error GoTo ERRORBOX
               mWriteXmlHeaderStationMarketron
               mWriteXmlSellerBuyerStationMarketron
               mWriteXmlOrderStartMarketron
    '           blWriteDemo = mWriteXmlDemoHeader()
                slRepVehicle = gRepLineName(ilRepCounter)
                For c = 0 To UBound(smClfInfo) - 1
                   If gAllowRepLine(c, ilRepCounter, slRepVehicle) Then
                       '8940 added blSupressRate
                       If mParseForXmlLine(c, smClfInfo, blDaily, slDemoForline, blSupressRate) Then
                           mWriteXmlBuyLineMarketron blDaily, blWriteDemo, c
                         '  slDemoHtml(c) = slDemoForline
                       End If
                   End If
               Next c
    '           If blWriteDemo Then
    '               mWriteXmlDemoExtra slDemoHtml
    '           Else
    '               mCSIXMLData "CT", "AvailList", ""
    '           End If
               mCloseXmlHeaderMarketron
               If Not csiXMLWrite(1) Then
                    ogStationXmlErrors.WriteWarning "Could not generate file for " & STATIONXMLMARKETRONNAME & ": " & smValues(RADIOCALLLETTERS) & "-" & smValues(RADIOBAND), True
               End If
                If bgStationXmlWebServiceAllowed(StationXmlIndexes.MARKETRONEO) Then
                    If ogEOrder.IsStation(smValues(RADIOCALLLETTERS), smValues(RADIOBAND)) Then
                        If ogEOrder.SendOrder() Then
                             mWriteXml = True
                             If ogEOrder.isTest Then
                                 ogStationXmlErrors.WriteFacts "order would have been sent to " & STATIONXMLMARKETRONNAME & ": " & smValues(RADIOCALLLETTERS) & "-" & smValues(RADIOBAND), True
                             Else
                                 ogStationXmlErrors.WriteFacts "order sent to " & STATIONXMLMARKETRONNAME & ": " & smValues(RADIOCALLLETTERS) & "-" & smValues(RADIOBAND), True
                             End If
                        Else
                             mWriteXml = False
                             ogStationXmlErrors.WriteWarning "Could not send to " & STATIONXMLMARKETRONNAME & ": " & smValues(RADIOCALLLETTERS) & "-" & smValues(RADIOBAND) & ": " & ogEOrder.ErrorMessage, True
                             gMsgBox "Could not send to " & STATIONXMLMARKETRONNAME & ": " & smValues(RADIOCALLLETTERS) & "-" & smValues(RADIOBAND) & ": " & ogEOrder.ErrorMessage
                        End If
                    Else
                         mWriteXml = False
                         ogStationXmlErrors.WriteWarning "Could not send to " & STATIONXMLMARKETRONNAME & ": " & smValues(RADIOCALLLETTERS) & "-" & smValues(RADIOBAND) & " to Marketron-" & ogEOrder.ErrorMessage, True
                         gMsgBox "Could not send to " & STATIONXMLMARKETRONNAME & ": " & smValues(RADIOCALLLETTERS) & "-" & smValues(RADIOBAND) & " to Marketron-" & ogEOrder.ErrorMessage
                    End If
                Else
                    mWriteXml = True
                    ogStationXmlErrors.WriteFacts "User chose not to send to " & STATIONXMLMARKETRONNAME & ": " & smValues(RADIOCALLLETTERS) & "-" & smValues(RADIOBAND), True
                End If
                ogEOrder.Clear True
            Else
                 mWriteXml = False
                ogStationXmlErrors.WriteWarning "Could not send to " & STATIONXMLMARKETRONNAME & ": " & smValues(RADIOCALLLETTERS) & "-" & smValues(RADIOBAND) & " to Marketron-" & ogEOrder.ErrorMessage, True
                gMsgBox "Could not send to " & STATIONXMLMARKETRONNAME & ": " & smValues(RADIOCALLLETTERS) & "-" & smValues(RADIOBAND) & " to Marketron-" & ogEOrder.ErrorMessage
            End If 'len
        End Select
    End If 'station/wo vs. agency xml
CLEANUP:
    csiXMLEnd
    Exit Function
ERRORBOX:
    mWriteXml = False
    slXmlErrorMessage = err.Source & " " & err.Description
End Function
Private Sub mWriteXmlDemoExtra(slDemoforHtml() As String, blSupress As Boolean)
    'ttp4930 adds blSupress
        'TTP5230
    Dim slGrossImpressions As String
    Dim c As Integer
    
    mCSIXMLData "CT", "AvailList", ""
    '5890
    rsDemoValues.Filter = "DemoIndex = 0"
    If rsDemoValues.EOF Then
        Exit Sub
    End If
    If tgSpf.sSAudData = "H" Then
        slGrossImpressions = gLongToStrDec(lmCPriDemoGrImp, 1)
    ElseIf tgSpf.sSAudData = "N" Then
        slGrossImpressions = gLongToStrDec(lmCPriDemoGrImp, 2)
    ElseIf tgSpf.sSAudData = "U" Then
        slGrossImpressions = gLongToStrDec(lmCPriDemoGrImp, 3)
    Else
        slGrossImpressions = gLongToStrDec(lmCPriDemoGrImp, 0)
    End If
    mCSIXMLData "OT", "SalesInformation", "title='Demo Information' salesInfoId='SI0'"
    mCSIXMLData "OT", "Content", ""
    mCSIXMLData "OT", "html", ""
    mCSIXMLData "OT", "body", "leftmargin = '0' topmargin = '0' marginwidth = '0' marginheight = '0'"
    mCSIXMLData "OT", "table", "border ='0' cellpadding ='10' width = '80%'"
    mCSIXMLData "OT", "tr", "align='center'"
    csiXMLDataEx "<th colspan = '4'>Demo Information</th>"
    mCSIXMLData "CT", "tr", ""
    'header
    mCSIXMLData "OT", "tr", "bgcolor='#CCCC99'"
    mCSIXMLData "OT", "th", "bgcolor='#FFFFFF'"
    mCSIXMLData "CT", "th", ""
    mCSIXMLData "CD", "th", "Book"
    mCSIXMLData "CD", "th", "Rating"
    mCSIXMLData "CD", "th", "Audience"
    mCSIXMLData "CD", "th", "Population"
    mCSIXMLData "CD", "th", "GrImp"
    mCSIXMLData "CD", "th", "GRP"
    '8940
    If Not blSupress Then
        mCSIXMLData "CD", "th", "CPM"
        mCSIXMLData "CD", "th", "CPP"
    End If
    mCSIXMLData "CT", "tr", ""
    'data
    For c = UBound(slDemoforHtml) To 0 Step -1
        csiXMLDataEx slDemoforHtml(c)
    Next c
   ' last row
    mCSIXMLData "OT", "tr", "align='center'"
    csiXMLDataEx "<th bgcolor='#CCCC99'>" & "Proposal Average" & "</th>"
    mCSIXMLData "CD", "td", ""
    mCSIXMLData "CD", "td", gIntToStrDec(imCPriDemoRating, 1)
    mCSIXMLData "CD", "td", ""
    mCSIXMLData "CD", "td", ""
    mCSIXMLData "CD", "td", slGrossImpressions
    mCSIXMLData "CD", "td", gLongToStrDec(lmCPriDemoGRP, 1)
    '8940
    If Not blSupress Then
        mCSIXMLData "CD", "td", gLongToStrDec(lmCPriDemoCPM, 2)
        mCSIXMLData "CD", "td", gLongToStrDec(lmCPriDemoCPP, 0)
    End If
    mCSIXMLData "CT", "tr", ""
   'close
    mCSIXMLData "CT", "table", ""
    mCSIXMLData "CT", "body", ""
    mCSIXMLData "CT", "html", ""
    mCSIXMLData "CT", "Content", ""
    mCSIXMLData "CT", "SalesInformation", ""
End Sub

Private Function mParseForXmlLine(ilIndex As Integer, slArray() As String, blDaily As Boolean, slDemoForline As String, blSupressRate As Boolean) As Boolean
'   7803  pass in array
    'I index to smClfInfo
    'I blMultiple--will there be multiple availLines? , blMultiple As Boolean  UNUSED!
    'O  able to parse true false
    'O blDaily  daily or weekly?
    'O slDemoForline
    Dim slLine() As String
    Dim slDates() As String
    Dim slTimes() As String
    Dim slDelimiter As String
    Dim ilLineSize As Integer
    Dim ilDateSize As Integer
    Dim c As Integer
    Dim blSafe As Boolean
    Dim slPresentKey As String
    Dim slPreviousKey As String
    Dim slTempLine As String
    ' blAnotherAvailLine..is the next line distinct from previous, or just a change of times or days? This changes whether we read certain values later in method
    Dim blAnotherAvailLine As Boolean
    Dim slComment As String
    Dim slRowNumber As String
    Dim ilTab As Integer
    '6472
    Dim blSpotInLine As Boolean
    
'0   1   2   3   4   5   6   7    8       9                10        11   12      13  14      15  16     17   18   19     20     21     22           23  spots by weeks
'0   1   1   1   1   1   1   1   False   NCAA Football   12AM-12AM   30  340.00  870 8900.00 4.3 112.4   79  .39 22620   20130   N   Spring 2000 ARB 26  24  1   0   0   0   0   0   1
'Const XMLCLFLINE As Integer = 0
'Const XMLMON As Integer = 1
'Const XMLTUE As Integer = 2
'Const XMLWED As Integer = 3
'Const XMLTHU As Integer = 4
'Const XMLFRI As Integer = 5
'Const XMLSAT As Integer = 6
'Const XMLSUN As Integer = 7
'Const XMLDAILY As Integer = 8
'Const XMLVEHICLE As Integer = 9
'Const XMLTIMES As Integer = 10
'Const XMLSPOTLENGTH As Integer = 11
'Const XMLRATE As Integer = 12
'Const XMLAVGAUD As Integer = 13
'Const XMLRATING As Integer = 15
'Const XMLGRP As Integer = 16
'Const XMLCPP As Integer = 17
'Const XMLCPM As Integer = 18
'Const XMLGRIMP As Integer = 19
'Const XMLPOPULATION As Integer = 20
'Const XMLPACKAGE As Integer = 21 '20
'Const XMLRATINGSOURCE As Integer = 22 '21
'Const XMLTOTALSPOTS As Integer = 23 '22
    slDelimiter = Chr(9)
    '7803 changed
    slLine = Split(slArray(ilIndex), slDelimiter)
    slDates = Split(smClfDates, slDelimiter)
    'for both arrays, upper is blank
    ilLineSize = UBound(slLine)
    ilDateSize = UBound(slDates)
    If ilLineSize >= 1 Then
        slPresentKey = slLine(XMLCLFLINE)
        'don't test first one...will cause error
      If ilIndex > 0 Then
            'ttp 5311: not always one character! line could be #84
            '7803
            slPreviousKey = mGetClfLineNumber(slArray(ilIndex - 1))
           ' slPreviousKey = mGetClfLineNumber(ilIndex - 1)
           ' slPreviousKey = Mid(smClfInfo(ilIndex - 1), 1, 1)
            If slPresentKey = slPreviousKey Then
                blAnotherAvailLine = True
            Else
               blAnotherAvailLine = False
            End If
        Else
            blAnotherAvailLine = False
        End If
        ' make sure array is safe.  should be a determined size up to xmltotalspots, and then the weeks after that. so add them together
        'add one.  ildatesize, since is 1 larger than it what it contains, is actually the count.
        If ilLineSize = XMLTOTALSPOTS + ilDateSize + 1 Then
            blSafe = True
        Else
            blSafe = False
        End If
    Else
        blSafe = False
    End If
    If blSafe Then
        If Not blAnotherAvailLine Then
            smValues(DAYAVAILNAME) = slLine(XMLVEHICLE)
            mGetRdfNames slPresentKey, smValues(DAYNAME)
            'Dan m 4/26/2012 error on ""
            'smValues(DAYSPOTLENGTH) = TimeSerial(0, 0, slLine(XMLSPOTLENGTH))
            '8032 marketron needs original integer, so move this somewhere else
'            smValues(DAYSPOTLENGTH) = TimeSerial(0, 0, Val(slLine(XMLSPOTLENGTH)))
'            smValues(DAYSPOTLENGTH) = Format(smValues(DAYSPOTLENGTH), XMLSPOTTIME)
            smValues(DAYSPOTLENGTH) = Val(slLine(XMLSPOTLENGTH))
            '6AM-2PM,3PM-7PM,etc.
            slTimes = Split(slLine(XMLTIMES), "-")
            ' ttp 5205
'            If UBound(slTimes) = 1 Then
'                smValues(DAYTIMESTARTTIME) = mXmlAdjustMidnight(slTimes(0))
'                smValues(DAYTIMEENDTIME) = mXmlAdjustMidnight(slTimes(1))
'            Else
'            'Dan todo is this really ok? or get out?
'                smValues(DAYTIMESTARTTIME) = ""
'                smValues(DAYTIMEENDTIME) = ""
'            End If
            '9126 moved to mAdjustTimes
'            smValues(DAYTIMESTARTTIME) = mXmlAdjustMidnight(slTimes(0))
'            smValues(DAYTIMEENDTIME) = mXmlAdjustMidnight(slTimes(UBound(slTimes)))
            'ttp 5644
'            smValues(DAYTIMEENDTIME) = mXmlAdjustOverMidnight(smValues(DAYTIMESTARTTIME), smValues(DAYTIMEENDTIME))
            smValues(DAYTIMESTARTTIME) = slTimes(0)
            smValues(DAYTIMEENDTIME) = slTimes(UBound(slTimes))
            mAdjustDaypartTimes smValues(DAYTIMESTARTTIME), smValues(DAYTIMEENDTIME), smValues(DAYAVAILNAME)
            smValues(PERIODDEMOVALUE) = slLine(XMLRATING)
            'TTP 5526
            'smValues(PERIODDEMOGRIMP) = slLine(XMLGRIMP)
            smValues(PERIODDEMOAUD) = slLine(XMLAVGAUD)
        End If
        For c = XMLMON To XMLSUN
            If slLine(c) = 1 Then
                smValues(c) = "Y"
            Else
                smValues(c) = "N"
            End If
        Next c
        'ttp 5230
    On Error Resume Next
        'D.S. 5/30/18 added conditional compile
        #If programmatic <> 1 Then
            slComment = smLnComment(slPresentKey + 1)
            On Error GoTo 0
            slComment = Trim$(Replace(slComment, vbCrLf, ""))
        #Else
            slComment = ""
        #End If
        'ttp 5412, I stumbled upon this
       ' smValues(COMMENTPROPOSAL) = gFileNameFilter(slComment)
       'dan 8/2/16 this gets done later.  don't do now
        smValues(COMMENTPROPOSAL) = slComment 'gXMLNameFilter(slComment)
        If Len(slLine(XMLRATINGSOURCE)) > 0 Then
            '8940 added blSupressRate
            slDemoForline = mXmlDemoHtml(slLine, slPresentKey, blSupressRate)
        Else
            slDemoForline = ""
        End If
        smValues(PERIODSTARTDATE) = ""
        smValues(PERIODENDDATE) = ""
        '8940
       ' smValues(PERIODRATE) = slLine(XMLRATE)
        If blSupressRate Then
            smValues(PERIODRATE) = 0
        Else
            smValues(PERIODRATE) = slLine(XMLRATE)
        End If
        'build string of spots to split later
        For c = XMLTOTALSPOTS + 1 To ilLineSize - 1
            slTempLine = slTempLine & slLine(c) & slDelimiter
            If slLine(c) > 0 Then
                blSpotInLine = True
            End If
        Next c
        smValues(PERIODSPOTSPERWEEK) = slTempLine
        blDaily = slLine(XMLDAILY)
        'mParseForXmlLine = True
'        If Not blSpotInLine Then
'            blSpotInLine = blSpotInLine
'        End If
        mParseForXmlLine = blSpotInLine
    Else
        mParseForXmlLine = False
    End If
End Function
Private Sub mAdjustDaypartTimes(slStartTime As String, slEndTime As String, slVehicleName As String)
    mAdjustTimesForTimeZone slStartTime, slEndTime, slVehicleName
    slStartTime = mXmlAdjustMidnight(slStartTime)
    slEndTime = mXmlAdjustMidnight(slEndTime)
   'ttp 5644
'    slEndTime = mXmlAdjustOverMidnight(slStartTime, slEndTime)
    '9633
    mXmlAdjustEarlyTimes slStartTime, slEndTime
End Sub
Private Sub mAdjustTimesForTimeZone(slStartTime As String, slEndTime As String, slVehName As String)
    '9126
    Dim slSiteTimeZoneChoice As String
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim ilTimeAdj As Integer
    Dim llTime As Long
    Dim ilZone As Integer
    
    Dim slDate As String
    'does not increment/decrement date. '9442 discovered and changed here
    slSiteTimeZoneChoice = tgSpf.sInvSpotTimeZone
    If (slSiteTimeZoneChoice = "E") Or (slSiteTimeZoneChoice = "C") Or (slSiteTimeZoneChoice = "M") Or (slSiteTimeZoneChoice = "P") Then
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If StrComp(Trim$(tgMVef(ilVef).sName), slVehName, vbTextCompare) = 0 Then
                ilVpf = gBinarySearchVpf(tgMVef(ilVef).iCode)
                If ilVpf <> -1 Then
                    For ilZone = 0 To 4 Step 1
                        If Left$(tgVpf(ilVpf).sGZone(ilZone), 1) = slSiteTimeZoneChoice Then
                            ilTimeAdj = tgVpf(ilVpf).iGLocalAdj(ilZone)
                            Exit For
                        End If
                    Next ilZone
                End If
                Exit For
            End If
        Next ilVef
        If ilTimeAdj <> 0 Then
            llTime = gTimeToLong(slStartTime, False) + (CLng(ilTimeAdj) * 3600)
            If llTime < 0 Then
                llTime = 86400 + llTime
                slDate = gDecOneDay(slDate)
            ElseIf llTime > 86399 Then
                llTime = llTime - 86400
            End If
            slStartTime = gFormatTimeLong(llTime, "A", "1")
            llTime = gTimeToLong(slEndTime, False) + (CLng(ilTimeAdj) * 3600)
            If llTime < 0 Then
                llTime = 86400 + llTime
                slDate = gDecOneDay(slDate)
            ElseIf llTime > 86399 Then
                llTime = llTime - 86400
                slDate = gIncOneDay(slDate)
            End If
            slEndTime = gFormatTimeLong(llTime, "A", "1")
        End If
    End If
End Sub
Private Function mGetClfLineNumber(slInLine As String) As String
    '7803 from index to line
    'ttp 5311: not always one character! line could be #84
    Dim ilTab As Integer
    Dim slKey As String
    
    slKey = ""
    ilTab = InStr(1, slInLine, vbTab, vbBinaryCompare)
    If ilTab > 0 Then
        slKey = Mid(slInLine, 1, ilTab - 1)
    Else
       slKey = Mid(slInLine, 1, 1)
    End If
    mGetClfLineNumber = slKey
End Function
'Private Function mGetClfLineNumber(ilIndex As Integer) As String
'    'ttp 5311: not always one character! line could be #84
'    Dim ilTab As Integer
'    Dim slKey As String
'
'    slKey = ""
'    ilTab = InStr(1, smClfInfo(ilIndex), vbTab, vbBinaryCompare)
'    If ilTab > 0 Then
'        slKey = Mid(smClfInfo(ilIndex), 1, ilTab - 1)
'    Else
'       slKey = Mid(smClfInfo(ilIndex), 1, 1)
'    End If
'    mGetClfLineNumber = slKey
'End Function
Private Function mXmlDemoHtml(slLine() As String, slPresentKey As String, blSupress As Boolean) As String
    Dim slTempString As String
    Dim slLineNumber As String
    'slLine = line of info to parse
    'slPresentKey = index to clf. to get counterpoint line # shown
'write out single line of demo values in html code
'slDemoForline = " Line Demo values-- Rating:" & slLine(XMLRATING) & "; Book: " & slLine(XMLRATINGSOURCE) & "; Audience:" & slLine(XMLAVGAUD) & "; Population:" & slLine(XMLPOPULATION) & "; GrImp:" & slLine(XMLGRIMP) & "; GRP:" & slLine(XMLGRP) & "; CPM:" & slLine(XMLCPM) & "; CPP:" & slLine(XMLCPP)
'<tr align='center'>
'<th bgcolor='#CCCC99'>Line 1</th>
'<td>Spring 2000 ARB</td>
'<td>4.3</td>
'<td>870</td>
'<td>21030</td>
'<td>22620</td>
'<td>112.4</td>
'<td>.39</td>
'<td>79</td>
'</tr>
If Val(Trim$(slLine(XMLPOPULATION))) <= 0 Then
    slLine(XMLPOPULATION) = ""
End If
slLineNumber = tgClfCntr(slPresentKey).ClfRec.iLine
'ttp 5412
'D.S. 5/30/18 conditional compile
'slTempString = "<tr align='center'><th bgcolor='#CCCC99'>Line " & slLineNumber & " </th><td>" & slLine(XMLRATINGSOURCE) & "</td><td>" & slLine(XMLRATING) & "</td>"
slTempString = "<tr align='center'><th bgcolor='#CCCC99'>Line " & slLineNumber & " </th><td>" & gXMLNameFilter(slLine(XMLRATINGSOURCE)) & "</td><td>" & slLine(XMLRATING) & "</td>"
'8940
'slTempString = slTempString & "<td>" & slLine(XMLAVGAUD) & "</td><td>" & slLine(XMLPOPULATION) & "</td><td>" & slLine(XMLGRIMP) & "</td><td>" & slLine(XMLGRP) & "</td><td>" & slLine(XMLCPM) & "</td><td>" & slLine(XMLCPP) & "</td></tr>"
slTempString = slTempString & "<td>" & slLine(XMLAVGAUD) & "</td><td>" & slLine(XMLPOPULATION) & "</td><td>" & slLine(XMLGRIMP) & "</td><td>" & slLine(XMLGRP)
If blSupress Then
    slTempString = slTempString & "</td></tr>"
Else
    slTempString = slTempString & "</td><td>" & slLine(XMLCPM) & "</td><td>" & slLine(XMLCPP) & "</td></tr>"
End If
mXmlDemoHtml = slTempString
End Function
Private Function mXmlAdjustMidnight(slTime As String) As String
    'midnight as 24:00
    Dim slTempTime As String
    '9589
    slTempTime = slTime
    If StrComp(slTempTime, "12M", vbTextCompare) = 0 Then
        slTempTime = "12AM"
    ElseIf StrComp(slTempTime, "12N", vbTextCompare) = 0 Then
        slTempTime = "12PM"
    End If
   ' slTempTime = Format(slTime, XMLTIME)
    slTempTime = Format(slTempTime, XMLTIME)
    If slTempTime = Format("00:00:00", XMLTIME) Then
        slTempTime = Format("24:00", XMLTIME)
    End If
    mXmlAdjustMidnight = slTempTime
End Function
Private Sub mXmlAdjustEarlyTimes(slStartInput As String, slEndInput As String)
    'ttp 9633 replaces mXmlAdjustOverMidnight
    ' 1 am to 4 am not allowed.  Also, end must be later than start
    ' if not valid or cannot read, does not stop
    Dim ilHour As Integer
    Dim slMinutes As String
    Dim ilPos As Integer
    Dim ilTest As Integer
    
    'start time is before 5.  End time is before 5, or before 7 (if before 7, turns 5 and 6 to 29 and 30.  Use this if start > 7
    ilTest = 5
    If IsDate(slStartInput) Then
        ilHour = DatePart("h", slStartInput)
        If ilHour < ilTest Then
            ilHour = ilHour + 24
            ilPos = InStr(1, slStartInput, ":")
            If ilPos = 2 Or ilPos = 3 Then
                slMinutes = Mid(slStartInput, ilPos)
                slStartInput = Format(ilHour & slMinutes, XMLTIME)
            End If
        End If
        If ilHour > 7 Then
            ilTest = 7
        End If
    ElseIf InStr(1, slStartInput, "24:") > 0 Then
        ilTest = 7
    End If
    If IsDate(slEndInput) Then
        ilHour = DatePart("h", slEndInput)
        If ilHour < ilTest Then
            ilHour = ilHour + 24
            ilPos = InStr(1, slEndInput, ":")
            If ilPos = 2 Or ilPos = 3 Then
                slMinutes = Mid(slEndInput, ilPos)
                slEndInput = Format(ilHour & slMinutes, XMLTIME)
            End If
        End If
    End If
End Sub
'Private Function mXmlAdjustOverMidnight(ByVal slStartInput As String, ByVal slEndInput As String) As String
'    'ttp 5644
'    ' if times cross midnight, add 24 to endtime.
'    ' if not valid or cannot read, does not stop
'    Dim ilStartDate As Integer
'    Dim ilEndDate As Integer
'    Dim ilNewEndDate As Integer
'    Dim slMinutes As String
'    Dim ilPos As Integer
'
'    If IsDate(slStartInput) Then
'       ilStartDate = DatePart("h", slStartInput)
'        If IsDate(slEndInput) Then
'            ilEndDate = DatePart("h", slEndInput)
'            If ilStartDate > ilEndDate And ilEndDate < 7 Then
'                ilNewEndDate = ilEndDate + 24
'                ilPos = InStr(1, slEndInput, ":")
'                'hour can only be 1-6, so always ilpos = 2. but in case "06"
'                If ilPos = 2 Or ilPos = 3 Then
'                    slMinutes = Mid(slEndInput, ilPos)
'                    slEndInput = Format(ilNewEndDate & slMinutes, XMLTIME)
'                End If
'            End If
'        'end not date
'        End If
'    ' 24:00 is not a valid time!
'    ElseIf InStr(1, slStartInput, "24:") > 0 Then
'        If IsDate(slEndInput) Then
'            ilEndDate = DatePart("h", slEndInput)
'            If ilEndDate < 7 Then
'                ilNewEndDate = ilEndDate + 24
'                ilPos = InStr(1, slEndInput, ":")
'                'hour can only be 1-6, so always ilpos = 2. but in case "06"
'                If ilPos = 2 Or ilPos = 3 Then
'                    slMinutes = Mid(slEndInput, ilPos)
'                    slEndInput = Format(ilNewEndDate & slMinutes, XMLTIME)
'                End If
'            End If
'        End If
'    'start not date
'    End If
'    mXmlAdjustOverMidnight = slEndInput
'End Function
Private Sub mGetRdfNames(slClf As String, slDay As String)
    Dim ilRdf As Integer
    
    ilRdf = gBinarySearchRdf(tgClfCntr(slClf).ClfRec.iRdfCode)
    If ilRdf <> -1 Then
        slDay = tgMRdf(ilRdf).sName
    End If
End Sub
Private Function mFindLineDates(slIndex As String, blFindEarliest As Boolean) As String
    ' going into tgClfCntr by line number    'the row is 1 greater than array
    Dim slTempDate As String
    
    If blFindEarliest Then
        gUnpackDate tgClfCntr(slIndex).ClfRec.iStartDate(0), tgClfCntr(slIndex).ClfRec.iStartDate(1), slTempDate
    Else
        gUnpackDate tgClfCntr(slIndex).ClfRec.iEndDate(0), tgClfCntr(slIndex).ClfRec.iEndDate(1), slTempDate
    End If
    mFindLineDates = Format(slTempDate, XMLDATE)
End Function
Private Sub mWriteXmlAvailLine(blDaily As Boolean, blDemo As Boolean, ilLineIndex As Integer)
    'ilLineIndex is index for smClfInfo.  ilClfIndex is the CLFLine in the array(tgClfCntr(ilClfIndex)).
    ' The row # is tgClfCntr(ilClfIndex).ClfRec.iLine
    Dim ilClfIndex As Integer
    Dim c As Integer
    Dim slACT1Info As String
    '8227
    Dim slSpotLength As String
    '8032  '8227 was resetting smvalues and that caused problem
    slSpotLength = TimeSerial(0, 0, Val(smValues(DAYSPOTLENGTH)))
    slSpotLength = Format(slSpotLength, XMLSPOTTIME)
    'ttp 5311: not always one character! line could be #84
    'ilClfIndex = Mid(smClfInfo(ilLineIndex), 1, 1)
    '7803 added 'smClfInfo()
    ilClfIndex = mGetClfLineNumber(smClfInfo(ilLineIndex))
    mCSIXMLData "OT", "AvailLineWithDetailedPeriods", ""
    mCSIXMLData "CA", "OutletReference", mXmlAttribute(AVAILOUTLET)
    mCSIXMLData "OT", "DayTimes", ""
    mCSIXMLData "OT", "DayTime", ""
    mCSIXMLData "CD", smFields(DAYTIMESTARTTIME), smValues(DAYTIMESTARTTIME)
    mCSIXMLData "CD", smFields(DAYTIMEENDTIME), smValues(DAYTIMEENDTIME)
    mCSIXMLData "OT", "Days", ""
    For c = XMLMON To XMLSUN
        mCSIXMLData "CD", smFields(c), smValues(c)
    Next c
    mCSIXMLData "CT", "Days", ""
    mCSIXMLData "CT", "DayTime", ""
    mCSIXMLData "CT", "DayTimes", ""
    mCSIXMLData "CD", smFields(DAYNAME), smValues(DAYNAME)
    mCSIXMLData "CD", smFields(DAYAVAILNAME), smValues(DAYAVAILNAME)
    '8227
    mCSIXMLData "CD", smFields(DAYSPOTLENGTH), slSpotLength
'    'ttp 5230
    mCSIXMLData "OT", "Comment", ""
    'line number to match for demo values
    ' added 6143 custom demo info
    mCSIXMLData "CD", smFields(COMMENTPROPOSAL), "Proposal line #: " & CStr(tgClfCntr(ilClfIndex).ClfRec.iLine) & " " & smDemoComment
    '9158 added gXmlNameFilter
    'this is line comment
    If Len(smValues(COMMENTPROPOSAL)) > 0 Then
        mCSIXMLData "CD", smFields(COMMENTPROPOSAL), " Line comment: " & gXMLNameFilter(smValues(COMMENTPROPOSAL))
    End If
    slACT1Info = mAddACT1Lineup(ilClfIndex)
    If slACT1Info <> "" Then
        mCSIXMLData "CD", smFields(COMMENTPROPOSAL), slACT1Info
    End If
    mCSIXMLData "CT", "Comment", ""
    mCSIXMLData "OT", "Periods", ""
    mWriteXmlPeriods blDemo, blDaily, ilClfIndex
    mCSIXMLData "CT", "Periods", ""
    mCSIXMLData "CT", "AvailLineWithDetailedPeriods", ""
End Sub
Private Sub mWriteXmlPeriods(blDemo As Boolean, blDaily As Boolean, ilClfIndex As Integer)

    Dim slLine As String
    Dim c As Integer
    Dim slWeeks() As String
    Dim slDates() As String
    Dim ilCounter As Integer
    Dim slDelimiter As String
   '8/28/18: Handle non-monday sunday flights
    Dim slStartDate As String
    Dim slEndDate As String
    '8/28/18: Handle non-monday sunday flights
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilDay As Integer
    Dim slPeriodName As String
    '6778
    Dim slRateToUse As String
    '7112
    Dim llCff As Long
    Dim llSpotsStart As Long
    Dim llSpotsEnd As Long
    Dim llDate As Long
    '7112 only write out for spots between these 2 values
    llSpotsStart = 0
    llSpotsEnd = 0
    llCff = tgClfCntr(ilClfIndex).lFirstCff
    Do While llCff <> -1
        '7596
        'If tgCffCntr(llCff).CffRec.iSpotsWk > 0 Then
        If tgCffCntr(llCff).CffRec.iSpotsWk > 0 Or blDaily Then
            If llSpotsStart = 0 Then
                llSpotsStart = tgCffCntr(llCff).lStartDate
            End If
            llSpotsEnd = tgCffCntr(llCff).lEndDate
        End If
        llCff = tgCffCntr(llCff).lNextCff
    Loop
    slRateToUse = ""
    slDelimiter = Chr(9)
    If blDaily Then
        slPeriodName = "DayDetailedPeriod"
    Else
        slPeriodName = "DetailedPeriod"
    End If
    slWeeks = Split(smValues(PERIODSPOTSPERWEEK), slDelimiter)
    slDates = Split(smClfDates, slDelimiter)
    'sldates is 1 larger than it actually contains
    For ilCounter = 0 To UBound(slDates) - 1
        '7112
        llDate = gDateValue(slDates(ilCounter))
        If llDate >= llSpotsStart And llDate <= llSpotsEnd Then
            If slWeeks(ilCounter) > 0 Then
                slLine = ""
                'smValues(PERIODSTARTDATE) = slDates(ilCounter)
                '8/28/18: Handle non-monday sunday flights
                slStartDate = slDates(ilCounter)
                llStartDate = gDateValue(slStartDate)
                slEndDate = DateAdd("d", 6, slStartDate)
                llEndDate = llStartDate + 6
                If tgExportAgf.sXMLDates = "A" Then
                    'Find flight that bracket start date
                    llCff = tgClfCntr(ilClfIndex).lFirstCff
                    Do While llCff <> -1
                        '7596
                        'If tgCffCntr(llCff).CffRec.iSpotsWk > 0 Then
                        If tgCffCntr(llCff).CffRec.iSpotsWk > 0 Or blDaily Then
                            If llStartDate >= tgCffCntr(llCff).lStartDate And llStartDate <= tgCffCntr(llCff).lEndDate Then
                                'gUnpackDateLong tgCffCntr(llCff).CffRec.iStartDate(0), tgCffCntr(llCff).CffRec.iStartDate(1), llStartDate
                                For ilDay = gWeekDayLong(tgCffCntr(llCff).lStartDate) To 6 Step 1
                                    If tgCffCntr(llCff).CffRec.iDay(ilDay) > 0 Then
                                        Exit For
                                    End If
                                    llStartDate = llStartDate + 1
                                Next ilDay
                                'gUnpackDateLong tgCffCntr(llCff).CffRec.iEndDate(0), tgCffCntr(llCff).CffRec.iEndDate(1), llEndDate
                                For ilDay = gWeekDayLong(tgCffCntr(llCff).lEndDate) To 0 Step -1
                                    If tgCffCntr(llCff).CffRec.iDay(ilDay) > 0 Then
                                        Exit For
                                    End If
                                    llEndDate = llEndDate - 1
                                Next ilDay
                                Exit Do
                            End If
                        End If
                        llCff = tgCffCntr(llCff).lNextCff
                    Loop
                    slStartDate = Format(llStartDate, XMLDATE)
                    slEndDate = Format(llEndDate, XMLDATE)
                End If
                smValues(PERIODSTARTDATE) = slStartDate 'slDates(ilCounter)

                'slEndDate = DateAdd("d", 6, smValues(PERIODSTARTDATE))
                smValues(PERIODENDDATE) = Format(slEndDate, XMLDATE)
                For c = PERIODSTARTDATE To PERIODENDDATE
                    slLine = slLine & mXmlAttribute(c) & " "
                Next c
                mCSIXMLData "OT", slPeriodName, slLine
                mCSIXMLData "CD", smFields(PERIODRATE), smValues(PERIODRATE)
                '6778 get rate here to use for skipped weeks
                slRateToUse = smValues(PERIODRATE)
'                '7318
'                If InStr(slRateToUse, ".") = 1 Then
'                    slRateToUse = "0" & slRateToUse
'                End If
                If Not blDaily Then
                    mCSIXMLData "CD", smFields(PERIODSPOTSPERWEEK), slWeeks(ilCounter)
                Else
                    mCSIXMLData "OT", "SpotsByDay", ""
                    mDailyDays ilClfIndex, smValues(PERIODSTARTDATE)
                    mCSIXMLData "CT", "SpotsByDay", ""
                End If
                If blDemo Then
                '5908
    ''                'ttp5230
                    mWriteXmlDemosPerPeriod ilClfIndex
    '                mCSIXMLData "OT", "DemoValues", ""
    '                csiXMLDataEx "<DemoValue demoRef='" & smValues(DEMOAVGAUDID) & "'>" & smValues(PERIODDEMOAUD) & "</DemoValue>"
    '                csiXMLDataEx "<DemoValue demoRef='" & smValues(DEMOID) & "'>" & smValues(PERIODDEMOVALUE) & "</DemoValue>"
    '                mCSIXMLData "CT", "DemoValues", ""
                End If
                mCSIXMLData "CT", slPeriodName, ""
            Else
    '        '6525
               '6778 new: if rate is 0, then this isn't a skipped week...it's before a week has started.  Skip it.
                If Len(slRateToUse) > 0 Then
                    slEndDate = DateAdd("d", 6, slDates(ilCounter))
                    mCSIXMLData "OT", slPeriodName, "startDate='" & slDates(ilCounter) & "' endDate='" & Format(slEndDate, XMLDATE) & "'"
                    mCSIXMLData "CD", "Rate", slRateToUse
                    mCSIXMLData "CT", slPeriodName, ""
               End If
               ' mCSIXMLData "CD", "Rate", "0.00"
              '  mCSIXMLData "CD", "SpotsPerWeek", "0"
            End If
        End If '7112
    Next ilCounter
End Sub
Private Sub mWriteXmlDemosPerPeriod(ilClfIndex As Integer)
'    Dim c As Integer
    Dim slAvAud As String

    mCSIXMLData "OT", "DemoValues", ""
'    primary demo here
'    csiXMLDataEx "<DemoValue demoRef='" & smValues(DEMOAVGAUDID) & "'>" & smValues(PERIODDEMOAUD) & "</DemoValue>"
'    csiXMLDataEx "<DemoValue demoRef='" & smValues(DEMOID) & "'>" & smValues(PERIODDEMOVALUE) & "</DemoValue>"
    If Not rsDemoValues Is Nothing Then
        rsDemoValues.Filter = " LineIndex = " & ilClfIndex
        'added 5890. And this was not being done before
        Do While Not rsDemoValues.EOF
            If tgSpf.sSAudData = "H" Then
                slAvAud = gLongToStrDec(rsDemoValues!audience, 1)
            ElseIf tgSpf.sSAudData = "N" Then
                slAvAud = gLongToStrDec(rsDemoValues!audience, 2)
            ElseIf tgSpf.sSAudData = "U" Then
                slAvAud = gLongToStrDec(rsDemoValues!audience, 3)
            Else
                slAvAud = gLongToStrDec(rsDemoValues!audience, 0)
            End If
'            csiXMLDataEx "<DemoValue demoRef='DM" & CStr(rsDemoValues!DEMOINDEX) & "'>" & gLongToStrDec(rsDemoValues!rating, 1) & "</DemoValue>"
 '           csiXMLDataEx "<DemoValue demoRef='DM" & CStr(rsDemoValues!DEMOINDEX + XMLIMPRESSIONS) & "'>" & gLongToStrDec(rsDemoValues!audience, 1) & "</DemoValue>"
            csiXMLDataEx "<DemoValue demoRef='DM" & CStr(rsDemoValues!DEMOINDEX) & "'>" & Trim$(rsDemoValues!rating) & "</DemoValue>"
            csiXMLDataEx "<DemoValue demoRef='DM" & CStr(rsDemoValues!DEMOINDEX + XMLIMPRESSIONS) & "'>" & slAvAud & "</DemoValue>"
            rsDemoValues.MoveNext
        Loop
    End If
    mCSIXMLData "CT", "DemoValues", ""
End Sub
Private Sub mDailyDays(ilClfIndex As Integer, slDate As String)
    Dim ilWkCff As Integer
    ReDim ilOVDays(0 To 6) As Integer
    ReDim ilWkDays(0 To 6) As Integer
    ReDim ilDailySpots(0 To 6) As Integer
    Dim llWkSpots As Long
    Dim slWDFlag As String
    Dim llDate As Long
    Dim ilClfRowNo As Integer
 '   Dim slClf As String
    '2/17/12 Dan  Simplified.  Just need ilClfIndex, so pass that. Then add one for mGetFlightSpotDays, which wants the clfLine, not index
    '7/26 I finally got this figured out.  ilIndex is index for smClfInfo.  slClf is the CLFLine in the array(tgClfCntr(slclf)).
    ' But mGetFlightSpotsDays wants the CLFLine, which is the array + 1.  It then subtracts 1 to use tgClfCntr.
    'Finally there is the row no.  This isn't used here, but sometimes a ClfLine and the row # displayed will be different. The row # is
    'tgClfCntr(slClf).ClfRec.iLine
    llDate = gDateValue(slDate)
  '  slClf = Mid(smClfInfo(ilIndex), 1, 1)
  '  ilLnRowNo = slClf + 1
    ilClfRowNo = ilClfIndex + 1
    'D.S. 5/30/18 conditional compile
    #If programmatic <> 1 Then
        ilWkCff = mGetFlightSpotsDays(ilClfRowNo, llDate, ilOVDays(), llWkSpots, ilWkDays(), slWDFlag, ilDailySpots())
    #Else
        ilWkCff = mPBGetFlightSpotsDays(ilClfRowNo, llDate, ilOVDays(), llWkSpots, ilWkDays(), slWDFlag, ilDailySpots())
        'Set values if it's a daily buy
        'Use prog buy CFF
    #End If
    mCSIXMLData "CD", "Monday", CStr(ilDailySpots(0))
    mCSIXMLData "CD", "Tuesday", CStr(ilDailySpots(1))
    mCSIXMLData "CD", "Wednesday", CStr(ilDailySpots(2))
    mCSIXMLData "CD", "Thursday", CStr(ilDailySpots(3))
    mCSIXMLData "CD", "Friday", CStr(ilDailySpots(4))
    mCSIXMLData "CD", "Saturday", CStr(ilDailySpots(5))
    mCSIXMLData "CD", "Sunday", CStr(ilDailySpots(6))
End Sub
Private Sub mWriteXmlBuyLineMarketron(blDaily As Boolean, blDemo As Boolean, ilLineIndex As Integer)
    'ilLineIndex is index for smClfInfo.  ilClfIndex is the CLFLine in the array(tgClfCntr(ilClfIndex)).
    ' The row # is tgClfCntr(ilClfIndex).ClfRec.iLine
    Dim ilClfIndex As Integer
    Dim c As Integer
    Dim llCff As Long
    Dim llSpotsStart As Long
    Dim llSpotsEnd As Long
    Dim slTradeFlag As String
    Dim slEndTime As String
    Dim ilPos As Integer
    Dim ilTime As Integer
    
    ilClfIndex = mGetClfLineNumber(smClfInfo(ilLineIndex))
    llCff = tgClfCntr(ilClfIndex).lFirstCff
    Do While llCff <> -1
        If tgCffCntr(llCff).CffRec.iSpotsWk > 0 Or blDaily Then
            If llSpotsStart = 0 Then
                llSpotsStart = tgCffCntr(llCff).lStartDate
            End If
            llSpotsEnd = tgCffCntr(llCff).lEndDate
        End If
        llCff = tgCffCntr(llCff).lNextCff
    Loop
    '9p to 3am?  other shows 21 to 27...this way is 21 and 3!
    ilPos = InStr(smValues(DAYTIMEENDTIME), ":")
    If ilPos > 0 Then
        slEndTime = Mid(smValues(DAYTIMEENDTIME), 1, ilPos - 1)
        If IsNumeric(slEndTime) Then
            ilTime = Val(slEndTime)
            If ilTime > 24 Then
                ilTime = ilTime - 24
                smValues(DAYTIMEENDTIME) = ilTime & Mid(smValues(DAYTIMEENDTIME), ilPos)
            End If
        End If
    End If
    If smValues(DAYTIMEENDTIME) = "24:00" Then
        smValues(DAYTIMEENDTIME) = "00:00"
    End If
    If smValues(DAYTIMESTARTTIME) = "24:00" Then
        smValues(DAYTIMESTARTTIME) = "00:00"
    End If
    If smValues(DAYTIMEENDTIME) = "24:00" Then
        smValues(DAYTIMEENDTIME) = "00:00"
    End If
    slTradeFlag = "Cash"
    If tgChfCntr.iPctTrade = 100 Then
        slTradeFlag = "Trade"
    End If
    mCSIXMLData "OT", "BuyLine", "tradeFlag=""" & slTradeFlag & """"
    mCSIXMLData "CD", "BuyLineNumber", CStr(tgClfCntr(ilClfIndex).ClfRec.iLine)
    mCSIXMLData "CD", smFields(DAYSPOTLENGTH), smValues(DAYSPOTLENGTH) 'Format(smValues(DAYSPOTLENGTH), "ss")
    mCSIXMLData "CD", "StartDate", Format(llSpotsStart, "yyyy-mm-dd")
    mCSIXMLData "CD", "EndDate", Format(llSpotsEnd, "yyyy-mm-dd")
    mCSIXMLData "OT", "Daypart", ""
    mCSIXMLData "CD", smFields(DAYNAME), smValues(DAYNAME)
    mCSIXMLData "CT", "Daypart", ""
    mCSIXMLData "OT", "DayTime", ""
    mCSIXMLData "CD", smFields(DAYTIMESTARTTIME), smValues(DAYTIMESTARTTIME)
    mCSIXMLData "CD", smFields(DAYTIMEENDTIME), smValues(DAYTIMEENDTIME)
    mCSIXMLData "OT", "rdo:Days", ""
    For c = XMLMON To XMLSUN
        mCSIXMLData "CD", smFields(c), "Y" 'smValues(c)
    Next c
    mCSIXMLData "CT", "rdo:Days", ""
    mCSIXMLData "CT", "DayTime", ""
    mWriteXmlPeriodsMarketron blDaily, ilClfIndex, llSpotsStart, llSpotsEnd
    If Len(smValues(COMMENTPROPOSAL)) > 0 Then
        mCSIXMLData "CA", "Comment", "owner=""Rep"""
        mCSIXMLData "CD", smFields(COMMENTPROPOSAL), smValues(COMMENTPROPOSAL)
        mCSIXMLData "CT", "Comment", ""
    End If
    mCSIXMLData "CT", "BuyLine", ""
End Sub
Private Sub mWriteXmlPeriodsMarketron(blDaily As Boolean, ilClfIndex As Integer, llSpotsStart As Long, llSpotsEnd As Long)

    Dim c As Integer
    Dim slWeeks() As String
    Dim slDates() As String
    Dim ilCounter As Integer
    Dim slDelimiter As String
    Dim slEndDate As String
    Dim slPeriodName As String
    '6778
    Dim slRateToUse As String
    Dim llDate As Long
    slRateToUse = ""
    slDelimiter = Chr(9)
    If blDaily Then
        slPeriodName = "DailyBuyPeriod"
    Else
        mCSIXMLData "OT", "RotationDays", ""
        For c = XMLMON To XMLSUN
            mCSIXMLData "CD", smFields(c), smValues(c)
        Next c
        mCSIXMLData "CT", "RotationDays", ""
        slPeriodName = "WeeklyBuyPeriod"
    End If
    slWeeks = Split(smValues(PERIODSPOTSPERWEEK), slDelimiter)
    slDates = Split(smClfDates, slDelimiter)
    'sldates is 1 larger than it actually contains
    For ilCounter = 0 To UBound(slDates) - 1
        llDate = gDateValue(slDates(ilCounter))
        If llDate >= llSpotsStart And llDate <= llSpotsEnd Then
            smValues(PERIODSTARTDATE) = slDates(ilCounter)
            slEndDate = DateAdd("d", 6, smValues(PERIODSTARTDATE))
            smValues(PERIODENDDATE) = Format(slEndDate, XMLDATE)
            If slWeeks(ilCounter) > 0 Then
                mCSIXMLData "OT", slPeriodName, ""
                mCSIXMLData "CD", smFields(PERIODSTARTDATE), smValues(PERIODSTARTDATE)
                mCSIXMLData "CD", smFields(PERIODENDDATE), smValues(PERIODENDDATE)
                mCSIXMLData "CD", smFields(PERIODRATE), smValues(PERIODRATE)
                '6778 get rate here to use for skipped weeks
                slRateToUse = smValues(PERIODRATE)
                If Not blDaily Then
                    mCSIXMLData "CD", smFields(PERIODSPOTSPERWEEK), slWeeks(ilCounter)
                Else
                    mCSIXMLData "OT", "DailySpots", ""
                    mDailyDays ilClfIndex, smValues(PERIODSTARTDATE)
                    mCSIXMLData "CT", "DailySpots", ""
                End If
                mCSIXMLData "CT", slPeriodName, ""
            Else
            ' marketron: don't write out hiatus weeks!
'               '6778 new: if rate is 0, then this isn't a skipped week...it's before a week has started.  Skip it.
'                If Len(slRateToUse) > 0 Then
'                    mCSIXMLData "OT", slPeriodName, ""
'                    mCSIXMLData "CD", smFields(PERIODSTARTDATE), smValues(PERIODSTARTDATE)
'                    mCSIXMLData "CD", smFields(PERIODENDDATE), smValues(PERIODENDDATE)
'                    mCSIXMLData "CD", smFields(PERIODRATE), slRateToUse
'                    If Not blDaily Then
'                        mCSIXMLData "CD", smFields(PERIODSPOTSPERWEEK), slWeeks(ilCounter)
'                    Else
'                        mCSIXMLData "OT", "DailySpots", ""
'                        mDailyDays ilClfIndex, smValues(PERIODSTARTDATE)
'                        mCSIXMLData "CT", "DailySpots", ""
'                    End If
'                    mCSIXMLData "CT", slPeriodName, ""
'               End If
            End If
        End If '7112
    Next ilCounter
End Sub
Public Sub mPrintOutFileStoreXml(hlTo As Integer, slRecord As String, Optional blStoreXml As Boolean)
    Dim ilLen As Integer
    
    If Not bmXmlOutput Then
        Print #hlTo, slRecord
    ElseIf blStoreXml Then
        'smclfinfo is 1 longer than really used.
        ilLen = UBound(smClfInfo)
        ReDim Preserve smClfInfo(ilLen + 1)
        smClfInfo(ilLen) = slRecord
    End If
End Sub
Private Sub mWriteXmlListStart()
    Dim c As Integer
    Dim slLine As String
    Dim ilCounter As Integer
    
    For c = AVAILID To AVAILISPACKAGE
        slLine = slLine & " " & mXmlAttribute(c)
    Next c
    mCSIXMLData "OT", "AvailList", slLine
    mCSIXMLData "CD", smFields(AVAILNAME), smValues(AVAILNAME)
    mCSIXMLData "OT", "OutletReferences", ""
    slLine = ""
    For c = OUTLETREF To OUTLETID
        slLine = slLine & mXmlAttribute(c) & " "
    Next c
    mCSIXMLData "OT", "OutletReference", slLine
    mCSIXMLData "CT", "OutletReference", ""
    mCSIXMLData "CT", "OutletReferences", ""
End Sub
Private Sub mWriteXmlStations()
    Dim slLine As String
    Dim ilCounter As Integer
        
    mCSIXMLData "OT", "Outlets", ""
    For ilCounter = RADIOCALLLETTERS To RADIOOUTLETID
        slLine = slLine & mXmlAttribute(ilCounter) & " "
    Next ilCounter
    mCSIXMLData "CA", "RadioStation", slLine
    mCSIXMLData "CT", "Outlets", ""
End Sub
Private Sub mWriteXmlSellerBuyer()
'required bcompany name, salesname, scompanyname, advname: if they don't exist, don't write out pertinent element: seller,buyer, advertiser
    Dim slLine As String
    Dim c As Integer
    Dim blWriteSeller As Boolean
    Dim blWriteBuyer As Boolean
    Dim blWriteAdv As Boolean
    
    If Len(smValues(SCOMPANYNAME)) > 0 And Len(smValues(SALESNAME)) > 0 Then
        blWriteSeller = True
    End If
    If Len(smValues(BCOMPANYNAME)) > 0 Then
        blWriteBuyer = True
    End If
    If Len(smValues(ADVNAME)) > 0 Then
        blWriteAdv = True
    End If
    For c = PROPOSALID To PENDDATE
        slLine = slLine & " " & mXmlAttribute(c)
    Next c
    mCSIXMLData "OT", "Proposal", slLine
    If blWriteSeller Then
        mCSIXMLData "OT", "Seller", mXmlAttribute(SCOMPANYNAME)
        If Len(Trim$(smValues(SOFFICENAME))) > 0 Then
            mCSIXMLData "CD", smFields(SOFFICENAME), smValues(SOFFICENAME)
        End If
        mCSIXMLData "OT", "Salesperson", mXmlAttribute(SALESNAME)
        mCSIXMLData "CT", "Salesperson", ""
        mCSIXMLData "CT", "Seller", ""
    End If
    If blWriteBuyer Then
        mCSIXMLData "OT", "Buyer", mXmlAttribute(BCOMPANYNAME)
        If Len(smValues(BOFFICENAME)) > 0 Then
            mCSIXMLData "CD", smFields(BOFFICENAME), smValues(BOFFICENAME)
        End If
        If Len(Trim$(smValues(BNAME))) > 0 Then
            mCSIXMLData "CD", smFields(BNAME), smValues(BNAME)
        End If
        mCSIXMLData "CT", "Buyer", ""
    End If
    If blWriteAdv Then
        mCSIXMLData "OT", "Advertiser", mXmlAttribute(ADVNAME)
        If Len(Trim$(smValues(ProductName))) > 0 Then
            mCSIXMLData "CA", "Product", mXmlAttribute(ProductName)
        End If
        mCSIXMLData "CT", "Advertiser", ""
    End If
End Sub
Private Sub mWriteXmlSellerBuyerStationMarketron()
'required bcompany name, salesname, scompanyname, advname: if they don't exist, don't write out pertinent element: seller,buyer, advertiser
    Dim slLine As String
    Dim ilPos As Integer
    Dim blWriteSeller As Boolean
    Dim blWriteBuyer As Boolean
    Dim blWriteAdv As Boolean
    Dim slFirstName As String
    Dim slLastName As String
    
    If Len(smValues(SALESNAME)) > 0 Then 'And Len(smValues(SALESNAME)) > 0 Then
        blWriteSeller = True
    End If
    If Len(smValues(BNAME)) > 0 Then
        ilPos = InStr(smValues(BNAME), " ")
        If ilPos > 0 Then
            slFirstName = Mid(smValues(BNAME), 1, ilPos - 1)
            slLastName = Mid(smValues(BNAME), ilPos + 1)
            If Len(slFirstName) > 0 And Len(slLastName) > 0 Then
                blWriteBuyer = True
            End If
        End If
    End If
    If Len(smValues(ADVNAME)) > 0 Then
        blWriteAdv = True
    End If
    'don't have email, so don't do buyer or seller
    blWriteBuyer = False
    blWriteSeller = False
    mCSIXMLData "OT", "rdo:Header", ""
    mCSIXMLData "OT", "rdo:BuyingCompany", "companyId=""CO01""" 'mXmlAttribute(BCOMPANYNAME)
    mCSIXMLData "CD", smFields(BCOMPANYNAME), smValues(BCOMPANYNAME)
    mCSIXMLData "OT", "rdo:Office", "" ' "officeId=""OFtODO"""
    mCSIXMLData "CD", smFields(BOFFICENAME), smValues(BOFFICENAME)
    mCSIXMLData "CT", "rdo:Office", ""
    mCSIXMLData "CT", "rdo:BuyingCompany", ""
    If blWriteBuyer Then
        mCSIXMLData "OT", "rdo:Buyer", "" '"contactId=""PStODO"""
        mCSIXMLData "CD", "rdo:FirstName", slFirstName
        mCSIXMLData "CD", "rdo:LastName", slLastName
        mCSIXMLData "CD", "rdo:Email", "MUSTDO@mustdo.com"
        mCSIXMLData "CT", "rdo:Buyer", ""
    End If
        mCSIXMLData "OT", "rdo:SellingCompany", "companyId='CO02'" 'mXmlAttribute(SCOMPANYNAME)
        mCSIXMLData "CD", smFields(SCOMPANYNAME), smValues(SCOMPANYNAME)
        mCSIXMLData "OT", "rdo:Office", "" ' "officeId=""OFtODO"""
        mCSIXMLData "CD", smFields(SOFFICENAME), smValues(SOFFICENAME)
        mCSIXMLData "CT", "rdo:Office", ""
        mCSIXMLData "CT", "rdo:SellingCompany", ""
    If blWriteSeller Then
        ilPos = InStr(smValues(SALESNAME), " ")
        If ilPos > 0 Then
            slFirstName = Mid(smValues(SALESNAME), 1, ilPos - 1)
            slLastName = Mid(smValues(SALESNAME), ilPos + 1)
            If Len(slFirstName) > 0 And Len(slLastName) > 0 Then
                mCSIXMLData "OT", "rdo:Seller", "" ' mXmlAttribute(SALESNAME)
                mCSIXMLData "CD", "rdo:FirstName", slFirstName
                mCSIXMLData "CD", "rdo:LastName", slLastName
                mCSIXMLData "CD", "rdo:Email", "MUSTDO@thistoo.com"
                mCSIXMLData "CT", "rdo:Seller", ""
            End If
        End If

    End If
    mCSIXMLData "CD", "rdo:UniqueAgencyOrderId", smValues(SELLERREFERENCE)
    mCSIXMLData "OT", "rdo:OrderReferences", ""
    csiXMLDataEx "<rdo:SystemCode owner=""Rep"">" & smValues(SELLERREFERENCE) & "</rdo:SystemCode>"
  '  mCSIXMLData "CD", smFields(SELLERREFERENCE), smValues(SELLERREFERENCE)
    mCSIXMLData "CT", "rdo:OrderReferences", ""
    If blWriteAdv Then
        mCSIXMLData "OT", "rdo:Advertiser", ""
        mCSIXMLData "CD", smFields(ADVNAME), smValues(ADVNAME)
        mCSIXMLData "CT", "rdo:Advertiser", ""
    End If
    If Len(Trim$(smValues(ProductName))) > 0 Then
        mCSIXMLData "OT", "rdo:Product", ""
        mCSIXMLData "CD", smFields(ProductName), smValues(ProductName)
        mCSIXMLData "CT", "rdo:Product", ""
    End If
    slLine = Trim$(tgChfCntr.sAgyEstNo) & Trim$(tgChfCntr.sTitle)
    If Len(slLine) = 0 Then
        'default
        slLine = "Unknown"
    End If
    mCSIXMLData "OT", "rdo:Estimate", ""
    csiXMLDataEx "<rdo:SystemCode owner=""Rep"">" & slLine & "</rdo:SystemCode>"
    mCsiXMLDataExAlternate "<rdo:SystemCode owner=""Rep"">" & slLine & "</rdo:SystemCode>"
    mCSIXMLData "CT", "rdo:Estimate", ""
    mCSIXMLData "OT", "rdo:Outlet", ""
    mCSIXMLData "CD", smFields(RADIOCALLLETTERS), smValues(RADIOCALLLETTERS)
    mCSIXMLData "CD", smFields(RADIOBAND), smValues(RADIOBAND)
    csiXMLDataEx "<rdo:SystemCode owner=""Rep"">" & smValues(RADIOCALLLETTERS) & "-" & smValues(RADIOBAND) & "</rdo:SystemCode>"
    mCsiXMLDataExAlternate "<rdo:SystemCode owner=""Rep"">" & smValues(RADIOCALLLETTERS) & "-" & smValues(RADIOBAND) & "</rdo:SystemCode>"
   ' mCSIXMLData "CD", "rdo:SystemCode", smValues(RADIOCALLLETTERS) & "-" & smValues(RADIOBAND)
  '  mCSIXMLData "CT", "rdo:SystemCode"
    mCSIXMLData "CT", "rdo:Outlet", ""
    
    
    mCSIXMLData "CT", "rdo:Header", ""
End Sub
Private Sub mWriteXmlOrderStartMarketron()
    mCSIXMLData "OT", "Order", mXmlAttribute(PVERSION) & " unwired=""Y"""
    mCSIXMLData "CD", "NationalLocal", "national"
    mCSIXMLData "CD", "OrderStatus", "New"
    mCSIXMLData "CD", smFields(PSTARTDATE), smValues(PSTARTDATE)
    mCSIXMLData "CD", smFields(PENDDATE), smValues(PENDDATE)
   ' mCSIXMLData "CD", smFields(PWEEKSTARTDAY), smValues(PWEEKSTARTDAY)
End Sub
Private Function mXmlAttribute(ilIndex As Integer) As String
    mXmlAttribute = smFields(ilIndex) & "='" & gXMLNameFilter(smValues(ilIndex)) & "'"
End Function



Public Sub mWriteXmlHeader()
    Dim slAction As String
    Dim slString As String
    
    '7580
    'If Val(smValues(PVERSION)) > 1 Then
    If tgChfCntr.iExtRevNo > 0 Then
        slAction = "Revised"
    Else
        slAction = "New"
    End If
    slString = "xmlns=""http://www.AAAA.org/schemas/spotTVCableProposal"" " & vbCrLf
    slString = slString & "xmlns:tvb = ""http://www.AAAA.org/schemas/spotTV"" " & vbCrLf
    slString = slString & "xmlns:tvb-tp = ""http://www.AAAA.org/schemas/TVBGeneralTypes"" " & vbCrLf
    slString = slString & "xmlns:xsi = ""http://www.w3.org/2001/XMLSchema-instance"" " & vbCrLf
    slString = slString & "xsi:schemaLocation = ""http://www.AAAA.org/schemas/spotTVCableProposal spotTVCableProposal-0.3.0.5R.xsd"" "
    csiXMLDataEx "<?xml version=""1.0"" encoding=""UTF-8""?>"
    mCSIXMLData "OT", "AAAA-Message", slString
    mCSIXMLData "OT", "AAAA-Values", ""
    mCSIXMLData "CD", "SchemaName", "SpotTVCableProposal"
    mCSIXMLData "CD", "SchemaVersion", "0.3.0.5R"
    mCSIXMLData "CD", "Media", "SpotRadio"
    mCSIXMLData "CD", "BusinessObject", "Proposal"
    mCSIXMLData "CD", "Action", slAction
    mCSIXMLData "CD", "UniqueMessageID", "3F2504E0-4F89-11D3-9A0C-0305E82C3301"
    mCSIXMLData "CT", "AAAA-Values", ""
End Sub
Public Sub mWriteXmlHeaderStationMarketron()
    Dim slString As String
    
    slString = "xmlns=""http://www.rabinsite.org/ediData/spotRadioOrder"" " & vbCrLf
    slString = slString & "xmlns:tp = ""http://www.rabinsite.org/ediData/Types"" " & vbCrLf
    slString = slString & "xmlns:rdo = ""http://www.rabinsite.org/ediData/spotRadio"" " & vbCrLf
    slString = slString & "xsi:schemaLocation = ""http://www.rabinsite.org/ediData/spotRadioOrder spotRadioOrder-0.8.0.4_DDS.xsd"" " & vbCrLf
    slString = slString & "xmlns:xsi = ""http://www.w3.org/2001/XMLSchema-instance"" "
    csiXMLDataEx "<?xml version=""1.0"" encoding=""UTF-8""?>"
    mCsiXMLDataExAlternate "<?xml version=""1.0"" encoding=""UTF-8""?>"
    mCSIXMLData "OT", "AAAA-Message", slString
    mCSIXMLData "OT", "AAAA-Values", ""
    mCSIXMLData "CD", "SchemaName", "SpotRadioOrder"
    mCSIXMLData "CD", "SchemaVersion", "0.8.0.4"
    mCSIXMLData "CD", "Media", "SpotRadio"
    mCSIXMLData "CD", "TransactionType", "Order"
    mCSIXMLData "CT", "AAAA-Values", ""
End Sub
Private Sub mCloseXmlHeader()
   ' mCSIXMLData "CT", "AvailList", ""
    mCSIXMLData "CT", "Proposal", ""
    mCSIXMLData "CT", "AAAA-Message", ""
End Sub
Private Sub mCloseXmlHeaderMarketron()
    mCSIXMLData "CT", "Order", ""
    mCSIXMLData "CT", "AAAA-Message", ""
End Sub
Private Sub mCSIXMLData(slInCommand As String, slInTag As String, slInData As String)
    '8032
    If Not ogEOrder Is Nothing Then
        slInCommand = UCase(slInCommand)
        If slInCommand = "CD" Then
            slInData = gXMLNameFilter(slInData)
        End If
        Select Case slInCommand
            Case "OT"
                If Len(slInData) > 0 Then
                    If Not ogEOrder.AddToBody("<" & slInTag & " " & slInData & ">") Then
                    
                    End If
                Else
                    If Not ogEOrder.AddToBody("<" & slInTag & ">") Then
                    
                    End If
                End If
            Case "CT"
                If Not ogEOrder.AddToBody("</" & slInTag & ">") Then
                
                End If
            Case "CD"
                If Not ogEOrder.AddToBody("<" & slInTag & ">" & slInData & "</" & slInTag & ">") Then
                
                End If
            'csiXMLData "CA", "Delete TransmissionID='" & slTransmissionID & "'", ""
            Case "CA"
                If Not ogEOrder.AddToBody("<" & slInTag & "/>") Then
                
                End If
        End Select
    End If
    'cd, ca, ot, ct
    slInCommand = UCase(slInCommand)
    If slInCommand = "CD" And ogEOrder Is Nothing Then
        slInData = gXMLNameFilter(slInData)
    End If
    csiXMLData slInCommand, slInTag, slInData
    DoEvents
End Sub
Private Sub mXmlSetFields()
Dim c As Integer

'8032 moved
'For c = 0 To UBound(smValues)
'    smValues(c) = ""
'Next c
'these fields are required to have at least one character
'smValues(PROPOSALNAME) = XMLUNKNOWN
'smValues(AVAILNAME) = XMLUNKNOWN
'smValues(DAYAVAILNAME) = XMLUNKNOWN
'constants
'radioOutletId and outletRef must match
smValues(RADIOOUTLETID) = "OUT0"
smValues(OUTLETREF) = "OUT0"
smValues(AVAILID) = "AL0"
smValues(OUTLETID) = "OUL0"
smValues(AVAILOUTLET) = "OUL0"
smValues(PWEEKSTARTDAY) = "Mo"
'D.S. 06/08/18 added conditional compile below
#If programmatic <> 1 Then
    smValues(AVAILISPACKAGE) = "N"
#Else
    smValues(AVAILISPACKAGE) = "Y"
#End If
smValues(DEMOID) = "DM0"
smValues(DEMOAVGAUDID) = "DM5"
'tag names
smFields(PROPOSALID) = "uniqueIdentifier"
smFields(BNAME) = "BuyerName"
smFields(BOFFICENAME) = "OfficeName"
smFields(BCOMPANYNAME) = "buyingCompanyName"
smFields(SALESNAME) = "name"
smFields(SOFFICENAME) = "OfficeName"
smFields(SCOMPANYNAME) = "companyName"
smFields(PENDDATE) = "endDate"
smFields(PSTARTDATE) = "startDate"
smFields(PWEEKSTARTDAY) = "weekStartDay"
smFields(PSENDDATETIME) = "sendDateTime"
smFields(PVERSION) = "version"
smFields(OUTLETID) = "outletForListId"
smFields(OUTLETREF) = "outletFromProposalRef"
smFields(AVAILNAME) = "Name"
smFields(AVAILISPACKAGE) = "isPackage"
smFields(AVAILENDDATE) = "endDate"
smFields(AVAILSTARTDATE) = "startDate"
smFields(AVAILID) = "identifier"
smFields(RADIOOUTLETID) = "outletId"
smFields(RADIOBAND) = "band"
smFields(RADIOCALLLETTERS) = "callLetters"
smFields(SELLERREFERENCE) = "SellerReference"
smFields(PROPOSALNAME) = "Name"
smFields(ProductName) = "name"
smFields(ADVNAME) = "name"
smFields(COMMENTPROPOSAL) = "tvb-tp:CommentLine"
smFields(DEMOID) = "DemoId"
smFields(DEMOAVGAUDID) = "DemoId"
smFields(DEMOREF) = "demoRef"
smFields(DEMOTYPE) = "tvb:DemoType"
smFields(DEMOGROUP) = "tvb:Group"
smFields(DEMOAGEFROM) = "tvb:AgeFrom"
smFields(DEMOAGETO) = "tvb:AgeTo"
smFields(DAYTIMESTARTTIME) = "StartTime"
smFields(DAYTIMEENDTIME) = "EndTime"
smFields(DAYNAME) = "DaypartName"
smFields(DAYAVAILNAME) = "AvailName"
smFields(DAYSPOTLENGTH) = "SpotLength"
smFields(COMMENTLINE) = "tvb-tp:CommentLine"
smFields(PERIODSTARTDATE) = "startDate"
smFields(PERIODENDDATE) = "endDate"
smFields(PERIODRATE) = "Rate"
smFields(PERIODSPOTSPERWEEK) = "SpotsPerWeek"
smFields(PERIODDEMOVALUE) = "DemoValue"
smFields(PERIODDEMOREF) = "demoRef"
smFields(AVAILOUTLET) = "outletFromListRef"
smFields(XMLMON) = "tvb-tp:Monday"
smFields(XMLTUE) = "tvb-tp:Tuesday"
smFields(XMLWED) = "tvb-tp:Wednesday"
smFields(XMLTHU) = "tvb-tp:Thursday"
smFields(XMLFRI) = "tvb-tp:Friday"
smFields(XMLSAT) = "tvb-tp:Saturday"
smFields(XMLSUN) = "tvb-tp:Sunday"
End Sub
Private Sub mXmlSetFieldsAgencyMarketron()

smFields(PROPOSALID) = "uniqueIdentifier"
smFields(BNAME) = "rdo:Name"
smFields(BOFFICENAME) = "rdo:Name"
smFields(BCOMPANYNAME) = "rdo:Name"
smFields(SALESNAME) = "rdo:Name"
smFields(SOFFICENAME) = "rdo:Name"
smFields(SCOMPANYNAME) = "rdo:Name"
smFields(ProductName) = "rdo:Name"
smFields(ADVNAME) = "rdo:Name"
smFields(RADIOBAND) = "rdo:Band"
smFields(RADIOCALLLETTERS) = "rdo:CallLetters"
smFields(SALESNAME) = "rdo:Seller"
smFields(SELLERREFERENCE) = "rdo:SystemCode"
smFields(DAYNAME) = "rdo:DaypartName"
smFields(PVERSION) = "version"
smFields(DAYSPOTLENGTH) = "SpotLength"
smFields(PENDDATE) = "OrderEndDate"
smFields(PSTARTDATE) = "OrderStartDate"
'smFields(PWEEKSTARTDAY) = "WeekStartDay"
smFields(DAYTIMESTARTTIME) = "rdo:StartTime"
smFields(DAYTIMEENDTIME) = "rdo:EndTime"
smFields(XMLMON) = "tp:Monday"
smFields(XMLTUE) = "tp:Tuesday"
smFields(XMLWED) = "tp:Wednesday"
smFields(XMLTHU) = "tp:Thursday"
smFields(XMLFRI) = "tp:Friday"
smFields(XMLSAT) = "tp:Saturday"
smFields(XMLSUN) = "tp:Sunday"
smFields(PERIODSTARTDATE) = "StartDate"
smFields(PERIODENDDATE) = "EndDate"
smFields(PERIODRATE) = "Rate"
smFields(PERIODSPOTSPERWEEK) = "WeeklySpots"
smFields(COMMENTPROPOSAL) = "tp:CommentLine"
'these fields are required to have at least one character
'smValues(BNAME) = XMLUNKNOWN
'smavales(S








''these fields are required to have at least one character
'smValues(PROPOSALNAME) = XMLUNKNOWN
'smValues(DAYAVAILNAME) = XMLUNKNOWN
''constants
''radioOutletId and outletRef must match
'smValues(RADIOOUTLETID) = "OUT0"
'smValues(OUTLETREF) = "OUT0"
'smValues(AVAILID) = "AL0"
'smValues(OUTLETID) = "OUL0"
'smValues(AVAILOUTLET) = "OUL0"
'smValues(PWEEKSTARTDAY) = "Mo"
'smValues(AVAILISPACKAGE) = "N"
'smValues(DEMOID) = "DM0"
'smValues(DEMOAVGAUDID) = "DM5"
'tag names
'smFields(AVAILNAME) = "Name"
'
'smFields(PSENDDATETIME) = "sendDateTime"
'smFields(OUTLETID) = "outletForListId"
'smFields(OUTLETREF) = "outletFromProposalRef"
'smFields(AVAILISPACKAGE) = "isPackage"
'smFields(AVAILENDDATE) = "endDate"
'smFields(AVAILSTARTDATE) = "startDate"
'smFields(AVAILID) = "identifier"
'smFields(RADIOOUTLETID) = "outletId"
'
'smFields(PROPOSALNAME) = "Name"
'smFields(COMMENTPROPOSAL) = "tvb-tp:CommentLine"
'smFields(DEMOID) = "DemoId"
'smFields(DEMOAVGAUDID) = "DemoId"
'smFields(DEMOREF) = "demoRef"
'smFields(DEMOTYPE) = "tvb:DemoType"
'smFields(DEMOGROUP) = "tvb:Group"
'smFields(DEMOAGEFROM) = "tvb:AgeFrom"
'smFields(DEMOAGETO) = "tvb:AgeTo"
'smFields(COMMENTLINE) = "tvb-tp:CommentLine"
'smFields(PERIODDEMOVALUE) = "DemoValue"
'smFields(PERIODDEMOREF) = "demoRef"
'smFields(AVAILOUTLET) = "outletFromListRef"
End Sub

Public Function mTestVehiclesForSpotInsertion(blDontRegenerate As Boolean) As Boolean
    '8803 blDontRegenerate  came from save, so don't ask any questions.
    Dim blRet As Boolean
    Dim ilRowNo As Integer
    Dim slRegenerate As String
    Dim slRegenDate As String
    '7132
    Dim blIsFailStationVehicleTest As Boolean
    Dim blAtLeastOneStationVehicleFail As Boolean
    
    blAtLeastOneStationVehicleFail = False
    blIsFailStationVehicleTest = False
    '7806
    blRet = False
    'must be order
    If tgChfCntr.sStatus = "O" Or tgChfCntr.sStatus = "N" Then
        blRet = True
        'already sent? if not regenerating, should always be true
        If tgChfCntr.iXMLSentExtRevNo = tgChfCntr.iExtRevNo Then
            'for message later
            gUnpackDate tgChfCntr.iXMLSentDate(0), tgChfCntr.iXMLSentDate(1), slRegenDate
            ' 0 means never revised, so let's test date
            If tgChfCntr.iXMLSentExtRevNo = 0 Then
                If tgChfCntr.iXMLSentDate(1) <> 0 Then
                    blRet = False
                    slRegenerate = "Original Order"
                End If
            Else
                blRet = False
                slRegenerate = "Revision " & tgChfCntr.iXMLSentExtRevNo & " Order"
            End If
        End If
' not needed: if in here from save, the contract ext rev no will be different.  If from 'resend', why test?  just send it.
'if I decided to only send those that 'need' to go, then uncomment
'        'revision  maybe not
'        If Not blRet Then
'            For ilRowNo = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
'                'the line ext is not up to date
'                If tgClfCntr(ilRowNo).ClfRec.iXMLSentExtRevNo < tgChfCntr.iXMLSentExtRevNo Then
'                    If mIsSpotInsertionVehicle(tgClfCntr(ilRowNo).ClfRec.iVefCode) Then
'                        Select Case igStationXmlChoice
'                            Case STATIONXMLMARKETRON
'                                If bgStationXmlOrderAllowed(StationXmlIndexes.MARKETRONEO) Then
'                                    blRet = True
'                                    Exit For
'                                End If
'                            Case STATIONXMLWIDEORBIT
'                                If bgStationXmlOrderAllowed(StationXmlIndexes.WO) Then
'                                    blRet = True
'                                    Exit For
'                                End If
'                        End Select
'                    End If
'                End If
'            Next ilRowNo
'        End If
        If Not blRet And Not blDontRegenerate Then
            slRegenerate = slRegenerate & " has already been generated on " & slRegenDate & ", regenerate?"
            If MsgBox(slRegenerate, vbYesNo, "Regenerate?") = vbYes Then
                blRet = True
            End If
        End If
    ElseIf Not blDontRegenerate Then
        MsgBox "Not a saved order.  Cannot send as station xml.", vbOKOnly, "Not an Order"
    End If
    'passed previous tests
    If blRet Then
        blRet = False
        For ilRowNo = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
            '7132
            blRet = mIsSpotInsertionVehicle(tgClfCntr(ilRowNo).ClfRec.iVefCode, blIsFailStationVehicleTest)
            If blRet Then
                Exit For
            ElseIf blIsFailStationVehicleTest Then
                blAtLeastOneStationVehicleFail = True
            End If
        Next ilRowNo
        'if didn't find rep
        If Not blRet And Not blDontRegenerate Then
            '8132
            If blAtLeastOneStationVehicleFail Then
                gMsgBox "Station Xml-could not create files.  Some vehicles are marked as station xml, but they are not properly defined"
            Else
                MsgBox "This order contains no vehicles assigned as station xml", vbOKOnly, "No Station Xml Vehicles"
            End If
        End If
    End If
    mTestVehiclesForSpotInsertion = blRet
End Function
Private Function mIsSpotInsertionVehicle(ilVefCode As Integer, Optional blTestStationVehicleFailed As Boolean = False) As Boolean
    Dim slVefName As String
    Dim blRet As Boolean
    Dim ilVef As Integer
    Dim ilLimit As Integer
    
    blRet = False
    blTestStationVehicleFailed = False
    If ilVefCode > 0 Then
        '7959
        If gIsInsertionExport(ilVefCode) Then
            ilVef = gBinarySearchVef(ilVefCode)
            If ilVef >= LBound(tgMVef) Then
                If tgMVef(ilVef).sType = "R" Then
                    blRet = True
                '8132
                ElseIf bgStationAreVehicles Then
                    If gGetStationFromVehicle(ilVefCode) > 0 Then
                        blRet = True
                    Else
                        blTestStationVehicleFailed = True
                    End If
                End If
            End If
        End If
'        If gIsInsertionExport(ilVefCode) Then
'            blRet = True
'        Else
'            ilVef = gBinarySearchVef(ilVefCode)
'            If ilVef > 0 Then
'                If tgMVef(ilVef).sType = "R" Then
'                    blRet = True
'                End If
'            End If
'        End If
    End If
    mIsSpotInsertionVehicle = blRet
End Function
Public Function gIsSpotInsertionVehicleByLine(ilLineNo As Integer, blIsPackage As Boolean) As Boolean
    Dim slVefName As String
    Dim blRet As Boolean
    Dim ilVef As Integer
    Dim ilLimit As Integer
    
    blRet = False
    blIsPackage = False
    If ilLineNo > 0 Then
        'D.S. 5/30/18 conditional compile
        #If programmatic <> 1 Then
            slVefName = smLnSave(1, ilLineNo)
        #Else
            'get vehicle name from vehicle line from the CLF
            ilVef = gBinarySearchVef(tgClfCntr(ilLineNo - 1).ClfRec.iVefCode)
            If ilVef <> -1 Then
                slVefName = Trim(tgMVef(ilVef).sName)
            Else
                slVefName = "Vehicle Not Found"
            End If
        #End If
        ilLimit = UBound(tgMVef)
        For ilVef = LBound(tgMVef) To ilLimit Step 1
            If slVefName = Trim$(tgMVef(ilVef).sName) Then
                '7959
                If tgMVef(ilVef).sType = "R" Then
                    If gIsInsertionExport(tgMVef(ilVef).iCode) Then
                        blRet = True
                    End If
                ElseIf tgMVef(ilVef).sType = "P" Then
                    blIsPackage = True
                    blRet = mIsRepPackage(ilLineNo)
                '8132
                ElseIf bgStationAreVehicles Then
                    If gGetStationFromVehicle(tgMVef(ilVef).iCode) > 0 Then
                        If gIsInsertionExport(tgMVef(ilVef).iCode) Then
                            blRet = True
                        End If
                    End If
                End If
                Exit For
            End If
        Next ilVef
    End If
    gIsSpotInsertionVehicleByLine = blRet
End Function
Private Function mIsRepPackage(ilLineNo As Integer) As Boolean
    Dim blRet As Boolean
    Dim ilRowNo As Integer
    blRet = False
    For ilRowNo = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        If tgClfCntr(ilRowNo).ClfRec.iPkLineNo = ilLineNo Then
            blRet = mIsSpotInsertionVehicle(tgClfCntr(ilRowNo).ClfRec.iVefCode)
            If blRet Then
                Exit For
            End If
        End If
    Next ilRowNo
    mIsRepPackage = blRet
End Function

Public Function gLoadFromIni(slSection As String, slKey As String, slPath As String, slvalue As String) As Boolean
    'get values from ini file.
    'I-slSection ("XDigital"),  slKey("Host"), slPath (path to ini file)
    'o-slValue. Value from ini file. If not found, = "Not Found"
    'o- boolean.  True if found.
    On Error GoTo ERR_gLoadOption
    Dim BytesCopied As Integer
    Dim sBuffer As String * 128
    
    slvalue = "Not Found"
    gLoadFromIni = False
    '8886
    'If Dir(slPath) > "" Then
    If gFileExist(slPath) = FILEEXISTS Then
        BytesCopied = GetPrivateProfileString(slSection, slKey, "Not Found", sBuffer, 128, slPath)
        If BytesCopied > 0 Then
            If InStr(1, sBuffer, "Not Found", vbTextCompare) = 0 Then
                slvalue = Left(sBuffer, BytesCopied)
                gLoadFromIni = True
            End If
        End If
    End If 'slPath not valid?
    Exit Function

ERR_gLoadOption:
    ' return now if an error occurs
End Function
Public Function mBuildAgencyXmlPath() As String
    Dim slPath As String
    Dim slPathToINI As String
    Dim slAgencyXMLPath As String
    
    slAgencyXMLPath = sgRootDrive & "csi\"
    If igDirectCall = -1 Then
        slPathToINI = sgIniPath & "Traffic.Ini"
    Else
        slPathToINI = CurDir$ & "\Traffic.Ini"
    End If
    If gLoadFromIni("Locations", "ProposalXMLPath", slPathToINI, slPath) Then
On Error GoTo CONTINUE
        mBuildDirectoriesAsNeeded (slPath)
On Error GoTo 0
        slAgencyXMLPath = gSetPathEndSlash(slPath, True)
    End If
CONTINUE:
    mBuildAgencyXmlPath = slAgencyXMLPath
End Function
Public Sub mBuildStationXMLPaths(slStationXMLPath As String, slWideOrbitPath As String)
    Dim slPath As String
    Dim slPathToINI As String
    '8078
    Dim slLocation As String
    Dim blIsMarketron As Boolean
    Dim blIsWOIssue As Boolean
    
    slStationXMLPath = ""
    slWideOrbitPath = ""
    If igDirectCall = -1 Then
        slPathToINI = sgIniPath & "Traffic.Ini"
    Else
        slPathToINI = CurDir$ & "\Traffic.Ini"
    End If
    '8078
    slLocation = "Locations"
    If igTestSystem Then
        slLocation = "TestLocations"
    End If
    If gLoadFromIni(slLocation, "InsertionXMLPath", slPathToINI, slPath) Then
On Error GoTo CONTINUE
        blIsWOIssue = True
        mBuildDirectoriesAsNeeded (slPath)
        blIsWOIssue = False
On Error GoTo 0
        slWideOrbitPath = gSetPathEndSlash(slPath, True)
    ElseIf gLoadFromIni(slLocation, "WideOrbitStationXMLPath", slPathToINI, slPath) Then
On Error GoTo CONTINUE
        blIsWOIssue = True
        mBuildDirectoriesAsNeeded (slPath)
        blIsWOIssue = False
On Error GoTo 0
        slWideOrbitPath = gSetPathEndSlash(slPath, True)
    End If
    If Len(slWideOrbitPath) > 0 Then
        mStartStopStationErrorLog True
        bgStationXmlSiteAllowed(StationXmlIndexes.WO) = True
    End If
CONTINUE:
    If blIsWOIssue Then
        If ogStationXmlErrors Is Nothing Then
            mStartStopStationErrorLog True
        End If
        ogStationXmlErrors.WriteError "Could not create WO station xml path- " & slPath
    End If
    'only build folders if using at least one service
    'false if no heading in xml.ini, or url is blank
    blIsMarketron = mStartStopMarketronOrder(True)
    ' add more as add station xml services
    If blIsMarketron Then
        If ogStationXmlErrors Is Nothing Then
            mStartStopStationErrorLog True
        End If
        slPath = sgExportPath & "Station XML Exports"
        '8886
        'If Dir(slPath, vbDirectory + vbHidden + vbSystem) = vbNullString Then
        If Not gFolderExist(slPath) Then
    On Error GoTo ERRBOX
            MkDir slPath
    On Error GoTo 0
        End If
        slStationXMLPath = gSetPathEndSlash(slPath, True)
        'I don't need to make this path, just build the folder
        'test each station xml and build as needed
        If blIsMarketron Then
            slPath = slStationXMLPath & STATIONXMLMARKETRONNAME
            '8886
            'If Dir(slPath, vbDirectory + vbHidden + vbSystem) = vbNullString Then
            If Not gFolderExist(slPath) Then
        On Error GoTo ERRBOX
                MkDir slPath
        On Error GoTo 0
            End If
            bgStationXmlSiteAllowed(StationXmlIndexes.MARKETRONEO) = True
            If ogEOrder.isTest Then
                ogStationXmlErrors.WriteWarning STATIONXMLMARKETRONNAME & ": running in 'Test Mode'.  Those orders will not be sent.", True
                gMsgBox STATIONXMLMARKETRONNAME & ": running in 'Test Mode'.  Those orders will not be sent."
            End If
        End If
    End If
    Exit Sub
ERRBOX:
    slStationXMLPath = ""
    ogStationXmlErrors.WriteError "Could not create station xml path- " & slPath
End Sub
Private Function mBuildDirectoriesAsNeeded(slPath As String) As Boolean
    'error caught by calling function
    '8137
    Dim slDirs() As String
    Dim c As Integer
    Dim slTempPath As String
    Dim blRet As Boolean
    
    blRet = False
    slTempPath = ""
    If InStrRev(slPath, "\") = Len(slPath) Then
        slPath = Mid(slPath, 1, Len(slPath) - 1)
    End If
    slDirs = Split(slPath, "\")
    For c = 0 To UBound(slDirs)
        If c = 0 Then
            slTempPath = slDirs(c)
        Else
            slTempPath = slTempPath & "\" & slDirs(c)
        End If
        '8886
        'If Dir(slTempPath, vbDirectory + vbHidden + vbSystem) = vbNullString Then
        If Not gFolderExist(slTempPath) Then
            MkDir slTempPath
        End If
    Next c
    blRet = True
    mBuildDirectoriesAsNeeded = blRet
End Function

'Public Sub mBuildXmlPaths(slAgencyXMLPath As String, slStationXMLPath As String, slWideOrbitPath As String)
'    'O: proposal path - if not defined in traffic.ini, set to default.  insertion/rep path--if not defined in traffic.ini, return blank!
'    Dim slPath As String
'    Dim slPathToINI As String
'
'    slWideOrbitPath = ""
'    slAgencyXMLPath = sgRootDrive & "csi\"
'    If igDirectCall = -1 Then
'        slPathToINI = sgIniPath & "Traffic.Ini"
'    Else
'        slPathToINI = CurDir$ & "\Traffic.Ini"
'    End If
'
'    If gLoadFromIni("Locations", "ProposalXMLPath", slPathToINI, slPath) Then
'        If Dir(slPath, vbDirectory + vbHidden + vbSystem) = vbNullString Then
'On Error GoTo CONTINUE
'            MkDir slPath
'On Error GoTo 0
'        End If
'        slAgencyXMLPath = gSetPathEndSlash(slPath, True)
'    End If
'CONTINUE:
'    If gLoadFromIni("Locations", "InsertionXMLPath", slPathToINI, slPath) Then
'        If Dir(slPath, vbDirectory + vbHidden + vbSystem) = vbNullString Then
'On Error GoTo CONTINUE2
'            MkDir slPath
'On Error GoTo 0
'        End If
'        slWideOrbitPath = gSetPathEndSlash(slPath, True)
'    End If
'CONTINUE:
'    slPath = sgExportPath & "StationXML"
'    If Dir(slPath, vbDirectory + vbHidden + vbSystem) = vbNullString Then
'On Error GoTo CLEANUP
'        MkDir slPath
'On Error GoTo 0
'    End If
'    slStationXMLPath = gSetPathEndSlash(slPath, True)
'CLEANUP:
'
'End Sub
Public Function mReturnLineNumber(ilIndex As Integer) As Integer
    Dim slDelimiter As String
    Dim slLine() As String
    Dim ilLineSize As Integer
    Dim ilPresentLine As Integer
    
    slDelimiter = Chr(9)
On Error GoTo ERRORBOX
    slLine = Split(smClfInfo(ilIndex), slDelimiter)
    ilLineSize = UBound(slLine)
    If ilLineSize >= 1 Then
        ilPresentLine = slLine(XMLCLFLINE)
    End If
    mReturnLineNumber = ilPresentLine + 1
CLEANUP:
    Erase slLine
    Exit Function
ERRORBOX:
    mReturnLineNumber = 0
    MsgBox "Problem with exporting-mReturnLineNumber:" & err.Description, vbExclamation, "Error"
    GoTo CLEANUP
End Function
Public Sub mDeletePrevious(slPath As String)
    Dim slFileName As String
    Dim myFs As FileSystemObject
    Dim myFolder As Folder
    Dim MyFile As file
On Error GoTo ERRORBOX
'   8886
'    slFileName = Dir(slPath & Trim$(str$(tgChfCntr.lCntrNo)) & "-O*.*")
'    Do While slFileName > ""
'        Kill slPath & slFileName
'        slFileName = Dir$
'    Loop
    Set myFs = New FileSystemObject
    If myFs.FolderExists(slPath) Then
        slFileName = myFs.BuildPath(slPath, Trim$(str$(tgChfCntr.lCntrNo)) & "-O*.*")
        myFs.DeleteFile slFileName, True
    End If
ERRORBOX:
    
End Sub
Public Sub gUpdateStationXml(hlChf As Integer, hlClf As Integer)
    Dim tlChfSrchKey As LONGKEY0  'CHF key record image
    Dim ilChfRecLen As Integer      'CHF record length
    Dim ilRet As Integer
    Dim tlChf As CHF
    Dim ilEntryDate(0 To 1) As Integer
    Dim ilEntryTime(0 To 1) As Integer
    Dim slStr As String
    Dim ilClf As Integer
    Dim blIsPackage As Boolean
    Dim ilClfRecLen As Integer
    Dim tlClf As CLF
    Dim tlClfSearchKey As LONGKEY0
    
    '8032 at least one was successful don't need this because test to get in here
  '  If mIsStationXmlAllowedForOrder() Then
        ilClfRecLen = Len(tlClf)
        ilChfRecLen = Len(tlChf)
        tlChfSrchKey.lCode = tgChfCntr.lCode
        ilRet = btrGetEqual(hlChf, tlChf, ilChfRecLen, tlChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            tlChf.iXMLSentExtRevNo = tlChf.iExtRevNo
            slStr = Format$(gNow(), "m/d/yy")
            gPackDate slStr, ilEntryDate(0), ilEntryDate(1)
            slStr = Format$(gNow(), "h:m:s AM/PM")
            gPackTime slStr, ilEntryTime(0), ilEntryTime(1)
            tlChf.iXMLSentDate(0) = ilEntryDate(0)
            tlChf.iXMLSentDate(1) = ilEntryDate(1)
            tlChf.iXMLSentTime(0) = ilEntryTime(0)
            tlChf.iXMLSentTime(1) = ilEntryTime(1)
            tlChf.iXMLSentUrfCode = tgUrf(0).iCode
            ilRet = btrUpdate(hlChf, tlChf, ilChfRecLen)
            'update tgchf!
            If ilRet = BTRV_ERR_NONE Then
                tgChfCntr.iXMLSentDate(0) = ilEntryDate(0)
                tgChfCntr.iXMLSentDate(1) = ilEntryDate(1)
                tgChfCntr.iXMLSentTime(0) = ilEntryTime(0)
                tgChfCntr.iXMLSentTime(1) = ilEntryTime(1)
                tgChfCntr.iXMLSentExtRevNo = tlChf.iXMLSentExtRevNo
                tgChfCntr.iXMLSentUrfCode = tlChf.iXMLSentUrfCode
            End If
            '7959 for reports, update each line
            For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                blIsPackage = False
                If gIsSpotInsertionVehicleByLine(ilClf + 1, blIsPackage) Then
                    If Not blIsPackage Then
                        '8032
                    '    If igStationXmlChoice = ilStationXmlChoice Then
                            'if something goes wrong in sending/creating, this becomes false
                            If mIsStationXmlAllowedForOrder(igStationXmlChoice, True) Then
                                tlClfSearchKey.lCode = tgClfCntr(ilClf).ClfRec.lCode
                                ilRet = btrGetEqual(hlClf, tlClf, ilClfRecLen, tlClfSearchKey, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                                If ilRet = BTRV_ERR_NONE Then
                                    tlClf.iXMLSentExtRevNo = tlChf.iXMLSentExtRevNo
                                    ilRet = btrUpdate(hlClf, tlClf, ilClfRecLen)
                                End If
                            End If
                     '   End If
                    End If
                End If
            Next ilClf
        End If
  '  End If
End Sub
Public Function mGetBuyerEmail(ilBuyerEmail As Integer, slBuyerName As String, hlPnf As Integer) As String
'O-slBuyerName, buyer's email address
    Dim llCefCode As Long
    Dim slToAddress As String
        If ilBuyerEmail > 0 Then
            llCefCode = mGetXmlPnfInfo(ilBuyerEmail, slBuyerName, hlPnf)
            'to emailExists
            If llCefCode > 0 Then
                slToAddress = ogEmailer.GetEmailAddress(llCefCode)
            End If
        End If
        mGetBuyerEmail = slToAddress
End Function
Private Function mGetXmlPnfInfo(ilPnfCode As Integer, slBuyerName As String, hlPnf As Integer) As Long
'O--slBuyerName
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim tlPnf As PNF
    
    slBuyerName = ""
    mGetXmlPnfInfo = 0
    ilRecLen = Len(tlPnf)
    ilRet = btrGetEqual(hlPnf, tlPnf, ilRecLen, ilPnfCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        mGetXmlPnfInfo = tlPnf.lEMailCefCode
        slBuyerName = Trim$(tlPnf.sName)
    End If
End Function
Public Sub mStartStopStationErrorLog(blIsStart As Boolean)
    Dim blRet As Boolean
    Dim slMessagePath As String
    Dim slErrorMessage As String
    
    If blIsStart Then
        slMessagePath = sgDBPath & "Messages\"
        Set ogStationXmlErrors = New CLogger
        ogStationXmlErrors.LogPath = ogStationXmlErrors.CreateLogName(slMessagePath & "StationXml")
    Else
        If Not ogStationXmlErrors Is Nothing Then
            Set ogStationXmlErrors = Nothing
        End If
    End If
End Sub
Public Function mStartStopMarketronOrder(blIsStart As Boolean) As Boolean
    Dim blRet As Boolean
    Dim slMessagePath As String
    Dim slErrorMessage As String
    
    blRet = True
    If blIsStart Then
        slMessagePath = sgDBPath & "Messages\"
        Set ogEOrder = New CEOExport
        If Len(ogEOrder.ErrorMessage) > 0 Then
            'this isn't an error!  Not turned on
            blRet = False
            '8278
            If Not ogEOrder Is Nothing Then
                Set ogEOrder = Nothing
            End If
        Else
            If ogEOrder.isProxy Then
                ogStationXmlErrors.WriteFacts STATIONXMLMARKETRONNAME & ": Proxy is set", True
            End If
            ogEOrder.LogPath = ogEOrder.CreateLogName(slMessagePath & STATIONXMLMARKETRONNAME & "Debug")
            ogEOrder.LogStart
            If Len(ogEOrder.SoapUrl) = 0 Then
                blRet = False
            End If
        End If
    Else
        If Not ogEOrder Is Nothing Then
            Set ogEOrder = Nothing
        End If
    End If
    mStartStopMarketronOrder = blRet
End Function
Private Sub mCsiXMLDataExAlternate(slWrite As String)
    If Not ogEOrder Is Nothing Then
        ogEOrder.AddToBody (slWrite)
    End If
End Sub
Public Sub mStationXmlAllowedStart()
    mStationXmlAllowedOrderStart
    bgStationXmlSiteAllowed(StationXmlIndexes.WO) = False
    bgStationXmlSiteAllowed(StationXmlIndexes.MARKETRONEO) = False
End Sub
Public Sub mStationXmlAllowedOrderStart()
    bgStationXmlOrderAllowed(StationXmlIndexes.WO) = False
    bgStationXmlOrderAllowed(StationXmlIndexes.MARKETRONEO) = False
    
    bgStationXmlWebServiceAllowed(StationXmlIndexes.WO) = True
    bgStationXmlWebServiceAllowed(StationXmlIndexes.MARKETRONEO) = True
End Sub
Public Function mStationXmlAllowedOrder(ilVefCode As Integer) As Boolean
    'return true except where 'site allowed' doesn't allow order
    Dim blRet As Boolean
    blRet = True
    If gIsInsertionExport(ilVefCode) Then
        Select Case igStationXmlChoice
            Case STATIONXMLWIDEORBIT
                If bgStationXmlSiteAllowed(StationXmlIndexes.WO) Then
                    bgStationXmlOrderAllowed(StationXmlIndexes.WO) = True
                Else
                    blRet = False
                End If
            Case STATIONXMLMARKETRON
                If bgStationXmlSiteAllowed(StationXmlIndexes.MARKETRONEO) Then
                    bgStationXmlOrderAllowed(StationXmlIndexes.MARKETRONEO) = True
                Else
                    blRet = False
                End If
        End Select
    End If
    mStationXmlAllowedOrder = blRet
End Function
Public Function mIsStationXmlAllowedForOrder(Optional ilTestOne As Integer = -1, Optional blTestService = False) As Boolean
'if ilTestOne < 0 then return true if any allowed.  Otherwise, test that specific one
' bgStationXmlOrderAllowed will only be true if bgStation..Site is true.  Note also testing bgStation..WebService, which defaults to true.
    Dim blRet As Boolean
    Dim c As Integer
    
    blRet = False
    If tgChfCntr.sStatus = "O" Or tgChfCntr.sStatus = "N" Then
        If ilTestOne < 0 Then
            For c = StationXmlIndexes.WO To StationXmlIndexes.MARKETRONEO
                If bgStationXmlOrderAllowed(c) Then
                    If blTestService Then
                        If bgStationXmlWebServiceAllowed(c) Then
                            blRet = True
                        End If
                    Else
                        blRet = True
                    End If
                    Exit For
                End If
            Next c
        Else
            Select Case ilTestOne
                Case STATIONXMLWIDEORBIT
                    If bgStationXmlOrderAllowed(StationXmlIndexes.WO) Then ' And bgStationXmlWebServiceAllowed(StationXmlIndexes.WO) Then
                        If blTestService Then
                            If bgStationXmlWebServiceAllowed(StationXmlIndexes.WO) Then
                                blRet = True
                            End If
                        Else
                            blRet = True
                        End If
                    End If
                Case STATIONXMLMARKETRON
                    If bgStationXmlOrderAllowed(StationXmlIndexes.MARKETRONEO) Then ' And bgStationXmlWebServiceAllowed(StationXmlIndexes.MARKETRONEO) Then
                        If blTestService Then
                            If bgStationXmlWebServiceAllowed(StationXmlIndexes.MARKETRONEO) Then
                                blRet = True
                            End If
                        Else
                            blRet = True
                        End If
                    End If
            End Select
        End If
    End If
    mIsStationXmlAllowedForOrder = blRet
End Function
Public Function mIsStationXmlAllowedForSite(Optional ilTestOne As Integer = -1) As Boolean
    Dim blRet As Boolean
    Dim c As Integer
    
    blRet = False
    If ilTestOne < 0 Then
        For c = StationXmlIndexes.WO To StationXmlIndexes.MARKETRONEO
            If bgStationXmlSiteAllowed(c) Then
                blRet = True
                Exit For
            End If
        Next c
    Else
        Select Case ilTestOne
            Case STATIONXMLWIDEORBIT
                If bgStationXmlSiteAllowed(StationXmlIndexes.WO) Then
                    blRet = True
                End If
            Case STATIONXMLMARKETRON
                If bgStationXmlSiteAllowed(StationXmlIndexes.MARKETRONEO) Then
                    blRet = True
                End If
        End Select
    End If
    mIsStationXmlAllowedForSite = blRet
End Function
Public Function mIsStationXmlAllowedForWebService(ilTestOne As Integer) As Boolean
'if ilTestOne < 0 then return true if any allowed.  Otherwise, test that specific one
    Dim blRet As Boolean
    Dim c As Integer
    
    blRet = False
    Select Case ilTestOne
        Case STATIONXMLWIDEORBIT
            'I could just return false since it isn't a webervice
            If bgStationXmlWebServiceAllowed(StationXmlIndexes.WO) Then
                blRet = True
            End If
        Case STATIONXMLMARKETRON
            If bgStationXmlWebServiceAllowed(StationXmlIndexes.MARKETRONEO) Then
                blRet = True
            End If
    End Select
    mIsStationXmlAllowedForWebService = blRet
End Function
'Dan M 10/25/11 replace the original mcExport_Click so that email Proposal can call function below before sending email
'Dan M 11/28/11 now is called from proposals(to export file as normal), and also from email.
'8940 from email, option to not include rate
Public Function mExportProposal(slExportType As String, Optional blChooseWhere As Boolean, Optional blCallFromSave As Boolean = False, Optional blZeroRates As Boolean = False) As String
    Dim ilClf As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilTest As Integer
    Dim slToFile As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim hlTo As Integer
    Dim slRecord As String
    Dim slDelimiter As String
    Dim slDateTime As String
    Dim hlSlf As Integer        'Salesperson
    Dim ilSlfRecLen As Integer     'Slf record length
    Dim tlSlf As SLF
    Dim tlSlfSrchKey As INTKEY0
    Dim hlSof As Integer        'Sales office
    Dim ilSofRecLen As Integer     'Sof record length
    Dim tlSof As SOF
    Dim tlSofSrchKey As INTKEY0
    Dim hlDnf As Integer        'Salesperson
    Dim ilDnfRecLen As Integer     'Slf record length
    Dim tlDnf As DNF
    Dim tlDnfSrchKey As INTKEY0
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim ilLnRowNo As Integer
    Dim llCntrPop As Long
    Dim slLineType As String
    Dim llStartDate As Long
    Dim llMoStartDate As Long
    Dim llEndDate As Long
    Dim llDate As Long
    Dim ilDay As Integer
    Dim llWkSpots As Long
    Dim slWDFlag As String
    ReDim ilOVDays(0 To 6) As Integer
    ReDim ilWkDays(0 To 6) As Integer
    ReDim ilDailySpots(0 To 6) As Integer
    Dim ilStdWkCff As Integer
    ReDim ilStdWkDays(0 To 6) As Integer
    Dim ilLnGen As Integer
    Dim slDate As String
    Dim ilUpper As Integer
    Dim ilMatch As Integer
    Dim ilWkCff As Integer
    Dim ilNoSpots As Integer
    Dim llWeek As Long
    Dim ilRdf As Integer
    Dim ilPos As Integer
    Dim slDays As String
    Dim ilLastLnNo As Integer
    ReDim slCffXDay(0 To 6) As String * 1
    Dim slEDIDays As String
    Dim slTime As String
    '9128 to long
    'Dim ilSTNoSpots As Integer
    Dim llSTNoSpots As Long
    'Dim llLnNoSpots As Integer
    Dim llLnNoSpots As Long
    Dim slDaysTimes As String
    Dim ilItemNo As Integer
    Dim slInputStr As String
    Dim slClientName As String          '2-18-05
    '2/15/05: Soc Eco added
    Dim ilMnfSocEco As Integer
    Dim slSocEcoName As String
    Dim slNameCode As String
    Dim slCode As String
    '4/29/2011 Dan xml proposal
    Dim blDailySpots As Boolean
    Dim slXmlErrorMessage As String
    Dim slBuyerName As String
    '6768
    Dim ilWeekSpan As Integer
    '7673 'treat as reps' because if rep vehicles but want to generate original proposalxml, it is false
    Dim blTreatAsReps As Boolean
    Dim c As Integer
    Dim slRepFileTemp As String
    
    Dim blContinuePrint As Boolean
    '7687
    Dim slRepVehicle As String
    Dim ilRepCounter As Integer
    Dim ilNextCounter As Integer
    Dim slTemp As String
    Dim slTestForDelete As String
    Dim slOldDelete() As String
    '7803
    Dim blIsRepPackage As Boolean
    '7959
    Dim blWasRun As Boolean
    '8032
    Dim slAgencyErrorMessage As String
    Dim blIsMissingStationRights As Boolean
    '8915
    Dim slAdvertiserName As String
    Dim slDemoName As String
    Dim ilVef As Integer
    Dim ilAgf As Integer
    '9459
    Dim blCSIAndHidden As Boolean
    Dim ilLineNumber As Integer
    Const CSIANDHIDDEN As String = "Z"
    '9202
    Const KILLISSUE As Integer = 4
    'D.S. D.M. 06/6/18
    blWasRun = False
    Const REPREPLACE = "-DanXReplaceForRepXDick"
    '9459
    blCSIAndHidden = False
    ilLineNumber = 0
    If slExportType = CSIANDHIDDEN Then
        blCSIAndHidden = True
        slExportType = "C"
    End If
    blIsMissingStationRights = False
    blTreatAsReps = False
    '8364
'    If imDiffMode Then
    If tgExportFacts.imDiffMode Then
        Exit Function
    End If
    slAgencyErrorMessage = ""
    'xml proposal
    smClfDates = ""
    'Allow export as a proposal function only.  Jim 10/14/03
    '8033 this now gets the button on screen (and the save from orders, which would be handled below also)
    If (sgCntrScreen = "Orders") Then
        '7958
       ' Exit Function
       blTreatAsReps = True
    End If

    If UBound(tgClfCntr) <= 0 Then
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
#If programmatic = 1 Then
    '8915
'    tgExportAgf.sXMPProposalBand = ""
'    ilAgf = gBinarySearchAgf(tgChfCntr.iAgfCode)
'    If ilAgf <> -1 Then
'        tgExportAgf.sName = tgCommAgf(ilAgf).sName
'    Else
'        tgExportAgf.sName = "Missing"
'    End If
    gAgfSQLtoUDT tgChfCntr.iAgfCode, tgExportAgf
    Set myRst = gSQLSelectCall("Select adfname from ADF_Advertisers where adfCode = " & tgChfCntr.iAdfCode)
    If Not myRst.EOF Then
        slAdvertiserName = Trim$(myRst!adfName)
    Else
        slAdvertiserName = ""
    End If
    Set myRst = gSQLSelectCall("Select mnfName from MNF_Multi_Names where mnfCode = " & tgChfCntr.iMnfDemo(0))
    If Not myRst.EOF Then
        slDemoName = Trim$(myRst!mnfName)
    Else
        slDemoName = ""
    End If
#End If
    '8033
    If Not blCallFromSave Then
        '8364
        'mAllSetShow 0
'        If Not mMoveLnCtrlToRec(True) Then
'            Screen.MousePointer = vbDefault
'            Exit Function
'        End If
        If blTreatAsReps Then
            'Dan cbcDifference-if not current, can't select button! I'll leave it in anyway
            'if invisible, then we have the first version (and lbcRevNo isn't 0) so if invisible, always allow. but if visible, then we have to have lbcRevno = 0
            '8364
           ' If (lbcRevNo.ListIndex <> 0 And cbcDifference.Visible) Then
'8915
#If programmatic <> 1 Then
            If (Contract.lbcRevNo.ListIndex <> 0 And Contract.cbcDifference.Visible) Then
                gMsgBox "Can only generate latest version. Exporting halted."
                Screen.MousePointer = vbDefault
                Exit Function
            End If
#End If
           ' 8032 set to false
            mStationXmlAllowedOrderStart
        End If
    Else
        blTreatAsReps = True
    End If
    '7806 contract orders not tied to 'export type'
    'no 'reps' if path not defined in traffic.ini
    '7958  now, don't export insertion orders from proposal screen
'    If Len(smStationXMLWOPath) > 0 Then
'        blTreatAsReps = True
'    End If
    'xml proposal Dan M save as needed before creating file; if ediCallLetters don't exist, don't write
    bmXmlOutput = mTestForProposalXml(slExportType)
    '7673 1) if blchoosewhere is false, then this is an email...don't get to choose reps--set to false 2) otherwise, if they choose reps, then they don't get to choose where!
    '7958 never true!
'    If blTreatAsReps And blChooseWhere Then
'        If MsgBox("Export as insertion order?", vbYesNo, "Output") = vbYes Then
'            'no reps? don't want to resend? not an order?
'            If mTestVehiclesForSpotInsertion() Then
'                blChooseWhere = False
'                bmXmlOutput = True
'                slExportType = "X"
'            Else
'                Screen.MousePointer = vbDefault
'                Exit Function
'            End If
'        Else
'            blTreatAsReps = False
'        End If
    '7958
    If blTreatAsReps Then
        'no reps? don't want to resend? not an order? 8132 hitting here
        If mTestVehiclesForSpotInsertion(blCallFromSave) Then
            blChooseWhere = False
            bmXmlOutput = True
            slExportType = "X"
        Else
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    Else
        'this takes care of email of reps...don't send as reps!
        blTreatAsReps = False
    End If
    DoEvents
    If bmXmlOutput Then
        'edi not defined, don't create xml file.
        If Len(tgSpf.sEDICallLetter) <= 0 Then
            MsgBox "Cannot create an xml file until EDI call letters are defined in site options", vbOKOnly + vbCritical + vbApplicationModal, "No EDI call letters defined"
            Screen.MousePointer = vbDefault
            Exit Function
        End If
        '6339 moved here
        '8364
        'ilRet = mXmlDemoInfo(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imStep1Index)
        With tgExportFacts
        ilRet = mXmlDemoInfo(.hmDrf, .hmMnf, .hmDpf, .hmDef, .hmRaf, .tmRegionCode(), .imStep1Index)
        End With
        If ilRet = 0 Then
            '7214
'            MsgBox "No demo information.  Proposal xml will not be generated", vbOkOnly + vbCritical + vbApplicationModal, "No demo info"
'            Screen.MousePointer = vbDefault
'            Exit Function
        ElseIf ilRet = -1 Then
            MsgBox "Custom demo invalid.  Proposal xml will not be generated", vbOKOnly + vbCritical + vbApplicationModal, "Can't read custom demo info"
            Screen.MousePointer = vbDefault
            Exit Function
        End If
' '9/14/2011  changes are now moved from ctrl to memory by setting 'mAllSetShow' above mMoveLnCtrlToRec!
        'did they make changes? exit if true
'        If (cmcSaveStay.Enabled = True) Or (cmcUpdate.Enabled = True) Then
'            MsgBox "Cannot create an xml file until proposal is saved", vbOkOnly + vbCritical + vbApplicationModal, "Save before continuing."
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
    End If
    'D.S. 5/30/18 conditional compile
    #If programmatic <> 1 Then
        mSplitFlightsIntoWeeks -1
    #Else
        'Dan, why the splitting it?
    #End If

    '2-18-05 use short client name vs long client name
    slClientName = Trim$(tgSpf.sGClient)
    If tgSpf.iMnfClientAbbr > 0 Then
    'Dan 8364
    With tgExportFacts
       ' tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
        .tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
    '    ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        'D.S. 5/30/18 conditional compile
        #If programmatic <> 1 Then
            ilRet = btrGetEqual(.hmMnf, .tmMnf, .imMnfRecLen, .tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            slClientName = Trim$(.tmMnf.sName)
        #Else
            slClientName = Trim$(tgSpf.sGClient)
        #End If
    End With
    End If

'    mAllSetShow 0
'Dan 11/14/11 removed and now called before going to email form
'    mMoveCtrlToRec False
'    If tgChfCntr.iAgfCode <= 0 Then
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'    tmAgfSrchKey.iCode = tgChfCntr.iAgfCode
'    ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'    If ilRet <> BTRV_ERR_NONE Then
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'    If tgChfCntr.lCntrNo <= 0 Then
'        Screen.MousePointer = vbDefault
'        ilRet = MsgBox("You must Save this Proposal before generating an Export File", vbOkOnly + vbExclamation, "Message")
'        Exit Function
'    End If
    slDelimiter = Chr(9)
    For ilDay = 0 To 6 Step 1
        slCffXDay(ilDay) = " "
    Next ilDay
    'Set line types
    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        'If tgClfCntr(ilClf).iSPChg And (tgUrf(0).sUseComputeCMC = "I") Then
        '    mLnTotals ilClf + 1, imLnTotalType, PRICEINDEX, hmDrf, hmMnf
        '    mGetLnPriceInfo ilClf + 1, slStr, imLnTotalType, PRICEINDEX     'Formats Price within Show
        'End If
        '8033 for 'resend' of orders, have to set what is allowed. Doesn't really serve a purpose here, but needed if from 'save'
        If Not blCallFromSave And blTreatAsReps Then
            mStationXmlAllowedOrder tgClfCntr(ilClf).ClfRec.iVefCode
        End If
        'D.S. 5/30/18 conditional compile
        llLnNoSpots = 0
        #If programmatic <> 1 Then
            If (sgCntrScreen <> "Orders") Then    'Not setting for new lines'And (imStep1Index <> 4) Then
               '8364
                imInResearch = True 'Only set Books
            'Dan 8364
            With tgExportFacts
                'mInitBookDemoForLn ilClf + 1, imCppCpm, imVefCode, imVpfIndex, imTerminate, imChgMode
                mInitBookDemoForLn ilClf + 1, .imCppCpm, .imVefCode, .imVpfIndex, .imTerminate, .imChgMode
            End With
                '8364
                imInResearch = False
            End If
            tgClfCntr(ilClf).ClfRec.sType = "S"
            tgClfCntr(ilClf).ClfRec.iPkLineNo = 0
            For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                If tgClfCntr(ilClf).ClfRec.iLine = tgPackLn(ilLoop).iLineNo Then
                    tgClfCntr(ilClf).ClfRec.sType = tgPackLn(ilLoop).sType
                    If (tgPackLn(ilLoop).sType <> "O") And (tgPackLn(ilLoop).sType <> "A") And (tgPackLn(ilLoop).sType <> "E") Then
                        For ilTest = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                            If (tgPackLn(ilTest).sType = "A") Or (tgPackLn(ilTest).sType = "O") Or (tgPackLn(ilTest).sType = "E") Then
                                ilIndex = tgPackLn(ilTest).iIndex
                                Do While ilIndex >= 0
                                    If tgClfCntr(ilClf).ClfRec.iLine = tgPackLn(ilIndex).iLineNo Then
                                        tgClfCntr(ilClf).ClfRec.iPkLineNo = tgPackLn(ilTest).iLineNo
                                        Exit For
                                    End If
                                    ilIndex = tgPackLn(ilIndex).iIndex
                                Loop
                            End If
                        Next ilTest
                    End If
                    Exit For
                End If
            Next ilLoop
        #End If
    Next ilClf
    '2/15/05:  Get Soc Eco Name if used
    '8364 added contract.
'8915
        ilMnfSocEco = 0
        slSocEcoName = ""
#If programmatic <> 1 Then
    ''8364
    Contract.mSetHdDate False, True
    Contract.mCntrTotal
    If Contract.lbcDBSocEco.ListIndex >= 1 Then
        slNameCode = tgSocEcoCode(Contract.lbcDBSocEco.ListIndex - 1).sKey  'Traffic!lbcSocEcoCode.List(lbcDBSocEco.ListIndex - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilMnfSocEco = Val(slCode)
        slSocEcoName = Trim$(Contract.lbcDBSocEco.List(Contract.lbcDBSocEco.ListIndex))
    End If
'    '5908
'    If bmXmlOutput Then
'        mXmlDemoInfo hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imStep1Index
'    End If
    If ilMnfSocEco <= 0 Then
        'Set Contract Gross Impressions and GRP
        '8364
       ' mSetPriDemoResearch hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmRegionCode(), imStep1Index, lmCPriDemoGRP, lmCPriDemoGrImp, lmCPriDemoCPP, lmCPriDemoCPM, imCPriDemoRating
       With tgExportFacts
        mSetPriDemoResearch .hmDrf, .hmMnf, .hmDpf, .hmDef, .hmRaf, .tmRegionCode(), .imStep1Index, .lmCPriDemoGRP, .lmCPriDemoGrImp, .lmCPriDemoCPP, .lmCPriDemoCPM, .imCPriDemoRating
       End With
    Else
        'Get values from line as they are set if Soc Eco value defined
        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
            tgClfCntr(ilClf).lPriDemoPop = tgClfCntr(ilClf).lPop
            tgClfCntr(ilClf).lPriDemoAvgAud = tgClfCntr(ilClf).lAvgAud
            tgClfCntr(ilClf).iPriDemoBook = tgClfCntr(ilClf).ClfRec.iDnfCode
            tgClfCntr(ilClf).iPriDemoAvgRating = tgClfCntr(ilClf).iAvgRating
            tgClfCntr(ilClf).lPriDemoGRP = tgClfCntr(ilClf).lGRP
            tgClfCntr(ilClf).lPriDemoCPP = tgClfCntr(ilClf).lCPP
            tgClfCntr(ilClf).lPriDemoCPM = tgClfCntr(ilClf).lCPM
            tgClfCntr(ilClf).lPriDemoGrImp = tgClfCntr(ilClf).lGrImp
        Next ilClf
    End If
#End If
    Screen.MousePointer = vbHourglass  'Wait
    'Create Output
    'Dan M 7/12/18 Dick and I think this was a mistake- V81 only version 34 6/6/18.  Go back to 0
   ' ilRet = 3
    ilRet = 0
    'On Error GoTo imcExportErr:
        '8364 Dan
'8915
#If programmatic <> 1 Then
    If Contract.lbcAdvt.ListIndex < 0 Then
        slStr = ""
    Else
'        slStr = gFileNameFilter(Trim$(lbcAdvt.List(lbcAdvt.ListIndex)))
        slStr = gFileNameFilter(Trim$(Contract.lbcAdvt.List(Contract.lbcAdvt.ListIndex)))
    End If
#Else
    slStr = gFileNameFilter(slAdvertiserName)
#End If
    slStr = slStr & "-" & gFileNameFilter(Trim$(tgChfCntr.sProduct))
    '2-18-05 use short client name (if exists) vs long client name
    '4-27-11 Dan M add proposal xml
    '9/12/11 Dan M. lose save file box.
    '11-28-11 replaced
    If blChooseWhere Then
        'only agency xml
        If bmXmlOutput = True Then
            '7580
            'slToFile = Trim$(str$(tgChfCntr.lCntrNo)) & "-" & Trim$(str$(tgChfCntr.iPropVer)) & slStr & ".xml"  '& gFileNameFilter(Trim$(slClientName))
            'Dan 8364
           ' slToFile = mXmlContractAndVersionOrRevision(imStep1Index) & "-" & slStr & ".xml"
            slToFile = mXmlContractAndVersionOrRevision(tgExportFacts.imStep1Index) & "-" & slStr & ".xml"
'Dan 8364 changed all cdcSetup to contract.cdcSetup
            Traffic!cdcSetup.Filter = "*.xml|*.xml"
                'cdcSetup.InitDir = sgExportPath
        Else
            Traffic!cdcSetup.Filter = "*.txt|*.txt"
            slToFile = Trim$(str$(tgChfCntr.lCntrNo)) & "-" & Trim$(str$(tgChfCntr.iPropVer)) & slStr & gFileNameFilter(Trim$(slClientName)) & ".TXT"
        End If
        If igOnCitrix Then
            slToFile = sgExportPath & slToFile
        Else
            '5676
            'slToFile = "c:\csi\" & slToFile
            '7673
            'slToFile = sgRootDrive & "csi\" & slToFile
            'Dan 8364
'            slToFile = smAgencyXMLPath & slToFile
            'D.S. 5/30/18 conditional compile
            #If programmatic <> 1 Then
                slToFile = tgExportFacts.smAgencyXMLPath & slToFile
            #Else
                tgExportFacts.smAgencyXMLPath = sgExportPath
                slToFile = sgExportPath & slToFile
            #End If
        End If
        Traffic!cdcSetup.flags = cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
        Traffic!cdcSetup.fileName = slToFile
        If igOnCitrix Then
            Traffic!cdcSetup.InitDir = Left$(sgExportPath, Len(sgExportPath) - 1)
        Else
            'cdcSetup.InitDir = "c:\csi"
            '7673 extra
'            cdcSetup.InitDir = sgRootDrive & "csi"
            'Dan 8364
           ' contract.cdcSetup.InitDir = smAgencyXMLPath
            Traffic!cdcSetup.InitDir = tgExportFacts.smAgencyXMLPath
        End If
        Traffic!cdcSetup.CancelError = True
        On Error Resume Next
        Traffic!cdcSetup.Action = 2    'DLG_FILE_SAVE
        'cdcSetup.ShowSave
        If err.Number = cdlCancel Then
            ilRet = 1
        End If
'D.S D.M. 6/6/18
#If programmatic <> 1 Then
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If
#End If
        slToFile = Traffic!cdcSetup.fileName
        '6808 remove this test
'        '6768  we will test xml after know it is not over 53 weeks.
'        If bmXmlOutput = False Then
            'slDateTime = FileDateTime(slToFile)
            ilRet = gFileExist(slToFile)
            If ilRet = 0 Then
                Kill slToFile
            End If
'        End If
        On Error GoTo 0
        ilRet = 0
    Else
        'Dan I'm not sure what this is for
        slToFile = sgExportPath & Trim$(str$(tgChfCntr.lCntrNo)) & "-v" & Trim$(str$(tgChfCntr.iPropVer)) & "-" & slStr & "-" & gFileNameFilter(Trim$(slClientName))
        'dan M 11-28-11 moved here now that added cdcSetup back in.
        If bmXmlOutput Then
            '7580
            '7673 added bltreatasreps
            If blTreatAsReps Then
                '8032
                ' Dan 8364
               ' slToFile = mXmlContractAndVersionOrRevision(imStep1Index) & slStr & "-" & gFileNameFilter(Trim$(slClientName)) & "-" & REPREPLACE & ".xml"
               slToFile = mXmlContractAndVersionOrRevision(tgExportFacts.imStep1Index) & slStr & "-" & gFileNameFilter(Trim$(slClientName)) & "-" & REPREPLACE & ".xml"
                slRepFileTemp = slToFile
                'mDeletePrevious smStationXMLWOPath
                '8364
                mDeletePrevious tgExportFacts.smStationXMLWOPath
               'don't delete previous marketron
               ' mDeletePrevious smStationXMLBasePath & STATIONXMLMARKETRONNAME & "\"
'                slToFile = smStationXMLWOPath & mXmlContractAndVersionOrRevision(imStep1Index) & slStr & "-" & gFileNameFilter(Trim$(slClientName)) & "-" & REPREPLACE & ".xml"
'                slRepFileTemp = slToFile
'                '7959 delete previous order versions.  This makes other 'kill' later in function unneeded
'                mDeletePrevious smStationXMLWOPath
            Else
                'Dan 8364
                'slToFile = sgExportPath & mXmlContractAndVersionOrRevision(imStep1Index) & slStr & "-" & gFileNameFilter(Trim$(slClientName)) & ".xml"
                '8933
                'slToFile = sgExportPath & mXmlContractAndVersionOrRevision(tgExportFacts.imStep1Index) & slStr & "-" & gFileNameFilter(Trim$(slClientName)) & ".xml"
                slToFile = sgExportPath & mXmlContractAndVersionOrRevision(tgExportFacts.imStep1Index) & "-" & slStr & "-" & gFileNameFilter(Trim$(slClientName)) & ".xml"
                'slToFile = slToFile & ".xml"
            End If
        Else
            slToFile = slToFile & ".txt"
            'slDateTime = FileDateTime(slToFile)
            ilRet = gFileExist(slToFile)
            If ilRet = 0 Then
                Kill slToFile
            End If
        End If
        '8032 move back to not xml
'        '7083 remove 'if'
''        '6768  we will test xml after know it is not over 53 weeks.
''        If bmXmlOutput = False Then
'        slDateTime = FileDateTime(slToFile)
'        If ilRet = 0 Then
'            Kill slToFile
'        End If
''        End If
        On Error GoTo 0
        ilRet = 0
    End If
    'ttp 5230. moved from below because need for xml and csi. We have a value if  llCntrPop > 0
    '8/24/16 moved lower
'    llCntrPop = -1
'    For ilClf = ilStartRow To ilEndRow Step 1
'        ilLnRowNo = ilClf + 1
'        slLineType = mGetLineType(ilLnRowNo)
'        If (Not tgClfCntr(ilClf).iCancel) And ((slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E")) Then
'            If llCntrPop = -1 Then
'                llCntrPop = tgClfCntr(ilClf).lPriDemoPop
'            ElseIf llCntrPop > 0 Then
'                If llCntrPop <> tgClfCntr(ilClf).lPriDemoPop Then
'                    llCntrPop = 0
'                End If
'            End If
'        End If
'    Next ilClf
    'On Error GoTo imcExportErr:
    '4-27-11 Dan M add proposal xml
    If bmXmlOutput Then
        'slToFile = slToFile & ".xml"
        '6865 added "" to end
        '6768 move to lower
'        If csiXMLStart("", "", "F", slToFile, Chr$(13) + Chr$(10), "") <> 0 Then
'            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'            Exit Function
'        End If
    Else
      '  slToFile = slToFile & ".txt"
        'hlTo = FreeFile
        'Open slToFile For Output As hlTo
        ilRet = gFileOpen(slToFile, "Output", hlTo)
        If ilRet <> 0 Then
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            Exit Function
        End If
    End If
    'Output Header information
'    ilStartRow = imLB1Or2   'LBound(smLnSave, 2)
'    ilEndRow = UBound(smLnSave, 2) - 1
    '8364 Dan
    'D.S. 5/30/18 conditional compile
    #If programmatic <> 1 Then
            ilStartRow = tgExportFacts.imLB1Or2   'LBound(smLnSave, 2)
            ilEndRow = UBound(tgExportFacts.smLnSave, 2) - 1
            ilStartRow = ilStartRow - 1
            ilEndRow = ilEndRow - 1
    #Else
        ilStartRow = LBound(tgClfCntr)
        ilEndRow = UBound(tgClfCntr) - 1
    #End If
    'ttp 5230. moved from below because need for xml and csi. We have a value if  llCntrPop > 0
    '8/24/16 moved from above to here
    llCntrPop = -1
    For ilClf = ilStartRow To ilEndRow Step 1
        ilLnRowNo = ilClf + 1
        'D.S. 5/30/18 conditional compile
        #If programmatic <> 1 Then
            slLineType = mGetLineType(ilLnRowNo)
        #Else
            slLineType = mPBGetLineType(ilClf)
        #End If
        If (Not tgClfCntr(ilClf).iCancel) And ((slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E")) Then
            If llCntrPop = -1 Then
                llCntrPop = tgClfCntr(ilClf).lPriDemoPop
            ElseIf llCntrPop > 0 Then
                If llCntrPop <> tgClfCntr(ilClf).lPriDemoPop Then
                    llCntrPop = 0
                End If
            End If
        End If
    Next ilClf
    slRecord = ""
    'If tmAgf.sCntrExptForm = "O" Then
    If slExportType = "O" Then
'8915
#If programmatic <> 1 Then
        If Contract.lbcAdvt.ListIndex < 0 Then
            slRecord = ""
        Else
        'Dan 3/28 17 added contract. to lbcAdvt and lbcDemo
            slRecord = Trim$(Contract.lbcAdvt.List(Contract.lbcAdvt.ListIndex))
        End If
#Else
        slRecord = slAdvertiserName
#End If
        slRecord = slRecord & "-" & Trim$(tgChfCntr.sProduct) & slDelimiter
#If programmatic <> 1 Then
        If Contract.lbcDemo(0).ListIndex >= 1 Then
            'slRecord = slRecord & lbcDemo(0).List(lbcDemo(0).ListIndex) & slDelimiter
            If Trim$(slSocEcoName) = "" Then
                slRecord = slRecord & Contract.lbcDemo(0).List(Contract.lbcDemo(0).ListIndex) & slDelimiter
            Else
                slRecord = slRecord & Contract.lbcDemo(0).List(Contract.lbcDemo(0).ListIndex) & ", " & Trim$(slSocEcoName) & slDelimiter
            End If
        Else
            slRecord = slRecord & "" & slDelimiter
        End If
#Else
        If tgChfCntr.iMnfDemo(0) > 0 Then
            If Trim$(slSocEcoName) = "" Then
                slRecord = slRecord & slDemoName & slDelimiter
            Else
                slRecord = slRecord & slDemoName & ", " & Trim$(slSocEcoName) & slDelimiter
            End If
        Else
            slRecord = slRecord & "" & slDelimiter
        End If
#End If
        'Dan 8364
        'slStr = Trim$(smTotalDollars)  'Trim$(edcGross.Text)
        'D.S. 5/30/18 condtional compile
        #If programmatic <> 1 Then
            slStr = Trim$(smTotalDollars)
        #Else
            slStr = gLongToStrDec(tgChfCntr.lInputGross, 2)
        #End If
        If slStr = "" Then
            slStr = "0.00"
        Else
            If InStr(1, slStr, ".", vbTextCompare) <= 0 Then
                slStr = slStr & ".00"
            End If
        End If
        slRecord = slRecord & slStr & slDelimiter
        gUnpackDate tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1), slStr
        slRecord = slRecord & slStr & slDelimiter
        gUnpackDate tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1), slStr
        slRecord = slRecord & slStr
        Print #hlTo, slRecord
    'Dan 4/27/2011 add proposal xml
    ElseIf bmXmlOutput Then
        ReDim smClfInfo(0)
        ReDim smClfHiddenInfo(0)
        '8915
#If programmatic <> 1 Then
        If Contract.lbcAdvt.ListIndex < 0 Then
            slStr = ""
        Else
            slStr = Trim$(Contract.lbcAdvt.List(Contract.lbcAdvt.ListIndex))
        End If
#Else
        'mPrintOutFileStoreXml 7, "1,2,3," & slDelimiter, True
        slStr = slAdvertiserName
#End If
    Else
        Print #hlTo, "Proposal #:" & slDelimiter & slDelimiter & Trim$(str$(tgChfCntr.lCntrNo))
        Print #hlTo, "Version #:" & slDelimiter & slDelimiter & Trim$(str$(tgChfCntr.iPropVer))
        Print #hlTo, "Date Prepared:" & slDelimiter & slDelimiter & Format(Now, "m/d/yy")
        gUnpackDate tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1), slStr
        Print #hlTo, "Start Date:" & slDelimiter & slDelimiter & slStr
        gUnpackDate tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1), slStr
        Print #hlTo, "End Date:" & slDelimiter & slDelimiter & slStr
        slStr = slStr & "/" & Trim$(tgChfCntr.sProduct)
        Print #hlTo, "Advertiser/Product:" & slDelimiter & slDelimiter & slStr
        If tgChfCntr.iSlfCode(0) > 0 Then
            tlSlfSrchKey.iCode = tgChfCntr.iSlfCode(0)
            hlSlf = CBtrvTable(ONEHANDLE)
            ilRet = btrOpen(hlSlf, "", sgDBPath & "Slf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                ilSlfRecLen = Len(tlSlf)  'btrRecordLength(hlSlf)  'Get and save record length
                ilRet = btrGetEqual(hlSlf, tlSlf, ilSlfRecLen, tlSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slStr = Trim$(tlSlf.sFirstName) & " " & Trim$(tlSlf.sLastName) & ", " & gFormatPhoneNo(tlSlf.sPhone)
                    Print #hlTo, "A/E:" & slDelimiter & slDelimiter & slStr
                    If tlSlf.iSofCode > 0 Then
                        hlSof = CBtrvTable(ONEHANDLE)
                        ilRet = btrOpen(hlSof, "", sgDBPath & "Sof.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                        If ilRet = BTRV_ERR_NONE Then
                            ilSofRecLen = Len(tlSof)  'btrRecordLength(hlSlf)  'Get and save record length
                            tlSofSrchKey.iCode = tlSlf.iSofCode
                            ilRet = btrGetEqual(hlSof, tlSof, ilSofRecLen, tlSofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
'                                slStr = ""
'                                For ilLoop = 2 To 0 Step -1
'                                    ilPos = InStr(1, tlSof.sAddr(ilLoop), ",", vbTextCompare)
'                                    If ilPos > 0 Then
'                                        slStr = Left$(tlSof.sAddr(ilLoop), ilPos - 1)
'                                        Exit For
'                                    End If
'                                Next ilLoop
'                                If slStr <> "" Then
                                    '2-18-05 reference short client name if exists vs long client name
                                    Print #hlTo, "Sales Office:" & slDelimiter & slDelimiter & Trim$(slClientName) & ", " & Trim$(tlSof.sName)
'                                Else
'                                    Print #hlTo, "Sales Office:" & slDelimiter & slDelimiter & Trim$(tlSof.sName)
'                                End If
                            Else
                                Print #hlTo, "Sales Office:" & slDelimiter & slDelimiter & Trim$(slClientName)
                            End If
                        Else
                            Print #hlTo, "Sales Office:" & slDelimiter & slDelimiter & Trim$(slClientName)
                        End If
                        ilRet = btrClose(hlSof)
                        btrDestroy hlSof
                    Else
                        Print #hlTo, "Sales Office:" & slDelimiter & slDelimiter & Trim$(slClientName)
                    End If
                End If
            End If
            ilRet = btrClose(hlSlf)
            btrDestroy hlSlf
        End If
'8915
#If programmatic <> 1 Then
        If Contract.lbcDemo(0).ListIndex >= 1 Then
            'Print #hlTo, "Demo:" & slDelimiter & slDelimiter & lbcDemo(0).List(lbcDemo(0).ListIndex)
            If Trim$(slSocEcoName) = "" Then
                Print #hlTo, "Demo:" & slDelimiter & slDelimiter & Contract.lbcDemo(0).List(Contract.lbcDemo(0).ListIndex)
            Else
                Print #hlTo, "Demo:" & slDelimiter & slDelimiter & Contract.lbcDemo(0).List(Contract.lbcDemo(0).ListIndex) & ", " & Trim$(slSocEcoName)
            End If
        Else
            Print #hlTo, "Demo:"
        End If
#Else
        If tgChfCntr.iMnfDemo(0) > 0 Then
            'Print #hlTo, "Demo:" & slDelimiter & slDelimiter & lbcDemo(0).List(lbcDemo(0).ListIndex)
            If Trim$(slSocEcoName) = "" Then
                Print #hlTo, "Demo:" & slDelimiter & slDelimiter & slDemoName
            Else
                Print #hlTo, "Demo:" & slDelimiter & slDelimiter & slDemoName & ", " & Trim$(slSocEcoName)
            End If
        Else
            Print #hlTo, "Demo:"
        End If
#End If
        'moved for ttp5230. xml needs also
'        llCntrPop = -1
'        For ilClf = ilStartRow To ilEndRow Step 1
'            ilLnRowNo = ilClf + 1
'            slLineType = mGetLineType(ilLnRowNo)
'            If (Not tgClfCntr(ilClf).iCancel) And ((slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E")) Then
'                If llCntrPop = -1 Then
'                    llCntrPop = tgClfCntr(ilClf).lPriDemoPop
'                ElseIf llCntrPop > 0 Then
'                    If llCntrPop <> tgClfCntr(ilClf).lPriDemoPop Then
'                        llCntrPop = 0
'                    End If
'                End If
'            End If
'        Next ilClf
        If llCntrPop > 0 Then
            Print #hlTo, "Population:" & slDelimiter & slDelimiter & Trim$(str$(llCntrPop))
        End If
'        If lbcAgency.ListIndex < 0 Then
'            slStr = ""
'        Else
'            slStr = lbcAgency.List(lbcAgency.ListIndex)
'        End If
'        tmAgfSrchKey.iCode = tgChfCntr.iAgfCode
'        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet = BTRV_ERR_NONE Then
            'Dan 8364
'            slStr = Trim$(tmAgf.sName)
            With tgExportFacts
            slStr = Trim$(tgExportAgf.sName)
            For ilLoop = 2 To 0 Step -1
                ilPos = InStr(1, tgExportAgf.sCntrAddr(ilLoop), ",", vbTextCompare)
                If ilPos > 0 Then
                    slStr = slStr & ", " & Left$(tgExportAgf.sCntrAddr(ilLoop), ilPos - 1)
                    Exit For
                End If
            Next ilLoop
            End With
'            For ilLoop = 2 To 0 Step -1
'                ilPos = InStr(1, tmAgf.sCntrAddr(ilLoop), ",", vbTextCompare)
'                If ilPos > 0 Then
'                    slStr = slStr & ", " & Left$(tmAgf.sCntrAddr(ilLoop), ilPos - 1)
'                    Exit For
'                End If
'            Next ilLoop
'        End If
        Print #hlTo, "Agency:" & slDelimiter & slDelimiter & slStr
        slStr = Trim$(tgChfCntr.sBuyer) & ", " & gFormatPhoneNo(tgChfCntr.sPhone)
        Print #hlTo, "Buyer:" & slDelimiter & slDelimiter & slStr
    End If
    'Get date span
    gUnpackDate tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1), slDate
    llMoStartDate = gDateValue(gObtainPrevMonday(slDate))
    gUnpackDateLong tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1), llStartDate
    gUnpackDateLong tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1), llEndDate
    'Dan M xmlproposal
'    If bmXmlOutput Then
'    ElseIf tmAgf.sCntrExptForm <> "O" Then
    'If tmAgf.sCntrExptForm <> "O" And Not bmXmlOutput Then
    If slExportType <> "O" And Not bmXmlOutput Then
        'Blank line
         Print #hlTo, ""
        'Titles
        '9459 add new 'hidden' column
        If blCSIAndHidden Then
            slRecord = "Line #" & slDelimiter & "Vehicle" & slDelimiter & "Pkg #" & slDelimiter & "Days & Times" & slDelimiter
        Else
            slRecord = "Line #" & slDelimiter & "Vehicle" & slDelimiter & "Days & Times" & slDelimiter
        End If
        slRecord = slRecord & "Spot Length" & slDelimiter & "Unit Rate" & slDelimiter & "Avg Aud" & slDelimiter & "Gross Total" & slDelimiter
        slRecord = slRecord & "Rating" & slDelimiter & "GRP" & slDelimiter & "CPP" & slDelimiter & "CPM" & slDelimiter
        slRecord = slRecord & "Gross Impr" & slDelimiter & "Rating Source" & slDelimiter & "Total Spots" & slDelimiter
        For llWeek = llMoStartDate To llEndDate Step 7
            slRecord = slRecord & Format(llWeek, "m/d/yy") & slDelimiter
        Next llWeek
        Print #hlTo, slRecord
    End If
    hlDnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlDnf, "", sgDBPath & "Dnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    ilUpper = (llEndDate - llMoStartDate) \ 7
    ReDim ilTNoSpots(0 To ilUpper) As Integer
    For ilLoop = 0 To ilUpper Step 1
        ilTNoSpots(ilLoop) = 0
    Next ilLoop
    For ilClf = ilStartRow To ilEndRow Step 1
        ilLnRowNo = ilClf + 1
        'D.S. 5/30/18 conditional compile
        #If programmatic <> 1 Then
            slLineType = mGetLineType(ilLnRowNo)
        #Else
            slLineType = mPBGetLineType(ilClf)
        #End If
        If (Not tgClfCntr(ilClf).iCancel) Then ' And ((slLineType <> "H")) Then
            ilStdWkCff = -1
            ilLnGen = False
            ilLastLnNo = -1
            ilUpper = (llEndDate - llMoStartDate) \ 7
            ReDim ilWkProc(0 To ilUpper) As Integer
            For ilLoop = 0 To ilUpper Step 1
                ilWkProc(ilLoop) = False
            Next ilLoop
            Do
                ilLnGen = False
                For llDate = llMoStartDate To llEndDate Step 7
                    'xml Dan 5/2/2011
                    blDailySpots = False
                    If (ilWkProc((llDate - llMoStartDate) \ 7) = False) Then
                        'D.S. 5/30/18
                        #If programmatic <> 1 Then
                            ilWkCff = mGetFlightSpotsDays(ilLnRowNo, llDate, ilOVDays(), llWkSpots, ilWkDays(), slWDFlag, ilDailySpots())
                        #Else
                            ilWkCff = mPBGetFlightSpotsDays(ilLnRowNo, llDate, ilOVDays(), llWkSpots, ilWkDays(), slWDFlag, ilDailySpots())
                        #End If
                        If (ilWkCff <> -1) And (ilWkProc((llDate - llMoStartDate) \ 7) = False) Then
                            'Test if week match
                            If ilLnGen Then
                                'Test if Week matched
                                ilMatch = True
                                If ilStdWkCff <> ilWkCff Then
                                    'D.S. 5/30/18 conditional compile
                                    #If programmatic <> 1 Then
                                        '10212 remove test of game
                                        'ilRet = mIsLnGame(ilLnRowNo)
                                        ilRet = 0
                                    #Else
                                        ilRet = mPBIsLnGame(ilClf)
                                    #End If
                                    'If Not mIsLnGame(ilLnRowNo) Then
                                    If Not ilRet Then
                                        If tgCffCntr(ilStdWkCff).CffRec.lActPrice = tgCffCntr(ilWkCff).CffRec.lActPrice Then
                                            For ilDay = 0 To 6 Step 1
                                                If ilStdWkDays(ilDay) <> ilWkDays(ilDay) Then
                                                    ilMatch = False
                                                    Exit For
                                                End If
                                            Next ilDay
                                        Else
                                            ilMatch = False
                                        End If
                                    End If
                                End If
                            End If
                            If (Not ilLnGen) Or ((ilLnGen) And (ilMatch)) Then
                                ilWkProc((llDate - llMoStartDate) \ 7) = True
                                ilStdWkCff = ilWkCff
                                slDays = ""
                                'D.S. 5/30/18 conditional compile
                                #If programmatic <> 1 Then
                                    ilRet = mIsLnGame(ilLnRowNo)
                                #Else
                                    ilRet = mPBIsLnGame(ilClf)
                                #End If
                                'If Not mIsLnGame(ilLnRowNo) Then
                                If Not ilRet Then
                                    If slWDFlag = "D" Then  'Daily
                                        'xml Dan 5/2/2011
                                        blDailySpots = True
                                        ilNoSpots = 0
                                        For ilDay = 0 To 6 Step 1
                                            If ilWkDays(ilDay) Then
                                                slDays = slDays & "1" & slDelimiter
                                            Else
                                                slDays = slDays & "0" & slDelimiter
                                            End If
                                            ilStdWkDays(ilDay) = ilWkDays(ilDay)
                                            'ilNoSpots = ilNoSpots + ilStdWkDays(ilDay)
                                            ilNoSpots = ilNoSpots + ilDailySpots(ilDay)
                                        Next ilDay
                                    Else
                                        For ilDay = 0 To 6 Step 1
                                            If ilWkDays(ilDay) Then
                                                slDays = slDays & "1" & slDelimiter
                                            Else
                                                slDays = slDays & "0" & slDelimiter
                                            End If
                                            ilStdWkDays(ilDay) = ilWkDays(ilDay)
                                        Next ilDay
                                        ilNoSpots = tgCffCntr(ilStdWkCff).CffRec.iSpotsWk 'llWkSpots
                                    End If
                                Else
                                    For ilDay = 0 To 6 Step 1
                                        slDays = slDays & "1" & slDelimiter
                                        ilStdWkDays(ilDay) = True
                                    Next ilDay
                                    ilNoSpots = llWkSpots
                                End If
                               ' ilTNoSpots((llDate - llMoStartDate) \ 7) = ilTNoSpots((llDate - llMoStartDate) \ 7) + ilNoSpots
                                '9490
                                If slLineType <> "H" Then
                                    ilTNoSpots((llDate - llMoStartDate) \ 7) = ilTNoSpots((llDate - llMoStartDate) \ 7) + ilNoSpots
                                    llSTNoSpots = llSTNoSpots + ilNoSpots
                                End If
                                llLnNoSpots = llLnNoSpots + ilNoSpots
                                'If tmAgf.sCntrExptForm = "O" Then
                                If slExportType = "O" Then
                                    If Not ilLnGen Then
'                                        'D.S. 5/30/18 conditional compile
'                                        #If programmatic <> 1 Then
'                                            slRecord = Trim$(smLnSave(1, ilLnRowNo)) & slDelimiter    'Vehicle Name
'                                        #Else
'                                            'get vehicle name from tgclfcntr(illnrowno - 1)
'                                            ilVef = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
'                                            If ilVef <> -1 Then
'                                                slRecord = Trim(tgMVef(ilVef).sName) & slDelimiter
'                                            Else
'                                                slRecord = "Vehicle Not Found" & slDelimiter
'                                            End If
'                                        #End If
                                        slRecord = slRecord & slDays    'part of days "& slDelimiter"
                                        ReDim ilRDFPercent(0 To 0) As Integer
                                        ReDim slRDFStartTime(0 To 0) As String
                                        ReDim slRDFEndTime(0 To 0) As String
                                        If ((tgClfCntr(ilClf).ClfRec.iStartTime(0) = 1) And (tgClfCntr(ilClf).ClfRec.iStartTime(1) = 0)) Then
                                            ilRdf = gBinarySearchRdf(tgClfCntr(ilClf).ClfRec.iRdfCode)
                                            If ilRdf <> -1 Then
                                                For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                                    If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                                        gUnpackTime tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), "A", "1", slStr
                                                        If StrComp(slStr, "12M", vbTextCompare) = 0 Then
                                                            slStr = "12AM"
                                                        ElseIf StrComp(slStr, "12N", vbTextCompare) = 0 Then
                                                            slStr = "12PM"
                                                        End If
                                                        '7787
                                                        slStr = Format$(slStr, "h:mmAM/PM")
                                                        'If UBound(ilRDFPercent) = 1 Then
                                                        If UBound(ilRDFPercent) = 0 Then
                                                            slRecord = slRecord & slStr & slDelimiter
                                                        End If
                                                        slRDFStartTime(UBound(ilRDFPercent)) = slStr
                                                        gUnpackTime tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), "A", "1", slStr
                                                        If StrComp(slStr, "12M", vbTextCompare) = 0 Then
                                                            slStr = "12AM"
                                                        ElseIf StrComp(slStr, "12N", vbTextCompare) = 0 Then
                                                            slStr = "12PM"
                                                        End If
                                                        '7787
                                                        slStr = Format$(slStr, "h:mmAM/PM")
                                                        'If UBound(ilRDFPercent) = 1 Then
                                                        If UBound(ilRDFPercent) = 0 Then
                                                            slRecord = slRecord & slStr & slDelimiter
                                                        End If
                                                        slRDFEndTime(UBound(ilRDFPercent)) = slStr
                                                        ilRDFPercent(UBound(ilRDFPercent)) = tgMRdf(ilRdf).iSpotPct(ilIndex)
                                                        ReDim Preserve slRDFStartTime(0 To UBound(ilRDFPercent) + 1) As String
                                                        ReDim Preserve slRDFEndTime(0 To UBound(ilRDFPercent) + 1) As String
                                                        ReDim Preserve ilRDFPercent(0 To UBound(ilRDFPercent) + 1) As Integer
                                                        'Exit For
                                                    End If
                                                Next ilIndex
                                            Else
                                                slRecord = slRecord & slDelimiter
                                                slRecord = slRecord & slDelimiter
                                            End If
                                        Else
                                            gUnpackTime tgClfCntr(ilClf).ClfRec.iStartTime(0), tgClfCntr(ilClf).ClfRec.iStartTime(1), "A", "1", slStr
                                            If StrComp(slStr, "12M", vbTextCompare) = 0 Then
                                                slStr = "12AM"
                                            ElseIf StrComp(slStr, "12N", vbTextCompare) = 0 Then
                                                slStr = "12PM"
                                            End If
                                            '7787
                                            slStr = Format$(slStr, "h:mmAM/PM")
                                            slRecord = slRecord & slStr & slDelimiter
                                            gUnpackTime tgClfCntr(ilClf).ClfRec.iEndTime(0), tgClfCntr(ilClf).ClfRec.iEndTime(1), "A", "1", slStr
                                            If StrComp(slStr, "12M", vbTextCompare) = 0 Then
                                                slStr = "12AM"
                                            ElseIf StrComp(slStr, "12N", vbTextCompare) = 0 Then
                                                slStr = "12PM"
                                            End If
                                            '7787
                                            slStr = Format$(slStr, "h:mmAM/PM")
                                            slRecord = slRecord & slStr & slDelimiter
                                        End If
                                        slRecord = slRecord & Trim$(str$(tgClfCntr(ilClf).ClfRec.iLen)) & slDelimiter
                                        'AQH
                                        'Add AQH, Book Name, Unit Rate
'                                        If tgSpf.sSAudData <> "H" Then
'                                            slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 0)
'                                        Else
'                                            slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 1)
'                                        End If
                                        If tgSpf.sSAudData = "H" Then
                                            slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 1)
                                        ElseIf tgSpf.sSAudData = "N" Then
                                            slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 2)
                                        ElseIf tgSpf.sSAudData = "U" Then
                                            slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 3)
                                        Else
                                            slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 0)
                                        End If
                                        slRecord = slRecord & slStr & slDelimiter
                                        slStr = ""
                                        'If package line, then determine book from hidden line
                                        If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
                                            For ilLoop = ilStartRow To ilEndRow Step 1
                                                If (Not tgClfCntr(ilLoop).iCancel) And (tgClfCntr(ilClf).ClfRec.iLine = tgClfCntr(ilLoop).ClfRec.iPkLineNo) Then
                                                    If tgClfCntr(ilLoop).iPriDemoBook > 0 Then
                                                        tlDnfSrchKey.iCode = tgClfCntr(ilLoop).iPriDemoBook
                                                        ilDnfRecLen = Len(tlDnf)  'btrRecordLength(hlSlf)  'Get and save record length
                                                        ilRet = btrGetEqual(hlDnf, tlDnf, ilDnfRecLen, tlDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            If InStr(1, Trim$(tlDnf.sBookName), "Radar", vbTextCompare) > 0 Then
                                                                If slStr = "" Then
                                                                    slStr = "R"
                                                                Else
                                                                    If slStr <> "R" Then
                                                                        slStr = "M"
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            Else
                                                                If slStr = "" Then
                                                                    slStr = "A"
                                                                Else
                                                                    If slStr <> "A" Then
                                                                        slStr = "M"
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        If slStr <> "" Then
                                                            If slStr <> "N" Then
                                                                slStr = "M"
                                                                Exit For
                                                            End If
                                                        End If
                                                        slStr = "N"
                                                    End If
                                                End If
                                            Next ilLoop
                                            If slStr = "N" Then
                                                slStr = ""
                                            End If
                                        Else
                                            If tgClfCntr(ilClf).iPriDemoBook > 0 Then
                                                tlDnfSrchKey.iCode = tgClfCntr(ilClf).iPriDemoBook
                                                ilDnfRecLen = Len(tlDnf)  'btrRecordLength(hlSlf)  'Get and save record length
                                                ilRet = btrGetEqual(hlDnf, tlDnf, ilDnfRecLen, tlDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet = BTRV_ERR_NONE Then
                                                    slStr = Trim$(tlDnf.sBookName)
                                                    If InStr(1, slStr, "Radar", vbTextCompare) > 0 Then
                                                        slStr = "R"
                                                    Else
                                                        slStr = "A"
                                                    End If
                                                End If
                                            End If
                                        End If
                                        slRecord = slRecord & slStr & slDelimiter
                                        slStr = gLongToStrDec(tgCffCntr(ilStdWkCff).CffRec.lActPrice, 2)
                                        slRecord = slRecord & slStr & slDelimiter
                                        '% US Coverage
                                        slRecord = slRecord & slDelimiter
                                        'Number of Affiliates
                                        slRecord = slRecord & slDelimiter
                                        'Week spot count of zero
                                        For llWeek = llMoStartDate To llDate - 1 Step 7
                                            ilWkProc((llWeek - llMoStartDate) \ 7) = True
                                            slRecord = slRecord & "0" & slDelimiter
                                        Next llWeek
                                    End If
                                    ilLnGen = True
                                    slRecord = slRecord & Trim$(str$(ilNoSpots)) & slDelimiter
                                Else 'not type 'O'
                                    If Not ilLnGen Then
                                        'dan proposal xml
                                        If bmXmlOutput Then
                                            slRecord = ilClf & slDelimiter
                                        Else
                                            slRecord = Trim$(str$(tgClfCntr(ilClf).ClfRec.iLine)) & slDelimiter
                                            '9459 if this is a package line, we need this line # to print hidden lines
                                            If slLineType <> "H" Then
                                                ilLineNumber = tgClfCntr(ilClf).ClfRec.iLine
                                            End If
                                        End If
                                        For ilDay = 0 To 6 Step 1
                                            If ilStdWkDays(ilDay) = False Then
                                                ilWkDays(ilDay) = 0
                                            Else
                                                ilWkDays(ilDay) = 1
                                            End If
                                        Next ilDay
                                        slStr = gDayNames(ilWkDays(), slCffXDay(), 1, slEDIDays) & " "
                                        'dan M 5/2/11 xml add days of avail
                                        If bmXmlOutput Then
                                            slRecord = slRecord & slDays & blDailySpots & slDelimiter
                                            'just write out times
                                            slStr = ""
                                        End If
                                        If ((tgClfCntr(ilClf).ClfRec.iStartTime(0) = 1) And (tgClfCntr(ilClf).ClfRec.iStartTime(1) = 0)) Then
                                            ilRdf = gBinarySearchRdf(tgClfCntr(ilClf).ClfRec.iRdfCode)
                                            If ilRdf <> -1 Then
                                                For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                                    If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                                        gUnpackTime tgMRdf(ilRdf).iStartTime(0, ilIndex), tgMRdf(ilRdf).iStartTime(1, ilIndex), "A", "1", slTime
                                                        If StrComp(slTime, "12M", vbTextCompare) = 0 Then
                                                            slTime = "12AM"
                                                        ElseIf StrComp(slTime, "12N", vbTextCompare) = 0 Then
                                                            slTime = "12PM"
                                                        End If
                                                        slStr = slStr & slTime & "-"
                                                        gUnpackTime tgMRdf(ilRdf).iEndTime(0, ilIndex), tgMRdf(ilRdf).iEndTime(1, ilIndex), "A", "1", slTime
                                                        If StrComp(slTime, "12M", vbTextCompare) = 0 Then
                                                            slTime = "12AM"
                                                        ElseIf StrComp(slTime, "12N", vbTextCompare) = 0 Then
                                                            slTime = "12PM"
                                                        End If
                                                        slStr = slStr & slTime & ", "
                                                    End If
                                                Next ilIndex
                                                slStr = Left(slStr, Len(slStr) - 2)
                                            End If
                                        Else
                                            gUnpackTime tgClfCntr(ilClf).ClfRec.iStartTime(0), tgClfCntr(ilClf).ClfRec.iStartTime(1), "A", "1", slTime
                                            If StrComp(slTime, "12M", vbTextCompare) = 0 Then
                                                slTime = "12AM"
                                            ElseIf StrComp(slTime, "12N", vbTextCompare) = 0 Then
                                                slTime = "12PM"
                                            End If
                                            slStr = slStr & slTime & "-"
                                            gUnpackTime tgClfCntr(ilClf).ClfRec.iEndTime(0), tgClfCntr(ilClf).ClfRec.iEndTime(1), "A", "1", slTime
                                            If StrComp(slTime, "12M", vbTextCompare) = 0 Then
                                                slTime = "12AM"
                                            ElseIf StrComp(slTime, "12N", vbTextCompare) = 0 Then
                                                slTime = "12PM"
                                            End If
                                            slStr = slStr & slTime
                                        End If
                                        slDaysTimes = slStr
                                        If ilLastLnNo = -1 Then
                                            ilLastLnNo = tgClfCntr(ilClf).ClfRec.iLine
                                            'D.S. 5/30/18 conditional compile
                                            'slRecord = slRecord & Trim$(smLnSave(1, ilLnRowNo)) & slDelimiter    'Vehicle Name
                                            #If programmatic <> 1 Then
                                                '9459 add "H" if hidden line
                                                If slLineType = "H" And blCSIAndHidden Then
                                                    slRecord = slRecord & Trim$(smLnSave(1, ilLnRowNo)) & slDelimiter & ilLineNumber & slDelimiter    'Vehicle Name
                                                ElseIf blCSIAndHidden Then
                                                    slRecord = slRecord & Trim$(smLnSave(1, ilLnRowNo)) & slDelimiter & "" & slDelimiter
                                                Else
                                                    slRecord = slRecord & Trim$(smLnSave(1, ilLnRowNo)) & slDelimiter    'Vehicle Name
                                                End If
                                            #Else
                                                'get vehicle name from tgclfcntr(illnrowno - 1)
                                                ilVef = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
                                                If ilVef <> -1 Then
                                                    slRecord = slRecord & Trim(tgMVef(ilVef).sName) & slDelimiter
                                                Else
                                                    slRecord = slRecord & "Vehicle Not Found" & slDelimiter
                                                End If
                                            #End If
                                            
                                            slRecord = slRecord & slDaysTimes & slDelimiter
                                            slRecord = slRecord & Trim$(str$(tgClfCntr(ilClf).ClfRec.iLen)) & slDelimiter
                                            slStr = gLongToStrDec(tgCffCntr(ilStdWkCff).CffRec.lActPrice, 2)
                                            slRecord = slRecord & slStr & slDelimiter
'                                            If tgSpf.sSAudData <> "H" Then
'                                                slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 0)
'                                            Else
'                                                slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 1)
'                                            End If
                                            If tgSpf.sSAudData = "H" Then
                                                slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 1)
                                            ElseIf tgSpf.sSAudData = "N" Then
                                                slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 2)
                                            ElseIf tgSpf.sSAudData = "U" Then
                                                slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 3)
                                            Else
                                                slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 0)
                                            End If
                                            slRecord = slRecord & slStr & slDelimiter
                                            slStr = gLongToStrDec(tgClfCntr(ilClf).lCost, 2)
                                            slRecord = slRecord & slStr & slDelimiter
                                            '10/28/14: Handle 2 place rating
                                            'slStr = gIntToStrDec(tgClfCntr(ilClf).iPriDemoAvgRating, 1)
                                            'D.S. 5/30/18 conditional compile
                                            #If programmatic <> 1 Then
                                                slStr = smCntr1or2PlaceRating
                                            #Else
                                                slStr = gSet1or2PlaceRating(tgChfCntr.iAgfCode)
                                            #End If
                                            'If smCntr1or2PlaceRating <> "2" Then
                                            If slStr <> "2" Then
                                                slStr = gIntToStrDec(tgClfCntr(ilClf).iPriDemoAvgRating, 1)
                                            Else
                                                slStr = gIntToStrDec(tgClfCntr(ilClf).iPriDemoAvgRating, 2)
                                            End If
                                            slRecord = slRecord & slStr & slDelimiter
                                            slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoGRP, 1)
                                            slRecord = slRecord & slStr & slDelimiter
                                            slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoCPP, 0)
                                            slRecord = slRecord & slStr & slDelimiter
                                            slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoCPM, 2)
                                            slRecord = slRecord & slStr & slDelimiter
'                                            If tgSpf.sSAudData <> "H" Then
'                                                slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoGrImp, 0)
'                                            Else
'                                                slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoGrImp, 1)
'                                            End If
                                            If tgSpf.sSAudData = "H" Then
                                                slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoGrImp, 1)
                                            ElseIf tgSpf.sSAudData = "N" Then
                                                slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoGrImp, 2)
                                            ElseIf tgSpf.sSAudData = "U" Then
                                                slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoGrImp, 3)
                                            Else
                                                slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoGrImp, 0)
                                            End If
                                            slRecord = slRecord & slStr & slDelimiter
                                            slStr = " "
                                            If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
                                                'DAN XML and ttp 5230 insert population
                                                If bmXmlOutput Then
                                                    slRecord = slRecord & llCntrPop & slDelimiter & "Y" & slDelimiter
                                                End If
                                                For ilLoop = ilStartRow To ilEndRow Step 1
                                                    If (Not tgClfCntr(ilLoop).iCancel) And (tgClfCntr(ilClf).ClfRec.iLine = tgClfCntr(ilLoop).ClfRec.iPkLineNo) Then
                                                        If tgClfCntr(ilLoop).iPriDemoBook > 0 Then
                                                            tlDnfSrchKey.iCode = tgClfCntr(ilLoop).iPriDemoBook
                                                            ilDnfRecLen = Len(tlDnf)  'btrRecordLength(hlSlf)  'Get and save record length
                                                            ilRet = btrGetEqual(hlDnf, tlDnf, ilDnfRecLen, tlDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                            If ilRet = BTRV_ERR_NONE Then
                                                                If slStr = " " Then
                                                                    slStr = Trim$(tlDnf.sBookName)
                                                                Else
                                                                    If InStr(1, slStr, Trim$(tlDnf.sBookName), vbTextCompare) <= 0 Then
                                                                        slStr = slStr & " & " & Trim$(tlDnf.sBookName)
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Next ilLoop
                                            Else
                                                'dan xml ttp5230 insert population
                                                If bmXmlOutput Then
                                                    slRecord = slRecord & llCntrPop & slDelimiter & "N" & slDelimiter
                                                End If
                                                If tgClfCntr(ilClf).iPriDemoBook > 0 Then
                                                    tlDnfSrchKey.iCode = tgClfCntr(ilClf).iPriDemoBook
                                                    ilDnfRecLen = Len(tlDnf)  'btrRecordLength(hlSlf)  'Get and save record length
                                                    ilRet = btrGetEqual(hlDnf, tlDnf, ilDnfRecLen, tlDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                    If ilRet = BTRV_ERR_NONE Then
                                                        slStr = Trim$(tlDnf.sBookName)
                                                    End If
                                                End If
                                            End If 'linetype O A E (package?)
                                            slRecord = slRecord & slStr & slDelimiter
                                            'dan m calculate spots in case not scheduled...5/3/2011
                                            'D.S. 5/30/18 conditional compile
                                            #If programmatic <> 1 Then
                                                mCalculateTotalSpots
                                            #Else
                                                tgClfCntr(ilClf).lTSpots = llLnNoSpots
                                            #End If
                                            slStr = gLongToStrDec(tgClfCntr(ilClf).lTSpots, 0)
                                            slRecord = slRecord & slStr & slDelimiter
                                        Else
                                            slRecord = slRecord & slDelimiter   'Vehicle
                                            slRecord = slRecord & slDaysTimes & slDelimiter   'Days & Times
                                            slRecord = slRecord & slDelimiter   'Spot Length
                                            slStr = gLongToStrDec(tgCffCntr(ilStdWkCff).CffRec.lActPrice, 2)
                                            slRecord = slRecord & slStr & slDelimiter
                                            slRecord = slRecord & slDelimiter   'Avg Aud
                                            slRecord = slRecord & slDelimiter   'Total Cost
                                            slRecord = slRecord & slDelimiter   'Rating
                                            slRecord = slRecord & slDelimiter   'GRP
                                            slRecord = slRecord & slDelimiter   'CPP
                                            slRecord = slRecord & slDelimiter   'CPM
                                            slRecord = slRecord & slDelimiter   'GrImp
                                            'dan xmlProposal, fill for population & package
                                            If bmXmlOutput Then
                                                slRecord = slRecord & slDelimiter & slDelimiter
                                            End If
                                            slRecord = slRecord & slDelimiter   'Rating Source
                                            slRecord = slRecord & slDelimiter   'Total Spots
                                        End If
                                        For llWeek = llMoStartDate To llDate - 1 Step 7
                                            ilWkProc((llWeek - llMoStartDate) \ 7) = True
                                            slRecord = slRecord & "0" & slDelimiter
                                        Next llWeek
                                    End If ' if not ilGen
                                    ilLnGen = True
                                    slRecord = slRecord & Trim$(str$(ilNoSpots)) & slDelimiter
                                End If 'type of output O and not O
                            Else
                                If ilLnGen Then
                                    'Non air week or does not match pattern (rate or days)
                                    slRecord = slRecord & "0" & slDelimiter
                                End If
                            End If
                        Else
                            If ilLnGen Then
                                'Non air week or does not match pattern (rate or days)
                                slRecord = slRecord & "0" & slDelimiter
                            End If
                        End If
                    Else
                        If ilLnGen Then
                            'Non air week or does not match pattern (rate or days)
                            slRecord = slRecord & "0" & slDelimiter
                        End If
                    End If
                Next llDate
                If ilLnGen Then
'                    If tmAgf.sCntrExptForm = "O" Then
'                        'Add AQH, Book Name, Unit Rate
'                        If tgSpf.sSAudData <> "H" Then
'                            slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 0)
'                        Else
'                            slStr = gLongToStrDec(tgClfCntr(ilClf).lPriDemoAvgAud, 1)
'                        End If
'                        slRecord = slRecord & slStr & slDelimiter
'                        slStr = ""
'                        'If package line, then determine book from hidden line
'                        If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
'                            For ilLoop = ilStartRow To ilEndRow Step 1
'                                If (Not tgClfCntr(ilLoop).iCancel) And (tgClfCntr(ilClf).ClfRec.iLine = tgClfCntr(ilLoop).ClfRec.iPkLineNo) Then
'                                    If tgClfCntr(ilLoop).iPriDemoBook > 0 Then
'                                        tlDnfSrchKey.iCode = tgClfCntr(ilLoop).iPriDemoBook
'                                        ilDnfRecLen = Len(tlDnf)  'btrRecordLength(hlSlf)  'Get and save record length
'                                        ilRet = btrGetEqual(hlDnf, tlDnf, ilDnfRecLen, tlDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                                        If ilRet = BTRV_ERR_NONE Then
'                                            If InStr(1, Trim$(tlDnf.sBookName), "Radar", vbTextCompare) > 0 Then
'                                                If slStr = "" Then
'                                                    slStr = "R"
'                                                Else
'                                                    If slStr <> "R" Then
'                                                        slStr = "M"
'                                                        Exit For
'                                                    End If
'                                                End If
'                                            Else
'                                                If slStr = "" Then
'                                                    slStr = "A"
'                                                Else
'                                                    If slStr <> "A" Then
'                                                        slStr = "M"
'                                                        Exit For
'                                                    End If
'                                                End If
'                                            End If
'                                        End If
'                                    Else
'                                        If slStr <> "" Then
'                                            If slStr <> "N" Then
'                                                slStr = "M"
'                                                Exit For
'                                            End If
'                                        End If
'                                        slStr = "N"
'                                    End If
'                                End If
'                            Next ilLoop
'                            If slStr = "N" Then
'                                slStr = ""
'                            End If
'                        Else
'                            If tgClfCntr(ilClf).iPriDemoBook > 0 Then
'                                tlDnfSrchKey.iCode = tgClfCntr(ilClf).iPriDemoBook
'                                ilDnfRecLen = Len(tlDnf)  'btrRecordLength(hlSlf)  'Get and save record length
'                                ilRet = btrGetEqual(hlDnf, tlDnf, ilDnfRecLen, tlDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                                If ilRet = BTRV_ERR_NONE Then
'                                    slStr = Trim$(tlDnf.sBookName)
'                                    If InStr(1, slStr, "Radar", vbTextCompare) > 0 Then
'                                        slStr = "R"
'                                    Else
'                                        slStr = "A"
'                                    End If
'                                End If
'                            End If
'                        End If
'                        slRecord = slRecord & slStr & slDelimiter
'                        slStr = gLongToStrDec(tgCffCntr(ilStdWkCff).CffRec.lActPrice, 2)
'                        slRecord = slRecord & slStr & slDelimiter
'                        slRecord = slRecord & slDelimiter
'                        slRecord = slRecord & slDelimiter
'                    End If
                   ' If tmAgf.sCntrExptForm = "O" Then
                   '7803 I've loaded hidden lines. remove here!
                   '9459 allow hidden lines in some cases
'                   If slLineType <> "H" Then
                    If slLineType <> "H" Or blCSIAndHidden Then
                        If slExportType = "O" Then
                            'If UBound(ilRDFPercent) > 2 Then
                            If UBound(ilRDFPercent) > 1 Then
                                slInputStr = slRecord
                                'For ilRdf = 1 To UBound(ilRDFPercent) - 1 Step 1
                                For ilRdf = 0 To UBound(ilRDFPercent) - 1 Step 1
                                    'Parse, then remake records
                                    'Field 9=Start Time, Field 10=End Time, Field 17=First week
                                    slRecord = ""
                                    For ilItemNo = 1 To 8 Step 1
                                        ilRet = gParseItem(slInputStr, ilItemNo, slDelimiter, slStr)
                                        slRecord = slRecord & slStr & slDelimiter
                                    Next ilItemNo
                                    ilRet = gParseItem(slInputStr, 9, slDelimiter, slStr)
                                    slRecord = slRecord & slRDFStartTime(ilRdf) & slDelimiter
                                    ilRet = gParseItem(slInputStr, 10, slDelimiter, slStr)
                                    slRecord = slRecord & slRDFEndTime(ilRdf) & slDelimiter
                                    For ilItemNo = 11 To 16 Step 1
                                        ilRet = gParseItem(slInputStr, ilItemNo, slDelimiter, slStr)
                                        slRecord = slRecord & slStr & slDelimiter
                                    Next ilItemNo
                                    ilItemNo = 17
                                    Do
                                        ilRet = gParseItem(slInputStr, ilItemNo, slDelimiter, slStr)
                                        If ilRet <> CP_MSG_NONE Then
                                            Exit Do
                                        End If
                                        slStr = gMulStr(slStr, gIntToStrDec(ilRDFPercent(ilRdf), 2))
                                        slStr = gDivStr(slStr, "100.00")
                                        slRecord = slRecord & slStr & slDelimiter
                                        ilItemNo = ilItemNo + 1
                                    Loop
                                    Print #hlTo, slRecord
                                Next ilRdf
                            ' C X
                            Else
                                Print #hlTo, slRecord
                            End If
                        Else
                            'Dan M 4-27-11 don't print if xml-but record passed is needed clf info
                            'Print #hlTo, slRecord
                            mPrintOutFileStoreXml hlTo, slRecord, True
                        End If
                        'write hidden lines
                    ElseIf blTreatAsReps Then
                        mPrintOutFileStoreXml hlTo, slRecord, True
                    End If
                End If
            Loop While ilLnGen
        End If ' not cancelled
    Next ilClf
    'exclude xml dan m 5/02/11
    'If tmAgf.sCntrExptForm <> "O" And Not bmXmlOutput Then
    If slExportType <> "O" And Not bmXmlOutput Then
        'Genertae totals
        slRecord = slDelimiter & slDelimiter & "Totals" & slDelimiter & slDelimiter & slDelimiter & slDelimiter
        '9459
        If blCSIAndHidden Then
            slRecord = slRecord & slDelimiter
        End If
        'Dan 8364
        'D.S. 5/30/18 condtional compile
        ' slStr = Trim$(smTotalDollars)  'Trim$(edcGross.Text)
        #If programmatic <> 1 Then
            slStr = Trim$(smTotalDollars)
        #Else
            slStr = gLongToStrDec(tgChfCntr.lInputGross, 2)
        #End If
        If slStr = "" Then
            slStr = "0.00"
        Else
            If InStr(1, slStr, ".", vbTextCompare) <= 0 Then
                slStr = slStr & ".00"
            End If
        End If
        slRecord = slRecord & slStr & slDelimiter
        '10/28/14: Handle 2 place rating
        'slStr = gIntToStrDec(imCPriDemoRating, 1)
        '9460
        #If programmatic <> 1 Then
            With tgExportFacts
                imCPriDemoRating = .imCPriDemoRating
                lmCPriDemoGRP = .lmCPriDemoGRP
                lmCPriDemoCPP = .lmCPriDemoCPP
                lmCPriDemoCPM = .lmCPriDemoCPM
                lmCPriDemoGrImp = .lmCPriDemoGrImp
            End With
        #End If
        'D.S. 5/30/18 conditional compile
        #If programmatic <> 1 Then
            slStr = smCntr1or2PlaceRating
        #Else
            slStr = gSet1or2PlaceRating(tgChfCntr.iAgfCode)
        #End If
        'If smCntr1or2PlaceRating <> "2" Then
        If slStr <> "2" Then
            slStr = gIntToStrDec(imCPriDemoRating, 1)
        Else
            slStr = gIntToStrDec(imCPriDemoRating, 2)
        End If
        slRecord = slRecord & slStr & slDelimiter
        slStr = gLongToStrDec(lmCPriDemoGRP, 1)
        slRecord = slRecord & slStr & slDelimiter
        slStr = gLongToStrDec(lmCPriDemoCPP, 0)
        slRecord = slRecord & slStr & slDelimiter
        slStr = gLongToStrDec(lmCPriDemoCPM, 2)
        slRecord = slRecord & slStr & slDelimiter
'        If tgSpf.sSAudData <> "H" Then
'            slStr = gLongToStrDec(lmCPriDemoGrImp, 0)
'        Else
'            slStr = gLongToStrDec(lmCPriDemoGrImp, 1)
'        End If
        If tgSpf.sSAudData = "H" Then
            slStr = gLongToStrDec(lmCPriDemoGrImp, 1)
        ElseIf tgSpf.sSAudData = "N" Then
            slStr = gLongToStrDec(lmCPriDemoGrImp, 2)
        ElseIf tgSpf.sSAudData = "U" Then
            slStr = gLongToStrDec(lmCPriDemoGrImp, 3)
        Else
            slStr = gLongToStrDec(lmCPriDemoGrImp, 0)
        End If
        slRecord = slRecord & slStr & slDelimiter
        slRecord = slRecord & slDelimiter
        '9128
        'slStr = gIntToStrDec(ilSTNoSpots, 0)
        slStr = gLongToStrDec(llSTNoSpots, 0)
        slRecord = slRecord & slStr & slDelimiter
        For llDate = llMoStartDate To llEndDate Step 7
            slStr = gIntToStrDec(ilTNoSpots((llDate - llMoStartDate) \ 7), 0)
            slRecord = slRecord & slStr & slDelimiter
        Next llDate
        Print #hlTo, slRecord
        'D.S. 5/30/18 conditional compile
        #If programmatic <> 1 Then
            If (Trim$(smComment(5)) <> "") And (((imCommShowOn(5, 0) = 0) And (sgCntrScreen <> "Orders")) Or ((imCommShowOn(5, 1) = 0) And (sgCntrScreen = "Orders"))) Then
                slRecord = Trim$(smComment(5))
                Print #hlTo, slRecord
            End If
        #Else
            slRecord = Trim(ProgrammaticBuy.edcComment.Text)
            Print #hlTo, slRecord
        #End If
    End If
    'Dan 4/27/2011 add proposal xml
    If bmXmlOutput Then
    'whew we made it to here.  I have a slRecord filled with the information from above.
    'agency xml is relatively simple.  It falls through most of this early stuff
        '6768
        ilWeekSpan = DateDiff("W", llMoStartDate, llEndDate)
        ilWeekSpan = DateAdd("W", 1, ilWeekSpan)
        If ilWeekSpan > 53 Then
            If MsgBox("ProposalXml cannot accept files that cover more than 53 weeks.  This file spans " & ilWeekSpan & " weeks. Continue anyway?", vbExclamation + vbYesNo, "ProposalXml Issue") = vbNo Then
                bmXmlOutput = False
                GoTo CLEANUP
            End If
        End If
        '8032 send to web?
        If mIsStationXmlAllowedForOrder(STATIONXMLMARKETRON) Then
            If MsgBox("Send " & STATIONXMLMARKETRONNAME & " to Web? 'Yes' will send and create a text file.  'No' will only create the text file.", vbYesNo, "Send Station Xml to web service?") = vbNo Then
                bgStationXmlWebServiceAllowed(StationXmlIndexes.MARKETRONEO) = False
            End If
        End If
        '7673  may have to write out each line as separate file
       ' ilNextCounter = 0
       'AgencyXml not needed.  We'll do first and then get out
        For c = 0 To UBound(smClfInfo) - 1
            'V81 further complicated:  skip if line is not a rep
            blContinuePrint = False
            '7803 packages' hidden lines get displayed, not the package itself
            blIsRepPackage = False
            If blTreatAsReps Then
                ilRet = mReturnLineNumber(c)
                blContinuePrint = gIsSpotInsertionVehicleByLine(ilRet, blIsRepPackage)
            Else
                blContinuePrint = True
            End If
            '7803 ignore packages, we've got them individually now
            If blContinuePrint And blIsRepPackage = False Then
                '7687  I need to combine the same vehicle if reps.  I handle this in mWriteXml, but then I end up with a file with no lines in it for the nth times that vehicle is in a line.
                'this stops the file from being created.
                If Not blTreatAsReps Then
                    ilRepCounter = -1
                Else
                    ilRepCounter = c
                End If
                slRepVehicle = gRepLineName(ilRepCounter)
                'agency xml.  Allowed to pass through!
                If gAllowRepLine(c, ilRepCounter, slRepVehicle) Then
                    ilNextCounter = ilNextCounter + 1
                    '8032 move below
'                    '6768 delete old and open new now
'                    On Error GoTo imcExportErr
'                    'StationXml each vehicle is separate
'                    If blTreatAsReps Then
'                        '8032
''                        slToFile = Replace(slRepFileTemp, REPREPLACE, ilNextCounter)
'                         slToFile = Replace(slRepFileTemp, REPREPLACE, ilNextCounter)
'                        Select Case igStationXmlChoice
'                            Case STATIONXMLWIDEORBIT
'                                slToFile = smStationXMLWOPath & slToFile
'                            Case STATIONXMLMARKETRON
'                                slToFile = smStationXMLBasePath & STATIONXMLMARKETRONNAME & "\" & slToFile
'                        End Select
'                    End If
                    '8032 pass agency, but make sure station is allowed
                    If igStationXmlChoice = AGENCYXML Or mIsStationXmlAllowedForOrder(igStationXmlChoice) Then
                        '6768 delete old and open new now
                        'On Error GoTo imcExportErr
                        'StationXml each vehicle is separate
                        If blTreatAsReps Then
                            '8032
    '                        slToFile = Replace(slRepFileTemp, REPREPLACE, ilNextCounter)
                             slToFile = Replace(slRepFileTemp, REPREPLACE, ilNextCounter)
                             '8364 Dan added tgExportfacts
                            Select Case igStationXmlChoice
                                Case STATIONXMLWIDEORBIT
                                    slToFile = tgExportFacts.smStationXMLWOPath & slToFile
                                Case STATIONXMLMARKETRON
                                    slToFile = tgExportFacts.smStationXMLBasePath & STATIONXMLMARKETRONNAME & "\" & slToFile
                            End Select
                        End If
                        ilRet = 0
                        'slDateTime = FileDateTime(slToFile)
                        ilRet = gFileExist(slToFile)
                        If ilRet = FILEEXISTS Then
                        '9202
                        On Error GoTo KILLERR1
                            Kill slToFile
                            If ilRet = KILLISSUE Then
                                On Error GoTo KILLERR2
                                Sleep 2000
                                If gFileExist(slToFile) = FILEEXISTS Then
                                    Kill slToFile
                                End If
                            End If
                        End If
                        On Error GoTo 0
                        ilRet = 0
                        '9744 move to after validate
'                        If csiXMLStart("", "", "F", slToFile, Chr$(13) + Chr$(10), "") <> 0 Then
'                            gMsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
'                            Exit Function
'                        End If
                        '7580 must pass new value
                        'If mXmlSetBasicAndValidate(llMoStartDate, llEndDate, slDelimiter, slStr, tmAgf) Then
                        '7673 pass c-only run on '0'
                        'If mXmlSetBasicAndValidate(llMoStartDate, llEndDate, slDelimiter, slStr, tmAgf, imStep1Index) Then
                        '7959 fix 7673. passing c wan't good enough. now pass boolean to keep track
                        'AgencyXml doesn't care about stepIndex or blWasRun: it only comes here one time.  I set the counter to -1 to say 'treat as normal'
                       ' If mXmlSetBasicAndValidate(slRepVehicle, llMoStartDate, llEndDate, slDelimiter, slStr, tmagf, imStep1Index, blWasRun) Then
                       'dan to do
                        If mXmlSetBasicAndValidate(slRepVehicle, llMoStartDate, llEndDate, slDelimiter, slStr, tgExportAgf, tgExportFacts.imStep1Index, blWasRun) Then
                            '9744
                            If csiXMLStart("", "", "F", slToFile, Chr$(13) + Chr$(10), "") <> 0 Then
                                gMsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
                                Exit Function
                            End If
                            '7673 must pass which line for reps; normal gets -1
                           ' If Not mWriteXml(slXmlErrorMessage) Then
                            If Not blTreatAsReps Then
                                c = -1
                            End If
                            '7803 added package info
                            '8940 supress rates
                          '  If Not mWriteXml(slXmlErrorMessage, c) Then
                            If Not mWriteXml(slXmlErrorMessage, c, blZeroRates) Then
                                MsgBox "Xml file not created.  There was a problem creating the xml file: " & slXmlErrorMessage, vbExclamation, "Error"
                                GoTo CLEANUP
                            End If
                        Else
                            'agencyxml get out. stationxml, go to next vehicle
                            Select Case igStationXmlChoice
                                Case AGENCYXML
                                    '9744
'                                    slStr = "Failed "
'                                    gMsgBox slStr & " to validate.  Xml file(s) not created.", vbExclamation, "Error"
                                    gMsgBox "Failed to validate.  Xml file(s) not created. " & slStr, vbExclamation, "Error"
                                    GoTo CLEANUP
                                Case STATIONXMLWIDEORBIT
                                    '8032 won't update
                                    '10455 it's blocking any vehicle after an invalid one.
                                   ' bgStationXmlOrderAllowed(StationXmlIndexes.WO) = False
                                    slAgencyErrorMessage = slAgencyErrorMessage & " WideOrbit files failed-" & slStr
                                Case STATIONXMLMARKETRON
                                    '10455
                                   ' bgStationXmlOrderAllowed(StationXmlIndexes.MARKETRONEO) = False
                                    slAgencyErrorMessage = slAgencyErrorMessage & " Marketron files failed-" & slStr
                            End Select
                        End If
                    Else
                        blIsMissingStationRights = True
                    End If
                    '7673 this is how we only write 1 (original)
                    'AgencyXml we're done!
                    If Not blTreatAsReps Then
                        Exit For
                    End If
                End If
            End If ' continue print
        Next c
        '7687 now write to chf that xml was created
        'Dan 11/16/15 don't update if nothing written due to error:  added ilNextCounter
        '8032- one station file type fails, still update- but won't update that one!
'        If Len(slAgencyErrorMessage) > 0 Then
'            gMsgBox slAgencyErrorMessage, vbExclamation, "Error"
'        Else
'            If blTreatAsReps And ilNextCounter > 0 Then
'                gUpdateStationXml hmChf, hmClf
'            End If
'        End If
        '8032
        If blIsMissingStationRights Then
            slAgencyErrorMessage = slAgencyErrorMessage & " Some vehicles marked as station xml did not go out."
            If Not ogStationXmlErrors Is Nothing Then
                ogStationXmlErrors.WriteWarning slAgencyErrorMessage, True
            End If
        End If
        If Len(slAgencyErrorMessage) > 0 Then
            gMsgBox slAgencyErrorMessage, vbExclamation, "Error"
        End If
        If mIsStationXmlAllowedForOrder(, True) Then
            If blTreatAsReps And ilNextCounter > 0 Then
                '8364
               ' gUpdateStationXml hmCHF, hmClf
                gUpdateStationXml tgExportFacts.hmCHF, tgExportFacts.hmClf
            End If
        End If
    Else
        Close #hlTo
    End If
    'Dan 9/12/11 all proposals get emailed
    'file exists  '8032  this is the last file created, so if it fails but previous successful, never see this message.  But this should be rare.
    '8886
   ' If Dir(slToFile) > "" Then
    If gFileExist(slToFile) = FILEEXISTS Then
        '7673 change file back to normal for message
        If blTreatAsReps Then
            slToFile = Replace(slRepFileTemp, "-" & REPREPLACE, "")
            If Len(slAgencyErrorMessage) > 0 Then
                slToFile = "Some files were not exported as " & slToFile
            Else
                '8120
                'slToFile = "Files exported as " & slToFile
            End If
            '8032
            If Not ogStationXmlErrors Is Nothing Then
                If Len(slAgencyErrorMessage) > 0 Then
                    ogStationXmlErrors.WriteWarning slAgencyErrorMessage, True
                Else
                    For c = StationXmlIndexes.WO To StationXmlIndexes.MARKETRONEO
                        If mIsStationXmlAllowedForOrder(c, True) Then
                            ogStationXmlErrors.WriteFacts "File exported: " & slToFile, True
                        End If
                    Next c
                End If
            End If
        Else
            '8120
            'slToFile = "File exported to " & slToFile
        End If
        mExportProposal = slToFile
    End If

CLEANUP:
    ilRet = btrClose(hlDnf)
    btrDestroy hlDnf
    If Not bmXmlOutput Then
        Close #hlTo
    End If
    Screen.MousePointer = vbDefault
    Exit Function
'9202
KILLERR1:
    ilRet = KILLISSUE
    Resume Next
KILLERR2:
    gMsgBox "ContractXML-mExportProposal  could not delete file located at " & slToFile & ".  Rename or delete this file, and attempt export again.", vbExclamation, "Deletion issue"
    Screen.MousePointer = vbDefault
'imcExportErr:
'    ilRet = Err.Number
'    Resume Next

End Function


Private Function mAddACT1Lineup(ilClf As Integer) As String
    Dim slStr As String
    slStr = ""
    If Trim$(tgClfCntr(ilClf).ClfRec.sACT1LineupCode) <> "" Then
        slStr = slStr & "ACT1code=" & Trim$(tgClfCntr(ilClf).ClfRec.sACT1LineupCode)
        slStr = slStr & " ACT1stored=" & Trim$(tgClfCntr(ilClf).ClfRec.sACT1StoredTime) & Trim$(tgClfCntr(ilClf).ClfRec.sACT1StoredSpots) & Trim$(tgClfCntr(ilClf).ClfRec.sACT1StoreClearPct) & Trim$(tgClfCntr(ilClf).ClfRec.sACT1DaypartFilter)
        If (tgClfCntr(ilClf).ClfRec.sLiveCopy = "L") Or (tgClfCntr(ilClf).ClfRec.sLiveCopy = "M") Then
            slStr = slStr & " ACT1qual=" & "L"
        ElseIf (tgClfCntr(ilClf).ClfRec.sLiveCopy = "P") Or (tgClfCntr(ilClf).ClfRec.sLiveCopy = "Q") Then
            slStr = slStr & " ACT1qual=" & "V"
        End If
    End If
    mAddACT1Lineup = slStr
End Function

Private Function mPBGetLineType(ilClfIndex As Integer) As String

    mPBGetLineType = tgClfCntr(ilClfIndex).ClfRec.sType
End Function

Private Function mPBGetFlightSpotsDays(ilLnRowNo As Integer, llDate As Long, ilOVDays() As Integer, llSpotCount As Long, ilDays() As Integer, slWDFlag As String, ilDailySpots() As Integer) As Integer
    
    'Note: This function was copied from contract.bas, the only difference is we added the PB in the name
    Dim ilCff As Integer
    Dim ilLoop As Integer
    Dim llFlStartDate As Long
    Dim llFlEndDate As Long
    Dim llFlMoStartDate As Long
    Dim llFlSuEndDate As Long
    Dim ilNoSpots As Integer
    Dim ilDaysFd As Integer
    Dim ilIndex As Integer
    ilNoSpots = 0
    slWDFlag = ""
    ilDaysFd = False
    ilCff = tgClfCntr(ilLnRowNo - 1).iFirstCff
    Do While ilCff <> -1
        If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
            'gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llFlStartDate    'Week Start date
            'gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llFlEndDate    'Week Start date
            llFlStartDate = tgCffCntr(ilCff).lStartDate
            llFlEndDate = tgCffCntr(ilCff).lEndDate
            If llFlEndDate < llFlStartDate Then
                Exit Do
            End If
            llFlMoStartDate = llFlStartDate
            Do While gWeekDayLong(llFlMoStartDate) <> 0
                llFlMoStartDate = llFlMoStartDate - 1
            Loop
            llFlSuEndDate = llFlEndDate
            Do While gWeekDayLong(llFlSuEndDate) <> 6
                llFlSuEndDate = llFlSuEndDate + 1
            Loop

            If (llDate >= llFlMoStartDate) And (llDate <= llFlSuEndDate) Then
                If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                    slWDFlag = "D"
                    For ilLoop = 0 To 6 Step 1
                        ilNoSpots = ilNoSpots + tgCffCntr(ilCff).CffRec.iDay(ilLoop)
                        If tgCffCntr(ilCff).CffRec.iDay(ilLoop) > 0 Then
                            ilDays(ilLoop) = True
                        Else
                            ilDays(ilLoop) = False
                        End If
                        ilDailySpots(ilLoop) = tgCffCntr(ilCff).CffRec.iDay(ilLoop)
                    Next ilLoop
                Else
                    slWDFlag = "W"
                    ilNoSpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                    For ilLoop = 0 To 6 Step 1
                        If (tgCffCntr(ilCff).CffRec.iDay(ilLoop) > 0) Or (tgCffCntr(ilCff).CffRec.sXDay(ilLoop) = "1") Then
                            ilDays(ilLoop) = True
                        Else
                            ilDays(ilLoop) = False
                        End If
                    Next ilLoop
                End If
                llSpotCount = ilNoSpots
                mPBGetFlightSpotsDays = ilCff
                Exit Function
            End If
        End If
        ilCff = tgCffCntr(ilCff).iNextCff
    Loop
    ilIndex = LBound(ilDays)
    For ilLoop = LBound(ilOVDays) To UBound(ilOVDays) Step 1
        If ilOVDays(ilLoop) = 1 Then
            ilDays(ilIndex) = False
        Else
            ilDays(ilIndex) = True
        End If
        ilIndex = ilIndex + 1
    Next ilLoop
    'Check if daily defined
    ilCff = tgClfCntr(ilLnRowNo - 1).iFirstCff
    Do While ilCff <> -1
        If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
            'Set daily flag was all week will be the same.  Set here incase no week found
            If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                slWDFlag = "D"
                For ilLoop = 0 To 6 Step 1
                    ilDailySpots(ilLoop) = tgCffCntr(ilCff).CffRec.iDay(ilLoop)
                Next ilLoop
                Exit Do
            End If
        End If
        ilCff = tgCffCntr(ilCff).iNextCff
    Loop
    llSpotCount = 0
    mPBGetFlightSpotsDays = -1
End Function



Private Function mPBIsLnGame(ilClfIndex As Integer) As Integer
    mPBIsLnGame = False
End Function

Private Sub mSetXMLResearchNo()

    Dim ilClf As Integer
    Dim slPkLineType As String
    Dim slCntr1or2PlaceRating As String
    Dim ilAvgRating As Integer
    Dim llAvgAud As Long
    
    slCntr1or2PlaceRating = gSet1or2PlaceRating(tgChfCntr.iAgfCode)
    For ilClf = 0 To UBound(tgClfCntr) - 1 Step 1
        slPkLineType = tgClfCntr(ilClf).ClfRec.sType
        If (slPkLineType = "O") Or (slPkLineType = "A") Then
            If Not rsDemoValues Is Nothing Then
                ilAvgRating = tgClfCntr(ilClf).iAvgRating
                llAvgAud = tgClfCntr(ilClf).lAvgAud
                '10/28/14: Handle 2 place rating
                'rsDemo.AddNew Array("LineIndex", "DemoIndex", "Rating", "Audience"), Array(ilClf, ilDemoIndex, gIntToStrDec(igLnAvgRating(ilLnUpper), 1), llAvgAudAvg)
                If slCntr1or2PlaceRating <> "2" Then
                    rsDemoValues.AddNew Array("LineIndex", "DemoIndex", "Rating", "Audience"), Array(ilClf, 0, gIntToStrDec(ilAvgRating, 1), llAvgAud)
                Else
                    rsDemoValues.AddNew Array("LineIndex", "DemoIndex", "Rating", "Audience"), Array(ilClf, 0, gIntToStrDec(ilAvgRating, 2), llAvgAud)
                End If
                'rsDemo.AddNew Array("LineIndex", "DemoIndex", "Rating", "Audience"), Array(ilClf, ilDemoIndex, gIntToStrDec(igLnAvgRating(ilLnUpper), 1), gLongToStrDec(llAvgAudAvg, 1))
            End If
        End If
    Next ilClf
End Sub

