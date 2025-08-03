VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCsiXmlReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private emCreationUser As ContractUser
Private smFileName As String
Private smErrorMessage As String
Private smXmlDoc As String
Private lmContractNumber As Long
'must have these...fill in 'load'
Private smBuyer As String
Private smAdvertiser As String
Private smProduct As String
Private smSalesPerson As String
'optional...fill in 'load'
Private smAgency As String
Private smAgencyCity As String
'fill after getting flights
Private smStartDateOfAllForContract As String
Private smEndDateOfAllForContract As String
Private tmClfVehicles As Dictionary
Private tmClfDayParts As Dictionary

Private smDaysOfWeek(6) As String
Private myRsXml As ADODB.Recordset
Private hmCxf As Integer
Private bmFailToParse As Boolean

Private Const CLASSNAME As String = "CCsiXmlReader"
Private Const NODATE As String = "01/01/1970"
Private Const NORATE As Integer = -1
Public Property Get CreationUser() As ContractUser
    CreationUser = emCreationUser
End Property
Public Property Let CreationUser(ByVal elValue As ContractUser)
    emCreationUser = elValue
End Property
Public Property Get fileName() As String
    fileName = smFileName
End Property
Public Property Get ContractNumber() As Long
    ContractNumber = lmContractNumber
End Property
Public Property Get ErrorMessage() As String
    ErrorMessage = smErrorMessage
End Property
Public Property Get Buyer() As String
    Buyer = smBuyer
End Property
Public Property Get Advertiser() As String
    Advertiser = smAdvertiser
End Property
Public Property Get Agency() As String
    Agency = smAgency
End Property
Public Property Get Product() As String
    Product = smProduct
End Property
Public Property Get SalesPerson() As String
    SalesPerson = smSalesPerson
End Property
Public Property Get AgencyCity() As String
    AgencyCity = smAgencyCity
End Property
Public Property Get StartDate() As String
    StartDate = smStartDateOfAllForContract
End Property
Public Property Get EndDate() As String
    EndDate = smEndDateOfAllForContract
End Property
Private Sub Class_Initialize()
    Dim c As Integer
    Dim slTemp As String
    Dim ilRet As Integer
    
On Error GoTo ERRORBOX
    Set tmClfVehicles = New Dictionary
    Set tmClfDayParts = New Dictionary
    Clear True
    For c = 0 To 6
        Select Case c
            Case 0
                slTemp = "Monday"
            Case 1
                slTemp = "Tuesday"
            Case 2
                slTemp = "Wednesday"
            Case 3
                slTemp = "Thursday"
            Case 4
                slTemp = "Friday"
            Case 5
                slTemp = "Saturday"
            Case 6
                slTemp = "Sunday"
        End Select
        smDaysOfWeek(c) = slTemp
    Next c
    hmCxf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        smErrorMessage = "Issue opening Cxf Table"
        ilRet = btrClose(hmCxf)
        btrDestroy hmCxf
        Exit Sub
    End If
    Exit Sub
ERRORBOX:
    smErrorMessage = err.Description
    gHandleError "", "CCsiXmlReader-Initialize"
End Sub
Private Sub Class_Terminate()
    btrClose hmCxf
    btrDestroy hmCxf
End Sub
Public Sub Clear(Optional blAll As Boolean = False)
    smFileName = ""
    smErrorMessage = ""
    lmContractNumber = 0
    smXmlDoc = ""
    smAgency = ""
    smAgencyCity = ""
    smBuyer = ""
    smAdvertiser = ""
    smProduct = ""
    smSalesPerson = ""
    smStartDateOfAllForContract = ""
    smEndDateOfAllForContract = ""
    tmClfVehicles.removeAll
    tmClfDayParts.removeAll
    If blAll Then
        emCreationUser = UnknownUser
    End If
End Sub
Public Function Load(slPath As String, slName As String, tmVehicles As Dictionary, tmDayParts As Dictionary) As Boolean
    'set tgchf to default. Get header values and store in variables.  Line and flight go to respective arrays
    Dim blRet As Boolean
    Dim slFilePath As String
    Dim slElement As String
    Dim olFile As FileSystemObject
    Dim slTemp As String
    Dim MyFile As TextStream
    
    Clear
    blRet = True
On Error GoTo ERRORBOX
    Set olFile = New FileSystemObject
    slFilePath = olFile.BuildPath(slPath, slName)
    If olFile.FILEEXISTS(slFilePath) Then
        Set MyFile = olFile.OpenTextFile(slFilePath)
        smXmlDoc = MyFile.ReadAll
        If emCreationUser = Filemaker Then
            If Not mHeaderLoadBasic() Then
                blRet = False
            End If
            'I will load start and end dates for both chf and clf in here after getting from flights
            If Not mLinesAndFlightsLoadBasic() Then
                blRet = False
            End If
            'I will load start and end dates for ntr here
            If Not mNtrLoadBasic() Then
                blRet = False
            End If
        'future: other than filemaker goes here
        Else
            blRet = False
            smErrorMessage = " Contact Counterpoint. Creation user not set properly, cannot load and retrieve header."
        End If
    Else
        blRet = False
        smErrorMessage = "couldn't load file: " & slFilePath
    End If
    If blRet Then
        smFileName = slName
    End If
    Load = blRet
    Set tmVehicles = tmClfVehicles
    Set tmDayParts = tmClfDayParts
    Exit Function
ERRORBOX:
    Load = False
    smErrorMessage = err.Description
End Function
Private Function mParseSubElement(sllines As String, slNameOfTopElement As String, slNameofSubElement As String, llCurrentRow As Long, Optional blResetCurrentRowBecauseMovingToNextElement = True, Optional slSubArrayStopElement As String = "")
    'sub must be a single element.
    'arrays of top element: set last call of subelements to blreset = true, otherwise false, so keep going through same top element.  Because on the first of these calls we can't tell if
    'we found the top element (as we're not returning the llCurrentRow), test for NODATE, which means: no top element
    ' I test to make sure i haven't gone past the top element.  For sub arrays, that won't work, so I test the higher element passed as subarraystopelement
    Dim slRet As String
    Dim slWholeElement As String
    Dim llTemp As Long
    Dim llPos As Long
    Dim slLimiter As String
    Dim llStartLimiter As Long
    
    slLimiter = slNameOfTopElement
    If Len(slSubArrayStopElement) > 0 Then
        slLimiter = slSubArrayStopElement
    End If
    slRet = ""
    If llCurrentRow < 1 Then
        llStartLimiter = 1
    Else
        llStartLimiter = llCurrentRow + 1
    End If
    llPos = InStr(llStartLimiter, sllines, "</" & slLimiter & ">")
    llTemp = llCurrentRow
    slWholeElement = mParseXml(sllines, slNameOfTopElement, llTemp)
    If blResetCurrentRowBecauseMovingToNextElement Then
        llCurrentRow = llTemp
    End If
    If Len(slWholeElement) > 0 And llTemp <= llPos Then
        slRet = mParseXml(slWholeElement, slNameofSubElement, 0)
    ElseIf Not blResetCurrentRowBecauseMovingToNextElement Then
        slRet = NODATE
    End If
    mParseSubElement = slRet
End Function
Private Function mParseXml(sllines As String, slName As String, llCurrentRow As Long) As String
    Dim slStartElement As String
    Dim slEndElement As String
    Dim slvalue As String
    Dim llPos As Long
    Dim ilEndPos As Long
    Dim ilLength As Integer
    Dim ilStart As Long
    
    bmFailToParse = True
    If llCurrentRow <= 0 Then
        llCurrentRow = 1
    End If
    slStartElement = "<" & slName & ">"
    slEndElement = "</" & slName & ">"
    llPos = InStr(llCurrentRow, sllines, slStartElement)
    If llPos > 0 Then
        ilStart = llPos + Len(slStartElement)
        ilEndPos = InStr(ilStart, sllines, slEndElement)
        ilLength = ilEndPos - ilStart
        'this can fail!  I catch in calling function
        slvalue = Mid(sllines, ilStart, ilLength)
        llCurrentRow = ilEndPos
        mParseXml = Trim(mUnencodeXmlData(slvalue))
    'what was this for?
'    Else
'        slStartElement = "<" & slName
'        slEndElement = ">"
'        llPos = InStr(llCurrentRow, sllines, slStartElement)
'        If llPos > 0 Then
'            ilStart = llPos + Len(slStartElement)
'            ilEndPos = InStr(llCurrentRow + llPos, sllines, slEndElement)
'            ilLength = ilEndPos - ilStart
'            slvalue = Mid(sllines, ilStart, ilLength)
'            mParseXml = mUnencodeXmlData(slvalue)
'        Else
'            mParseXml = vbNullString
'        End If
    End If
    bmFailToParse = False
End Function
Private Function mUnencodeXmlData(slData As String) As String
    Dim slRet As String
    If InStr(1, slData, "&") > 0 Then
        slRet = Replace(slData, "&lt;", "<")
        slRet = Replace(slRet, "&gt;", ">")
        slRet = Replace(slRet, "&amp;", "&")
        slRet = Replace(slRet, "&apos;", "`")
        slRet = Replace(slRet, "&quot;", """")
        mUnencodeXmlData = slRet
    Else
        mUnencodeXmlData = slData
    End If
End Function
Private Function mDefaultCHF() As CHF
    Dim tlChf As CHF
    Dim slDate As String
    Dim c As Integer
    With tlChf
        .lCode = 0
        .sAgyEstNo = ""
        .iExtRevNo = 0
        .iXMLSentUrfCode = 0
        .iCntRevNo = 0
        .sType = "C"
        For c = 1 To 9
            .iSlfCode(c) = 0
        Next c
        For c = 1 To 9
            .lComm(c) = 0
        Next c
        .lComm(0) = 1000000
        For c = 1 To 9
            .iSlspCommPct(c) = 0
        Next c
        .iMnfComp(0) = 0
        .iMnfComp(1) = 0
        .iMnfExcl(0) = 0
        .iMnfExcl(1) = 0
        .iSlspCommPct(0) = 10000
        .iEDSSentUrfCode = 0
        .iPctBudget = 0
        For c = 0 To 4
            .iMnfRevSet(c) = 0
        Next c
        .sCppCpm = "N"
        .iMerchPct = 0
        .iPromoPct = 0
        For c = 0 To 3
            .iMnfDemo(c) = 0
        Next c
        For c = 0 To 3
            .lTarget(c) = 0
        Next c
        .iPctTrade = 0
        .sBillCycle = "S"
        .sInvGp = "A"
       ' .iPctTag = 0
        .lCxfCode = 0
        .lCxfChgR = 0
        .lCxfInt = 0
        .lCxfMerch = 0
        .lCxfProm = 0
        .lCxfCanc = 0
        .iPropVer = 0
        .sMGMiss = "G"
        .sStatus = "N"
        .sTitle = ""
        For c = 0 To 9
            .iMnfSubCmpy(c) = 0
        Next c
        .iMnfBus = 0
        .lGuar = 0
        'Dan changed 2/15/18
        .iEDSSentExtRevNo = -2
        .iEDSSentUrfCode = 0
        For c = 0 To 6
            .iMnfCmpy(c) = 0
        Next c
        For c = 0 To 6
            .iCmpyPct(c) = 0
        Next c
        .sResvNew = "N"
        .lChfCode = 0
        .iMnfPotnType = 0
        .sPrint = "N"
        .sDiscrep = "N"
        'not sure on this one
        .sNewBus = "N"
        .sSchStatus = "A"
        .sAgyCTrade = "N"
        .imnfSeg = 0
        .sCBSOrder = "N"
        .iHdChg = 0
        .iXMLSentExtRevNo = 0
        .sDelete = "N"
        .lSifCode = 0
        .sInstallDefined = "N"
        .sRepDBID = ""
        .sNRProcessed = "Y"
        .lEffCode = 0
        .sHideDemo = ""
        .lSpotChfCode = 0
        .sNoAssigned = "Y"
        .lGrImp = 0
        .lGRP = 0
        .sCppCpm = "N"
        .sResvNew = "N"
        'date/time should be current
        slDate = gNow()
        gPackDate slDate, .iOHDDate(0), .iOHDDate(1)
        gPackDate slDate, .iPropDate(0), .iPropDate(1)
        gPackTime slDate, .iPropTime(0), .iPropTime(1)
        gPackTime slDate, .iOHDTime(0), .iOHDTime(1)
        'date/time should be 1/1/70
        slDate = "01/01/1970 00:00:00"
        gPackDate slDate, .iEDSSentDate(0), .iEDSSentDate(1)
        gPackTime slDate, .iEDSSentTime(0), .iEDSSentTime(1)
        gPackDate slDate, .iDtNeed(0), .iDtNeed(1)
        'date/time should be blank but 1970 should be ok
        gPackDate slDate, .iPropDate(0), .iPropDate(1)
        gPackDate slDate, .iXMLSentDate(0), .iXMLSentDate(1)
        gPackTime slDate, .iXMLSentTime(0), .iXMLSentTime(1)
        'added 2/15/18
        gPackDate "", .iDtNeed(0), .iDtNeed(1)
    End With
    mDefaultCHF = tlChf
End Function
Private Function mDefaultCLF() As CLF
    Dim tlClf As CLF
    Dim c As Integer
    Dim slDate As String
    
    slDate = gNow()
    With tlClf
        .iCntRevNo = 0
        .iRpfCode = 0
        .sBB = ""
        .sExtra = ""
        .sPgmTime = ""
        .iBreak = 0
        .iPosition = 0
        .iNoGames = 0
        .iSpotsOrdered = 0
        .iSpotsBooked = 0
        .iSpotsWrite = 0
        .sPreempt = "P"
        For c = 0 To 6
            .sPrefDays(c) = ""
        Next c
        .iPctAllocation = 0
        .lAcquisitionCost = 0
        .sSoloAvail = ""
        .sLiveCopy = ""
        .sSportsByWeek = ""
        'Dan changed 2/15/18
        .iEDSSentExtRevNo = -2
        .iXMLSentExtRevNo = 0
        .sLineChgd = "Y"
        .sHideCBS = "N"
        .sPriceType = "T"
        .iPriority = -1
        .sSchStatus = "N"
        .sDelete = "N"
        .iMnfDemo = 0
        .lCPM = 0
        .lCPP = 0
        .lGrImp = 0
        .iDnfCode = 0
        .iPropVer = 0
        .iMnfSocEco = 0
        .sType = "S"
        .iPkLineNo = 0
        .iAdvtSepFlag = 0
        .iBBOpenLen = 0
        .iBBCloseLen = 0
        .sOV2DefinedBits = Chr(0)
        .lCode = 0
        .lghfcode = 0
        .sGameLayout = ""
        .lRafCode = 0
        .sACT1LineupCode = ""
        .sACT1StoredTime = ""
        .sACT1StoredSpots = ""
        .sACT1StoreClearPct = ""
        .sACT1DaypartFilter = ""
        'Dan added 2/15/18
        .sUnused = ""
        gPackTime "", .iPrefStartTime(0), .iPrefStartTime(1)
        gPackTime "", .iPrefStartTime(0), .iPrefStartTime(1)
        gPackTime "", .iPrefEndTime(0), .iPrefEndTime(1)
        gPackDate slDate, .iEntryDate(0), .iEntryDate(1)
        gPackTime slDate, .iEntryTime(0), .iEntryTime(1)
    End With
    mDefaultCLF = tlClf
End Function
Private Function mDefaultCFF() As CFF
    Dim tlCff As CFF
    Dim c As Integer
    
    With tlCff
        .iCntRevNo = 0
        .iXSpotsWk = 0
        For c = 0 To 6
            .sXDay(c) = ""
        Next c
        .lBBPrice = 0
        .sPriceType = "T"
        .iPropVer = 0
        .lAdjPrice = 0
        .lCode = 0
        .lPropPrice = 0
        'added 2/15/18
        .sUnused = ""
        .sDelete = "N"
    End With
    mDefaultCFF = tlCff
End Function
Private Function mDefaultNTR() As SBF
    Dim tlNtr As SBF
    Dim slDate As String
    
    With tlNtr
'        .lChfCode = 0
'        .iNoItems = 0
'        'from mnf?
'        .lGross = 0
'        .iDate(0) = 0
'        .iUrfCode = 0
'        .sDescr = ""
'        .sAgyComm = ""
'        'for agency
'        .iCommPct = 0
'        .iAirVefCode = 0
'        .iBillVefCode = 0
'        .lAcquisitionCost = 0
'        .iImportDate(0) = 0
'        'ntr type
'        .iMnfItem = 0
        .lCode = 0
        .sSlsComm = ""
        .lSpotPrice = 0
        .lOGross = 0
        .sTranType = "I"
        .iCalCarryBonus = 0
        .sBilled = "N"
        .sCashTrade = ""
        .iBonusNoSpots = 0
        .iTrfCode = 0
        .iAirNoSpots = 0
        .lRefInvNo = 0
        .iCombineID = 0
        .iRecDate(0) = 0
        .iPostDate(0) = 0
        .sPostStatus = ""
        .iCalCarryOver = 0
        .iMissCarryOver = 0
        .sInserted = ""
        .sBarterPaid = "N"
        .iSpotLen = 0
        .iIhfCode = 0
        .iLineNo = 0
        '1970
        slDate = "1970-01-01"
        gPackDate slDate, .iExportDate(0), .iExportDate(1)
        gPackDate slDate, .iPrintInvDate(0), .iPrintInvDate(1)
        'current date
        slDate = gNow()
        gPackDate slDate, .iImportDate(0), .iImportDate(1)
    End With
    mDefaultNTR = tlNtr
End Function
Private Function mHeaderLoadBasic() As Boolean
    Dim slElement As String
    Dim slTemp As String
    Dim blRet As Boolean
    
    blRet = True
On Error GoTo ERRORBOX
    tgChf = mDefaultCHF()
    slElement = "Source"
    slTemp = mParseXml(smXmlDoc, slElement, 0)
    If slTemp <> "Relevant" Then
        blRet = False
        smErrorMessage = "'" & slElement & "' of '" & slTemp & "' is not allowed."
    End If
    slElement = "Number"
    slTemp = mParseXml(smXmlDoc, slElement, 0)
    If Len(slTemp) > 0 Then
        If IsNumeric(slTemp) Then
            lmContractNumber = slTemp
        Else
            blRet = False
            smErrorMessage = smErrorMessage & " 'Number' of " & slTemp & " is not a number."
        End If
    Else
        blRet = False
        smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist or is blank."
    End If
    slElement = "Advertiser"
    slTemp = mParseSubElement(smXmlDoc, slElement, "Name", 0)
    If Len(slTemp) > 0 Then
        smAdvertiser = slTemp
    Else
        blRet = False
        smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist or is blank."
    End If
    slElement = "Agency"
    slTemp = mParseSubElement(smXmlDoc, slElement, "Name", 0)
    If Len(slTemp) > 0 Then
        smAgency = slTemp
        slElement = "CityID"
        slTemp = mParseSubElement(smXmlDoc, "Agency", slElement, 0)
        If Len(slTemp) > 0 Then
            smAgencyCity = slTemp
        Else
            blRet = False
            smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist or is blank."
        End If
    End If
    slElement = "Buyer"
    slTemp = mParseXml(smXmlDoc, slElement, 0)
    If Len(slTemp) > 0 Then
        smBuyer = slTemp
    Else
        blRet = False
        smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist or is blank."
    End If
    'not element is subelement on this call...just so name if error makes more sense
    '8809 optional
    slElement = "Product"
    slTemp = mParseSubElement(smXmlDoc, "Advertiser", slElement, 0)
    If Len(slTemp) > 0 Then
        smProduct = slTemp
    Else
        smProduct = ""
'        blRet = False
'        smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist or is blank."
    End If
    slElement = "Salesperson"
    slTemp = mParseXml(smXmlDoc, slElement, 0)
    If Len(slTemp) > 0 Then
        smSalesPerson = slTemp
    Else
        blRet = False
        smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist or is blank."
    End If
    mHeaderLoadBasic = blRet
    Exit Function
ERRORBOX:
    mHeaderLoadBasic = False
    If bmFailToParse Then
        smErrorMessage = slElement & " not properly set in file."
    Else
        smErrorMessage = "Contact Counterpoint.  Error in loading header " & slElement & ":" & err.Description
    End If
End Function
Private Function mLinesAndFlightsLoadBasic() As Boolean
    Dim blRet As Boolean
    Dim slElement As String
    Dim slTempLine As String
    Dim ilCount As Integer
    Dim slStartDateOfAllForLine As String
    Dim slEndDateOfAllForLine As String
    Dim tlClf As CLF
    Dim tlCff As CFF
    Dim llCurrentRow As Long
    Dim ilDailyThenWeekly As Integer
    
    Const NEVERRAN As Integer = -3
    Const DAILY As Integer = 0
    Const WEEKLY As Integer = 1
On Error GoTo ERRORBOX
    
    blRet = True
    smStartDateOfAllForContract = NODATE
    smEndDateOfAllForContract = NODATE
    tmClfVehicles.removeAll
    tmClfDayParts.removeAll
    ReDim tgClf(0) As CLFLIST
    ReDim tgCff(0) As CFFLIST
    tlClf = mDefaultCLF()
    tlCff = mDefaultCFF()
    'daily and weekly are so close.  the only differences are in the periods.
    For ilDailyThenWeekly = DAILY To WEEKLY
        llCurrentRow = 1
        Do While llCurrentRow > -1
            slStartDateOfAllForLine = NODATE
            slEndDateOfAllForLine = NODATE
            'these are overrides only, so clear out from previous
            gPackTime "", tlClf.iStartTime(0), tlClf.iStartTime(1)
            gPackTime "", tlClf.iEndTime(0), tlClf.iEndTime(1)
            If ilDailyThenWeekly = DAILY Then
                slElement = "DailyLine"
                tlClf.sCntPct = "0"
                tlCff.iSpotsWk = 0
                'this is what tells whether to use daily or weekly in periods
                tlCff.sDyWk = "D"
            Else
                slElement = "WeeklyLine"
                tlClf.sCntPct = "1"
                tlCff.sDyWk = "W"
            End If
            slTempLine = mParseXml(smXmlDoc, slElement, llCurrentRow)
            If Len(slTempLine) > 0 Then
                If Not mClfLoadHeaderBasic(tlClf, slTempLine) Then
                    blRet = False
                    'error!
                    If InStr(smErrorMessage, "Contact Counterpoint") > 0 Then
                        Exit Function
                    End If
                End If
                tlCff.iClfLine = tlClf.iLine
                If Not mCffLoadPeriodsDayOrWeek(tlCff, slTempLine, slStartDateOfAllForLine, slEndDateOfAllForLine) Then
                    blRet = False
                    'error!
                    If InStr(smErrorMessage, "Contact Counterpoint") > 0 Then
                        Exit Function
                    End If
                End If
                'finish clf with what we got from cff
                gPackDate slStartDateOfAllForLine, tlClf.iStartDate(0), tlClf.iStartDate(1)
                gPackDate slEndDateOfAllForLine, tlClf.iEndDate(0), tlClf.iEndDate(1)
                smStartDateOfAllForContract = mDateReturnLeastOrMost(smStartDateOfAllForContract, slStartDateOfAllForLine, False)
                smEndDateOfAllForContract = mDateReturnLeastOrMost(smEndDateOfAllForContract, slEndDateOfAllForLine, True)
                ilCount = UBound(tgClf)
                tgClf(ilCount).ClfRec = tlClf
                ReDim Preserve tgClf(ilCount + 1)
            Else
                llCurrentRow = -1
            End If
        Loop
    Next ilDailyThenWeekly
    mLinesAndFlightsLoadBasic = blRet
    Exit Function
ERRORBOX:
    mLinesAndFlightsLoadBasic = False
        If bmFailToParse Then
        smErrorMessage = slElement & " not properly set in file."
    Else
        smErrorMessage = "Contact Counterpoint.  Error in loading lines and periods " & slElement & ":" & err.Description
    End If
End Function
Public Function mNtrLoadBasic() As Boolean
    Dim blRet As Boolean
    Dim slElement As String
    Dim tlNtr As SBF
    Dim llCurrentRow As Long
    Dim slTempLine As String
    Dim slTemp As String
    Dim ilCount As Integer
    'negative # starting at -1
    Dim ilDictionaryKey As Integer
    blRet = True
    ReDim tgNTRInfo(0) As NTRINFO
    tlNtr = mDefaultNTR()
    llCurrentRow = 1
    ilCount = 0
    ilDictionaryKey = 0
    
    If InStr(smXmlDoc, "<NTRs>") > 0 Then
        If InStr(smXmlDoc, "<NTR>") > 0 Then
            Do While llCurrentRow > -1
                ilDictionaryKey = ilDictionaryKey - 1
                'optional fields--set to blank so know to get from our database
                With tlNtr
                    .sAgyComm = ""
                    .sSlsComm = ""
                    'salestaxable ?
                    'because 0 could be passed
                    .iCommPct = NORATE
                    .lGross = NORATE
                    .lAcquisitionCost = NORATE
                End With
                slElement = "NTR"
                slTempLine = mParseXml(smXmlDoc, slElement, llCurrentRow)
                If Len(slTempLine) > 0 Then
                    slElement = "Vehicle"
                    slTemp = mParseXml(slTempLine, slElement, 0)
                    If Len(slTemp) > 0 Then
                        If Not tmClfVehicles.Exists(ilDictionaryKey) Then
                            tmClfVehicles.Add ilDictionaryKey, slTemp
                        Else
                            blRet = False
                            smErrorMessage = smErrorMessage & "Contact Counterpoint: key not unique for: " & slTemp '"NTRLine of " & slTemp & " is not unique."
                        End If
                    Else
                        blRet = False
                        smErrorMessage = smErrorMessage & "'" & slElement & "' does not exist or is blank. NTR#:" & ilCount + 1
                    End If
                    slElement = "BillDate"
                    slTemp = mParseXml(slTempLine, slElement, 0)
                    If mIsValidDate(slTemp) Then
                       gPackDate slTemp, tlNtr.iDate(0), tlNtr.iDate(1)
                       smStartDateOfAllForContract = mDateReturnLeastOrMost(smStartDateOfAllForContract, slTemp, False)
                       smEndDateOfAllForContract = mDateReturnLeastOrMost(smEndDateOfAllForContract, slTemp, True)
                    Else
                        blRet = False
                        smErrorMessage = smErrorMessage & " '" & slElement & "' of " & slTemp & " is not a Date."
                    End If
                    slElement = "Description"
                    slTemp = mParseXml(slTempLine, slElement, 0)
                    If Len(slTemp) > 0 Then
                        tlNtr.sDescr = slTemp
                    Else
                        blRet = False
                        smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist or is blank."
                    End If
                    slElement = "NTRType"
                    slTemp = mParseXml(slTempLine, slElement, 0)
                    If Len(slTemp) > 0 Then
                        If Not tmClfDayParts.Exists(ilDictionaryKey) Then
                            tmClfDayParts.Add ilDictionaryKey, slTemp
                        Else
                            blRet = False
                            smErrorMessage = smErrorMessage & "Contact Counterpoint: key not unique for: " & slTemp
                        End If
                    Else
                        blRet = False
                        smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist or is blank."
                    End If
                    slElement = "NumberOfItems"
                    slTemp = mParseXml(slTempLine, slElement, 0)
                    If Len(slTemp) > 0 Then
                        If IsNumeric(slTemp) Then
                            tlNtr.iNoItems = slTemp
                        Else
                            blRet = False
                            smErrorMessage = smErrorMessage & " '" & slElement & "' of " & slTemp & " is not a number."
                        End If
                    Else
                        blRet = False
                        smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist or is blank."
                    End If
                     'optional
                    slElement = "AgencyCommissionable"
                    slTemp = mParseXml(slTempLine, slElement, 0)
                    If Len(slTemp) > 0 Then
                        If slTemp = "Y" Or slTemp = "N" Then
                            tlNtr.sAgyComm = slTemp
                        Else
                            blRet = False
                            smErrorMessage = smErrorMessage & " '" & slElement & "' of " & slTemp & " is not a 'Y' or 'N'."
                        End If
                    End If
                    slElement = "SalesCommission"
                    slTemp = mParseXml(slTempLine, slElement, 0)
                    If Len(slTemp) > 0 Then
                        If mIsValidPercent(slTemp) Then
                            tlNtr.iCommPct = slTemp
                        Else
                            blRet = False
                            smErrorMessage = smErrorMessage & " '" & slElement & "' of " & slTemp & " is not a number."
                        End If
                    End If
                    slElement = "AmountPerItem"
                    slTemp = mParseXml(slTempLine, slElement, 0)
                    If Len(slTemp) > 0 Then
                        If mIsValidCurrency(slTemp) Then
                            tlNtr.lGross = slTemp
                        Else
                            blRet = False
                            smErrorMessage = smErrorMessage & " '" & slElement & "' of " & slTemp & " is not a number."
                        End If
                    Else
                    
                    End If
                    slElement = "AcquisitionCostPerItem"
                    slTemp = mParseXml(slTempLine, slElement, 0)
                    If Len(slTemp) > 0 Then
                        If mIsValidCurrency(slTemp) Then
                            tlNtr.lAcquisitionCost = slTemp
                        Else
                            blRet = False
                            smErrorMessage = smErrorMessage & " '" & slElement & "' of " & slTemp & " is not a number."
                        End If
                    End If
                    ilCount = UBound(tgNTRInfo)
                    tgNTRInfo(ilCount).tNTR = tlNtr
                    ReDim Preserve tgNTRInfo(ilCount + 1)
                    'not currently supported
                    slElement = "SalesTaxable"
                    slTemp = mParseXml(slTempLine, slElement, 0)
                    If Len(slTemp) > 0 Then
                        blRet = False
                        smErrorMessage = smErrorMessage & " '" & slElement & "' is not currently supported."
        '                If slTemp = "Y" Or slTemp = "N" Then
        '                   ' tlNtr.iAirVefCode = slTemp
        '                Else
        '                    blRet = False
        '                    smErrorMessage = smErrorMessage & " '" & slElement & "' of " & slTemp & " is not a 'Y' or an 'N'."
        '                End If
                    End If
        
                Else
                    llCurrentRow = -1
                End If
            Loop
        Else
            blRet = False
            smErrorMessage = smErrorMessage & " 'NTR' does not exist or is blank."
        End If
    End If
    mNtrLoadBasic = blRet
    Exit Function
ERRORBOX:
    mNtrLoadBasic = False
    If bmFailToParse Then
        smErrorMessage = slElement & " not properly set in file."
    Else
        smErrorMessage = "Contact Counterpoint.  Error in loading NTR " & slElement & ":" & err.Description
    End If
End Function
Private Function mDateReturnLeastOrMost(slDateOne As String, slDateTwo As String, blReturnGreatest As Boolean) As String
    'nodate' means hasn't been set.
    Dim slRet As String
    Dim llDifference As Long
    
    If slDateOne = NODATE Then
        slRet = slDateTwo
    ElseIf slDateTwo = NODATE Then
        slRet = slDateOne
    Else
        ' negative means 1 > 2
        llDifference = DateDiff("d", slDateOne, slDateTwo)
        If blReturnGreatest Then
            If llDifference > 0 Then
                slRet = slDateTwo
            Else
                slRet = slDateOne
            End If
        Else
            If llDifference > -1 Then
                slRet = slDateOne
            Else
                slRet = slDateTwo
            End If
        End If
    End If
    mDateReturnLeastOrMost = slRet
End Function
Private Function mIsValidDate(slDate As String) As Boolean
    Dim blRet As Boolean
'    Dim slDates() As String
'    Dim slTemp As String
    
    blRet = False
    If IsDate(slDate) Then
        blRet = True
    End If
'    slDates = Split(slDate, "-")
'    If UBound(slDates) = 2 Then
'        blRet = True
'        'year
'        slTemp = slDates(0)
'        If Not IsNumeric(slTemp) Then
'            blRet = False
'        End If
'        If blRet Then
'            'month
'            slTemp = slDates(1)
'            If IsNumeric(slTemp) Then
'                If slTemp < 1 Or slTemp > 12 Then
'                    blRet = False
'                End If
'            Else
'                blRet = False
'            End If
'        End If
'        If blRet Then
'            'day
'            slTemp = slDates(2)
'            If IsNumeric(slTemp) Then
'                If slTemp < 1 Or slTemp > 31 Then
'                    blRet = False
'                End If
'            Else
'                blRet = False
'            End If
'        End If
'    End If
'    Erase slDates
    mIsValidDate = blRet
End Function
Private Function mIsValidTime(slTime As String, Optional blConvertToSeconds As Boolean) As Boolean
    'returns slTime as integer if blConvertToSeconds
    'converts 30 hour clock
    'allows 30 :30 19:04  19:04:30
    Dim blRet As Boolean
    Dim ilHour As Integer
    Dim ilMinutes As Integer
    Dim ilSeconds As Integer
    Dim slOriginal As String
    Dim slTimes() As String
    Dim c As Integer
    
    slOriginal = slTime
    blRet = False
    '8807
    If blConvertToSeconds = False And Trim$(slTime) = "24:00:00" Then
        slTime = "12A"
        blRet = True
    Else
        slTimes = Split(slTime, ":")
        Select Case UBound(slTimes)
            Case 0
                If IsNumeric(slTimes(0)) Then
                    ilSeconds = slTimes(0)
                    blRet = True
                End If
            ':30 or 19:30-which is hours and minutes
            Case 1
                 If IsNumeric(slTimes(0)) And IsNumeric(slTimes(1)) Then
                    ilHour = slTimes(0)
                    ilMinutes = slTimes(1)
                    ilSeconds = 0
                    blRet = True
                ElseIf IsNumeric(slTimes(1)) Then
                    ilSeconds = slTimes(1)
                    blRet = True
                End If
            Case 2
                If IsNumeric(slTimes(0)) And IsNumeric(slTimes(1)) And IsNumeric(slTimes(2)) Then
                    ilHour = slTimes(0)
                    ilMinutes = slTimes(1)
                    ilSeconds = slTimes(2)
                    blRet = True
                End If
        End Select
        If blRet Then
            '120 is valid for seconds, but 1:12:120 is not
            If ilMinutes > 0 Or ilHour > 0 Then
                If ilSeconds > 59 Or ilMinutes > 59 Or ilHour > 24 Then
                    blRet = False
                End If
            End If
    '        If ilSeconds > 60 Or ilMinutes > 60 Then
    '            blRet = False
    '        ElseIf ilHour > 29 Then
    '            blRet = False
    ''        ElseIf ilHour > 23 Then
    ''            ilHour = ilHour - 24
    '        End If
        End If
        If blRet And blConvertToSeconds Then
            ilMinutes = ilHour * 60 + ilMinutes
            ilSeconds = ilMinutes * 60 + ilSeconds
            slTime = ilSeconds
        End If
    End If
    Erase slTimes
    mIsValidTime = blRet
End Function
Private Function mIsValidCurrency(slCurrency As String) As Boolean
    'O: slCurrency changed to long
    Dim blRet As Boolean
    Dim ilPos As Integer
    
    blRet = False
    If IsNumeric(slCurrency) Then
        ilPos = InStr(slCurrency, ".")
        If ilPos > 0 Then
             'if have decimal, must have no more than 2 as mantissa
             If Len(slCurrency) < ilPos + 3 Then
                slCurrency = Replace(slCurrency, ".", "")
                blRet = True
             End If
        Else
            slCurrency = slCurrency * 100
            blRet = True
        End If
    End If
    mIsValidCurrency = blRet
End Function
Private Function mIsValidPercent(slCurrency As String) As Boolean
    'O: slCurrency changed to integer 1.25 becomes 125 100 becomes 10000
    Dim blRet As Boolean
    Dim ilPos As Integer
    
    blRet = False
    If IsNumeric(slCurrency) Then
        ilPos = InStr(slCurrency, ".")
        If ilPos > 0 Then
             'if have decimal, must have no more than 2 as mantissa
             If Len(slCurrency) < ilPos + 3 Then
                slCurrency = Replace(slCurrency, ".", "")
                blRet = True
             End If
        Else
            slCurrency = slCurrency * 100
            If slCurrency <= 10000 Then
                blRet = True
            End If
        End If
    End If
    mIsValidPercent = blRet
End Function
Private Function mAddComentLine(slComment As String) As Long
    'for line only
    Dim tlCxf As CXF
    Dim ilRet As Integer
    Dim ilCxfRecLen As Integer
    
    mAddComentLine = 0
    slComment = Trim(slComment)
    If Len(slComment) > 0 Then
        Do  'Loop until record updated or added
            tlCxf.sComment = ""
            tlCxf.sComType = "L"
            tlCxf.sShProp = "N"
            tlCxf.sShOrder = "N"
            tlCxf.sShSpot = "N"
            tlCxf.sShInv = "N"
            tlCxf.sShInsertion = "N"
            tlCxf.sComment = slComment & Chr$(0)
            ilCxfRecLen = Len(tlCxf)
            tlCxf.lCode = 0 'Autoincrement
            tlCxf.iRemoteID = tgUrf(0).iRemoteUserID
            tlCxf.iSourceID = tgUrf(0).iRemoteUserID
            tlCxf.lAutoCode = tlCxf.lCode
            ilRet = gCXFInsert(hmCxf, tlCxf, ilCxfRecLen, INDEXKEY0)
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet = BTRV_ERR_NONE Then
            Do
                tlCxf.lAutoCode = tlCxf.lCode
                gPackDate gNow, tlCxf.iSyncDate(0), tlCxf.iSyncDate(1)
                gPackTime gNow, tlCxf.iSyncTime(0), tlCxf.iSyncTime(1)
                ilCxfRecLen = Len(tlCxf)
                ilRet = gCXFUpdate(hmCxf, tlCxf, ilCxfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet = BTRV_ERR_NONE Then
                mAddComentLine = tlCxf.lCode
            End If
        Else
            mAddComentLine = -1
            smErrorMessage = "mAddComentLine had an issue.  Could not insert."
        End If
    End If
End Function
Private Function mClfLoadHeaderBasic(tlClf As CLF, slTempLine As String) As Boolean
    'returns tlClf
    Dim slElement As String
    Dim slTemp As String
    Dim blRet As Boolean
    
    blRet = True
    tlClf.sOV2DefinedBits = Chr(0)
    slElement = "LineNumber"
On Error GoTo ERRORBOX
    slTemp = mParseXml(slTempLine, slElement, 0)
    If Len(slTemp) > 0 Then
        If IsNumeric(slTemp) Then
            tlClf.iLine = slTemp
        Else
            blRet = False
            tlClf.iLine = 0
            smErrorMessage = smErrorMessage & " 'LineNumber' of " & slTemp & " is not a number."
        End If
    Else
        blRet = False
        tlClf.iLine = 0
        smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist or is blank."
    End If
    slElement = "CommentLine"
    slTemp = mParseXml(slTempLine, slElement, 0)
    tlClf.lCxfCode = mAddComentLine(slTemp)
    If tlClf.lCxfCode = -1 Then
        blRet = False
        smErrorMessage = "Contact Counterpoint.  Error in mClfLoadHeaderBasic.  Could not add comment."
        Exit Function
    End If
    tlClf.lAcquisitionCost = 0
    slElement = "AcquisitionCost"
    slTemp = mParseXml(slTempLine, slElement, 0)
    If Len(slTemp) > 0 Then
        If mIsValidCurrency(slTemp) Then
            tlClf.lAcquisitionCost = slTemp
            tlClf.sOV2DefinedBits = Chr(Asc(tlClf.sOV2DefinedBits) Or LNACQUISITION)
        Else
            blRet = False
            smErrorMessage = smErrorMessage & " 'AcquistionCost' of " & slTemp & " is not a currency. Line #:" & tlClf.iLine
        End If
    End If
    slElement = "SpotLength"
    slTemp = mParseXml(slTempLine, slElement, 0)
    If Len(slTemp) > 0 Then
        If mIsValidTime(slTemp, True) Then
            tlClf.iLen = slTemp
        Else
            blRet = False
            smErrorMessage = smErrorMessage & " 'SpotLength' of " & slTemp & " is not a time. Line #:" & tlClf.iLine
        End If
    Else
        blRet = False
        smErrorMessage = smErrorMessage & "'" & slElement & "' does not exist or is blank. Line #:" & tlClf.iLine
    End If
    'optional
    slElement = "StartTime"
    slTemp = mParseXml(slTempLine, slElement, 0)
    If Len(slTemp) > 0 Then
        If mIsValidTime(slTemp, False) Then
            gPackTime slTemp, tlClf.iStartTime(0), tlClf.iStartTime(1)
        Else
            blRet = False
            smErrorMessage = smErrorMessage & "'StartTime' of " & slTemp & " is not a time. Line #:" & tlClf.iLine
        End If
    End If
    'optional
    slElement = "EndTime"
    slTemp = mParseXml(slTempLine, slElement, 0)
    If Len(slTemp) > 0 Then
        If mIsValidTime(slTemp) Then
            gPackTime slTemp, tlClf.iEndTime(0), tlClf.iEndTime(1)
        Else
            blRet = False
            smErrorMessage = smErrorMessage & "'EndTime' of " & slTemp & " is not a time. Line #:" & tlClf.iLine
        End If
    End If
    slElement = "VehicleName"
    slTemp = mParseXml(slTempLine, slElement, 0)
    If Len(slTemp) > 0 Then
        If Not tmClfVehicles.Exists(tlClf.iLine) Then
            tmClfVehicles.Add tlClf.iLine, slTemp
        Else
            blRet = False
            smErrorMessage = smErrorMessage & "Line # of " & tlClf.iLine & " for vehicle " & slTemp & " is not unique."
        End If
    Else
        blRet = False
        smErrorMessage = smErrorMessage & "'" & slElement & "' does not exist or is blank. Line #:" & tlClf.iLine
    End If
    slElement = "DaypartName"
    slTemp = mParseXml(slTempLine, slElement, 0)
    If Len(slTemp) > 0 Then
        If Not tmClfDayParts.Exists(tlClf.iLine) Then
            tmClfDayParts.Add tlClf.iLine, slTemp
        Else
            blRet = False
            smErrorMessage = smErrorMessage & "Line # of " & tlClf.iLine & " for daypartName " & slTemp & " is not unique."
        End If
    Else
        blRet = False
        smErrorMessage = smErrorMessage & "'" & slElement & "' does not exist or is blank. Line #:" & tlClf.iLine
    End If
    mClfLoadHeaderBasic = blRet
    Exit Function
ERRORBOX:
    mClfLoadHeaderBasic = False
    If bmFailToParse Then
        smErrorMessage = slElement & " not properly set in file."
    Else
        smErrorMessage = "Contact Counterpoint.  Error in ClfLoadHeaderBasic " & slElement & ":" & err.Description
    End If
End Function
Private Function mCffLoadPeriodsDayOrWeek(tlCff As CFF, slTempLine As String, slStartDateOfAllForLine As String, slEndDateOfAllForLine As String) As Boolean
    'return tlCff and start and end dates
    Dim blRet As Boolean
    Dim slElement As String
    Dim slStartDate As String
    Dim slTemp As String
    Dim llCurrentRowForPeriods As Long
    Dim llStayOnCurrentPeriod As Long
    Dim ilCount As Integer
    Const NEVERRAN As Integer = -3

    blRet = True
On Error GoTo ERRBOX
    slElement = "Periods"
    slTempLine = mParseXml(slTempLine, slElement, 0)
    If Len(slTempLine) > 0 Then
        'but if you have 'periods', have to have a 'period'
        llCurrentRowForPeriods = NEVERRAN
        'subelement will move to next period on each callwithout this!
        llStayOnCurrentPeriod = 0
        Do While llCurrentRowForPeriods > 0 Or llCurrentRowForPeriods = NEVERRAN
            slElement = "StartDate"
            'not setting currentforperiods until end. Since this is first subelement, test for NODATE
            slStartDate = mParseSubElement(slTempLine, "DetailedPeriod", slElement, llStayOnCurrentPeriod, False)
            If slStartDate = NODATE Or Len(slStartDate) = 0 Then
                Exit Do
            End If
            If mIsValidDate(slStartDate) Then
                gPackDate slStartDate, tlCff.iStartDate(0), tlCff.iStartDate(1)
            Else
                blRet = False
                smErrorMessage = smErrorMessage & "'" & slElement & "' of " & slTemp & " is not a date. Line #:" & tlCff.iClfLine
            End If
            slElement = "EndDate"
            slTemp = mParseSubElement(slTempLine, "DetailedPeriod", slElement, llStayOnCurrentPeriod, False)
            If Len(slTemp) > 0 Then
                If mIsValidDate(slTemp) Then
                    gPackDate slTemp, tlCff.iEndDate(0), tlCff.iEndDate(1)
                    'compare dates. end must be greater than start
                    If mDateReturnLeastOrMost(slStartDate, slTemp, True) = slTemp Then
                        slStartDateOfAllForLine = mDateReturnLeastOrMost(slStartDateOfAllForLine, slStartDate, False)
                        slEndDateOfAllForLine = mDateReturnLeastOrMost(slEndDateOfAllForLine, slTemp, True)
                    Else
                        blRet = False
                        smErrorMessage = smErrorMessage & "'EndDate' of " & slTemp & " is not greater than 'StartDate' of " & slStartDate & ". Line #:" & tlCff.iClfLine
                    End If
                Else
                    blRet = False
                    smErrorMessage = smErrorMessage & "'" & slElement & "' of " & slTemp & " is not a date. Line #:" & tlCff.iClfLine
                End If
            Else
                blRet = False
                smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist or is blank.  Line #:" & tlCff.iClfLine
            End If
            If tlCff.sDyWk = "D" Then
                For ilCount = 0 To 6
                    slElement = smDaysOfWeek(ilCount)
                    slTemp = mParseSubElement(slTempLine, "SpotsByDay", slElement, llStayOnCurrentPeriod, False, "DetailedPeriod")
                    If slTemp <> NODATE Then
                        If IsNumeric(slTemp) Then
                            tlCff.iDay(ilCount) = slTemp
                        Else
                            blRet = False
                            smErrorMessage = smErrorMessage & " A 'SpotsByDay' of " & slTemp & " is not a number.  Line #:" & tlCff.iClfLine
                        End If
                    Else
                        blRet = False
                        smErrorMessage = smErrorMessage & " 'SpotsByDay' does not exist or is blank.  Line #:" & tlCff.iClfLine
                        Exit For
                    End If
                Next ilCount
            Else
                For ilCount = 0 To 6
                    slElement = smDaysOfWeek(ilCount)
                    slTemp = mParseSubElement(slTempLine, "Days", slElement, llStayOnCurrentPeriod, False, "DetailedPeriod")
                    If slTemp <> NODATE And Len(slTemp) > 0 Then
                        If slTemp = "Y" Then
                            tlCff.iDay(ilCount) = 1
                        ElseIf slTemp = "N" Then
                            tlCff.iDay(ilCount) = 0
                        Else
                            blRet = False
                            smErrorMessage = smErrorMessage & " A 'Days' of " & slTemp & " is not a 'Y' or 'N'.  Line #:" & tlCff.iClfLine
                        End If
                    Else
                        blRet = False
                        smErrorMessage = smErrorMessage & " A 'Days' does not exist or is blank. Line #:" & tlCff.iClfLine
                        Exit For
                    End If
                Next ilCount
                slElement = "SpotsPerWeek"
                slTemp = mParseSubElement(slTempLine, "DetailedPeriod", slElement, llStayOnCurrentPeriod, False)
                If Len(slTemp) > 0 Then
                    If IsNumeric(slTemp) Then
                        tlCff.iSpotsWk = slTemp
                    Else
                        blRet = False
                        smErrorMessage = smErrorMessage & "'" & slElement & "' of " & slTemp & " is not a Number. Line #:" & tlCff.iClfLine
                    End If
                Else
                    blRet = False
                    smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist or is blank.  Line #:" & tlCff.iClfLine
                End If
            End If
            slElement = "Rate"
            'set the current row here, so will go to next detailedperiod next
            slTemp = mParseSubElement(slTempLine, "DetailedPeriod", slElement, llCurrentRowForPeriods, True)
            If slTemp <> "" Then
                If mIsValidCurrency(slTemp) Then
                    tlCff.lActPrice = slTemp
                Else
                    blRet = False
                    smErrorMessage = smErrorMessage & " '" & slElement & "' of " & slTemp & " is not a currency. Line #:" & tlCff.iClfLine
                End If
            Else
                blRet = False
                smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist or is blank.  Line #:" & tlCff.iClfLine
            End If
            ilCount = UBound(tgCff)
            tgCff(ilCount).CffRec = tlCff
            ReDim Preserve tgCff(ilCount + 1)
            'this will allow us to move to next period, but the parse sub will still look within that current period
            llStayOnCurrentPeriod = llCurrentRowForPeriods
        Loop
        'must make sure we got one period
        If llCurrentRowForPeriods = NEVERRAN Then
            blRet = False
            If slStartDate = NODATE Then
                smErrorMessage = smErrorMessage & " 'Periods' cannot be empty. Line #" & tlCff.iClfLine
            Else
                smErrorMessage = smErrorMessage & " 'StartDate does not exist or is blank. Line #" & tlCff.iClfLine
            End If
        End If
    Else
        blRet = False
        smErrorMessage = smErrorMessage & " '" & slElement & "' does not exist. Line #" & tlCff.iClfLine
    End If
    mCffLoadPeriodsDayOrWeek = blRet
    Exit Function
ERRBOX:
    mCffLoadPeriodsDayOrWeek = False
    If bmFailToParse Then
        smErrorMessage = slElement & " not properly set in file."
    Else
        smErrorMessage = "Contact Counterpoint.  Error in mCffLoadPeriodsDayOrWeek " & slElement & ":" & err.Description
    End If
End Function


