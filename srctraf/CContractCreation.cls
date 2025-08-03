VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CContractCreation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private emCreationUser As ContractUser
Private emContractToProcess As ContractProcess
Private emProcessResult As ProcessResult
Private smErrorMessage As String
Private myXml As CCsiXmlReader
Private hmCHF As Integer
Private imRecLenCHF As Integer
Private hmClf As Integer
Private imRecLenCLF As Integer
Private hmCff As Integer
Private imRecLenCFF As Integer
Private tmVehicles As Dictionary
Private tmDayParts As Dictionary
Private myRsXml As ADODB.Recordset
Private hmNTR As Integer
Private imRecLenNTR As Integer

Public Enum ContractUser
    UnknownUser
    Filemaker
    Coin
End Enum
Public Enum ContractProcess
    NoProcess
    Order
   ' Proposal
End Enum
Public Enum ProcessResult
    Cancelled
    ErrorResult
    Success
End Enum
Private Const NORATE As Integer = -1
Private Const NODATE As String = "01/01/1970"
Private Const CLASSNAME As String = "CContractCreation"
Public Property Get CreationUser() As ContractUser
    CreationUser = emCreationUser
End Property
Public Property Let CreationUser(ByVal elValue As ContractUser)
    emCreationUser = elValue
End Property
Public Property Get ProcessResult() As ProcessResult
    ProcessResult = emProcessResult
End Property
Public Property Let ProcessResult(ByVal ilValue As ProcessResult)
    emProcessResult = ilValue
End Property
Public Property Get ContractToProcess() As ContractProcess
    ContractToProcess = emContractToProcess
End Property
Public Property Let ContractToProcess(ByVal elValue As ContractProcess)
    emContractToProcess = elValue
End Property
Public Property Get fileName() As String
    fileName = myXml.fileName
End Property
'Public Property Let FileName(ByVal slvalue As String)
'    smFileName = slvalue
'End Property
Public Property Get ContractNumber() As Long
    ContractNumber = myXml.ContractNumber
End Property
Public Property Get ErrorMessage() As String
    ErrorMessage = smErrorMessage
End Property
'Public Property Let ErrorMessage(ByVal slvalue As String)
'    smErrorMessage = slvalue
'End Property
Public Property Get isFilemakerActive() As Boolean
    Dim blRet As Boolean
    Dim ilValue As Integer
    
    ilValue = Asc(tgSaf(0).sFeatures4)
    If ilValue And FILEMAKERIMPORT Then
        blRet = True
    Else
        blRet = False
    End If
    isFilemakerActive = blRet
End Property
Private Sub Class_Initialize()
    Dim ilRet As Integer
    Dim tlChf As CHF
    Dim tlClf As CLF
    Dim tlCff As CFF
    Dim tlNtr As SBF
    
On Error GoTo ERRORBOX
    Set myXml = New CCsiXmlReader
    If Len(myXml.ErrorMessage) > 0 Then
        smErrorMessage = myXml.ErrorMessage
    Else
        Set myRsXml = New ADODB.Recordset
        Set tmVehicles = New Dictionary
        Set tmDayParts = New Dictionary
        Clear True
        imRecLenCHF = Len(tlChf)
        hmCHF = CBtrvTable(TWOHANDLES)
        ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            smErrorMessage = "Issue opening Chf Table"
            ilRet = btrClose(hmCHF)
            btrDestroy hmCHF
            Exit Sub
        End If
        imRecLenCLF = Len(tlClf)
        hmClf = CBtrvTable(TWOHANDLES)
        ilRet = btrOpen(hmClf, "", sgDBPath & "CLF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            smErrorMessage = "Issue opening CLF Table"
            ilRet = btrClose(hmClf)
            btrDestroy hmClf
            Exit Sub
        End If
        imRecLenCFF = Len(tlCff)
        hmCff = CBtrvTable(TWOHANDLES)
        ilRet = btrOpen(hmCff, "", sgDBPath & "CFF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            smErrorMessage = "Issue opening CFF Table"
            ilRet = btrClose(hmCff)
            btrDestroy hmCff
            Exit Sub
        End If
        hmNTR = CBtrvTable(TWOHANDLES)
        ilRet = btrOpen(hmNTR, "", sgDBPath & "SBF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            smErrorMessage = "Issue opening SBF Table"
            ilRet = btrClose(hmNTR)
            btrDestroy hmNTR
            Exit Sub
        End If
        imRecLenNTR = Len(tlNtr)
    End If
    Exit Sub
ERRORBOX:
    smErrorMessage = Err.Description
    gHandleError "", "CContractCreation-Initialize"
End Sub
Private Sub Class_Terminate()
    btrClose hmCHF
    btrDestroy hmCHF
    btrClose hmClf
    btrDestroy hmClf
    btrClose hmCff
    btrDestroy hmCff
    btrDestroy hmNTR
    Set myXml = Nothing
End Sub
Public Sub Clear(Optional blAll As Boolean)
    'use 'clear true' instead of setting to nothing
    emContractToProcess = NoProcess
    emProcessResult = Cancelled
    smErrorMessage = ""
    tmVehicles.removeAll
    tmDayParts.removeAll
    If blAll Then
        emCreationUser = UnknownUser
        myXml.Clear True
        mClearContractArrays
    Else
        myXml.Clear
    End If
End Sub
Public Function LoadXml(slPath As String, slName As String) As Boolean
    'file exists and xml valid.
    Dim blRet As Boolean
    
    smErrorMessage = ""
    myXml.CreationUser = emCreationUser
    blRet = myXml.Load(slPath, slName, tmVehicles, tmDayParts)
    'I need this before 'load late' to get rate card
    If Not mNTRValidate() Then
        blRet = False
    End If
    If Not mHeaderLoadLate() Then
        blRet = False
    End If

    If Not blRet Then
        If Len(myXml.ErrorMessage) > 0 Then
            smErrorMessage = smErrorMessage & myXml.ErrorMessage
        End If
        myXml.Clear
    End If
On Error GoTo ERRORBOX
    LoadXml = blRet
    Exit Function
ERRORBOX:
    LoadXml = False
    myXml.Clear
    smErrorMessage = Err.Description
End Function
Public Function ValidateXML() As Boolean
    Dim blRet As Boolean
    
    smErrorMessage = ""
    blRet = False
On Error GoTo ERRORBOX
    If emCreationUser = Filemaker Then
        If UBound(tgNTRInfo) > 0 Or UBound(tgCff) > 0 Then
            blRet = mHeaderValidate()
            If Not mLineValidate() Then
                blRet = False
            End If
'            If Not mNTRValidate() Then
'                blRet = False
'            End If
        Else
            blRet = False
            smErrorMessage = "File must have at least one flight or NTR item defined."
        End If
    Else
        blRet = False
        smErrorMessage = "Creation user not set properly, cannot validate header.  Contact Counterpoint."
    End If
    If Not blRet Then
       ' lmContractNumber = 0
       ' smFileName = ""
       myXml.Clear
    End If
    ValidateXML = blRet
    Exit Function
ERRORBOX:
    ValidateXML = False
    smErrorMessage = "Contact Counterpoint.  Error in validating:" & Err.Description
    myXml.Clear
End Function
Public Function CreateOrder() As Boolean
    Dim blRet As Boolean
    Dim tlChf As CHF
    Dim llChfCode As Long
    Dim c As Integer
    
On Error GoTo ERRORBOX
    blRet = True
    smErrorMessage = ""
    tgChf.lCode = 0
    llChfCode = mInsertToCHF(tgChf)
    If llChfCode > 0 Then
        For c = 0 To UBound(tgClf) - 1
            tgClf(c).ClfRec.lCode = 0
            tgClf(c).ClfRec.lChfCode = llChfCode
            tgClf(c).ClfRec.iUrfCode = tgUrf(0).iCode
            If Not mInsertToCLF(tgClf(c).ClfRec) Then
                blRet = False
                Exit For
            End If
        Next c
        If blRet Then
            For c = 0 To UBound(tgCff) - 1
                tgCff(c).CffRec.lCode = 0
                tgCff(c).CffRec.lChfCode = llChfCode
                If Not mInsertToCFF(tgCff(c).CffRec) Then
                    blRet = False
                    Exit For
                End If
            Next c
        Else
        'clf failed
        End If
        'cff failed
        If blRet Then
            For c = 0 To UBound(tgNTRInfo) - 1
                tgNTRInfo(c).tNTR.lCode = 0
                tgNTRInfo(c).tNTR.lChfCode = llChfCode
                tgNTRInfo(c).tNTR.iUrfCode = tgUrf(0).iCode
                 If Not mInsertToNTR(tgNTRInfo(c).tNTR) Then
                    blRet = False
                    Exit For
                End If
            Next c
        End If
    'chf failed
    Else
        blRet = False
        smErrorMessage = smErrorMessage & " " & myXml.ErrorMessage
    End If
    CreateOrder = blRet
    Exit Function
ERRORBOX:
    CreateOrder = False
    smErrorMessage = Err.Description
End Function
Public Function DeleteOrder() As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim llRet As Long
    Dim c As Integer
    
    blRet = True
    smErrorMessage = ""
On Error GoTo ERRORBOX
    slSql = "delete from CFF_Contract_Flight where cffchfcode = " & tgChf.lCode
    If gSQLWaitNoMsgBox(slSql, False) = 0 Then
        'does not show false if comment is not removed
        For c = 0 To UBound(tgClf) - 1
            slSql = "delete from CXF_Hdr_Ln_Comments where cxfcode = " & tgClf(c).ClfRec.lCxfCode
             If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                gLogMsg "cContractCreation-DeleteOrders had an issue.  Could not delete comments for order #" & tgChf.lCntrNo, "TrafficErrors.txt", False
             End If
        Next c
        slSql = "delete from CLF_Contract_Line where clfchfcode = " & tgChf.lCode
        If gSQLWaitNoMsgBox(slSql, False) = 0 Then
            slSql = "delete from SBF_Special_Billing where sbfchfcode = " & tgChf.lCode
            If gSQLWaitNoMsgBox(slSql, False) = 0 Then
                'Dan this could easily be a sql call, but I'd already written it this way
                'blRet = mDeleteFromCHF(myXml.ContractNumber)
                slSql = "delete  from CHF_Contract_Header  where chfCntrNo = " & myXml.ContractNumber
                If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                   gLogMsg "cContractCreation-DeleteOrders had an issue.  Could not delete from Contract Header for order #" & tgChf.lCntrNo, "TrafficErrors.txt", False
                End If
            Else
                blRet = False
                smErrorMessage = "could not delete ntr lines"
            End If
        Else
            blRet = False
            smErrorMessage = "could not delete order lines"
        End If
    Else
        blRet = False
        smErrorMessage = "could not delete order flights"
    End If
    'delete ntr!
    DeleteOrder = blRet
    Exit Function
ERRORBOX:
    DeleteOrder = False
    smErrorMessage = Err.Description
    gHandleError "", CLASSNAME & ",mDeleteOrder"

End Function
Private Sub mClearContractArrays()
    Erase tgClf
    Erase tgCff
End Sub
Private Function mInsertToCHF(tlChf As CHF) As Long
    Dim ilRet As Integer
    
On Error GoTo ERRORBOX
    Do
        ilRet = btrInsert(hmCHF, tlChf, imRecLenCHF, INDEXKEY0)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet = BTRV_ERR_NONE Then
        mInsertToCHF = tlChf.lCode
    Else
        mInsertToCHF = -1
        smErrorMessage = "mInsertToCHF had an issue.  Could not insert."
    End If
    Exit Function
ERRORBOX:
    smErrorMessage = Err.Description
    mInsertToCHF = -1
    gHandleError "", CLASSNAME & ",mInsertToCHF"
End Function
Private Function mDeleteFromCHF(llContractToDelete As Long) As Boolean
    Dim ilRet As Integer
    Dim tlChf As CHF
    Dim ilFinder As CHFKEY1
On Error GoTo ERRORBOX
    
    ilFinder.lCntrNo = llContractToDelete
    ilFinder.iPropVer = 0
    ilFinder.iCntRevNo = 0
    ilRet = btrGetEqual(hmCHF, tlChf, imRecLenCHF, ilFinder, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        btrDelete hmCHF
        mDeleteFromCHF = True
    Else
        mDeleteFromCHF = False
        smErrorMessage = "mDeleteFromCHF had an issue.  Could not find record to delete."
    End If
    Exit Function
ERRORBOX:
    smErrorMessage = Err.Description
    mDeleteFromCHF = False
    gHandleError "", CLASSNAME & ",mDeleteFromCHF"
End Function
Private Function mInsertToCLF(tlClf As CLF) As Boolean
    Dim blRet As Boolean
    Dim ilRet As Integer
    
    blRet = True
On Error GoTo ERRORBOX
    Do
        ilRet = btrInsert(hmClf, tlClf, imRecLenCLF, INDEXKEY0)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        blRet = False
        smErrorMessage = "mInsertToCLF had an issue.  Could not insert."
    End If
    mInsertToCLF = blRet
    Exit Function
ERRORBOX:
    smErrorMessage = Err.Description
    mInsertToCLF = False
    gHandleError "", CLASSNAME & ",mInsertToCLF"
End Function
Private Function mInsertToCFF(tlCff As CFF) As Boolean
    Dim blRet As Boolean
    Dim ilRet As Integer
    
    blRet = True
On Error GoTo ERRORBOX
    Do
        ilRet = btrInsert(hmCff, tlCff, imRecLenCFF, INDEXKEY0)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        blRet = False
        smErrorMessage = "mInsertToCFF had an issue.  Could not insert."
    End If
    mInsertToCFF = blRet
    Exit Function
ERRORBOX:
    smErrorMessage = Err.Description
    mInsertToCFF = False
    gHandleError "", CLASSNAME & ",mInsertToCFF"
End Function
Private Function mInsertToNTR(tlNtr As SBF) As Boolean
    Dim blRet As Boolean
    Dim ilRet As Integer
    
    blRet = True
On Error GoTo ERRORBOX
    Do
        ilRet = btrInsert(hmNTR, tlNtr, imRecLenNTR, INDEXKEY0)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        blRet = False
        smErrorMessage = "mInsertToNTR had an issue.  Could not insert."
    End If
    mInsertToNTR = blRet
    Exit Function
ERRORBOX:
    smErrorMessage = Err.Description
    mInsertToNTR = False
    gHandleError "", CLASSNAME & ",mInsertToNTR"
End Function
Private Function mHeaderValidate() As Boolean

    Dim blRet As Boolean
    Dim slElement As String
    Dim blAgencyAndAdvOK As Boolean
    
    blRet = True
    blAgencyAndAdvOK = True
    slElement = "Contract Number"
    If mContractExists() Then
        blRet = False
        smErrorMessage = smErrorMessage & " " & slElement & " is not valid."
    End If
    slElement = "Advertiser Name"
    If Not mAdvertiserExists() Then
        blRet = False
        blAgencyAndAdvOK = False
        smErrorMessage = smErrorMessage & " " & slElement & " is not valid."
    Else
    '8809 will come back true if not in xml...it's optional
        slElement = "Product Name"
        If Not mProductExists() Then
            blRet = False
            smErrorMessage = smErrorMessage & " " & slElement & " is not valid."
        End If
    End If
    'will come back true if not in xml...it's optional
    slElement = "Agency Name or Agency City ID"
    If Not mAgencyExists() Then
        blRet = False
        blAgencyAndAdvOK = False
        smErrorMessage = smErrorMessage & " " & slElement & " is not valid."
    End If
    'needs adf/agf code, so run after them
    If blAgencyAndAdvOK Then
        slElement = "Buyer Name"
        If Not mBuyerExists() Then
            blRet = False
            smErrorMessage = smErrorMessage & " " & slElement & " is not valid."
        End If
    Else
        smErrorMessage = smErrorMessage & " Buyer Name cannot be tested."
    End If
    slElement = "Salesperson"
    If Not mSalesPersonExists() Then
        blRet = False
        smErrorMessage = smErrorMessage & " " & slElement & " is not valid."
    End If
    mHeaderValidate = blRet
End Function
Private Function mLineValidate() As Boolean
    Dim blRet As Boolean
    Dim slVehicle As String
    Dim slDayPart As String
    Dim c As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    
    blRet = True
    For c = 0 To UBound(tgClf) - 1
        If tmVehicles.Exists(tgClf(c).ClfRec.iLine) Then
            slVehicle = tmVehicles.Item(tgClf(c).ClfRec.iLine)
            If mVehicleExists(slVehicle, c) Then
                If Not mSpotLengthExists(tgClf(c).ClfRec.iVefCode, tgClf(c).ClfRec.iLen) Then
                    smErrorMessage = smErrorMessage & " Spot length for " & slVehicle & " is not valid."
                    blRet = False
                End If
            Else
                smErrorMessage = smErrorMessage & " VehicleName of " & slVehicle & " is not valid. SpotLength and DaypartName cannot be tested."
                slVehicle = ""
                blRet = False
            End If
        Else
            blRet = False
            slVehicle = ""
            smErrorMessage = smErrorMessage & " VehicleName is missing"
        End If
        If Len(slVehicle) > 0 Then
            If tmDayParts.Exists(tgClf(c).ClfRec.iLine) Then
                slDayPart = tmDayParts.Item(tgClf(c).ClfRec.iLine)
                If Not mDaypartExists(slDayPart, c) Then
                    smErrorMessage = smErrorMessage & " DaypartName of " & slDayPart & " is not valid."
                    blRet = False
                End If
            Else
                blRet = False
                smErrorMessage = smErrorMessage & " DaypartName is missing"
            End If
        End If
        gUnpackTimeLong tgClf(c).ClfRec.iStartTime(0), tgClf(c).ClfRec.iStartTime(1), False, llStartTime
        gUnpackTimeLong tgClf(c).ClfRec.iEndTime(0), tgClf(c).ClfRec.iEndTime(1), False, llEndTime
        If llEndTime < llStartTime Then
            blRet = False
            smErrorMessage = smErrorMessage & " EndTime cannot be earlier than StartTime.  Line #" & tgClf(c).ClfRec.iLine
        End If
    Next c
    mLineValidate = blRet
End Function
Private Function mNTRValidate() As Boolean
    Dim blRet As Boolean
    Dim slVehicle As String
    Dim slNTRType As String
    Dim c As Integer
    Dim ilDictionaryKey As Integer
    
    blRet = True
    For c = 0 To UBound(tgNTRInfo) - 1
        'negative numbers starting at -1!
        ilDictionaryKey = (c * -1) - 1
        If tmVehicles.Exists(ilDictionaryKey) Then
            slVehicle = tmVehicles.Item(ilDictionaryKey)
            If mVehicleExists(slVehicle, c, True) Then
'                If Not mSpotLengthExists(tgNTRInfo(c).ClfRec.iVefCode, tgNTRInfo(c).ClfRec.iLen) Then
'                    smErrorMessage = smErrorMessage & " Spot length of " & tgNTRInfo(c).ClfRec.iLen & " for " & slVehicle & " is not valid."
'                    blRet = False
'                End If
            Else
                smErrorMessage = smErrorMessage & " VehicleName of " & slVehicle & " is not valid."
                slVehicle = ""
                blRet = False
            End If
        Else
            blRet = False
            slVehicle = ""
            smErrorMessage = smErrorMessage & " VehicleName is missing"
        End If
        If Len(slVehicle) > 0 Then
            If tmDayParts.Exists(ilDictionaryKey) Then
                slNTRType = tmDayParts.Item(ilDictionaryKey)
                If Not mNTRTypeExistsWithOptionals(slNTRType, c) Then
                    smErrorMessage = smErrorMessage & " NTRType  of " & slNTRType & " is not valid."
                    blRet = False
                End If
            Else
                blRet = False
                smErrorMessage = smErrorMessage & " NTRType is missing"
            End If
        End If
    Next c
    mNTRValidate = blRet
End Function
Private Function mAdvertiserExists() As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    
    blRet = False
    slSql = "select adfCode as Code from Adf_Advertisers where adfName = '" & myXml.Advertiser & "'"
    Set myRsXml = gSQLSelectCall(slSql)
    If Not myRsXml.EOF Then
        blRet = True
        tgChf.iAdfCode = myRsXml!Code
    End If
    mAdvertiserExists = blRet
End Function
Private Function mProductExists() As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    
    If Len(myXml.Product) > 0 Then
        blRet = False
        If tgChf.iAdfCode > 0 Then
            slSql = "select prfCode as Code from Prf_Product_Names where prfName = '" & myXml.Product & "' AND prfadfcode = " & tgChf.iAdfCode
            Set myRsXml = gSQLSelectCall(slSql)
            If Not myRsXml.EOF Then
                blRet = True
                tgChf.sProduct = myXml.Product
               ' imPrfCode = myRsXml!Code
            End If
        End If
    Else
        blRet = True
        tgChf.sProduct = ""
    End If
    mProductExists = blRet
End Function
Private Function mAgencyExists() As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    
    blRet = False
    'optional
    If Len(myXml.Agency) = 0 And Len(myXml.AgencyCity) = 0 Then
        blRet = True
    Else
        slSql = "select agfCode from Agf_Agencies where agfName = '" & myXml.Agency & "' AND agfCity = '" & myXml.AgencyCity & "'"
        Set myRsXml = gSQLSelectCall(slSql)
        If Not myRsXml.EOF Then
            blRet = True
            tgChf.iAgfCode = myRsXml!agfCode
        End If
    End If
    mAgencyExists = blRet
End Function
Private Function mBuyerExists() As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    
    blRet = False
    slSql = "select pnfCode as Code,pnfPhone from Pnf_Personnel_Names where pnfType = 'B' AND pnfName = '" & myXml.Buyer & "'"
    If tgChf.iAgfCode > 0 Then
        slSql = slSql & " AND pnfagfCode = " & tgChf.iAgfCode
    ElseIf tgChf.iAdfCode > 0 Then
        slSql = slSql & " AND pnfadfCode = " & tgChf.iAdfCode
    Else
        mBuyerExists = False
        Exit Function
    End If
    Set myRsXml = gSQLSelectCall(slSql)
    If Not myRsXml.EOF Then
        blRet = True
        tgChf.iPnfBuyer = myRsXml!Code
        tgChf.sPhone = myRsXml!pnfPhone
        tgChf.sBuyer = myXml.Buyer
    End If
    mBuyerExists = blRet
End Function
Private Function mSalesPersonExists() As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim slFirst As String
    Dim slLast As String
    Dim ilPos As Integer
    
    blRet = False
    ilPos = InStr(myXml.SalesPerson, " ")
    If ilPos > 0 And ilPos < Len(myXml.SalesPerson) Then
        slFirst = Mid(myXml.SalesPerson, 1, ilPos)
        slLast = Mid(myXml.SalesPerson, ilPos + 1)
        slSql = "select slfCode as Code from SLF_Salespeople where slfFirstName = '" & slFirst & "' and slfLastName = '" & slLast & "'"
        Set myRsXml = gSQLSelectCall(slSql)
        If Not myRsXml.EOF Then
            blRet = True
            tgChf.iSlfCode(0) = myRsXml!Code
        End If
    End If
    mSalesPersonExists = blRet
End Function
Private Function mContractExists() As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim slFirst As String
    Dim slLast As String
    Dim ilPos As Integer
    
    blRet = False
    slSql = "select chfCode as Code from CHF_Contract_Header where chfCntrNo = " & myXml.ContractNumber
    Set myRsXml = gSQLSelectCall(slSql)
    If myRsXml.EOF Then
        tgChf.lCntrNo = myXml.ContractNumber
    Else
        blRet = True
    End If
    mContractExists = blRet
End Function
Private Function mVehicleExists(slVehicleName As String, ilIndex As Integer, Optional blIsNTR As Boolean = False) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    
    blRet = False
    If Len(slVehicleName) > 0 Then
        slSql = "Select  vefCode from VEF_Vehicles Where vefName = '" & slVehicleName & "'"
        Set myRsXml = gSQLSelectCall(slSql)
        If Not myRsXml.EOF Then
            blRet = True
            If blIsNTR Then
                tgNTRInfo(ilIndex).tNTR.iAirVefCode = myRsXml!vefCode
                tgNTRInfo(ilIndex).tNTR.iBillVefCode = myRsXml!vefCode
            Else
                tgClf(ilIndex).ClfRec.iVefCode = myRsXml!vefCode
            End If
        End If
    End If
    mVehicleExists = blRet
End Function
Private Function mDaypartExists(slDayPartName As String, ilIndex As Integer) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
        
    blRet = False
    If Len(slDayPartName) > 0 Then
        slSql = "select rdfCode as code from RDF_Standard_Daypart  inner join RIF_Rate_Card_Items  on rdfcode = rifrdfcode where rifrcfcode = " & tgChf.iRcfCode & " and rifvefcode = " & tgClf(ilIndex).ClfRec.iVefCode & " and rdfname = '" & slDayPartName & "'"
        Set myRsXml = gSQLSelectCall(slSql)
        If Not myRsXml.EOF Then
            blRet = True
            tgClf(ilIndex).ClfRec.iRdfCode = myRsXml!Code
        End If
    End If
    mDaypartExists = blRet
End Function
Private Function mRateCardExists(slStart As String) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    
    blRet = False
    slSql = "Select top 1 rcfCode from RCF_Rate_Card where rcfstartdate <=  '" & Format$(slStart, sgSQLDateForm) & "' order by rcfStartdate desc"
    Set myRsXml = gSQLSelectCall(slSql)
    If Not myRsXml.EOF Then
        blRet = True
        tgChf.iRcfCode = myRsXml!rcfCode
    End If
    mRateCardExists = blRet
End Function
Private Function mSpotLengthExists(ilVefCode As Integer, ilSpotLength As Integer) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim ilFound As Integer
    Dim ilSpotIndex As Integer
    
    blRet = False
    ilFound = gBinarySearchVpf(ilVefCode)
        If ilFound <> -1 Then
            For ilSpotIndex = 0 To UBound(tgVpf(ilFound).iSLen)
                If tgVpf(ilFound).iSLen(ilSpotIndex) = ilSpotLength Then
                    blRet = True
                    Exit For
                End If
            Next ilSpotIndex
        End If
    mSpotLengthExists = blRet
End Function
Private Function mNTRTypeExistsWithOptionals(slDayPartName As String, ilIndex As Integer) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim slTemp As String
    blRet = False
On Error GoTo ERRORBOX
    With tgNTRInfo(ilIndex).tNTR
        If Len(slDayPartName) > 0 Then
            slSql = "select mnfCode as code, mnfRPU as RatePerItem, mnfssComm as SalesCommission, mnfCost as Acquisition from MNF_Multi_Names where mnfType = 'I' AND mnfName = '" & slDayPartName & "'"
            Set myRsXml = gSQLSelectCall(slSql)
            If Not myRsXml.EOF Then
                blRet = True
                .iMnfItem = myRsXml!Code
                If .lGross = NORATE Then
                    .lGross = gStrDecToLong(myRsXml!rateperitem, 2)
                End If
                 If .iCommPct = NORATE Then
                    .iCommPct = gStrDecToInt(myRsXml!salescommission, 4)
                   ' gPDNToLong slTEmp, .iCommPct  'or gstrdectolong
                End If
                If .lAcquisitionCost = NORATE Then
                    .lAcquisitionCost = myRsXml!acquisition
                End If
            End If
        End If
    End With
    mNTRTypeExistsWithOptionals = blRet
    Exit Function
ERRORBOX:
    mNTRTypeExistsWithOptionals = False
    smErrorMessage = "mNTRTypeExists: " & Err.Description & " Contact Counterpoint"
End Function
Private Function mHeaderLoadLate() As Boolean
    Dim blRet As Boolean
    Dim c As Integer
    Dim llGross As Long
    Dim llGrossNtr As Long
    Dim ilDailyRate As Integer
    Dim blAtLeastOne As Boolean
    
    'input gross is ntr + airtime.
    'chf with what we got from cff
    blAtLeastOne = False
    blRet = mRateCardExists(myXml.StartDate)
    If blRet Then
        llGross = 0
        gPackDate myXml.StartDate, tgChf.iStartDate(0), tgChf.iStartDate(1)
        gPackDate myXml.EndDate, tgChf.iEndDate(0), tgChf.iEndDate(1)
        For c = 0 To UBound(tgCff) - 1
            With tgCff(c).CffRec
                blAtLeastOne = True
                If .sDyWk = "W" Then
                    llGross = llGross + (.lActPrice * .iSpotsWk)
                Else
                    For ilDailyRate = 0 To UBound(.iDay)
                        llGross = llGross + (.lActPrice * .iDay(ilDailyRate))
                    Next ilDailyRate
                End If
            End With
        Next c
        tgChf.lAirTimeGross = llGross
        If UBound(tgNTRInfo) > 0 Then
            blAtLeastOne = True
            llGross = 0
            For c = 0 To UBound(tgNTRInfo) - 1
                With tgNTRInfo(c).tNTR
                    llGross = llGross + (.lGross * .iNoItems)
                End With
            Next c
            tgChf.sNTRDefined = "Y"
            tgChf.lNTRGross = llGross
            tgChf.lInputGross = tgChf.lAirTimeGross + llGross
        Else
            tgChf.lInputGross = llGross
            tgChf.lNTRGross = 0
            tgChf.sNTRDefined = "N"
        End If
        If Not blAtLeastOne Then
            blRet = False
            smErrorMessage = smErrorMessage & " Must have at least one flight or item defined."
        End If
    Else
        If UBound(tgClf) = 0 And UBound(tgCff) = 0 And UBound(tgNTRInfo) = 0 Then
            smErrorMessage = smErrorMessage & " Must have at least one flight or item defined."
        Else
            If myXml.StartDate = NODATE Then
                smErrorMessage = smErrorMessage & " Issue with finding start date for order."
            Else
                smErrorMessage = smErrorMessage & " Rate card could not be found for date of " & myXml.StartDate
            End If
        End If
    End If
    mHeaderLoadLate = blRet
End Function
