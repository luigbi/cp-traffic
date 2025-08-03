Attribute VB_Name = "DigitalVendors"
Option Explicit
Private Const CURRENTAVAILABLEVENDORS = 7
Public Const NODATE As String = "01/01/1970"
Public Enum ContractMethodType
    None = 0
    Manual = 1
    Vendor = 2
    CSI = 3
End Enum
Public Enum ImpressionMethodType
    None = 0
    Manual = 1
    Amazon = 2
    CSI = 3
    Vendor = 4
End Enum
Public Enum DigitalVendorStatus
    DORMANT = 0
    ACTIVE = 1
   ' Paused = 2
End Enum
Public Enum ExternalVehicleIDType
    None = 0
    Allowed = 1
    ReadOnly = 2
End Enum
Public Type VendorInfo
    oContractMethod As ContractMethodType
    oImpressionMethod As ImpressionMethodType
    iCode As Integer
    sName As String
    sVendorUserName As String
    sVendorPassword As String
    oStatus As DigitalVendorStatus
    sVendorURL As String
    sOrgID As String
    sBucketName As String
    sBucketFolder As String
    sBucketRegion As String
    sBucketAccessKey As String
    sBucketPrivateKey As String
    sLastImportDate As String
    iPriorityStart As Integer
    iPriorityEnd As Integer
    bIsCustom As Boolean
    'currently handles OrgID and TimeZone
    bSecondaryID As Boolean
    oExternalVehicleIDType As ExternalVehicleIDType
    sTimeZone As String
End Type
Public Function gGetDigitalVendors(Optional blIncludeAvailable = False) As VendorInfo()
    'active and available?  or just active?
    Dim tlVendorsToReturn() As VendorInfo
    Dim tlAvailableVendors() As VendorInfo
    Dim tlTempVendors() As VendorInfo
    Dim ilIndex As Integer
    Dim ilSubIndex As Integer
    Dim blIsActive As Boolean

    ReDim tlTempVendors(0)
    tlVendorsToReturn = mGetActiveVendors()
    tlAvailableVendors = mGetAvailableVendors()
    'first pass-'find matching available...or not matching (custom). Set methods and if custom and clear fields from legacies
    For ilIndex = 0 To UBound(tlVendorsToReturn) - 1
        For ilSubIndex = 0 To UBound(tlAvailableVendors) - 1
            tlVendorsToReturn(ilIndex).bIsCustom = True
            tlVendorsToReturn(ilIndex).bSecondaryID = False
            tlVendorsToReturn(ilIndex).oContractMethod = ContractMethodType.Manual
            tlVendorsToReturn(ilIndex).oImpressionMethod = ImpressionMethodType.Manual
            tlVendorsToReturn(ilIndex).oExternalVehicleIDType = ExternalVehicleIDType.None
            ' case insensitive-for Legacy vendors  Here and setting name in this if statement are the only spots that it's tested
            If UCase(Trim$(tlAvailableVendors(ilSubIndex).sName)) = UCase(Trim$(tlVendorsToReturn(ilIndex).sName)) Then
                tlVendorsToReturn(ilIndex).oContractMethod = tlAvailableVendors(ilSubIndex).oContractMethod
                tlVendorsToReturn(ilIndex).oImpressionMethod = tlAvailableVendors(ilSubIndex).oImpressionMethod
                tlVendorsToReturn(ilIndex).oExternalVehicleIDType = tlAvailableVendors(ilSubIndex).oExternalVehicleIDType
                tlVendorsToReturn(ilIndex).bSecondaryID = tlAvailableVendors(ilSubIndex).bSecondaryID
                tlVendorsToReturn(ilIndex).bIsCustom = False
                'get name case right in case it's legacy
                tlVendorsToReturn(ilIndex).sName = Trim$(tlAvailableVendors(ilSubIndex).sName)
                tlVendorsToReturn(ilIndex) = mClearForLegacy(tlVendorsToReturn(ilIndex))
                Exit For
            End If
        Next
    Next
    '2nd pass.  what's available that's not active?
    If blIncludeAvailable Then
        For ilIndex = 0 To UBound(tlAvailableVendors) - 1
            blIsActive = False
            For ilSubIndex = 0 To UBound(tlVendorsToReturn) - 1
                If Trim$(tlAvailableVendors(ilIndex).sName) = Trim$(tlVendorsToReturn(ilSubIndex).sName) Then
                    blIsActive = True
                    Exit For
                End If
            Next
            ''AdsWizz' not active? Add to temp array and add to main array in next step
            If Not blIsActive Then
                tlTempVendors(UBound(tlTempVendors)) = tlAvailableVendors(ilIndex)
                tlTempVendors(UBound(tlTempVendors)).oStatus = DORMANT
                ReDim Preserve tlTempVendors(UBound(tlTempVendors) + 1)
            End If
        Next
        For ilIndex = 0 To UBound(tlTempVendors) - 1
            tlVendorsToReturn(UBound(tlVendorsToReturn)) = tlTempVendors(ilIndex)
            ReDim Preserve tlVendorsToReturn(UBound(tlVendorsToReturn) + 1)
        Next
    End If
    gGetDigitalVendors = tlVendorsToReturn
End Function
Public Function gIsVendor(slVendorName As String) As Boolean
    '11010
    Dim blRet As Boolean
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    SQLQuery = "SELECT count(avfName) as amount FROM avf_AdVendor where avfName = '" & slVendorName & "'"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If rst!amount = 1 Then
            blRet = True
        End If
    End If
    gIsVendor = blRet
    Exit Function
ErrHand:
     gHandleError "TrafficErrors.txt", "mIsVendor"
End Function
Private Function mGetActiveVendors() As VendorInfo()
    Dim slSql As String
    Dim rst As ADODB.Recordset
    Dim tlVendor As VendorInfo
    Dim tlActiveVendors() As VendorInfo
    Dim ilCount As Integer

    ilCount = 0
    ReDim tlActiveVendors(ilCount)
    slSql = "Select * from avf_AdVendor order by avfName"
    Set rst = gSQLSelectCall(slSql)
    Do While Not rst.EOF
        With tlVendor
            .iCode = rst!avfCode
            .sName = Trim(rst!avfName)
            .sVendorUserName = Trim(rst!avfVendorUserName)
            .sVendorPassword = Trim(rst!avfVendorPassword)
            If rst!avfStatus = "D" Then
                .oStatus = DORMANT
            Else
                .oStatus = ACTIVE
            End If
            .sVendorURL = Trim(rst!avfVendorURL)
            .sOrgID = Trim(rst!avfOrganizationID)
            .sTimeZone = Trim(rst!avfTimeZone)
            .sBucketName = Trim(rst!avfBucketName)
            .sBucketFolder = Trim(rst!avfBucketFolder)
            .sBucketRegion = Trim(rst!avfBucketRegion)
            .sBucketAccessKey = Trim(rst!avfBucketAccessKey)
            .sBucketPrivateKey = Trim(rst!avfBucketPrivatekey)
            If IsNull(rst!avfLastImportDate) Then
                .sLastImportDate = NODATE
            Else
                .sLastImportDate = Format(rst!avfLastImportDate, sgSQLDateForm)
            End If
            .iPriorityStart = rst!avfprioritystart
            .iPriorityEnd = rst!avfpriorityend
        End With
        tlActiveVendors(ilCount) = tlVendor
        ilCount = ilCount + 1
        ReDim Preserve tlActiveVendors(ilCount)
        rst.MoveNext
    Loop
    mGetActiveVendors = tlActiveVendors
End Function
Private Function mGetAvailableVendors() As VendorInfo()
    Dim ilTotal As Integer
    Dim tlVendor As VendorInfo
    Dim tlMyAvailableVendors() As VendorInfo
    Dim ilCount As Integer
    Dim ilCurrent As Integer
    'if some digital vendors are in site options
    Dim blAdd As Boolean

    ilTotal = CURRENTAVAILABLEVENDORS
    ReDim tlMyAvailableVendors(0)
    For ilCount = 1 To ilTotal
        blAdd = True
        With tlVendor
            .oExternalVehicleIDType = ExternalVehicleIDType.None
            .bSecondaryID = False
            .sTimeZone = ""
            Select Case ilCount
                Case 1
                    .sName = "AdsWizz"
                    .oContractMethod = ContractMethodType.Manual
                    .oImpressionMethod = ImpressionMethodType.Manual
                Case 2
                    .sName = "Boostr"
                    .oContractMethod = ContractMethodType.CSI
                    .oImpressionMethod = ImpressionMethodType.CSI
                    .oExternalVehicleIDType = ExternalVehicleIDType.ReadOnly
                Case 3
                    .sName = "Manual"
                    .oContractMethod = ContractMethodType.Manual
                    .oImpressionMethod = ImpressionMethodType.Manual
                Case 4
                    .sName = "Megaphone"
                    .oContractMethod = ContractMethodType.Vendor
                    .oImpressionMethod = ImpressionMethodType.Amazon
                    .oExternalVehicleIDType = ExternalVehicleIDType.Allowed
                    .bSecondaryID = True
                Case 5
                    .sName = "RAB"
                    .oContractMethod = ContractMethodType.None
                    .oImpressionMethod = ImpressionMethodType.None
                Case 6
                    .sName = "Spreaker"
                    .oContractMethod = ContractMethodType.Manual
                    .oImpressionMethod = ImpressionMethodType.Manual
                Case 7
                    .sName = "TAP"
                    .oContractMethod = ContractMethodType.Manual
                    .oImpressionMethod = ImpressionMethodType.Manual
            End Select
        End With
        If blAdd Then
            ilCurrent = UBound(tlMyAvailableVendors)
            tlMyAvailableVendors(ilCurrent) = tlVendor
            ReDim Preserve tlMyAvailableVendors(ilCurrent + 1)
        End If
    Next
    mGetAvailableVendors = tlMyAvailableVendors
End Function
Private Function mClearForLegacy(tlVendor As VendorInfo) As VendorInfo
    ' ImpressionMethod fields are not legacy.  Don't have to do, as well as contractMethod new fields vendorurl and orgID
    With tlVendor
        Select Case .oContractMethod
            Case ContractMethodType.Manual
                If .oImpressionMethod = ImpressionMethodType.Manual Then
                    .sVendorPassword = ""
                    .sVendorUserName = ""
                End If
            Case ContractMethodType.None
                .iPriorityStart = 0
                .iPriorityEnd = 0
        End Select
    End With
    mClearForLegacy = tlVendor
End Function


Private Function mPrepRecordsetVendors() As Recordset
    Dim myRs As ADODB.Recordset

    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "Code", adInteger
            .Append "Name", adChar, 255
            .Append "UserName", adChar, 255
            .Append "Password", adChar, 255
            .Append "Status", adChar, 1
            .Append "APIURL", adChar, 255
            .Append "OrganizationID", adChar, 70
            .Append "BucketName", adChar, 63
            .Append "BucketFolder", adChar, 80
            .Append "BucketRegion", adChar, 16
            .Append "BucketAccessKey", adChar, 20
            .Append "BucketPrivateKey", adChar, 40
            .Append "LastImportDate", adDate
            .Append "PriorityStart", adInteger
            .Append "PriorityEnd", adInteger
            .Append "ContractMethod", adInteger
            .Append "ImpressionMethod", adInteger
        End With
    myRs.Open
    myRs!Name.Properties("optimize") = True
    myRs.Sort = "Name"
    Set mPrepRecordsetVendors = myRs
End Function
