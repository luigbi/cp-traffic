Attribute VB_Name = "modVendors"
Option Explicit
'7967 does client have a web vendor set up?  not demomode and using web? Web database in sync?  If all true, allow connecting to web.  Set in function gVendorToWebAllowed
Public bgVendorToWebAllowed As Boolean
Public bgVendorExportSent As Boolean
'7701
Public Enum Vendors
    '8484
    AnalyticOwl = 25
    BSI = 118
    cBs = 1
    '9184
    stratus = 2
    Jelli = 3
    'clearchannel became iheart: 8324 I changed this name to iheart too.
    iheart = 4
    iDc = 111
    Wegener_Compel = 112
    Wegener_IPump = 113
    XDS_Break = 114
    XDS_ISCI = 115
    'now radio interchange by wide orbit
    WideOrbit = 21
    NetworkConnect = 22
    MrMaster = 116
    RadioTraffic = 23
    RCS = 24
    RadioWorkflow = 26
    Synchronicity = 117
    LeadsRX = 81
    Web = 0
    None = 0
End Enum
Public Enum MethodType
    None = 0
    csiStationService = 1
    csiVendorService = 2
    BothServices = 3
End Enum
Public Enum DeliveryType
    Audio = 0
    Logs = 1
    None = 2
End Enum
'9204
Public Enum VendorType
    Delivery = 0
    Informational = 1
End Enum
Public Enum VendorWvmAlert
    None = 0
    ImportTest = 1
    ImportMissed = 2
    ExportTest = 3
    ConnectionIssue = 4
    ExportRunning = 5
    WebVendorSent = 6
    ExportMissed = 7
    ExportError = 8
    NonMonitorIssue = 9
    ImportError = 10
    ManagerNotRunning = 11
    UserSetTime = 12
End Enum
Public Type VendorInfo
    iIdCode As Integer
    sName As String
    sApprovalPassword As String
    sVendorPassword As String
    sVendorUserName As String
    sAddress As String
    sStationPassword As String
    sStationUserName As String
    iExportMethod As Integer
    iImportMethod As Integer
    sSendUpdatesOnly As String
    iHierarchy As Integer
    sDeliveryType As String
    sIsOverridable As String
    bIsActive As Boolean
    sSourceName As String
    '8418
    iMinimumWebVersion As Integer
    '8457
    sAllowAutoPosting As String
    '8862
    bAutoVendingCSIAllowed As Boolean
    sAutoVendor As String
End Type
Private Enum ExportInfoIndex
    'Export order: code,attcode,vendoridcode,hasbeensent errormessage exportMondayDate,ProcessedDateTime EnteredDateTime
    code = 0
    attCode = 1
    vendoridcode = 2
    hasbeensent = 3
    result = 3
    Message = 4
    mondaydate = 5
    ProcessedDateTime = 6
    entereddatetime = 7
    Service = 7
    spotOkcount = 8
    spotErrorcount = 9
End Enum
Public Type ServiceControllerInfo
    Mode As String
    GenerateDebug As Boolean
    GenerateFile As Boolean
    ImportLast As String
    ImportFiles As String
    IsRunning As Boolean
    ImportSpan As String
End Type
'9818
Public Const SHAREDHEADENDID As String = "SHAREDHEADENDID"
Public Function gIsVendorWithAgreement(llAtt As Long, ilVendor As Integer) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim myRst As ADODB.Recordset
    
    blRet = False
    slSql = "select * from VAT_Vendor_Agreement where vatattcode = " & llAtt & " AND vatwvtVendorId = " & ilVendor
On Error GoTo Err_Handler
    Set myRst = gSQLSelectCall(slSql)
    If Not myRst.EOF Then
        blRet = True
    End If
Cleanup:
    If Not myRst Is Nothing Then
        If (myRst.State And adStateOpen) <> 0 Then
            myRst.Close
        End If
        Set myRst = Nothing
    End If
    gIsVendorWithAgreement = blRet
    Exit Function
Err_Handler:
    gHandleError "", "modVendors-gIsVendorWithAgreement"
    blRet = False
    GoTo Cleanup
End Function
Public Function gIsVendorMethodExportWithAgreement(llAtt As Long, ilVendorMethod As Integer) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim myRst As ADODB.Recordset
    Dim myVendors() As VendorInfo
    Dim iLoop As Integer
    
    blRet = False
    myVendors = gGetActiveDeliveryVendors()
    For iLoop = 0 To UBound(myVendors) - 1 Step 1
        If myVendors(iLoop).iExportMethod = ilVendorMethod Then
            slSql = "select * from VAT_Vendor_Agreement where vatattcode = " & llAtt & " AND vatwvtVendorId = " & myVendors(iLoop).iIdCode
On Error GoTo Err_Handler
            Set myRst = gSQLSelectCall(slSql)
            If Not myRst.EOF Then
                blRet = True
            End If
        End If
        If blRet Then
            Exit For
        End If
    Next iLoop
Cleanup:
    If Not myRst Is Nothing Then
        If (myRst.State And adStateOpen) <> 0 Then
            myRst.Close
        End If
        Set myRst = Nothing
    End If
    Erase myVendors
    gIsVendorMethodExportWithAgreement = blRet
    Exit Function
Err_Handler:
    gHandleError "", "modVendors-gIsVendorMethodExportWithAgreement"
    blRet = False
    GoTo Cleanup
End Function
'Public Function gSaveVAT(llAtt As Long, ilLog As Integer, ilAudio As Integer) As Boolean
'    Dim blRet As Boolean
'    Dim llCount As Long
'
'    blRet = True
'    If llAtt > 0 And (ilLog > 0 Or ilAudio > 0) Then
'        SQLQuery = "update VAT_Vendor_Agreement set vatWvtIdCodeLog = " & ilLog & ", vatwvtIDcodeaudio = " & ilAudio & " Where vatattcode = " & llAtt
'        If gSQLAndReturn(SQLQuery, False, llCount) <> 0 Then
'            '6/13/16: Replaced GoSub
'            'GoSub ErrHand:
'            gHandleError "AffErrorLog.txt", "modVendors-gSaveVAT"
'            gSaveVAT = False
'            Exit Function
'        End If
'        If llCount = 0 Then
'            SQLQuery = "Insert into VAT_Vendor_Agreement (vatattcode,vatwvtidcodelog,vatwvtidcodeaudio) Values ( " & llAtt & " , " & ilLog & " , " & ilAudio & " )"
'            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                '6/13/16: Replaced GoSub
'                'GoSub ErrHand:
'                gHandleError "AffErrorLog.txt", "modVendors-gSaveVAT"
'                gSaveVAT = False
'                Exit Function
'            End If
'        End If
'    End If
'    gSaveVAT = blRet
'    Exit Function
'ErrHand:
'    gSaveVAT = False
'    gHandleError "", "modVendors-gSaveVAT"
'End Function
Public Function gSaveVATMulti(llAtt As Long, ilLog() As Integer, ilAudio() As Integer) As Boolean
    Dim blRet As Boolean
    Dim llCount As Long
    Dim c As Integer
    Dim slSql As String
    Dim slSQLQuery As String
    Dim myRst As ADODB.Recordset
  '  Dim blSkip As Boolean
    Dim blFound As Boolean
    
    blRet = True
    If llAtt > 0 Then
On Error GoTo ErrHand
        'delete those missing. update those that didn't change.  Then add new ones
        slSql = "Select * from VAT_Vendor_Agreement where vatattcode = " & llAtt
        Set myRst = gSQLSelectCall(slSql)
        Do While Not myRst.EOF
            blFound = False
            For c = 0 To UBound(ilLog) - 1
                If ilLog(c) = myRst!vatwvtvendorid Then
                    blFound = True
                    Exit For
                End If
            Next c
            If Not blFound Then
                For c = 0 To UBound(ilAudio) - 1
                    If ilAudio(c) = myRst!vatwvtvendorid Then
                        blFound = True
                        Exit For
                    End If
                Next c
            End If
            'not found, delete
            If Not blFound Then
On Error Resume Next
                If bgVendorToWebAllowed And myRst!vatSentToWeb = "Y" Then
                    slSql = "delete from webVendors_Header where attcode = " & llAtt & " and vendorid = " & myRst!vatwvtvendorid
                    gExecWebSQLWithRowsEffected slSql
                End If
On Error GoTo ErrHand
                slSQLQuery = "delete from VAT_Vendor_Agreement WHERE vatattcode = " & llAtt & " AND vatwvtvendorid = " & myRst!vatwvtvendorid
                If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    gHandleError "AffErrorLog.txt", "modVendors-gSaveVATMulti"
                    gSaveVATMulti = False
                    Exit Function
                End If
            End If
            myRst.MoveNext
        Loop
        If Not myRst Is Nothing Then
            If (myRst.State And adStateOpen) <> 0 Then
                myRst.Close
            End If
            Set myRst = Nothing
        End If
        'now add new
        For c = 0 To UBound(ilLog) - 1
            If ilLog(c) > 0 Then
                If Not gIsVendorWithAgreement(llAtt, ilLog(c)) Then
                    slSQLQuery = "Insert into VAT_Vendor_Agreement (vatattcode,vatwvtvendorId) Values ( " & llAtt & " , " & ilLog(c) & " )"
'                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                        '6/13/16: Replaced GoSub
'                        'GoSub ErrHand:
'                        gHandleError "AffErrorLog.txt", "modVendors-gSaveVATMulti"
'                        gSaveVATMulti = False
'                        Exit Function
'                    End If
                '8344
                Else
                   slSQLQuery = "update Vat_Vendor_Agreement set vatsenttoweb = '' where vatattcode = " & llAtt & " AND vatwvtvendorid = " & ilLog(c)
                End If
                If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    gHandleError "AffErrorLog.txt", "modVendors-gSaveVATMulti"
                    gSaveVATMulti = False
                    Exit Function
                End If
            End If
        Next c
        For c = 0 To UBound(ilAudio) - 1
            If ilAudio(c) > 0 Then
                If Not gIsVendorWithAgreement(llAtt, ilAudio(c)) Then
                    slSQLQuery = "Insert into VAT_Vendor_Agreement (vatattcode,vatwvtvendorid) Values ( " & llAtt & " , " & ilAudio(c) & " )"
'                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                        '6/13/16: Replaced GoSub
'                        'GoSub ErrHand:
'                        gHandleError "AffErrorLog.txt", "modVendors-gSaveVATMulti"
'                        gSaveVATMulti = False
'                        Exit Function
'                    End If
                '8344
                Else
                   slSQLQuery = "update Vat_Vendor_Agreement set vatsenttoweb = '' where vatattcode = " & llAtt & " AND vatwvtvendorid = " & ilAudio(c)
                End If
                If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub ErrHand:
                    gHandleError "AffErrorLog.txt", "modVendors-gSaveVATMulti"
                    gSaveVATMulti = False
                    Exit Function
                End If
            End If
        Next c
    End If
    gSaveVATMulti = blRet
    Exit Function
ErrHand:
    gSaveVATMulti = False
    gHandleError "", "modVendors-gSaveVATMulti"
End Function
Public Function gGetAvailableVendors() As VendorInfo()
    Dim ilCount As Integer
    Dim ilTotal As Integer
    Dim tlVendor As VendorInfo
    Dim tlMyAvailableVendors() As VendorInfo
    Dim blAdd As Boolean
    Dim ilCurrent As Integer
    Dim blIsJelli As Boolean
    Dim blIsIPump As Boolean
    Dim blIsCompel As Boolean
    '1 = isci 2 = break 3 = both.  0 is issue
    Dim ilXDS As Integer
    'Change this when adding vendor! '8178
    ilTotal = 19
    '9204 block temporarily!
    ilTotal = 17
    If (StrComp(sgUserName, "Guide", 1) = 0) Then
        If Not bgLimitedGuide Then
            ilTotal = 19
        End If
    End If
    ' 'block' for cumulus/cbs not included here, but are below
    mAllowedVendors blIsJelli, blIsCompel, blIsIPump, ilXDS
    ReDim tlMyAvailableVendors(0)
    For ilCount = 1 To ilTotal
        blAdd = True
        With tlVendor
            .sApprovalPassword = ""
            Select Case ilCount
                Case 1
                    .iIdCode = Vendors.NetworkConnect
                    .sName = "Network Connect by Marketron"
                    .sDeliveryType = "L"
                    .iExportMethod = MethodType.csiVendorService
                    .iImportMethod = MethodType.csiVendorService
                    .iHierarchy = 2
                    .sApprovalPassword = "40881aa8-ec88-4c73-8366-c13a2aa591d5"
                    .sSourceName = "NC"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 1
                Case 2
                    .iIdCode = Vendors.WideOrbit
                    .sName = "WO Traffic Radio Interchange"
                    .sDeliveryType = "L"
                    .iExportMethod = MethodType.csiVendorService
                    .iImportMethod = MethodType.BothServices
                    .iHierarchy = 2
                    .sApprovalPassword = "3e47b796-f6ab-4f13-bddc-2968082b94d3"
                    .sSourceName = "WO"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 2
                Case 3
                    .iIdCode = Vendors.cBs
                    .sName = "CBS"
                    .sDeliveryType = "L"
                    .iExportMethod = MethodType.csiStationService
                    .iImportMethod = MethodType.csiStationService
                    .iHierarchy = 2
                    .sApprovalPassword = "60a302a5-147b-4df3-b615-f3ba03230251"
                    .sSourceName = "CB"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 1
                    blAdd = gUsingWeb
                Case 4
                    '9184
                    .iIdCode = Vendors.stratus
                    .sName = "Stratus"
                    .sDeliveryType = "L"
                    .iExportMethod = MethodType.csiStationService
                    .iImportMethod = MethodType.csiStationService
                    .iHierarchy = 2
                    .sApprovalPassword = "1f0f0bbc-5307-4524-ab86-5557bb239a21"
                    .sSourceName = "ST"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 1
                    blAdd = gUsingWeb
                Case 5
                    .iIdCode = Vendors.Jelli
                    .sName = "Jelli"
                    .sDeliveryType = "L"
                    .iExportMethod = MethodType.None
                    .iImportMethod = MethodType.None
                    .iHierarchy = 0
                    .sApprovalPassword = ""
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 1
                    blAdd = blIsJelli
                Case 6
                    .iIdCode = Vendors.iheart
                    .sName = "iHeart"
                    .sDeliveryType = "L"
                    '100125
                    .iExportMethod = MethodType.BothServices
                    .iImportMethod = MethodType.csiStationService
                    .iHierarchy = 2
                    .sApprovalPassword = "ba858df9-2606-4dac-bdc1-61ea21081e4b"
                    .sSourceName = "IH"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 2
                Case 7
                    .iIdCode = Vendors.Wegener_Compel
                    .sName = "Wegener-Compel"
                    .sDeliveryType = "A"
                    .iExportMethod = MethodType.None
                    .iImportMethod = MethodType.None
                    .iHierarchy = 0
                    .sApprovalPassword = ""
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 1
                    blAdd = blIsCompel
                Case 8
                    .iIdCode = Vendors.Wegener_IPump
                    .sName = "Wegener-IPump"
                    .sDeliveryType = "A"
                    .iExportMethod = MethodType.None
                    .iImportMethod = MethodType.None
                    .iHierarchy = 0
                    .sApprovalPassword = ""
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 1
                    blAdd = blIsIPump
                Case 9
                    .iIdCode = Vendors.XDS_Break
                    .sName = "X-Digital-Break"
                    .sDeliveryType = "A"
                    .iExportMethod = MethodType.None
                    '7912
                    .iImportMethod = MethodType.csiVendorService
                    .iHierarchy = 1
                    .sApprovalPassword = "b44125b1-04a8-4045-8463-f1aa6335d739"
                    .sSourceName = "XD"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 1
                    If ilXDS < 2 Then
                        blAdd = False
                    End If
                Case 10
                    .iIdCode = Vendors.XDS_ISCI
                    .sName = "X-Digital-ISCI"
                    .sDeliveryType = "A"
                    .iExportMethod = MethodType.None
                    .iImportMethod = MethodType.csiVendorService
                    .iHierarchy = 1
                    .sApprovalPassword = "bcf5a1fb-7d0a-4554-b2ce-b7ac11f61879"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 1
                    If Not (ilXDS = 1 Or ilXDS = 3) Then
                        blAdd = False
                    End If
                Case 11
                    .iIdCode = Vendors.iDc
                    .sName = "IDC"
                    .sDeliveryType = "A"
                    .iExportMethod = MethodType.None
                    .iImportMethod = MethodType.None
                    .iHierarchy = 1
                    .sApprovalPassword = ""
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 1
                Case 12
                    .iIdCode = Vendors.MrMaster
                    .sName = "Mr. Master"
                    .sDeliveryType = "A"
                    .iExportMethod = MethodType.csiStationService
                    .iImportMethod = MethodType.csiStationService
                    .iHierarchy = 1
                    .sApprovalPassword = "c44125c1-04a8-4045-8463-f1aa6335d739"
                    .sSourceName = "MM"
                    .iMinimumWebVersion = 2
                    .bAutoVendingCSIAllowed = True
                Case 13
                    .iIdCode = Vendors.RCS
                    .sName = "RCS"
                    .sDeliveryType = "L"
                    .iExportMethod = MethodType.csiStationService
                    .iImportMethod = MethodType.csiStationService
                    .iHierarchy = 2
                    .sApprovalPassword = "45c8cdc8-acd7-44a1-ac82-b383d01d3b9a"
                    .sSourceName = "RC"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 2
                Case 14
                    .iIdCode = Vendors.Synchronicity
                    .sName = "Synchronicity"
                    .sDeliveryType = "A"
                    .iExportMethod = MethodType.csiStationService
                    .iImportMethod = MethodType.csiStationService
                    .iHierarchy = 1
                    .sApprovalPassword = "1d3b58b2-b9b7-40bc-98de-0a8ea42a098b"
                    .sSourceName = "SY"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 2
                Case 15
                    .iIdCode = Vendors.RadioTraffic
                    .sName = "RadioTraffic.com"
                    .sDeliveryType = "L"
                    .iExportMethod = MethodType.csiStationService
                    .iImportMethod = MethodType.csiStationService
                    .iHierarchy = 2
                    .sApprovalPassword = "afdddf0b-4a49-46dc-bd73-f17439cc5f8e"
                    .sSourceName = "RT"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 2
                Case 16
                    .iIdCode = Vendors.BSI
                    .sName = "Broadcast Software Intl."
                    .sDeliveryType = "A"
                    .iExportMethod = MethodType.csiStationService
                    .iImportMethod = MethodType.csiStationService
                    .iHierarchy = 1
                    .sApprovalPassword = "eaeb520d-1aa3-42b1-9276-6f26201be03f"
                    .sSourceName = "BI"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 2
                Case 17
                    .iIdCode = Vendors.RadioWorkflow
                    .sName = "Radio Workflow"
                    .sDeliveryType = "L"
                    .iExportMethod = MethodType.csiStationService
                    .iImportMethod = MethodType.csiStationService
                    .iHierarchy = 2
                    .sApprovalPassword = "07927108-4cd5-4fb7-8201-5201f0b654ef"
                    .sSourceName = "RW"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 2
                Case 18
                    .iIdCode = Vendors.AnalyticOwl
                    .sName = "AnalyticOwl"
                    .sDeliveryType = "I"
                    .iExportMethod = MethodType.csiStationService
                    .iImportMethod = MethodType.None
                    .iHierarchy = 0
                    .sApprovalPassword = "4a66b973-d718-4629-bff9-f7a5516ee467"
                    .sSourceName = "AO"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 2
                Case 19
                    .iIdCode = Vendors.LeadsRX
                    .sName = "LeadsRX"
                    .sDeliveryType = "I"
                    .iExportMethod = MethodType.csiStationService
                    .iImportMethod = MethodType.None
                    .iHierarchy = 0
                    .sApprovalPassword = "01fcb6fb-d488-4890-b363-92c832b755e0"
                    .sSourceName = "RX"
                    .bAutoVendingCSIAllowed = False
                    .iMinimumWebVersion = 2
            End Select
        End With
        If blAdd Then
            ilCurrent = UBound(tlMyAvailableVendors)
            tlMyAvailableVendors(ilCurrent) = tlVendor
            ReDim Preserve tlMyAvailableVendors(ilCurrent + 1)
        End If
    Next
    gGetAvailableVendors = tlMyAvailableVendors
End Function
Public Function gGetActiveDeliveryVendors() As VendorInfo()
    
    Dim tlActiveVendor() As VendorInfo
    ReDim tlActiveVendor(50) As VendorInfo
    Dim tlAvailableVendors() As VendorInfo
    Dim ilCount As Integer
    Dim ilCountAvailable As Integer

On Error GoTo ErrHand
    tlAvailableVendors = gGetAvailableVendors()
    SQLQuery = "Select * From wvt_Vendor_Table where wvtDeliveryType = 'L' or wvtdeliverytype = 'A'"
    Set rst = gSQLSelectCall(SQLQuery)
    Do While Not rst.EOF
        For ilCountAvailable = 0 To UBound(tlAvailableVendors) - 1
            If tlAvailableVendors(ilCountAvailable).iIdCode = rst!wvtVendorId Then
                tlActiveVendor(ilCount).sName = Trim$(tlAvailableVendors(ilCountAvailable).sName)
                tlActiveVendor(ilCount).iIdCode = tlAvailableVendors(ilCountAvailable).iIdCode
                tlActiveVendor(ilCount).sDeliveryType = tlAvailableVendors(ilCountAvailable).sDeliveryType
                tlActiveVendor(ilCount).iMinimumWebVersion = tlAvailableVendors(ilCountAvailable).iMinimumWebVersion
                tlActiveVendor(ilCount).iHierarchy = tlAvailableVendors(ilCountAvailable).iHierarchy
                tlActiveVendor(ilCount).bAutoVendingCSIAllowed = tlAvailableVendors(ilCountAvailable).bAutoVendingCSIAllowed
                tlActiveVendor(ilCount).iExportMethod = tlAvailableVendors(ilCountAvailable).iExportMethod
                tlActiveVendor(ilCount).iImportMethod = tlAvailableVendors(ilCountAvailable).iImportMethod
                
                tlActiveVendor(ilCount).sApprovalPassword = rst!wvtApprovalPassword
                tlActiveVendor(ilCount).sVendorPassword = rst!wvtVendorPassword
                tlActiveVendor(ilCount).sVendorUserName = rst!wvtVendorUserName
                tlActiveVendor(ilCount).sStationPassword = rst!wvtStationPassword
                tlActiveVendor(ilCount).sStationUserName = rst!wvtStationUserName
                tlActiveVendor(ilCount).sAddress = rst!wvtAddress
                tlActiveVendor(ilCount).sSendUpdatesOnly = rst!wvtSendUpdatesOnly
                tlActiveVendor(ilCount).sIsOverridable = rst!wvtIsOverridable
                tlActiveVendor(ilCount).sAllowAutoPosting = rst!wvtAllowAutoPost
                tlActiveVendor(ilCount).sAutoVendor = rst!wvtVendorUpdateVAT
                ilCount = ilCount + 1
                Exit For
            End If
        Next ilCountAvailable
        rst.MoveNext
    Loop
    ReDim Preserve tlActiveVendor(ilCount)
    gGetActiveDeliveryVendors = tlActiveVendor
    Erase tlAvailableVendors
    Erase tlActiveVendor
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Agreement-mPopVendorList"
End Function
Public Function gUpdateWebVendorsOnWeb() As Long
    Dim llCount As Long
    Dim llAffected As Long
    Dim myRst As ADODB.Recordset
    Dim slSql As String
    Dim slStr As String
    Dim slComma As String
    Dim slQuotes As String
    Dim tlVendors() As VendorInfo
    Dim c As Integer
    Dim slSource As String
    '9204 note I only send for 'informational'  this is so non-updated web sites don't break.  At some point, I can remove all this and always send. 6/4 I updated
    'Dim slVendorTypeAsNeeded As String
    Dim slVendorType As String
    '9818
    Dim slAdditional As String
    Dim ilISCI As Integer
    Dim ilCue As Integer
    Dim ilAdditionalToUse As Integer
    
    On Error GoTo Err_Handler
    llCount = 0
    slComma = ","
    slQuotes = """"
    slSql = "select * from WVT_Vendor_Table " & mGetWebVendorSqlWhere()
    Set myRst = gSQLSelectCall(slSql)
    If Not myRst.EOF Then
        tlVendors = gGetAvailableVendors()
        Do While Not myRst.EOF
            slSource = ""
            ilAdditionalToUse = 0
            For c = 0 To UBound(tlVendors) - 1
                If tlVendors(c).iIdCode = myRst!wvtVendorId Then
                    slSource = tlVendors(c).sSourceName
                    '9818
                    If tlVendors(c).iIdCode = Vendors.XDS_Break Or tlVendors(c).iIdCode = Vendors.XDS_ISCI Then
                        gGetSharedHeadEnd ilISCI, ilCue
                        If tlVendors(c).iIdCode = Vendors.XDS_Break And ilCue > 0 Then
                            ilAdditionalToUse = ilCue
                        ElseIf ilISCI > 0 Then
                            ilAdditionalToUse = ilISCI
                        End If
                    End If
                    Exit For
                End If
            Next c
            If myRst!wvtDeliveryType = "I" Then
               ' slVendorTypeAsNeeded = ",VendorType ='I'"
               slVendorType = "I"
            Else
                slVendorType = "D"
            End If
            '8457 '8862  '9818
            slSql = "update webvendors set Name = '" & Trim(myRst!wvtName) & "',ApprovalPassword = '" & Trim(myRst!wvtApprovalPassword) & "',VendorUserName = '" & Trim(myRst!wvtVendorUserName) & "' ,VendorPassword = '" & Trim(myRst!wvtVendorPassword)
            slSql = slSql & "',ExportMethod =" & myRst!wvtExportMethod & ",VendorAddress = '" & Trim(myRst!wvtAddress) & "',StationUserName = '" & Trim(myRst!wvtStationUserName) & "',SendUpdatesOnly ='" & Trim(myRst!wvtSendUpdatesOnly)
            slSql = slSql & "',StationPassword = '" & Trim(myRst!wvtStationPassword) & "',ImportMethod =" & myRst!wvtImportMethod & ",isOverridable = '" & Trim(myRst!wvtIsOverridable) & "',Hierarchy =" & myRst!wvthierarchy & ", AllowAutoPost ='" & myRst!wvtAllowAutoPost & "'"
            slSql = slSql & ",SourceName = '" & slSource & "', VendorUpdateStatus = '" & Trim(myRst!wvtVendorUpdateVAT) & "',VendorType = '" & slVendorType & "', additionalInfo = '" & ilAdditionalToUse & "' WHERE idcode =" & myRst!wvtVendorId
            llAffected = gExecWebSQLWithRowsEffected(slSql)
            '9204 I no longer need to split "I" from "D" also 9818
            If llAffected = 0 Then
                slSql = "insert into webvendors (idcode,name,approvalPassword,VendorUserName,VendorPassword,ExportMethod,VendorAddress,StationUserName,SendUpdatesOnly,StationPassword,ImportMethod,isOverridable,Hierarchy,AllowAutoPost,sourceName,VendorUpdateStatus,VendorType,AdditionalInfo) values "
                slSql = slSql & "(" & myRst!wvtVendorId & ",'" & Trim(myRst!wvtName) & "','" & Trim(myRst!wvtApprovalPassword) & "','" & Trim(myRst!wvtVendorUserName) & "','" & Trim(myRst!wvtVendorPassword) & "'," & myRst!wvtExportMethod & ",'" & Trim(myRst!wvtAddress)
                slSql = slSql & "','" & Trim(myRst!wvtStationUserName) & "','" & Trim(myRst!wvtSendUpdatesOnly) & "','" & Trim(myRst!wvtStationPassword) & "'," & myRst!wvtImportMethod & ",'" & Trim(myRst!wvtIsOverridable) & "'," & myRst!wvthierarchy & ",'" & myRst!wvtAllowAutoPost & "','" & slSource & "','" & Trim(myRst!wvtVendorUpdateVAT) & "','" & slVendorType & "','" & ilAdditionalToUse & "')"

'                If myRst!wvtDeliveryType = "I" Then
'                    slSql = "insert into webvendors (idcode,name,approvalPassword,VendorUserName,VendorPassword,ExportMethod,VendorAddress,StationUserName,SendUpdatesOnly,StationPassword,ImportMethod,isOverridable,Hierarchy,AllowAutoPost,sourceName,VendorUpdateStatus,VendorType) values "
'                    slSql = slSql & "(" & myRst!wvtVendorId & ",'" & Trim(myRst!wvtName) & "','" & Trim(myRst!wvtApprovalPassword) & "','" & Trim(myRst!wvtVendorUserName) & "','" & Trim(myRst!wvtVendorPassword) & "'," & myRst!wvtExportMethod & ",'" & Trim(myRst!wvtAddress)
'                    slSql = slSql & "','" & Trim(myRst!wvtStationUserName) & "','" & Trim(myRst!wvtSendUpdatesOnly) & "','" & Trim(myRst!wvtStationPassword) & "'," & myRst!wvtImportMethod & ",'" & Trim(myRst!wvtIsOverridable) & "'," & myRst!wvthierarchy & ",'" & myRst!wvtAllowAutoPost & "','" & slSource & "','" & Trim(myRst!wvtVendorUpdateVAT) & "','I')"
'                Else
'                    slSql = "insert into webvendors (idcode,name,approvalPassword,VendorUserName,VendorPassword,ExportMethod,VendorAddress,StationUserName,SendUpdatesOnly,StationPassword,ImportMethod,isOverridable,Hierarchy,AllowAutoPost,sourceName,VendorUpdateStatus,AdditionalInfo) values "
'                    slSql = slSql & "(" & myRst!wvtVendorId & ",'" & Trim(myRst!wvtName) & "','" & Trim(myRst!wvtApprovalPassword) & "','" & Trim(myRst!wvtVendorUserName) & "','" & Trim(myRst!wvtVendorPassword) & "'," & myRst!wvtExportMethod & ",'" & Trim(myRst!wvtAddress)
'                    slSql = slSql & "','" & Trim(myRst!wvtStationUserName) & "','" & Trim(myRst!wvtSendUpdatesOnly) & "','" & Trim(myRst!wvtStationPassword) & "'," & myRst!wvtImportMethod & ",'" & Trim(myRst!wvtIsOverridable) & "'," & myRst!wvthierarchy & ",'" & myRst!wvtAllowAutoPost & "','" & slSource & "','" & Trim(myRst!wvtVendorUpdateVAT) & "'," & ilAdditionalToUse & ")"
'                End If
                llAffected = gExecWebSQLWithRowsEffected(slSql)
                If llAffected = -1 Then
                    llCount = -1
                End If
            ElseIf llAffected = -1 Then
                llCount = -1
            End If
            If llCount > -1 Then
                llCount = llCount + 1
            End If
            myRst.MoveNext
        Loop
    End If
Cleanup:
    If Not myRst Is Nothing Then
        If (myRst.State And adStateOpen) <> 0 Then
            myRst.Close
        End If
        Set myRst = Nothing
    End If
    If llCount = -1 Then
        gLogMsg "Error writing web vendor information in -gUpdateWebVendorsOnWeb", "AffErrorLog.txt", False
    End If
    gUpdateWebVendorsOnWeb = llCount
    Exit Function
Err_Handler:
    gHandleError "", "modVendors-gUpdateWebVendorsOnWeb"
    llCount = -1
    GoTo Cleanup
End Function
Public Function gUpdateWebVendorsHeaderOnWeb(flWebExport As Form) As Long
    Dim llCount As Long
    Dim myRst As ADODB.Recordset
    Dim rstList As ADODB.Recordset
    Dim slSql As String
    Dim slSlave As String
    Dim slAdditional As String
    '7942
    Dim slAdditional2 As String
    Dim slAdditional3 As String
    '9114
    Dim slAdditional4 As String
    Dim ilShtt As Integer
    Dim blPopulated As Boolean
    Dim ilZones() As Integer
    '8344
    Dim blContinue As Boolean
    '9256
    Dim blNeedToTestIfISCIORCue As Boolean
    Dim slISCISiteId As String
    Dim slCueSiteId As String
    '9915
    Dim blUseProgressBar As Boolean
    Dim ilPercent As Integer
    Dim llTotalNoRecs As Long
    '100125
    Dim slAdditional5 As String
    
    On Error GoTo Err_Handler
    'to be used in gUpdateVat
    ReDim sgVatUpdate(0)
    llCount = 0
    blPopulated = False
    '9915
    blUseProgressBar = False
    slSql = "select count (distinct vatAttCode) as amount from VAT_Vendor_Agreement, WVT_Vendor_Table " & mGetWebVendorSqlWhere() & " AND (wvtVendorID = vatWvtVendorID) AND vatSentToWeb <> 'Y'"
    Set rstList = gSQLSelectCall(slSql)
    If Not rstList.EOF Then
        llTotalNoRecs = rstList!amount
        If llTotalNoRecs > 100 Then
            blUseProgressBar = True
            flWebExport.mStartVendorProgress
        End If
    End If
'    Dim ilIndex As Integer
'    For ilIndex = 0 To 500
'        Sleep 100
'        DoEvents
'        frmWebExportSchdSpot.mUpdateVendorProgress ilIndex
'    Next ilIndex
'    frmWebExportSchdSpot.mStopVendorProgress
    '  where wvtApprovalPassword <> '' AND ((WvtExportmethod = 1 AND WvtStationuserName <> '' AND WvtStationPassword <> '') OR (wvtExportMethod = 2 AND  WvtVendorUserName <>'' AND WvtVendorPassword <>'' AND WvtAddress <> ''))
    slSql = "select distinct vatAttCode from VAT_Vendor_Agreement, WVT_Vendor_Table " & mGetWebVendorSqlWhere() & " AND (wvtVendorID = vatWvtVendorID) AND vatSentToWeb <> 'Y'"
    Set rstList = gSQLSelectCall(slSql)
    If Not rstList.EOF Then
        '9256
        blNeedToTestIfISCIORCue = mIsDualXDSProviderAndReceiverIdSources(slISCISiteId, slCueSiteId)
        Do While Not rstList.EOF
            '9915
            llCount = llCount + 1
            If blUseProgressBar Then
            On Error GoTo ERRLONG
                ilPercent = (llCount * CSng(100)) / llTotalNoRecs
                If ilPercent > 100 Then
                    ilPercent = 100
                End If
                flWebExport.mUpdateVendorProgress ilPercent
            On Error GoTo Err_Handler
                DoEvents
            End If
            'delete from web by attcode!
            slSql = "delete from webvendors_header where attcode = " & rstList!vatattcode
            If gExecWebSQLWithRowsEffected(slSql) <> -1 Then
                slSql = "Select vatwvtvendorid from vat_vendor_agreement,wvt_Vendor_Table " & mGetWebVendorSqlWhere() & " AND (wvtVendorID = vatWvtVendorID) and vatattcode = " & rstList!vatattcode
                Set myRst = gSQLSelectCall(slSql)
                If Not myRst.EOF Then
                    If gIsMulticastDontSend(rstList!vatattcode) Then
                        slSlave = "Y"
                    Else
                        slSlave = ""
                    End If
                    Do While Not myRst.EOF
                        blContinue = True
                        'send import info 7912
                        slAdditional = ""
                        slAdditional2 = ""
                        slAdditional3 = ""
                        slAdditional4 = ""
                        '100125
                        slAdditional5 = ""
                        Select Case myRst!vatwvtvendorid
                            Case Vendors.XDS_Break, Vendors.XDS_ISCI
                                If Not blPopulated Then
                                    ilZones = mHeadEndZoneAdjust()
                                    gPopShttInfo
                                    blPopulated = True
                                End If
                                slAdditional = mGetXDSiteId(rstList!vatattcode, ilShtt, blNeedToTestIfISCIORCue, slISCISiteId, slCueSiteId)
                                '+3 would be pacific 0 eastern
                                slAdditional2 = mStationAdjust(ilShtt, ilZones)
                                '9114
                                slAdditional3 = mGetVehicleCode5AndMergeInfo(rstList!vatattcode, slAdditional4)
                            Case Else
                                If Not mGetVehicleAndShttCodeAndNotServiceAgreement(rstList!vatattcode, slAdditional, slAdditional5) Then
                                    blContinue = False
                                End If
                                '10197
                                If blContinue And Len(slAdditional5) > 0 Then
                                    slAdditional2 = mGetHonorDaylight(slAdditional5)
                                End If
'                            Case Vendors.NetworkConnect
'                                '8344
'                                 If Not mGetVehicleCodeAndNotServiceAgreement(rstList!vatattcode, slAdditional) Then
'                                    blContinue = False
'                                 End If
'                            '8344
'                            Case Vendors.WideOrbit
'                                 If mGetVehicleCodeAndNotServiceAgreement(rstList!vatattcode, slAdditional) Then
'                                    'we don't want vefcode, just to know if service agreement
'                                    slAdditional = ""
'                                Else
'                                    blContinue = False
'                                 End If
                        End Select
                        'not a service agreement? continue
                        If blContinue Then
                            blContinue = False
                        'add all for that att!
                            slSql = "insert into webvendors_Header (attcode,vendorid,multicastdontsend,VendorSpecificInfo,VendorSpecificInfo2,VendorSpecificInfo3,VendorSpecificInfo4,vendorspecificinfo5) values (" & rstList!vatattcode & "," & myRst!vatwvtvendorid & ",'" & slSlave & "','" & slAdditional & "','" & slAdditional2 & "','" & slAdditional3 & "','" & slAdditional4 & "','" & slAdditional5 & "')"
                            If gExecWebSQLWithRowsEffected(slSql) <> -1 Then
                                blContinue = True
                            End If
                        'we'll update below even if we didn't send to web because service agreement!
                        Else
                            blContinue = True
                        End If
                        'even if didn't update
                        If blContinue Then
                            slSql = "update VAT_Vendor_Agreement set vatSentToWeb = 'Y' where vatAttCode  = " & rstList!vatattcode & "  AND vatwvtvendorId = " & myRst!vatwvtvendorid
                            If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                                '6/13/16: Replaced GoSub
                                'GoSub Err_Sql:
                                gHandleError "AffErrorLog.txt", "modVendors-gUpdateWebVendorsHeaderOnWeb"
                                gUpdateWebVendorsHeaderOnWeb = llCount
                                Exit Function
                            End If
                        End If
                        myRst.MoveNext
                    Loop
                End If
            End If 'new att set to go to web
            rstList.MoveNext
        Loop 'each att
    End If
Cleanup:
    '9915
    If blUseProgressBar Then
        flWebExport.mStopVendorProgress
        DoEvents
    End If
    If Not myRst Is Nothing Then
        If (myRst.State And adStateOpen) <> 0 Then
            myRst.Close
        End If
        Set myRst = Nothing
    End If
    If Not rstList Is Nothing Then
        If (rstList.State And adStateOpen) <> 0 Then
            rstList.Close
        End If
        Set rstList = Nothing
    End If
    Erase ilZones
    gUpdateWebVendorsHeaderOnWeb = llCount
    Exit Function
ERRLONG:
    Resume Next
Err_Handler:
    gHandleError "", "modVendors-gUpdateWebVendorsHeaderOnWeb"
    llCount = 0
    GoTo Cleanup
End Function
Private Function mGetWebVendorSqlWhere() As String
    'So getting the web vendors remains consistent
    'test appropriate username pasword (and address if needed)
    'if using for vat, add wvt to From and add AND (wvtVendorID = vatWvtVendorID)
    Dim slSql As String
    
    slSql = " where wvtApprovalPassword <> '' AND (((WvtExportmethod = 1 AND WvtStationuserName <> '' AND WvtStationPassword <> '') OR (wvtExportMethod = 2 AND WvtVendorUserName <>'' AND WvtVendorPassword <>'' AND WvtAddress <> '') OR (wvtExportMethod = 3 And WvtStationuserName <> '' AND WvtStationPassword <> '' AND  WvtVendorUserName <>'' AND WvtVendorPassword <>'' AND WvtAddress <> '')) "
    slSql = slSql & " OR  ((WvtImportmethod = 1 AND WvtStationuserName <> '' AND WvtStationPassword <> '') OR (wvtImportMethod = 2 AND  WvtVendorUserName <>'' AND WvtVendorPassword <>'' AND WvtAddress <> '') OR (wvtImportMethod = 3 And WvtStationuserName <> '' AND WvtStationPassword <> '' AND  WvtVendorUserName <>'' AND WvtVendorPassword <>'' AND WvtAddress <> '')))  "
    
    mGetWebVendorSqlWhere = slSql
End Function
Public Sub gSetByItemData(myCbo As ComboBox, llItem As Long)
    Dim ilLoop As Integer
    With myCbo
        For ilLoop = 0 To .ListCount - 1
            If .ItemData(ilLoop) = llItem Then
                .ListIndex = ilLoop
                Exit For
            End If
        Next
    End With
End Sub


Public Function gLoseLastLetterIfComma(slInput As String) As String
    Dim llLength As Long
    Dim slNewString As String
    Dim llLastLetter As Long
    
    llLength = Len(slInput)
    llLastLetter = InStrRev(slInput, ",")
    If llLength > 0 And llLastLetter = llLength Then
        slNewString = Mid(slInput, 1, llLength - 1)
    Else
        slNewString = slInput
    End If
    gLoseLastLetterIfComma = slNewString
End Function
Public Function gSQLAndReturn(sSQLQuery As String, iDoTrans As Integer, llAffectedRecords As Long) As Long
    'Dan m 9/19/09 changed llRet and function return to Long
    Dim llRet As Long
    Dim fStart As Single
    Dim iCount As Integer
    Dim hlMsg As Integer
    Dim ilRet As Integer
    On Error GoTo ErrHand

    '12/4/12: Check if activity should be logged
    'mLogActivityFileName sSQLQuery
    '12/4/12: end of change
    llAffectedRecords = 0
    iCount = 0
    Do
        llRet = 0
        If iDoTrans Then
            'cnn.BeginTrans
        End If
        'cnn.Execute sSQLQuery, rdExecDirect
        cnn.Execute sSQLQuery, llAffectedRecords
        If llRet = 0 Then
            If iDoTrans Then
                'cnn.CommitTrans
            End If
        ElseIf (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT) Then
            fStart = Timer
            Do While Timer <= fStart
                llRet = llRet
            Loop
            iCount = iCount + 1
            If iCount > igWaitCount Then
                'gMsgBox "A SQL error has occurred: " & "Error # " & llRet, vbCritical
                Exit Do
            End If
        End If
    Loop While (llRet = BTRV_ERR_REC_LOCKED) Or (llRet = BTRV_ERR_FILE_LOCKED) Or (llRet = BTRV_ERR_INCOM_LOCK) Or (llRet = BTRV_ERR_CONFLICT)
    gSQLAndReturn = llRet
    If llRet <> 0 Then
        If (bgIgnoreDuplicateError) And ((llRet = -4994) Or (llRet = BTRV_ERR_DUPLICATE_KEY)) Then
        Else
            On Error GoTo mOpenFileErr:
            'hlMsg = FreeFile
            'Open sgMsgDirectory & "AffErrorLog.txt" For Append As hlMsg
            ilRet = gFileOpen(sgMsgDirectory & "AffErrorLog.txt", "Append", hlMsg)
            If ilRet = 0 Then
                Print #hlMsg, sSQLQuery
                Print #hlMsg, "Error # " & llRet
            End If
            Close #hlMsg
        End If
    End If
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "modVendors-gSQLAndReturn"
mOpenFileErr:
    Resume Next
End Function
Public Function gIfNullInteger(ilValue As Variant) As Integer
    If IsNull(ilValue) Then
        gIfNullInteger = 0
    Else
        gIfNullInteger = ilValue
    End If
End Function
Private Sub mAllowedVendors(blIsJelli As Boolean, blIsCompel As Boolean, blIsPump As Boolean, ilXDS As Integer)
    Dim ilRet As Integer
    Dim slSQLQuery As String
    Dim sql_rst As ADODB.Recordset
    
    ilRet = 0
    ilXDS = 0
    blIsJelli = False
    blIsCompel = False
    blIsPump = False
On Error GoTo ErrHand
    slSQLQuery = "Select safFeatures1 From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set sql_rst = gSQLSelectCall(slSQLQuery)
    If Not sql_rst.EOF Then
        ilRet = Asc(sql_rst!safFeatures1)
        If (ilRet And JELLIEXPORT) = JELLIEXPORT Then
            blIsJelli = True
        End If
    End If
    slSQLQuery = "Select spfUsingFeatures10,spfUsingFeatures7 From SPF_Site_Options"
    Set sql_rst = gSQLSelectCall(slSQLQuery)
    If Not sql_rst.EOF Then
        ilRet = Asc(sql_rst!spfUsingFeatures10)
        If (ilRet And WEGENERIPUMP) = WEGENERIPUMP Then
            blIsPump = True
        End If
        ilRet = Asc(sql_rst!spfUsingFeatures7)
        If (ilRet And WEGENEREXPORT) = WEGENEREXPORT Then
            blIsCompel = True
        End If
    End If
    '1 = isci 2 = break 3 = both.  0 is issue
    ilXDS = gSiteISCIAndOrBreak()
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "modVendors-mAllowedVendors"
End Sub
Public Function gIsMulticastDontSend(llAtt As Long)
    'returns true if agreement is using a vendor that uses master/slave and this agreement is set as slave
    Dim blRet As Boolean
    Dim myRs As ADODB.Recordset
    Dim slSql As String
    Dim blNotSet As Boolean
    
    blRet = False
    '7701DanChange  '7957 7952 added xds/wo
    slSql = "select shttclustergroupId,shttMasterCluster,attMulticast from shtt inner join att on shttcode = attshfcode inner join VAT_Vendor_Agreement on vatattcode = attcode where vatwvtVendorId in (" & Vendors.NetworkConnect & "," & Vendors.WideOrbit & "," & Vendors.XDS_Break & ") AND attcode = " & llAtt
    Set myRs = gSQLSelectCall(slSql)
    If Not myRs.EOF Then
        blRet = gSlave(myRs!shttclustergroupId, myRs!shttMasterCluster, myRs!attMulticast, blNotSet, "")
        '10024 if not set, send anyway
        If blNotSet Then
            blRet = False
        End If
    End If
Cleanup:
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
        End If
        Set myRs = Nothing
    End If
    gIsMulticastDontSend = blRet
    Exit Function
ErrHandler:
    gHandleError "", "modVendors-gIsMulticastDontSend"
    blRet = False
    GoTo Cleanup
End Function
Public Function gSlave(llCluster As Long, slMaster As String, slIsMulticast As String, blNoSisterWarn As Boolean, slPathForgLogMsg As String) As Boolean
' it's a slave IF shttClusterGroupId > 0 AND there is a shttMasterCluster = 'Y' assigned for that same groupId
'3/28/12  ttp 5263 add slIsMulticast and blNoSisterWarn.  blNoSisterWarn is outgoing: if multi true but sister never set, warn user!
   '10024 for Marketron, set to 'true' if multi and no sister set up.  But if going to web, we want to send.  So test blNoSisterWarn
    Dim myRs As ADODB.Recordset
    Dim Sql As String
    'ttp 5263. not multicast, can't be slave.
    If slIsMulticast <> "Y" Then
        gSlave = False
        Exit Function
    End If
    'ttp 5263, now changes the fix for 5261... llCluster should not be <= 0.  that means multicast but no master/slave relationship.
    'warn user
    If llCluster <= 0 Then
        gSlave = True
        blNoSisterWarn = True
        Exit Function
    End If
'    'ttp 5261 sister station issue.
'    If llCluster <= 0 Then
'        gSlave = False
'        Exit Function
'    End If
    Select Case slMaster
        Case "y", "Y"
            gSlave = False
        'this assumes if it's marked "n" then it has a clusterGroupId
        Case "n", "N"
            gSlave = True
        Case Else
On Error GoTo ErrHandler
            If llCluster > 0 Then
                If igExportSource = 2 Then DoEvents
                Sql = "select count(*) as Amount from shtt where shttmasterCluster = 'Y' and shttclustergroupid = " & llCluster
                Set myRs = gSQLSelectCall(Sql)
                If myRs!amount > 0 Then
                    gSlave = True
                Else
                    gSlave = False
                End If
            Else
                gSlave = False
            End If
    End Select
Cleanup:
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
        End If
        'dan 5/21/2012
        Set myRs = Nothing
    End If
    Exit Function
ErrHandler:
    gHandleError slPathForgLogMsg, "modVendors-gSlave"
    gSlave = True
    GoTo Cleanup

End Function
Public Function gConvertToVat(ilVendor As Integer) As Long
    'retun -1 on error; otherwise # changed
    Dim blRet As Boolean
    Dim rsConvert As ADODB.Recordset
    Dim slSql As String
    Dim blConvert As Boolean
    Dim llCount As Long
    Dim llSubCount As Long
    
    blConvert = True
    llSubCount = 0
    Select Case ilVendor
        Case Vendors.NetworkConnect
            slSql = "select attcode from att where attExportToMarketron = 'Y'"
        Case Vendors.Jelli
            slSql = "select attcode from att where attExportToJelli = 'Y'"
        Case Vendors.iheart
            slSql = "select attcode from att where attExportToClearCh = 'Y'"
        Case Vendors.cBs
            slSql = "select attcode from att where attExportToCBS = 'Y'"
        Case Vendors.stratus
            slSql = "select attcode from att where attWebInterface = 'C'"
        Case Vendors.iDc
            slSql = "select attcode from att where attAudioDelivery = 'I'"
        Case Vendors.Wegener_Compel
            slSql = "select attcode from att where attAudioDelivery = 'W'"
        Case Vendors.Wegener_IPump
            slSql = "select attcode from att where attAudioDelivery = 'P'"
        Case Vendors.XDS_ISCI
            slSql = "select attcode from att where attAudioDelivery = 'X'"
        Case Vendors.XDS_Break
            slSql = "select attcode from att where attAudioDelivery = 'B'"
        Case Vendors.WideOrbit
            blRet = True
            llCount = 0
            GoTo Cleanup
        '8178
        Case Vendors.MrMaster
            blRet = True
            llCount = 0
            GoTo Cleanup
        Case Vendors.RadioTraffic
            blRet = True
            llCount = 0
            GoTo Cleanup
        Case Vendors.RCS
            blRet = True
            llCount = 0
            GoTo Cleanup
        Case Vendors.Synchronicity
            blRet = True
            llCount = 0
            GoTo Cleanup
        Case Vendors.BSI
            blRet = True
            llCount = 0
            GoTo Cleanup
        Case Vendors.AnalyticOwl
            blRet = True
            llCount = 0
            GoTo Cleanup
        Case Vendors.LeadsRX
            blRet = True
            llCount = 0
            GoTo Cleanup
        Case Vendors.RadioWorkflow
            blRet = True
            llCount = 0
            GoTo Cleanup
       Case Else
            blConvert = False
    End Select
    blRet = blConvert
    If blConvert Then
        Set rsConvert = gSQLSelectCall(slSql)
        Do While Not rsConvert.EOF
             slSql = "insert into VAT_Vendor_Agreement (vatattcode,vatwvtVendorID) Values ( " & rsConvert!attCode & "," & ilVendor & ")"
             If gSQLAndReturn(slSql, False, llSubCount) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub Err_Sql:
                gHandleError "AffErrorLog.txt", "modVendors-gConvertToVat"
                gConvertToVat = -1
                If Not rsConvert Is Nothing Then
                    If (rsConvert.State And adStateOpen) <> 0 Then
                        rsConvert.Close
                    End If
                End If
                Exit Function
            End If
            llCount = llCount + llSubCount
            rsConvert.MoveNext
        Loop
    End If
Cleanup:
    If Not blRet Then
        llCount = -1
    End If
    gConvertToVat = llCount
    If Not rsConvert Is Nothing Then
        If (rsConvert.State And adStateOpen) <> 0 Then
            rsConvert.Close
        End If
    End If
    Exit Function
ErrHand:
    gHandleError "", "modVendors-gConvertToVat"
    blRet = False
    GoTo Cleanup
End Function
Private Function mGetXDSiteId(llAtt As Long, ilShtt As Integer, Optional blNeedToTestIfISCIORCue As Boolean = False, Optional slISCISiteId As String = "", Optional slCueSiteId As String = "") As String
    '7942 added shtt and vef codes to go out '9256 may have to test which site id to get if using dual provider
    Dim llID As Long
    Dim myRst As ADODB.Recordset
    Dim slSql As String
    Dim slSiteChoiceForThisVehicle As String
    '9256
    Dim llShttSiteId As Long
    Dim llAttSiteId As Long
    Dim ilVefCode As Integer
    
    slSiteChoiceForThisVehicle = "B"
    slSql = "Select attXdReceiverId, shttStationId,attshfCode,attVefCode from att inner join shtt on attshfcode = shttcode where attcode = " & llAtt
    Set myRst = gSQLSelectCall(slSql)
    If Not myRst.EOF Then
        ilShtt = myRst!attshfcode
        llShttSiteId = myRst!shttStationId
        llAttSiteId = myRst!attXDReceiverId
        ilVefCode = myRst!attvefCode
        If blNeedToTestIfISCIORCue Then
            'break first          'if isci, use sliscisiteid to set slsitechoiceforthisvehicle
            slSql = "Select vffXDProgCodeId as ProgramCode from VFF_Vehicle_Features where VffvefCode = " & ilVefCode
            Set myRst = gSQLSelectCall(slSql)
            If Not myRst.EOF Then
                If Len(Trim$(myRst!ProgramCode)) > 0 Then
                    slSiteChoiceForThisVehicle = slCueSiteId
                Else
                    slSiteChoiceForThisVehicle = slISCISiteId
                End If
            Else
                slSiteChoiceForThisVehicle = slISCISiteId
            End If
        End If
        If slSiteChoiceForThisVehicle = "S" Then
            llID = llShttSiteId
        'I guess technically, if it's "A" we shouldn't get station...but that seems wrong too.
        Else
            llID = llAttSiteId
            If llID = 0 Then
                llID = llShttSiteId
            End If
        End If
    End If
    'saw this while woring on 9114
    mGetXDSiteId = Trim$(Str(llID))
End Function
'Private Function mGetVehicleCode(llAtt As Long) As String
'    Dim ilVefCode As Integer
'    Dim myRst As ADODB.Recordset
'    Dim slSql As String
'
'    slSql = "Select attvefCode from att where attcode = " & llAtt
'    Set myRst = cnn.Execute(slSql)
'    If Not myRst.EOF Then
'        ilVefCode = myRst!attvefCode
'    End If
'    mGetVehicleCode = Str(ilVefCode)
'End Function
Private Function mGetVehicleCodeAndNotServiceAgreement(llAtt As Long, slVefCodeReturn As String) As Boolean
    Dim ilVefCode As Integer
    Dim myRst As ADODB.Recordset
    Dim slSql As String
    Dim blRet As Boolean
    
    blRet = True
    slSql = "Select attvefCode,attServiceAgreement from att where attcode = " & llAtt
    Set myRst = gSQLSelectCall(slSql)
    If Not myRst.EOF Then
        ilVefCode = myRst!attvefCode
        If myRst!attServiceAgreement = "Y" Then
            blRet = False
        End If
    End If
    slVefCodeReturn = Str(ilVefCode)
    'saw this while doing 9114
    slVefCodeReturn = Trim$(slVefCodeReturn)
    mGetVehicleCodeAndNotServiceAgreement = blRet
End Function
Private Function mGetVehicleCode5AndMergeInfo(llAtt As Long, slMergeInfo As String) As String
    Dim ilVefCode As Integer
    Dim myRst As ADODB.Recordset
    Dim slSql As String
    Dim slVefCodeReturn As String
    Dim ilVpf As Integer
    Dim slProgCode As String
    
    slMergeInfo = ""
    slSql = "Select  attvefcode,vffXDProgCodeId,vpfUsingFeatures2 as HonorMerge from ATT left outer join VFF_Vehicle_Features on attvefcode = Vffvefcode left outer join VPF_Vehicle_Options on attvefcode = vpfVefKCode  where attcode = " & llAtt
    Set myRst = gSQLSelectCall(slSql)
    If Not myRst.EOF Then
        ilVefCode = myRst!attvefCode
        If Not IsNull(myRst!vffxdprogcodeid) Then
            slProgCode = myRst!vffxdprogcodeid
            If UCase(Trim$(slProgCode)) = "MERGE" Then
                slMergeInfo = "MERGEE"
            End If
        End If
        If Len(slMergeInfo) = 0 And ilVefCode > 0 Then
            ilVpf = gBinarySearchVpf(CLng(ilVefCode))
            If ilVpf <> -1 Then
                If (Asc(tgVpfOptions(ilVpf).sUsingFeatures2) And XDSAPPLYMERGE) = XDSAPPLYMERGE Then
                    slMergeInfo = "MERGER"
                End If
            End If
        End If
    End If
    slVefCodeReturn = Str(ilVefCode)
    slVefCodeReturn = Trim$(slVefCodeReturn)
    Do While Len(slVefCodeReturn) < 5
        slVefCodeReturn = "0" & slVefCodeReturn
    Loop
    mGetVehicleCode5AndMergeInfo = slVefCodeReturn
End Function
Private Function mStationAdjust(ilShttCode As Integer, imHDAdj() As Integer) As Integer
    
    Dim llShttRet As Long
    Dim slZone As String
On Error GoTo ERRORCODE
    llShttRet = gBinarySearchShtt(ilShttCode)
    If llShttRet <> -1 Then
        slZone = Trim$(tgShttInfo1(llShttRet).shttTimeZone)
    Else
        slZone = ""
    End If
    mStationAdjust = 0
    Select Case Left(slZone, 1)
        Case "E"
            mStationAdjust = imHDAdj(0)
        Case "C"
            mStationAdjust = imHDAdj(1)
        Case "M"
            mStationAdjust = imHDAdj(2)
        Case "P"
            mStationAdjust = imHDAdj(3)
    End Select
    Exit Function
ERRORCODE:
     mStationAdjust = 0
    gHandleError "AffErrorLog.txt", "ModVendors-mStationAdjust"

End Function
Private Function mHeadEndZoneAdjust() As Integer()
    Dim slSQLQuery As String
    Dim saf_rst As ADODB.Recordset
    Dim slHeadEndZone As String
    Dim ilHDAdj(0 To 3) As Integer
    
    ilHDAdj(0) = 0
    ilHDAdj(1) = 0
    ilHDAdj(2) = 0
    ilHDAdj(3) = 0
    slHeadEndZone = ""
    slSQLQuery = "Select safXDSHeadEndZone From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set saf_rst = gSQLSelectCall(slSQLQuery)
    If Not saf_rst.EOF Then
        slHeadEndZone = saf_rst!safXDSHeadEndZone
    End If
    If (slHeadEndZone = "") Then
        slHeadEndZone = "E"
    End If
    'Adjust time as times are relative to location of head end
    If slHeadEndZone = "E" Then
        ilHDAdj(0) = 0
        ilHDAdj(1) = -1
        ilHDAdj(2) = -2
        ilHDAdj(3) = -3
    ElseIf slHeadEndZone = "C" Then
        ilHDAdj(0) = 1
        ilHDAdj(1) = 0
        ilHDAdj(2) = -1
        ilHDAdj(3) = -2
    ElseIf slHeadEndZone = "M" Then
        ilHDAdj(0) = 2
        ilHDAdj(1) = 1
        ilHDAdj(2) = 0
        ilHDAdj(3) = -1
    ElseIf slHeadEndZone = "P" Then
        ilHDAdj(0) = 3
        ilHDAdj(1) = 2
        ilHDAdj(2) = 1
        ilHDAdj(3) = 0
    End If
    mHeadEndZoneAdjust = ilHDAdj
End Function
Public Sub gVatSetToGoToWebByShttCode(ilShttCode As Integer, ilVendor As Integer)
    '8824 vendor of 0 means update for all vendors
    Dim rst_pw As ADODB.Recordset
    Dim slTemp As String
    Dim slSQLQuery As String
    On Error GoTo ErrHand
    If ilShttCode = 0 Then
        Exit Sub
    End If
    slSQLQuery = "SELECT attCode From att where attshfcode = " & ilShttCode
    Set rst_pw = gSQLSelectCall(slSQLQuery)
    While Not rst_pw.EOF
        '10/3/18: Dan- I changed this code: Removed vendor ID in the where cause. Dan put back in 10/10, but added test of 0
        'TTP 8824 reopened
        'slSQLQuery = "UPDATE VAT_Vendor_Agreement SET vatSentToWeb = '' WHERE vatWvtVendorId = " & ilVendor & " AND vatattcode = " & rst_pw!attCode
       ' slSQLQuery = "UPDATE VAT_Vendor_Agreement SET vatSentToWeb = '' WHERE vatattcode = " & rst_pw!attCode
        If ilVendor > 0 Then
            slSQLQuery = "UPDATE VAT_Vendor_Agreement SET vatSentToWeb = '' WHERE vatWvtVendorId = " & ilVendor & " AND vatattcode = " & rst_pw!attCode
        Else
            slSQLQuery = "UPDATE VAT_Vendor_Agreement SET vatSentToWeb = '' WHERE vatattcode = " & rst_pw!attCode
        End If
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub ErrHand:
            gHandleError "AffErrorLog.txt", "modVendors-gVatSetToGoToWebByShttCode"
            Exit Sub
        End If
        rst_pw.MoveNext
    Wend
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "ModVendors-gVatSetToGoToWebByShttCode"
End Sub
Private Function mAlertsForWebVendors(blExport As Boolean, myServiceFacts() As ServiceControllerInfo) As VendorWvmAlert
    'return none if no issue, exportmissed,exporttest, connectionIssue, or exportrunning-which now means there was an error in the code!
    'gExecWebSqlForVendor returns 0 if 'demo' mode.
    Dim ilRet As VendorWvmAlert
    Dim slDateRealMonday As String
    Dim slDateMonday As String
    Dim slDateShouldImport As String
    Dim slSql As String
    Dim slData() As String
    Dim c As Long
    Dim llCount As Long
    Dim slCode As String
    Dim myRst As ADODB.Recordset
    Dim blContinue As Boolean
    '8133
    Dim slWVECode As String
    Dim slControllerLastImported As String
On Error GoTo errbox
    
    ilRet = VendorWvmAlert.None
    blContinue = True
    If blExport Then
        slDateMonday = dgWvImportLast
        slDateMonday = gObtainPrevMonday(slDateMonday)
        slDateMonday = Format(slDateMonday, sgSQLDateForm) & " " & Format(slDateMonday, sgSQLTimeForm)
        slDateRealMonday = gObtainPrevMonday(gNow())
        slDateRealMonday = Format(slDateRealMonday, sgSQLDateForm) & " " & Format(slDateRealMonday, sgSQLTimeForm)
        For c = 0 To 2
           If myServiceFacts(c).Mode = "E" Then
               If myServiceFacts(c).IsRunning Then
                   ilRet = VendorWvmAlert.ExportRunning
                   blContinue = False
                   Exit For
                ElseIf myServiceFacts(c).GenerateFile Then
                    ilRet = VendorWvmAlert.ExportTest
                    blContinue = False
               End If
               Exit For
           End If
        Next c
        If blContinue Then
                'didn't go out? Note I don't check senttoaff..there shoulnd't be any = 0
            slSql = "select count(*) as amount from webvendorexport where hasbeensent = 0 and exportMondayDate >='" & slDateRealMonday & "'"
            llCount = mExecWebSQLCount(slSql)
            If llCount > 0 Then
               ilRet = VendorWvmAlert.ExportMissed
            ElseIf llCount < 0 Then
                ilRet = VendorWvmAlert.ConnectionIssue
            Else
                'errors shown in export queue? First let's see if we can clear old
                slSql = "select count(*) as amount from webvendorexport where (hasbeensent =6 or hasbeensent < -1) and exportMondayDate >='" & slDateRealMonday & "'  AND sentToAff = 'Y' "
                llCount = mExecWebSQLCount(slSql)
                If llCount > 0 Then
                    'We have errors up on web from before. we haven't taken care of old errors
                    ilRet = VendorWvmAlert.NonMonitorIssue
                    slSql = "select count(*) as Amount FROM AUF_ALERT_USER WHERE aufType = 'V'  AND aufStatus = 'R' AND aufSubType = 'E' AND aufchfcode <> 0 and aufMoWeekDate >='" & slDateMonday & "' "
                    Set myRst = gSQLSelectCall(slSql)
                    'all the same, which would be normal (haven't done export yet) doesn't need to go through each one
                    If myRst!amount <> llCount Then
                        mAlertsClearWebVendors True
                        slSql = "Select code,attcode,vendoridcode,hasbeensent,errormessage,exportMondayDate,ProcessedDateTime,EnteredDateTime,senttoaff from webvendorexport WHERE (hasbeensent < -1 or hasbeensent = 6) AND sentToAff = 'Y' "
                        slSql = slSql & " AND exportMondayDate >= '" & slDateMonday & "'"
                        llCount = gExecWebSQLForVendor(slData, slSql, True)
                        If llCount > 1 Then
                            mAlertsClearWebVendors True
                            '0 is header
                            For c = 1 To llCount - 1
                                '8133
                                If Len(mAlertAddVendorIssue(True, slData(c))) = 0 Then
                                    ilRet = VendorWvmAlert.ExportError
                                    GoTo Cleanup
                                End If
'                                If mRetrieveFromVet(slData(c), slCode, slResult, slMondayOfRecord) Then
'                                    mAlertAddVendorIssue True, slMondayOfRecord, slCode, slResult
'                                Else
'                                    ilRet = VendorWvmAlert.ExportError
'                                    GoTo Cleanup
'                                End If
                            Next c
                        ElseIf llCount < 0 Then
                            ilRet = VendorWvmAlert.ConnectionIssue
                            GoTo Cleanup
                        End If
                    End If
                ElseIf llCount < 0 Then
                    ilRet = VendorWvmAlert.ConnectionIssue
                    GoTo Cleanup
                Else
                    'no old errors. Clear.
                    mAlertsClearWebVendors True
                End If
                slSql = "Select code,attcode,vendoridcode,hasbeensent,errormessage,exportMondayDate,ProcessedDateTime,EnteredDateTime,senttoaff from webvendorexport WHERE (hasbeensent < -1 or hasbeensent = 6) AND sentToAff <> 'Y' "
                slSql = slSql & " AND exportMondayDate >= '" & slDateRealMonday & "'"
                llCount = gExecWebSQLForVendor(slData, slSql, True)
                If llCount > 1 Then
                    ilRet = VendorWvmAlert.NonMonitorIssue
                    '0 is header
                    For c = 1 To llCount - 1
                        '8133
'                        'already senttoaff? return no slResult and false but now we only get !Y, so it's moot
'                        If mInsertToVet(slData(c), slCode, slResult, slMondayOfRecord) Then
'                            mAlertAddVendorIssue True, slMondayOfRecord, slCode, slResult
'                        ElseIf Len(slResult) > 0 Then
'                            ilRet = VendorWvmAlert.ExportError
'                            GoTo Cleanup
'                        End If
                        slWVECode = mAlertAddVendorIssue(True, slData(c))
                        If Len(slWVECode) > 0 Then
                            slSql = "update webvendorexport set senttoaff = 'Y' where code = " & slWVECode
                            If gExecWebSQLWithRowsEffected(slSql) <> 1 Then
                                ilRet = VendorWvmAlert.ExportError
                                GoTo Cleanup
                            End If
                        Else
                            ilRet = VendorWvmAlert.ExportError
                            GoTo Cleanup
                        End If
                    Next c
                ElseIf llCount < 0 Then
                    ilRet = VendorWvmAlert.ConnectionIssue
                End If
            End If 'web export has been sent = 0
        End If 'not test or 'running': continue
    'import
    Else
        blContinue = True
        For c = 0 To 2
           If myServiceFacts(c).Mode = "I" Then
                If myServiceFacts(c).ImportFiles <> "" Then
                    ilRet = VendorWvmAlert.ImportTest
                    blContinue = False
                Else
                    If IsDate(myServiceFacts(c).ImportLast) And IsDate(myServiceFacts(c).ImportSpan) Then
                        slControllerLastImported = myServiceFacts(c).ImportLast
                        'this is how I store this info.  if date = 1/1/1970 then use hour; otherwise, take date and - 1 1/2/1970 = 1 day
                        If Day(myServiceFacts(c).ImportSpan) > 1 Then
                            slDateShouldImport = DateAdd("d", Day(myServiceFacts(c).ImportSpan) - 1, myServiceFacts(c).ImportLast)
                        Else
                            slDateShouldImport = DateAdd("h", Day(myServiceFacts(c).ImportSpan), myServiceFacts(c).ImportLast)
                        End If
                        'if span + last is out in future from now, we are ok.  2nd > 1st is positive number
                        If DateDiff("n", dgWvImportLast, slDateShouldImport) < 0 Then
                            ilRet = VendorWvmAlert.ImportMissed
                            blContinue = False
                        End If
                    Else
                        ilRet = VendorWvmAlert.ImportMissed
                        blContinue = False
                    End If
                End If
                Exit For
           End If
        Next c
        If blContinue Then
            '8133 test each webvendor to make sure updated recently
            If Len(slControllerLastImported) > 0 Then
                slSql = "select idcode,lastimportdate from webvendors WHERE importMethod > 0 and lastimportdate < '" & Format(slControllerLastImported, sgSQLDateForm) & "' "
                llCount = gExecWebSQLForVendor(slData, slSql, True)
                If llCount > 1 Then
                    mAlertAddVendorIssue False, slData(c)
                    ilRet = VendorWvmAlert.NonMonitorIssue
                End If
            Else
                'didn't find 'I' controller!
                ilRet = VendorWvmAlert.ImportError
            End If
        End If
    End If
Cleanup:
    If Not myRst Is Nothing Then
        If (myRst.State And adStateOpen) <> 0 Then
            myRst.Close
        End If
        Set myRst = Nothing
    End If
    mAlertsForWebVendors = ilRet
    Exit Function
errbox:
    ilRet = VendorWvmAlert.ExportError
    gHandleError "", "mAlertsForWebVendors"
    GoTo Cleanup
End Function
'8133
Private Function mAlertAddVendorIssue(blIsExport As Boolean, slData As String) As String
    '8133 added vendor id and result to auf
    'add new only if it's an error.  I don't have to update because I already cleared them all!
    'EXPORT code,attcode,vendoridcode,hasbeensent errormessage exportMondayDate,ProcessedDateTime EnteredDateTime
    Dim slSql As String
    Dim slSub As String
    Dim slCode As String
    Dim slMondayDate As String
    Dim ilResult As Integer
    Dim ilVendorId As Integer
    Dim llAtt As Long
    Dim slRet As String
    
    slRet = ""
    If blIsExport Then
        If mParseWebVendorExport(slData, slCode, ilResult, slMondayDate, ilVendorId, llAtt) Then
            slRet = slCode
            slSub = "E"
        On Error GoTo ERRORBOX
            slSql = "insert into AUF_Alert_User (aufcode,aufentereddate,aufenteredtime,aufstatus,auftype,aufsubtype,aufchfcode,aufMoWeekDate,aufVefCode,aufCountdown,aufUlfCode) values (0,'" & Format(dgWvImportLast, sgSQLDateForm) & "','" & Format(dgWvImportLast, sgSQLTimeForm) & "','R','V','" & slSub & "'," & slCode & ",'" & slMondayDate & " '," & ilVendorId & "," & ilResult & "," & llAtt & ")"
            If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ERRORBOX:
                gHandleError "AffErrorLog.txt", "modVendors-mAlertAddVendorIssue"
                Exit Function
            End If
        End If
    Else
        slSub = "I"
        If mParseWebVendorImport(slData, slMondayDate, ilVendorId) Then
            slSql = "insert into AUF_Alert_User (aufcode,aufentereddate,aufenteredtime,aufstatus,auftype,aufsubtype,aufMoWeekDate,aufVefCode,aufCefCode) values (0,'" & Format(dgWvImportLast, sgSQLDateForm) & "','" & Format(dgWvImportLast, sgSQLTimeForm) & "','R','V','" & slSub & "','" & Format(dgWvImportLast, sgSQLDateForm) & " '," & ilVendorId & "," & VendorWvmAlert.ImportMissed & ")"
            If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ERRORBOX:
                gHandleError "AffErrorLog.txt", "modVendors-mAlertAddVendorIssue"
                Exit Function
            End If
        
        End If
    End If
    mAlertAddVendorIssue = slRet
    Exit Function
ERRORBOX:
    gHandleError "AffErrorLog.txt", "modVendors-mAlertAddVendorIssue"
    mAlertAddVendorIssue = ""
End Function
Private Function mParseWebVendorExport(slData As String, slCode As String, ilResult As Integer, slMondayDate As String, ilVendorId As Integer, llAtt As Long) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim slValues() As String
    Dim slVendor As String
    
    blRet = True
    ilResult = 0
    slCode = ""
    slMondayDate = ""
    ilVendorId = 0
On Error GoTo ErrHand
    slValues = Split(slData, ",")
    If UBound(slValues) >= 8 Then
        slCode = Trim$(Replace(slValues(ExportInfoIndex.code), """", ""))
        slVendor = Trim$(Replace(slValues(ExportInfoIndex.hasbeensent), """", ""))
        If IsNumeric(slVendor) Then
            ilResult = CInt(slVendor)
        Else
            blRet = False
        End If
        slMondayDate = Trim$(Replace(slValues(ExportInfoIndex.mondaydate), """", ""))
        slMondayDate = Format(slMondayDate, sgSQLDateForm)
        slVendor = Trim$(Replace(slValues(ExportInfoIndex.vendoridcode), """", ""))
        If IsNumeric(slVendor) Then
            ilVendorId = CInt(slVendor)
        Else
            blRet = False
        End If
        slVendor = Trim$(Replace(slValues(ExportInfoIndex.attCode), """", ""))
        If IsNumeric(slVendor) Then
            llAtt = CLng(slVendor)
        Else
            blRet = False
        End If
    Else
        blRet = False
    End If
Cleanup:
    Erase slValues
    mParseWebVendorExport = blRet
    Exit Function
ErrHand:
    blRet = False
    gHandleError "AffErrorLog.txt", "ModVendors-mParseWebVendorExport"
    GoTo Cleanup
End Function
Private Function mParseWebVendorImport(slData As String, slMondayDate As String, ilVendorId As Integer) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim slValues() As String
    Dim slVendor As String
    
    blRet = True
    slMondayDate = ""
    ilVendorId = 0
On Error GoTo ErrHand
    slValues = Split(slData, ",")
    If UBound(slValues) >= 1 Then
        slVendor = Trim$(Replace(slValues(0), """", ""))
        If IsNumeric(slVendor) Then
            ilVendorId = CInt(slVendor)
        Else
            blRet = False
        End If
        slMondayDate = Trim$(Replace(slValues(1), """", ""))
        slMondayDate = Format(slMondayDate, sgSQLDateForm)
    Else
        blRet = False
    End If
Cleanup:
    Erase slValues
    mParseWebVendorImport = blRet
    Exit Function
ErrHand:
    blRet = False
    gHandleError "AffErrorLog.txt", "ModVendors-mParseWebVendorImport"
    GoTo Cleanup
End Function
Public Function gExecWebSQLForVendor(aDataArray() As String, sSQL As String, blWantData As Boolean, Optional ilRetryMax As Integer = 0) As Long
    On Error GoTo ErrHandler
    Dim objXMLHTTP
    Dim llReturn As Long
    Dim slISAPIExtensionDLL As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slRegSection As String
    Dim alRecordsArray() As String
    Dim llErrorCode As Long
    Dim ilRetries As Integer
    Dim WebCmds As New WebCommands

    If igDemoMode Then
        gExecWebSQLForVendor = 0
        Exit Function
    End If
    gExecWebSQLForVendor = -1    ' -1 is an error condition.
    If Not gLoadOption(sgWebServerSection, "RootURL", slRootURL) Then
        gLogMsg "Error: gExecWebSQLForVendor: LoadOption RootURL Error", "AffErrorLog.Txt", False
        gMsgBox "Error: gExecWebSQLForVendor: LoadOption RootURL Error"
        Exit Function
    End If
    
    'D.S. 11/27/12 Strip and replace characters (URL enCoding) that cause
    'IIS to stop SQL calls from making it to the database
    sSQL = gUrlEncoding(sSQL)
    
    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    If gLoadOption(sgWebServerSection, "RegSection", slRegSection) Then
        slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & sSQL
    End If
    
    'We will retry every 2 seconds and wait up to 30 seconds
    For ilRetries = 0 To ilRetryMax
        llReturn = 1
        If bgUsingSockets Then
            slResponse = WebCmds.ExecSQL(sSQL)
            If Not Left(slResponse, 5) = "ERROR" Then
                llReturn = 200
            End If
        Else
            Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
            objXMLHTTP.Open "GET", slISAPIExtensionDLL, False
            objXMLHTTP.Send
            llReturn = objXMLHTTP.Status
            slResponse = objXMLHTTP.responseText
            Set objXMLHTTP = Nothing
        End If
    
        If llReturn = 200 Then
            If Not blWantData Then
                ' Caller does not want any data returned.
                gExecWebSQLForVendor = 0
                Exit Function
            End If
    
            ' Parse out the response we got.
            '
            alRecordsArray = Split(slResponse, vbCrLf)
            If Not IsArray(alRecordsArray) Then
                Exit Function
            End If
            ' We have to have back at least two records. The first one is the column headers.
            ' The rest of the entries are the data itself.
            If UBound(alRecordsArray) < 2 Then
                ' If the table is empty, we will get back at least one record containing the column
                ' definitions of the table itself, but no data records.
                gExecWebSQLForVendor = 0
                Exit Function
            End If
        
            ' Each record we get back is a comma delimited string. In this case were only interested
            ' in the first record.
          '  aDataArray = Split(alRecordsArray(1), ",")
            aDataArray = alRecordsArray
            If Not IsArray(aDataArray) Then
                Exit Function
            End If
            gExecWebSQLForVendor = UBound(aDataArray)
            'If gExecWebSQLForVendor < 1 Then
            '9749 removed!
            If gExecWebSQLForVendor > 1 Then
                ' The SQL Statement asked for two fields. DTStamp and PCName
                Exit Function
            End If
        End If
        If ilRetryMax > 0 Then
            Call gSleep(1)
        End If
    Next ilRetries
    
    ' We were never successful if we make it to here.
    gLogMsg "gExecWebSQLForVendor, " & ilRetryMax + 1 & " tries were exceeded. sSQL = " & sSQL & " llErrorCode = " & llErrorCode, "AffErrorLog.txt", False
    Exit Function
    
ErrHandler:
    llErrorCode = Err.Number
    Resume Next
End Function

Public Function gVendorToWebAllowed(Optional tmcWebConnectIssue As Timer) As Boolean
    'if timer exists, then coming from startup.  add message to alert!
    Dim blRet As Boolean
    Dim slSql As String
    Dim llRet As Long
    
On Error GoTo errbox
    blRet = False
    
    '9747
'    If gUsingWeb And igDemoMode = 0 Then
    If gWebAccessTestedOk And igDemoMode = 0 Then
        If gWebVendorIsUsed() Then
            'doesn't exist will return 0!
            slSql = "Select COL_LENGTH('webVendors_Header','attCode') as amount"
            llRet = mExecWebSQLCount(slSql)
            If llRet > 0 Then
                blRet = True
            ElseIf llRet < 0 Then
                mVendorToWebFailed tmcWebConnectIssue
                If Not tmcWebConnectIssue Is Nothing Then
                    mAlertMonitorIssue ConnectionIssue, True
                End If
            Else
                gLogMsg "Call Counterpoint. Web is out of sync: web vendor tables/fields missing.", "affErrorLog.txt", False
            End If
        End If
    End If
Cleanup:
    bgVendorToWebAllowed = blRet
    gVendorToWebAllowed = blRet
    Exit Function
errbox:
    blRet = False
    gHandleError "", "modVendors-gVendorToWebAllowed"
    GoTo Cleanup
End Function
'10320 made public
Public Function gWebVendorIsUsed() As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim sql_rst As ADODB.Recordset
    
    blRet = False
    slSql = "select wvtVendorId from WVT_Vendor_Table " & mGetWebVendorSqlWhere()
    Set sql_rst = gSQLSelectCall(slSql)
    If Not sql_rst.EOF Then
        blRet = True
    End If
    gWebVendorIsUsed = blRet
End Function
Private Function mExecWebSQLCount(sSQL As String, Optional ilRetryMax = 0) As Long
    On Error GoTo ErrHandler
    Dim objXMLHTTP
    Dim llReturn As Long
    Dim slISAPIExtensionDLL As String
    Dim slRootURL As String
    Dim slResponse As String
    Dim slRegSection As String
    Dim alRecordsArray() As String
    Dim llErrorCode As Long
    Dim ilRetries As Integer
    Dim aDataArray() As String
    Dim slCount As String
    Dim WebCmds As New WebCommands
    '9747
'    If igDemoMode Or Not gUsingWeb Then
    If igDemoMode Or Not gWebAccessTestedOk Then
        mExecWebSQLCount = 0
        GoTo Cleanup
    End If
    mExecWebSQLCount = -1    ' -1 is an error condition.
    If Not gLoadOption(sgWebServerSection, "RootURL", slRootURL) Then
        gLogMsg "Error: mExecWebSQLCount: LoadOption RootURL Error", "AffErrorLog.Txt", False
        gMsgBox "Error: mExecWebSQLCount: LoadOption RootURL Error"
        GoTo Cleanup
    End If
    
    'D.S. 11/27/12 Strip and replace characters (URL enCoding) that cause
    'IIS to stop SQL calls from making it to the database
    sSQL = gUrlEncoding(sSQL)
    
    slRootURL = gSetURLPathEndSlash(slRootURL, False)  ' Make sure the path has the final slash on it.
    ' RegSection is a parameter passed to all ISAPI extensions so it will know which section in the
    ' registry to gather additional information. This is necessary to run multiple databases on the
    ' same IIS platform. The password is hardcoded and never changes.
    If gLoadOption(sgWebServerSection, "RegSection", slRegSection) Then
        slISAPIExtensionDLL = slRootURL & "ExecuteSQL.dll" & "?ExecSQL?PW=jfdl" & Now() & "&RK=" & Trim(slRegSection) & "&SQL=" & sSQL
    End If
    
    'We will retry every 2 seconds and wait up to 30 seconds
    For ilRetries = 0 To ilRetryMax
        llReturn = 1
        If bgUsingSockets Then
            slResponse = WebCmds.ExecSQL(sSQL)
            If Not Left(slResponse, 5) = "ERROR" Then
                llReturn = 200
            End If
        Else
            Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
            objXMLHTTP.Open "GET", slISAPIExtensionDLL, False
            objXMLHTTP.Send
            llReturn = objXMLHTTP.Status
            slResponse = objXMLHTTP.responseText
            Set objXMLHTTP = Nothing
        End If
    
        If llReturn = 200 Then
            ' Parse out the response we got.
            alRecordsArray = Split(slResponse, vbCrLf)
            If Not IsArray(alRecordsArray) Then
                GoTo Cleanup
            End If
            ' We have to have back at least two records. The first one is the column headers.
            ' The rest of the entries are the data itself.
            If UBound(alRecordsArray) < 2 Then
                ' If the table is empty, we will get back at least one record containing the column
                ' definitions of the table itself, but no data records.
                mExecWebSQLCount = 0
                GoTo Cleanup
            End If
            slCount = Replace(alRecordsArray(1), """", "")
            If IsNumeric(slCount) Then
                mExecWebSQLCount = CLng(slCount)
                GoTo Cleanup
            Else
                mExecWebSQLCount = 0
                GoTo Cleanup
            End If
        End If
        If ilRetryMax > 0 Then
            Call gSleep(1)
        End If
    Next ilRetries
    
    ' We were never successful if we make it to here.
    gLogMsg "mExecWebSQLCount, " & ilRetryMax + 1 & " tries were exceeded. sSQL = " & sSQL & " llErrorCode = " & llErrorCode, "AffErrorLog.txt", False
Cleanup:
    Erase alRecordsArray
    Exit Function
    
ErrHandler:
    llErrorCode = Err.Number
    Resume Next
End Function

Private Function mAlertsClearWebVendors(blExport As Boolean) As Boolean
    'does not clear monitor's messages!
    Dim slSubType As String
    Dim slSql As String
    Dim blRet As Boolean
    Dim llCount As Long
    
    blRet = False
    If blExport Then
        slSubType = "'E'"
    Else
        slSubType = "'I'"
    End If
    slSql = "Update AUF_Alert_User set aufStatus = 'C', aufClearMethod = 'A', aufClearUstCode = " & igUstCode & ", aufcleardate = '" & Format(gNow(), sgSQLDateForm) & "', aufcleartime = '" & Format(gNow(), sgSQLTimeForm) & "' where auftype = 'V' and aufsubtype = " & slSubType & " and aufcefcode = 0 and aufStatus = 'R' "
    If gSQLAndReturn(slSql, False, llCount) <> 0 Then
        gHandleError "AffErrorLog.txt", "modVendors-gClearVendorAlters"
        Exit Function
    ElseIf llCount > 0 Then
        blRet = True
    End If
    mAlertsClearWebVendors = blRet
    Exit Function
ErrHand:
    gHandleError "", "modVendors-mAlertsClearWebVendors"
End Function
Private Sub mAlertMonitorIssue(ilWvmIssue As VendorWvmAlert, blAdd As Boolean)
    Dim llCefInsert As Long
    Dim slDate As String
    Dim slMon As String
    Dim slSubType As String
    Dim blAlreadyThere As Boolean
    
    'I don't think I need this since I clear all before writing
    blAlreadyThere = mAlertExistsWvmIssue(ilWvmIssue, slSubType)
    If blAdd And Not blAlreadyThere Then
        mAlertAddWvmIssue ilWvmIssue, slSubType
    ElseIf blAlreadyThere And Not blAdd Then
        mAlertClearWvmIssue ilWvmIssue, slSubType
    End If
End Sub
Private Function mAlertExistsWvmIssue(ilWvmIssue As VendorWvmAlert, slSubType As String) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim llCefInsert As Long
    Dim auf_rst As ADODB.Recordset
    
    blRet = False
    slSubType = mAlertSubType(ilWvmIssue)
    llCefInsert = ilWvmIssue
On Error GoTo errbox
    slSql = "select count(*) as Total from Auf_Alert_User where aufStatus = 'R' and auftype = 'V' and aufsubtype = '" & slSubType & "' and aufcefCode = " & llCefInsert
    Set auf_rst = gSQLSelectCall(slSql)
    If auf_rst!Total > 0 Then
        blRet = True
    End If
    mAlertExistsWvmIssue = blRet
    Exit Function
errbox:
    mAlertExistsWvmIssue = False
End Function
Private Sub mAlertAddWvmIssue(ilWvmIssue As VendorWvmAlert, slSubType As String)
    '8129 add direcly to auf..special messages only
    Dim slSql As String
    Dim llCefInsert As Long
    Dim slMon As String
    
    slMon = dgWvImportLast
    slMon = gObtainPrevMonday(slMon)
    slMon = Format(slMon, sgSQLDateForm)
    llCefInsert = ilWvmIssue
    slSql = "insert into AUF_Alert_User (aufcode,aufCreateUstCode,aufentereddate,aufenteredtime,aufstatus,auftype,aufsubtype,aufcefcode,aufMoWeekDate) values (0," & igUstCode & ",'" & Format(dgWvImportLast, sgSQLDateForm) & "','" & Format(dgWvImportLast, sgSQLTimeForm) & "','R','V','" & slSubType & "'," & llCefInsert & ",'" & slMon & "')"
    If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
       gHandleError "AffErrorLog.txt", "modVendors-mAlertAddWvmIssue"
    End If
End Sub
Private Function mAlertClearWvmIssue(ilWvmIssue As VendorWvmAlert, slSubType As String) As Boolean
    '8129  returns true if any were cleared
    Dim slSql As String
    Dim llCefInsert As Long
    Dim slTime As String
    Dim llRet As Long
    Dim blRet As Boolean
    Dim slDate As String
    
    blRet = False
    slTime = Format(dgWvImportLast, sgSQLTimeForm)
    slDate = Format(dgWvImportLast, sgSQLDateForm)
    llCefInsert = ilWvmIssue
    slSql = "Update AUF_Alert_User set aufStatus = 'C', aufClearMethod = 'A', aufClearUstCode = " & igUstCode & ", aufcleardate = '" & slDate & "', aufcleartime = '" & slTime & "' where aufStatus = 'R' and auftype = 'V' and aufsubtype = '" & slSubType & "' and aufcefCode = " & llCefInsert
    If gSQLAndReturn(slSql, False, llRet) <> 0 Then
        gHandleError "AffErrorLog.txt", "modVendors-mAlertClearWvmIssue"
    ElseIf llRet > 0 Then
        blRet = True
    End If
    mAlertClearWvmIssue = blRet
End Function
Private Function mAlertClearWVMBasic() As Boolean
    '8129 returns true if any cleared
    Dim slSql As String
    Dim slTime As String
    Dim slDate As String
    Dim slSubType As String
    Dim llRet As Long
    Dim blRet As Boolean
    
    blRet = False
    slTime = Format(dgWvImportLast, sgSQLTimeForm)
    slDate = Format(dgWvImportLast, sgSQLDateForm)
    slSql = "Update AUF_Alert_User set aufStatus = 'C', aufClearMethod = 'A', aufClearUstCode = " & igUstCode & ", aufcleardate = '" & slDate & "', aufcleartime = '" & slTime & "' where aufStatus = 'R' and auftype = 'V' and aufsubtype in ('I','E') and aufcefCode > 0"
    If gSQLAndReturn(slSql, False, llRet) <> 0 Then
        gHandleError "AffErrorLog.txt", "modVendors-mAlertClearWVMBasic"
    ElseIf llRet > 0 Then
        blRet = True
    End If
    mAlertClearWVMBasic = blRet
End Function

Private Function mAlertSubType(ilWvm As VendorWvmAlert) As String
    Dim slRet As String
    
    'note: VendorWvmAlert.ConnectionIssue shows as export
    slRet = ""
    Select Case ilWvm
        Case VendorWvmAlert.ImportTest, VendorWvmAlert.ImportMissed, VendorWvmAlert.ManagerNotRunning
            slRet = "I"
        Case Else
            slRet = "E"
    End Select
    mAlertSubType = slRet
End Function
Public Function gVendorName(ilVendor As Integer) As String
    Dim slRet As String
    
    Select Case ilVendor
        Case Vendors.NetworkConnect
            slRet = "Network Connect by Marketron"
        Case Vendors.Jelli
            slRet = "Jelli"
        Case Vendors.iheart
            slRet = "iHeart"
        Case Vendors.cBs
            slRet = "CBS"
        '9184
        Case Vendors.stratus
            slRet = "Stratus"
        Case Vendors.iDc
            slRet = "IDC"
        Case Vendors.Wegener_Compel
            slRet = "Wegener-Compel"
        Case Vendors.Wegener_IPump
            slRet = "Wegener-IPump"
        Case Vendors.XDS_ISCI
            slRet = "X-Digital-ISCI"
        Case Vendors.XDS_Break
            slRet = "X-Digital-Break"
        Case Vendors.WideOrbit
            slRet = "WO Traffic Radio Interchange"
         Case Vendors.MrMaster
            slRet = "Mr. Master"
        Case Vendors.RadioTraffic
            slRet = "RadioTraffic.com"
        Case Vendors.RCS
            slRet = "RCS"
        Case Vendors.Synchronicity
            slRet = "Synchronicity"
        Case Vendors.BSI
            slRet = "Business Software Intl."
        Case Vendors.AnalyticOwl
            slRet = "AnalyticOwl"
        Case Vendors.LeadsRX
            slRet = "LeadsRX"
        Case Vendors.RadioTraffic
            slRet = "Radio Workflow"
      Case Else
            slRet = ""
    End Select
    gVendorName = slRet
End Function
Public Function gVendorInitials(ilVendor As Integer) As String
    Dim slRet As String
    
    Select Case ilVendor
        Case Vendors.NetworkConnect
            slRet = "NC"
        Case Vendors.XDS_ISCI
            slRet = "XD"
        Case Vendors.WideOrbit
            slRet = "WO"
         Case Vendors.MrMaster
            slRet = "MM"
        Case Vendors.RadioTraffic
            slRet = "RT"
        Case Vendors.RCS
            slRet = "RC"
        Case Vendors.Synchronicity
            slRet = "SY"
        Case Vendors.iheart
            slRet = "IH"
        Case Vendors.cBs
            slRet = "CB"
        '9184
        Case Vendors.stratus
            slRet = "ST"
        Case Vendors.BSI
            slRet = "BI"
        Case Vendors.AnalyticOwl
            slRet = "AO"
        Case Vendors.LeadsRX
            slRet = "RX"
        Case Vendors.RadioWorkflow
            slRet = "RW"
      Case Else
            slRet = ""
    End Select
    gVendorInitials = slRet
End Function
Public Function gVendorIssue(blIsExport As Boolean, ilIssue As Integer) As String
    Dim slRet As String
    
    slRet = ""
    If blIsExport Then
        Select Case ilIssue
            Case -2
                slRet = "General error"
            Case -3
                slRet = "Problem with vendor"
            '9863
            Case -4
                slRet = "Station or agreement rejected"
            Case -5
                slRet = "Result unknown"
            Case 6
                slRet = "Result unknown"
           Case Else
                slRet = "Unknown issue"
        End Select
    Else
        Select Case ilIssue
            Case -2
                slRet = "General error"
            Case -3
                slRet = "Match failed"
            Case -4
                slRet = "Could not read"
            Case -5
                slRet = "Could not parse"
            Case -6
                slRet = "Multicast issue"
           Case Else
                slRet = "Unknown issue"
        End Select
    
    End If
    gVendorIssue = slRet
End Function
'7967 '8129
Public Function gMonitorVendor(tmcWebConnectIssue As Timer) As String
    'G for success, E for error, Y for other ('G'reen,I need to use E as it's what is expected,'Y'ellow)
    Dim slRet As String
    Dim slSql As String
    Dim ilFirst As Integer
    Dim dlNow As Date
    Dim llRet As Long
    Dim ilVendorIssue As VendorWvmAlert
    Dim blSetTodaysDate As Boolean
    Dim myServiceFacts(2) As ServiceControllerInfo
    Dim c As Integer
On Error GoTo errbox
    'return true unless error
    'when today - the elapse is greater than the last time we tested.... so when I run on 6/30 and then I test again on 6/30 - 1 day is 6/29...won't run until now is 6/31
    slRet = "G"
    blSetTodaysDate = False
    If bgVendorToWebAllowed Then
        ilFirst = 0
        If igWVImportElapsed = 0 Then
            'as minutes
            igWVImportElapsed = 30
            'testing!  Must remove!
           ' igWVImportElapsed = 2
            ilFirst = 1
        End If
        'dan is this ok?
        'dlNow = Now()
        dlNow = gNow()
        'it's noon. Last ran 11:45  11:30 > 11:45?
        If DateAdd("n", -igWVImportElapsed, dlNow) > dgWvImportLast Then
            dgWvImportLast = dlNow
            slRet = mAlertRunByOther()
            If slRet = "X" Then
                slRet = "G"
                'clears monitor issues, but not issues from vit/vet
                blSetTodaysDate = mAlertClearWVMBasic()
                'if this is false, could not connect
                If mQueryServiceController(myServiceFacts) Then
                     'are we currently exporting?
                     If Not bgVendorExportSent Then
                         'mAlerts...will write to the vet table (which gets picked up by auf)
                         'returns if in test mode or there are some waiting in queue
                         ilVendorIssue = mAlertsForWebVendors(True, myServiceFacts)
                         If ilVendorIssue <> VendorWvmAlert.None Then
                            'non monitor issue:  regular errors that were handled in function above.
                             If ilVendorIssue <> VendorWvmAlert.NonMonitorIssue Then
                                 mAlertMonitorIssue ilVendorIssue, True
                             End If
                             Select Case ilVendorIssue
                                 Case VendorWvmAlert.NonMonitorIssue, VendorWvmAlert.ExportMissed, VendorWvmAlert.ExportTest, VendorWvmAlert.ExportError
                                     slRet = "E"
                                 Case VendorWvmAlert.ConnectionIssue
                                     slRet = "Y"
                                     mVendorToWebFailed tmcWebConnectIssue
'                                 Case VendorWvmAlert.ExportError
'                                     slRet = "E"
                                 Case VendorWvmAlert.ExportRunning
                                     slRet = "Y"
                                     If mAlertsClearWebVendors(True) Then
                                        blSetTodaysDate = True
                                     End If
                             End Select
                         Else
                             If mAlertsClearWebVendors(True) Then
                                blSetTodaysDate = True
                            End If
                        End If
                    End If
                    'import
                    ilVendorIssue = mAlertsForWebVendors(False, myServiceFacts)
                    If ilVendorIssue <> VendorWvmAlert.None Then
                        If ilVendorIssue <> VendorWvmAlert.NonMonitorIssue Then
                            mAlertMonitorIssue ilVendorIssue, True
                        End If
                        Select Case ilVendorIssue
                            Case VendorWvmAlert.NonMonitorIssue, VendorWvmAlert.ImportTest, VendorWvmAlert.ImportMissed, VendorWvmAlert.ImportError
                                slRet = "E"
                            Case VendorWvmAlert.ConnectionIssue
                                slRet = "Y"
                                mVendorToWebFailed tmcWebConnectIssue
                        End Select
                    Else
                        mAlertsClearWebVendors False
                    End If
                    'now the manager!
                    For c = 0 To 2
                       If myServiceFacts(c).Mode = "M" Then
                           If Not myServiceFacts(c).IsRunning Then
                               slRet = "E"
                               mAlertMonitorIssue VendorWvmAlert.ManagerNotRunning, True
                           End If
                           Exit For
                        End If
                    Next c
                Else
                    If slRet <> "E" Then
                        slRet = "Y"
                    End If
                    mAlertMonitorIssue VendorWvmAlert.ConnectionIssue, True
                    mVendorToWebFailed tmcWebConnectIssue
                End If
                If Not bgVendorExportSent Then
                    If slRet = "G" Then
                        gUpdateTaskMonitor ilFirst, "WVM"
                        If Not blSetTodaysDate Then
                           'set a cleared date when didn't have to do anything to block other users
                           mAlertSetUserGeneric
                        End If
                   ' in general, the old issue is still set as old
                    ElseIf ilVendorIssue = VendorWvmAlert.NonMonitorIssue And Not blSetTodaysDate Then
                       mAlertSetUserGeneric
                    End If
                   Debug.Print "Testing WVM: Returns '" & slRet & "' at time: " & dgWvImportLast '& " and testing import span from : " & dlNow & " and timer span: " & igWVImportElapsed
                End If
            ElseIf Not bgVendorExportSent Then
                Debug.Print "Testing WVM: Returns '" & slRet & "' at time: " & dlNow & " doesn't need to run since previous user ran since " & Format(DateAdd("n", -igWVImportElapsed, dlNow), sgSQLTimeForm)
            End If 'another user hasn't run
            If bgVendorExportSent Then
                bgVendorExportSent = False
                slRet = "Y"
                mAlertsClearWebVendors True
                mAlertMonitorIssue VendorWvmAlert.WebVendorSent, True
                Debug.Print "Testing WVM: Returns '" & slRet & "' at time: " & dgWvImportLast '& " and testing import span from : " & dlNow & " and timer span: " & igWVImportElapsed
            End If
            gAlertForceCheck
        End If 'time to run
    Else
        slRet = "Y"
    End If
Cleanup:
    gMonitorVendor = slRet
    Exit Function
errbox:
    slRet = "E"
    GoTo Cleanup
End Function
Private Sub mAlertSetUserGeneric()
    Dim llCount As Long
    Dim ilCef As Integer
    Dim slDate As String
    Dim slTime As String
    Dim slSQLQuery As String
    slTime = Format(dgWvImportLast, sgSQLTimeForm)
    slDate = Format(dgWvImportLast, sgSQLDateForm)

    ilCef = VendorWvmAlert.UserSetTime
    'this is a fake record. always status C,sub E, cefcode above
    slSQLQuery = "Update AUF_Alert_User set aufStatus = 'C', aufClearMethod = 'A', aufClearUstCode = " & igUstCode & ", aufcleardate = '" & slDate & "', aufcleartime = '" & slTime & "' where aufStatus = 'C' and auftype = 'V' and aufsubtype  = 'E' and aufcefCode = " & ilCef
    If gSQLAndReturn(slSQLQuery, False, llCount) <> 0 Then
        gHandleError "AffErrorLog.txt", "modVendors-mAlertSetUserGeneric"
        Exit Sub
    End If
    If llCount = 0 Then
        slSQLQuery = "insert into AUF_Alert_User (aufcode,aufCreateUstCode,aufentereddate,aufenteredtime,aufstatus,auftype,aufsubtype,aufcefcode,aufCleardate,aufClearTime) values (0," & igUstCode & ",'" & slDate & "','" & slTime & "','R','V','E'," & ilCef & ",'" & slDate & "','" & slTime & "')"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            gHandleError "AffErrorLog.txt", "modVendors-mAlertSetUserGeneric"
        End If
    End If
End Sub
Private Sub mVendorToWebFailed(tmcWebConnectIssue As Timer)
    bgVendorToWebAllowed = False
    If Not tmcWebConnectIssue Is Nothing Then
       tmcWebConnectIssue.Enabled = True
    End If
End Sub
Public Function gAdjustAllowedExportsImports(Optional ilVendorChoice As Vendors = Vendors.None, Optional blAdjustMenu As Boolean = True) As Boolean
'8156  first time through, set all.  from vendor form, just reset as specified.
'note that I go through all possible vendors even though only a few will be set/unset.
'I could also speed things up by not getting available vendors and just using ilVendorChoice, but this won't be that often (vendor form)
' O= is it allowed?  Plan to use with ilVendorChoice, but will show if at least one is allowed if vendors = none
' blAdjustMenu  set to false if don't need to set main form's menus.
    Dim blRet As Boolean
    Dim ilCount As Integer
    Dim tlMyVendors() As VendorInfo
    Dim blAllowed As Boolean
    Dim myRst As ADODB.Recordset
    Dim ilVendorPass As Vendors
    Dim slSQLQuery As String
    blRet = False
    tlMyVendors = gGetAvailableVendors()
    For ilCount = 0 To UBound(tlMyVendors) - 1
        If ilVendorChoice = Vendors.None Or ilVendorChoice = tlMyVendors(ilCount).iIdCode Then
            slSQLQuery = "SELECT count(*) as amount FROM WVT_Vendor_Table WHERE wvtVendorID = " & tlMyVendors(ilCount).iIdCode
            'because isci after break, it will override it, setting xds to false when break just set it to true!
            If ilVendorChoice = Vendors.None And tlMyVendors(ilCount).iIdCode = XDS_ISCI Then
                slSQLQuery = slSQLQuery & " OR wvtvendorid = " & Vendors.XDS_Break
            End If
            Set myRst = gSQLSelectCall(slSQLQuery)
            blAllowed = False
            If Not myRst.EOF Then
                If myRst!amount > 0 Then
                    blAllowed = True
                    blRet = True
                    '8322 block Marketron if going to web
                    If tlMyVendors(ilCount).iIdCode = Vendors.NetworkConnect Then
                        If mIsWebVendor(Vendors.NetworkConnect) Then
                            blAllowed = False
                            blRet = False
                            '9/9/19 allow csi guide
                            If (StrComp(sgUserName, "Guide", 1) = 0) And Not bgLimitedGuide Then
                                blAllowed = True
                                blRet = True
                            End If
                        End If
                    End If
                End If
            End If
            If blAdjustMenu Then
                ilVendorPass = tlMyVendors(ilCount).iIdCode
                frmMain.gAllowedExportsImportsInMenu blAllowed, ilVendorPass 'tlMyVendors(ilCount).iIdCode
            End If
        End If
    Next ilCount
    If Not myRst Is Nothing Then
        If (myRst.State And adStateOpen) <> 0 Then
            myRst.Close
        End If
        Set myRst = Nothing
    End If
   gAdjustAllowedExportsImports = blRet
End Function
Public Function gIsWebVendor(ilVendorId As Integer) As Boolean
    gIsWebVendor = mIsWebVendor(ilVendorId)
End Function
Private Function mIsWebVendor(ilVendorChoice As Integer) As Boolean
    Dim myRst As ADODB.Recordset
    Dim blRet As Boolean
    Dim slSQLQuery As String
On Error GoTo errbox
    slSQLQuery = "select count(*) as amount from WVT_Vendor_Table " & mGetWebVendorSqlWhere() & " AND wvtvendorid = " & ilVendorChoice
    Set myRst = gSQLSelectCall(slSQLQuery)
    If Not myRst.EOF Then
        If myRst!amount > 0 Then
            blRet = True
        End If
    End If
    If Not myRst Is Nothing Then
        If (myRst.State And adStateOpen) <> 0 Then
            myRst.Close
        End If
        Set myRst = Nothing
    End If
    mIsWebVendor = blRet
    Exit Function
errbox:
    gHandleError "", "mIsWebVendor"
     mIsWebVendor = blRet
End Function
Private Function mAlertRunByOther() As String
    'O: G E Y or X, meaning need to run
    Dim blRet As Boolean
    Dim slRet As String
    Dim myRst As ADODB.Recordset
    Dim slSql As String
    Dim dlDate As Date
    Dim ilAlert As Integer
    
    blRet = False
    slRet = "X"
On Error GoTo errbox
'test will get first and then will not after the # of the last number
'igWVImportElapsed = igWVImportElapsed + (igWVImportElapsed * 1)
    dlDate = DateAdd("n", -igWVImportElapsed, dgWvImportLast)
'reset test
'igWVImportElapsed = igWVImportElapsed - (igWVImportElapsed * 1)
    slSql = "Select count(*) as amount from auf_alert_user where auftype = 'V' and aufstatus = 'C' and aufcleardate >= '" & Format(dlDate, sgSQLDateForm) & "' AND aufcleartime > '" & Format(dlDate, sgSQLTimeForm) & "'"
    Set myRst = gSQLSelectCall(slSql)
    If Not myRst.EOF Then
        If myRst!amount > 0 Then
            slSql = "select aufcefcode from auf_alert_user where auftype = 'V' and aufstatus = 'R'"
            Set myRst = gSQLSelectCall(slSql)
            If Not myRst.EOF Then
                Do While Not myRst.EOF
                    ilAlert = myRst!aufcefcode
                    If mParseMonitorIssueIsYellow(ilAlert) Then
                        slRet = "Y"
                        Exit Do
                    Else
                        slRet = "E"
                    End If
                    myRst.MoveNext
                Loop
            Else
                slRet = "G"
            End If
        Else
            slSql = "Select count(*) as amount from auf_alert_user where auftype = 'V' and aufstatus = 'R' and aufEntereddate >= '" & Format(dlDate, sgSQLDateForm) & "' AND aufEnteredtime > '" & Format(dlDate, sgSQLTimeForm) & "'"
            Set myRst = gSQLSelectCall(slSql)
            'not find?  then must run.  but slret already = X
            If Not myRst.EOF Then
                If myRst!amount > 0 Then
                    slSql = "select aufcefcode from auf_alert_user where auftype = 'V' and aufstatus = 'R'"
                    Set myRst = gSQLSelectCall(slSql)
                    If Not myRst.EOF Then
                        Do While Not myRst.EOF
                            ilAlert = myRst!aufcefcode
                            If mParseMonitorIssueIsYellow(ilAlert) Then
                                slRet = "Y"
                                Exit Do
                            Else
                                slRet = "E"
                            End If
                            myRst.MoveNext
                        Loop
                    Else
                        slRet = "G"
                    End If
                End If
            End If
        End If
    End If
Cleanup:
    If Not myRst Is Nothing Then
        If (myRst.State And adStateOpen) <> 0 Then
            myRst.Close
        End If
        Set myRst = Nothing
    End If
    mAlertRunByOther = slRet
    Exit Function
errbox:
    slRet = "X"
    GoTo Cleanup
End Function
Private Function mParseMonitorIssueIsYellow(ilAlert As Integer) As Boolean
    Dim blRet As Boolean
    
    blRet = False
    'none means regular alert issue--red!
    If ilAlert <> VendorWvmAlert.None Then
        Select Case ilAlert
            'these are yellow
            Case VendorWvmAlert.ConnectionIssue, VendorWvmAlert.ExportRunning, VendorWvmAlert.WebVendorSent
                blRet = True
        End Select
    End If
    mParseMonitorIssueIsYellow = blRet
End Function
Private Function mQueryServiceController(myServiceFacts() As ServiceControllerInfo) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim slData() As String
    Dim slValues() As String
    Dim llCount As Long
    Dim c As Integer
    Dim ilCurrent As Integer
    'service controller is always 3 records
    ilCurrent = 0
    blRet = True
    slSql = "Select Mode,GenerateDebug,GenerateFile,ImportLast,ImportFiles,isRunning,ImportSpan from vendorservicecontroller"
    llCount = gExecWebSQLForVendor(slData, slSql, True)
    If llCount > 0 Then
        For c = 1 To llCount - 1
            slValues = Split(slData(c), ",")
            If UBound(slValues) >= 6 Then
                ilCurrent = c - 1
                If ilCurrent < 3 Then
                    With myServiceFacts(ilCurrent)
                        .Mode = Trim$(Replace(slValues(0), """", ""))
                        If Trim$(Replace(slValues(1), """", "")) = "Y" Then
                            .GenerateDebug = True
                        End If
                        If Trim$(Replace(slValues(2), """", "")) = "Y" Then
                            .GenerateFile = True
                        End If
                        .ImportLast = Trim$(Replace(slValues(3), """", ""))
                        .ImportFiles = Trim$(Replace(slValues(4), """", ""))
                        If Trim$(Replace(slValues(5), """", "")) = "Y" Then
                            .IsRunning = True
                        End If
                         .ImportSpan = Trim$(Replace(slValues(6), """", ""))
                    End With
                End If
            End If
        Next c
    ElseIf llCount < 0 Then
        blRet = False
    End If
    Erase slData
    Erase slValues
    mQueryServiceController = blRet
End Function
'8259
Public Function gVendorReportInfoToText(blIsExport As Boolean, slStartDate As String, slEndDate As String) As String
    Dim slSql As String
    Dim slData() As String
    Dim slValues() As String
    Dim c As Integer
    Dim llCount As Long
    Dim slExportOrImport As String
    Dim slRet As String
    Dim olFileSys As FileSystemObject
    Dim olCsv As TextStream
    
    slRet = ""
    If blIsExport Then
        slExportOrImport = "E"
    Else
        slExportOrImport = "I"
    End If
    slSql = "Select attcode,vendorIdCode,SpotsDate,SpotsCount,ProcessedDateTime from webvendorcountarchive where ExportOrImport = '" & slExportOrImport & "' AND spotsdate between '" & Format$(slStartDate, sgSQLDateForm) & "' AND '" & Format$(slEndDate, sgSQLDateForm) & "'"
    llCount = gExecWebSQLForVendor(slData, slSql, True)
    If llCount > 0 Then
        Set olFileSys = New FileSystemObject
        slRet = sgImportDirectory & "VendorFile-" & sgUserName & ".txt"
On Error GoTo ERRORBOX
        If olFileSys.FolderExists(sgImportDirectory) Then
            Set olCsv = olFileSys.OpenTextFile(slRet, ForWriting, True)
            For c = 1 To llCount - 1
                'Debug.Print slData(c)
                olCsv.WriteLine slData(c)
            Next c
        Else
            slRet = ""
        End If
    End If
Cleanup:
    If Not olCsv Is Nothing Then
        olCsv.Close
    End If
    Set olCsv = Nothing
    Set olFileSys = Nothing
    gVendorReportInfoToText = slRet
    Exit Function
ERRORBOX:
    slRet = ""
    gHandleError "AffErrorLog.txt", "modVendors-gVendorReportInfoToText"
    GoTo Cleanup
End Function
Public Function gAllowVendorAlerts(blTestUser As Boolean) As Boolean
    '8273
    Dim blRet As Boolean
    Dim ilIndex As Integer
    Dim slSql As String
    Dim sql_rst As ADODB.Recordset

    
    blRet = False
    For ilIndex = 0 To UBound(tgTaskInfo) Step 1
        If Trim$(tgTaskInfo(ilIndex).sTaskCode) = "WVM" Then
            If tgTaskInfo(ilIndex).iMenuIndex > 0 Then
                blRet = True
            End If
            Exit For
        End If
    Next ilIndex
    If blRet And blTestUser Then
        'test user here
        blRet = True
        'csi
        If (StrComp(sgUserName, "Guide", 1) = 0) And Not bgLimitedGuide Then
        
        Else
            slSql = "select ustVendorAlert from ust where ustcode =" & igUstCode
            Set sql_rst = gSQLSelectCall(slSql, "modVendors-gAlloweVendorAlerts")
            If Not sql_rst.EOF Then
                If sql_rst!ustVendorAlert = "N" Then
                    blRet = False
                    
                End If
            End If
        End If
    End If
    gAllowVendorAlerts = blRet
End Function
'8133 8230
Public Function gVendorWvmIssue(llWvmIssueCode As Long) As String
    Dim slRet As String
    
    slRet = ""
    Select Case llWvmIssueCode
        Case VendorWvmAlert.ImportTest
            slRet = "'test mode'; not retrieving"
        Case VendorWvmAlert.ImportMissed
            slRet = "missed last scheduled import."
        Case VendorWvmAlert.ExportTest
            slRet = "'test mode'; not sending"
        Case VendorWvmAlert.ExportMissed
            slRet = "not all exports went out."
        Case VendorWvmAlert.ConnectionIssue
            slRet = "cannot connect to web database."
        Case VendorWvmAlert.ExportRunning
            slRet = "Export currently running."
        Case VendorWvmAlert.WebVendorSent
            slRet = "Exports queued to go out."
        Case VendorWvmAlert.ExportError
            slRet = "Error in the code."
        Case VendorWvmAlert.ManagerNotRunning
            slRet = "Manager not running."
    End Select
    gVendorWvmIssue = slRet
End Function
'8418
Public Function gVendorMinVersion(ilVendorCode As Integer, tlVendorInfo() As VendorInfo) As Integer
    Dim ilRet As Integer
    Dim c As Integer
    
    ilRet = 0
    For c = 0 To UBound(tlVendorInfo)
        If tlVendorInfo(c).iIdCode = ilVendorCode Then
            ilRet = tlVendorInfo(c).iMinimumWebVersion
            Exit For
        End If
    Next c
    gVendorMinVersion = ilRet
End Function
'8862
Public Function gUpdateVendorStatusAsNeeded(llAttCode As Long, ilVendorCodes() As Integer) As Boolean
    'returns true if no error. VendorCodes should be + 1
    Dim ilCount As Integer
    Dim slSql As String
    Dim myRst As ADODB.Recordset
    
On Error GoTo ERRORBOX
    For ilCount = 0 To UBound(ilVendorCodes) - 1 Step 1
        slSql = "SELECT count(*) as amount FROM VAT_Vendor_Agreement WHERE vatAttcode = " & llAttCode & " AND vatWVTVendorId = " & ilVendorCodes(ilCount)
        Set myRst = gSQLSelectCall(slSql)
        If Not myRst.EOF Then
            If myRst!amount = 0 Then
                slSql = "insert into Vat_Vendor_Agreement (vatAttCode,vatWvtVendorID,vatSentToWeb) values (" & llAttCode & "," & ilVendorCodes(ilCount) & ",'Y')"
                If gSQLWaitNoMsgBox(slSql, True) <> 0 Then
                    gUpdateVendorStatusAsNeeded = False
                    gHandleError "AffErrorLog.txt", "modVendors-gUpdateVendorStatusAsNeeded"
                    Exit Function
                End If
            End If
        End If
    Next ilCount
    gUpdateVendorStatusAsNeeded = True
    Exit Function
ERRORBOX:
    gUpdateVendorStatusAsNeeded = False
    gHandleError "AffErrorLog.txt", "modVendors-gUpdateVendorStatusAsNeeded"
    
End Function
'8862
Public Function gAutoDeliveryVendors() As VendorInfo()
    'vendors who are set to allow autosetting.  array is +1
    Dim tlVendorInfo() As VendorInfo
    Dim slSql As String
    Dim myRst As ADODB.Recordset
    Dim ilCount As Integer
    Dim tlAvailableVendors() As VendorInfo
    Dim ilAvailableCount As Integer
    
On Error GoTo ERRORBOX
    ilCount = 0
    tlAvailableVendors = gGetAvailableVendors()
    ReDim tlVendorInfo(ilCount) As VendorInfo
    slSql = "SELECT wvtVendorId, wvtName FROM WVT_Vendor_Table WHERE (wvtVendorUpdateVAT = 'Y')"
    Set myRst = gSQLSelectCall(slSql)
    Do While Not myRst.EOF
        For ilAvailableCount = 0 To UBound(tlAvailableVendors)
            If myRst!wvtVendorId = tlAvailableVendors(ilAvailableCount).iIdCode Then
                If tlAvailableVendors(ilAvailableCount).bAutoVendingCSIAllowed Then
                    tlVendorInfo(ilCount).sSourceName = tlAvailableVendors(ilAvailableCount).sSourceName
                    tlVendorInfo(ilCount).bAutoVendingCSIAllowed = tlAvailableVendors(ilAvailableCount).bAutoVendingCSIAllowed
                    tlVendorInfo(ilCount).sName = myRst!wvtName
                    tlVendorInfo(ilCount).sAutoVendor = "Y"
                    tlVendorInfo(ilCount).iIdCode = myRst!wvtVendorId
                    ilCount = ilCount + 1
                    ReDim Preserve tlVendorInfo(ilCount) As VendorInfo
                End If
                Exit For
            End If
        Next ilAvailableCount
        myRst.MoveNext
    Loop
    gAutoDeliveryVendors = tlVendorInfo
    Exit Function
ERRORBOX:
    gHandleError "AffErrorLog.txt", "modVendors-gAutoDeliveryVendors"
    gAutoDeliveryVendors = tlVendorInfo
End Function
'9256
Private Function mIsDualXDSProviderAndReceiverIdSources(slISCISiteId As String, slCueSiteId As String) As Boolean
   'not really if dual provider:  if dual AND setting site id outside of default.
    Dim blRet As Boolean
    Dim slChoice As String
    Dim slSection As String
    Dim ilCounter As Integer
    Dim ilUpper As Integer
    Dim slRet As String
    Dim slXMLINIInputFile As String
    
    blRet = False
    slCueSiteId = "B"
    slISCISiteId = "B"
    ilUpper = UBound(sgXDSSection)
    'dual provider.  now let's see if using ReceiverIDSource
    If ilUpper > 1 Then
        slXMLINIInputFile = gXmlIniPath(True)
        For ilCounter = 0 To ilUpper - 1
            'get name 'XDigital-CU'
            slSection = Mid(sgXDSSection(ilCounter), 2, Len(sgXDSSection(ilCounter)) - 2)
            gLoadFromIni slSection, STATIONXMLRECEIVERID, slXMLINIInputFile, slChoice
            If slChoice = "A" Or slChoice = "S" Then
                blRet = True
                gLoadFromIni slSection, "WebServiceURL", slXMLINIInputFile, slRet
                If InStr(slRet, "backoffice") > 0 Then
                    slCueSiteId = slChoice
                ElseIf InStr(slRet, "abcdService") > 0 Then
                    slISCISiteId = slChoice
                End If
            End If
        Next ilCounter
    End If
    mIsDualXDSProviderAndReceiverIdSources = blRet
End Function
'9818
Private Function mIsDualXDSProviderAndSharedHeadEndIDs(ilISCIHeadEndId As Integer, ilCueHeadEndId As Integer) As Boolean
   'not really if dual provider:  if dual AND setting site id outside of default.
    Dim blRet As Boolean
    Dim slChoice As String
    Dim slSection As String
    Dim ilCounter As Integer
    Dim ilUpper As Integer
    Dim slRet As String
    Dim slXMLINIInputFile As String
    Dim ilChoice As Integer
    
    blRet = False
    ilISCIHeadEndId = 0
    ilCueHeadEndId = 0
    ilUpper = UBound(sgXDSSection)
    'dual provider.  now let's see if using SharedHeadEnd
    If ilUpper > 1 Then
        slXMLINIInputFile = gXmlIniPath(True)
        For ilCounter = 0 To ilUpper - 1
            'get name 'XDigital-CU'
            slSection = Mid(sgXDSSection(ilCounter), 2, Len(sgXDSSection(ilCounter)) - 2)
            gLoadFromIni slSection, SHAREDHEADENDID, slXMLINIInputFile, slChoice
             If slChoice <> "Not Found" Then
                If IsNumeric(slChoice) Then
                    ilChoice = CInt(slChoice)
                    If ilChoice > 0 Then
                        blRet = True
                        gLoadFromIni slSection, "WebServiceURL", slXMLINIInputFile, slRet
                        If InStr(slRet, "backoffice") > 0 Then
                            ilCueHeadEndId = ilChoice
                        ElseIf InStr(slRet, "abcdService") > 0 Then
                            ilISCIHeadEndId = ilChoice
                        End If
                    End If
                End If
            End If
        Next ilCounter
    End If
    mIsDualXDSProviderAndSharedHeadEndIDs = blRet
End Function
'9818
Public Sub gGetSharedHeadEnd(ilIsciHeadEnd As Integer, ilCueHeadEnd As Integer)
    Dim slRet As String
    
    If mIsDualXDSProviderAndSharedHeadEndIDs(ilIsciHeadEnd, ilCueHeadEnd) Then
    Else
        gLoadFromIni "XDigital", SHAREDHEADENDID, gXmlIniPath(True), slRet
        If slRet <> "Not Found" Then
            If IsNumeric(slRet) Then
                ilIsciHeadEnd = CInt(slRet)
                ilCueHeadEnd = ilIsciHeadEnd
            End If
        End If
    End If
End Sub
'10125
Private Function mGetVehicleAndShttCodeAndNotServiceAgreement(llAtt As Long, slVefCodeReturn As String, slShttCodeReturn As String) As Boolean
    Dim ilVefCode As Integer
    Dim ilShttCode As Integer
    Dim myRst As ADODB.Recordset
    Dim slSql As String
    Dim blRet As Boolean
    
    blRet = True
    slSql = "Select attvefCode,attServiceAgreement,attshfcode from att where attcode = " & llAtt
    Set myRst = gSQLSelectCall(slSql)
    If Not myRst.EOF Then
        ilVefCode = myRst!attvefCode
        ilShttCode = myRst!attshfcode
        If myRst!attServiceAgreement = "Y" Then
            blRet = False
        End If
    End If
    slVefCodeReturn = Str(ilVefCode)
    slVefCodeReturn = Trim$(slVefCodeReturn)
    slShttCodeReturn = Str(ilShttCode)
    slShttCodeReturn = Trim$(slShttCodeReturn)
    mGetVehicleAndShttCodeAndNotServiceAgreement = blRet
End Function
Private Function mGetHonorDaylight(slShttCode As String) As String
    Dim myRst As ADODB.Recordset
    Dim slSql As String
    Dim slRet As String
    
    slRet = ""
    slSql = "Select shttAckDaylight from shtt where shttcode = " & slShttCode
    Set myRst = gSQLSelectCall(slSql)
    If Not myRst.EOF Then
        If myRst!shttAckDaylight = 1 Then
            slRet = "N"
        Else
            slRet = "Y"
        End If
    End If
    mGetHonorDaylight = slRet
End Function
