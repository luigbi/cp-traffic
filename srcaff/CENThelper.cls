VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CENThelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'by agreement.  If sending more than one agreement at a time (and thus don't know if it was successful at the time of ent creation), create by agreement with status 'I' and then update later.
'UpdateStatusLater: set property to true. send by agreement, as normal.  Then update status when done (UpdateAsSuccessful) what we are doing is 'creatingEnts' with status of error, and saving the entcodes.
'   on 'Update..', we change the status to 'S'
'Marketron import needs multicast and out of date bucket
Public ErrorLog As String
'new for 7509.  Only needed for web import/export.  Not sure how it will work there, so not fully implemented.  Right now, never cleared out, and only writes when to ent on creation when there is an error
Public EntErrorMsg As String
Private smFileName As String
Private smErrorMessage As String
'do we need to save the created entcodes? Only for 'updateAsSuccessful' or if user will call 'EntCodes' to get list (2,4,5) of entCodes.
Private bmReturnEntCode As Boolean
'used with above.  The only way to clear is to reset the boolean above with 'ReturnEntCode = ' and either true or false
Private smReturnCodes As String
'to set special rules. Checked in 'ProcessStart' and set in change to Party
Private bmFirstTime As Boolean
'for marketron import.  Multicast will be done by storing the alternate stations info
Private bmIsMulticast As Boolean
Dim lmAtts() As Long
Dim imShtts() As Integer
'end marketron import
'to help change status 'I' to 'E' or 'S'
Private bmCreateUniqueFilenames As Boolean
'import wegener sends right to web after importing, so need 2 sets of records
Private bmCopyForWeb As Boolean
Private smUniqueWebName As String
Dim myRsBuckets As ADODB.Recordset
Dim tmENT As ENT
' "Special Rules"
'imports compare ingest to received to determine status
Private bmIngestDeterminesStatus As Boolean
'7511
Private Const DELETEBEFOREMONTHS As Integer = 6
'Public Enum ThirdPartyEnum
'    Web = 0
'    Marketron = 1
'    Xds = 2
'    wegenerCompel = 3
'    wegenerIPumps = 4
'    cBs = 5
'    cumulusQ = 6
'    ClearChannel = 7
'    iDc = 8
'    Jelli = 9
'    Univision = 10
'End Enum
Public Enum TypesEnum
    Sentunpostedtoweb = 0
    Receivedpostedfromweb = 1
    Exportunposted3rdparty = 2
    Importposted3rdparty = 3
    senttoweb3rdpartyPosted = 4
    Forwardtoweb3rdparty = 5
    importedinToweb3party = 6
End Enum
Public Enum StatusEnum
    Successful = 0
    EntError = 1
    Incomplete = 2
    NotSent = 3
End Enum
Public Enum ProcessingEnum
    SentOrReceived = 0
    Replacement = 4
    MakeGood = 3
    Bonus = 5
    Ingested = 2
    Asts = 1
    Deleted = 6
End Enum
Public Enum UpdateExceptionEnum
    None
    AgreementCodes
    StationCodes
    VehicleCodes
    EntCode
End Enum
Public Enum XdsAllOrRegionalEnum
    All = 1
    Regional = 2
End Enum
Public Property Let CreateCopyForWeb(blValue As Boolean)
    bmCopyForWeb = blValue
    smUniqueWebName = ""
End Property
Public Property Let CreateUniqueFilenames(blValue As Boolean)
    bmCreateUniqueFilenames = blValue
End Property
Public Property Get EntCodes() As String
    '"1,2,3"
    EntCodes = mCleanReturnCodes(False)
End Property
Public Property Let ReturnEntCode(blValue As Boolean)
    'this is how you'd clear out the return codes if you never called 'updateAsSuccessful' and were going to reuse the class
    bmReturnEntCode = blValue
    smReturnCodes = ""
End Property
Public Property Get User() As Integer
    User = tmENT.iUstCode
End Property
Public Property Let User(ilUstCode As Integer)
    tmENT.iUstCode = ilUstCode
End Property
Public Property Get Vehicle() As Integer
    Vehicle = tmENT.iVefCode
End Property
Public Property Let Vehicle(ilVefCode As Integer)
    tmENT.iVefCode = ilVefCode
End Property
Public Property Get Station() As Integer
    Station = tmENT.iShttCode
End Property
Public Property Let Station(ilShttCode As Integer)
    tmENT.iShttCode = ilShttCode
End Property
Public Property Get Agreement() As Long
    Agreement = tmENT.lAttCode
End Property
Public Property Let Agreement(llAttCode As Long)
    tmENT.lAttCode = llAttCode
    smErrorMessage = ""
End Property
Public Property Get fileName() As String
    fileName = smFileName
End Property
Public Property Let fileName(slFileName As String)
    smFileName = slFileName
End Property
Public Property Let ThirdParty(elSource As Vendors)
    tmENT.s3rdParty = mDecodeParty(elSource)
End Property
Public Property Get ThirdParty() As Vendors
    ThirdParty = mEncodeParty(tmENT.s3rdParty)
End Property
Public Property Let TypeEnt(elType As TypesEnum)
    tmENT.sType = mDecodeType(elType)
    bmFirstTime = True
End Property
Public Property Get StatusEnt() As StatusEnum
    StatusEnt = tmENT.sStatus
End Property
Public Property Get ErrorMessage() As String
    If Len(smErrorMessage) > 0 Then
        ErrorMessage = "cENTHelper issue: " & smErrorMessage
        smErrorMessage = ""
    Else
        ErrorMessage = ""
    End If
End Property
'10099
Public Property Let XdsAllOrRegional(ilValue As XdsAllOrRegionalEnum)
    tmENT.iBonusCount = ilValue
End Property
Public Function Add(slFeedDate As String, llGsfCode As Long, Optional slProcess As ProcessingEnum = SentOrReceived, Optional ilValue As Integer = 1, Optional blDontUseLogic As Boolean = False) As Boolean
    'UseLogic?  every subset is added automatically to bucket above it.  example:  export:  if it's sent, then it must've been an astcode.  Both incremented
    'But if it's an astCode, doesn't mean it was sent. Only ast incremented.
    Dim slBucket As String
    Dim blRet As Boolean
    
On Error GoTo ERRORBOX
    blRet = True
    slFeedDate = Format(slFeedDate, sgSQLDateForm)
    slBucket = mDecodeProcessing(slProcess)
    With myRsBuckets
        .Filter = "FeedDate = '" & slFeedDate & "' AND Game = " & llGsfCode
        If .EOF Then
            .AddNew Array("FeedDate", "Game"), Array(slFeedDate, llGsfCode)
        End If
         myRsBuckets.Fields(slBucket).Value = myRsBuckets.Fields(slBucket).Value + ilValue
        'import - bonus,etc. will add to ingested, which will add to sentorReceived.  export-sentOrReveived adds to Ast
        If Not blDontUseLogic Then
            If tmENT.sType = mDecodeType(TypesEnum.Receivedpostedfromweb) Or tmENT.sType = mDecodeType(TypesEnum.Importposted3rdparty) Then
                If slProcess = Ingested Then
                    myRsBuckets.Fields("SentOrReceived").Value = myRsBuckets.Fields("SentOrReceived").Value + ilValue
                'mg,bonus,replacement
                ElseIf slProcess > 2 Then
                    myRsBuckets.Fields("SentOrReceived").Value = myRsBuckets.Fields("SentOrReceived").Value + ilValue
                    myRsBuckets.Fields("Ingested").Value = myRsBuckets.Fields("Ingested").Value + ilValue
                End If
            'exports: sent adds to ast  Make sure not an import!
            Else
                If slProcess = SentOrReceived Then
                    myRsBuckets.Fields("Asts").Value = myRsBuckets.Fields("Asts").Value + ilValue
                End If
            End If
        End If
        .Update
    End With
    Add = blRet
    Exit Function
ERRORBOX:
    smErrorMessage = smErrorMessage & " Problem in 'Add': couldn't add date " & slFeedDate & ": " & Err.Description & vbCrLf
    Add = False
End Function
Public Sub ProcessStart()
    Dim slNow As String
    
    'dont use gnow
    slNow = Now()
    tmENT.sEnteredDate = Format(slNow, sgSQLDateForm)
    tmENT.sEnteredTime = Format(slNow, sgSQLTimeForm)
    If bmFirstTime Then
        mBasicStart
        bmFirstTime = False
        If bmCreateUniqueFilenames Then
            mCreateUniqueFilename
        End If
    End If
End Sub
Public Function CreateEnts(Optional slStatus As StatusEnum = Successful) As Boolean
    Dim blRet As Boolean
    Dim slOriginalType As String
    Dim blOriginalIngestStatus As Boolean
    Dim slOriginalFileName As String
    
    blRet = mCreateEnts(slStatus, bmCopyForWeb)
    'no agreement?  then this can't go to web
    If tmENT.lAttCode > 0 And blRet Then
        If bmCopyForWeb Then
            'change type,don't compare treat as import, and change the ingested bucket to the send bucket
            slOriginalType = tmENT.sType
            blOriginalIngestStatus = bmIngestDeterminesStatus
            bmIngestDeterminesStatus = False
            tmENT.sType = mDecodeType(senttoweb3rdpartyPosted)
            'unique filename for easy update later.  Only create once
            If Len(smUniqueWebName) = 0 Then
                mCreateUniqueFilename
            End If
            slOriginalFileName = smFileName
            smFileName = smUniqueWebName
            If Not (myRsBuckets.EOF And myRsBuckets.BOF) Then
                myRsBuckets.MoveFirst
                Do While Not myRsBuckets.EOF
                    myRsBuckets("SentOrReceived") = myRsBuckets("Ingested")
                    myRsBuckets("Ingested") = 0
                    myRsBuckets.MoveNext
                Loop
            End If
            blRet = mCreateEnts(Incomplete, False)
            tmENT.sType = slOriginalType
            bmIngestDeterminesStatus = blOriginalIngestStatus
            smFileName = slOriginalFileName
        End If
    End If
    CreateEnts = blRet
End Function


Public Function UpdateAsSuccessful(Optional blIsNOTIn As Boolean = False, Optional slCommaSeparatedExceptionsList As String = "", Optional elExceptionsChoice As UpdateExceptionEnum = None) As Boolean
    'this works only with user choosing 'ReturnEntCode=true'.  All those ents saved get set to 'S', except for the following special case(s)
    'exceptionList:  "3,4,5"  if this is formatted badly, nothing gets updated!
    'blIsNOTIn.  Is the list to be a 'not in (xxx)' or a 'in (xxx)'?  if True, then what you sent will not be updated to 'S'; if false, only those in the list will be updated to 'S'
    'elExceptionChoice  vehicle? station? etc.
    Dim blRet As Boolean
    Dim slList As String
    Dim ilPreviousValue As Integer
    Dim ilPos As Integer
    Dim slSecondIn As String
    
    blRet = True
    If blIsNOTIn Then
        slSecondIn = " NOT in "
    Else
        slSecondIn = " in "
    End If
    If elExceptionsChoice <> UpdateExceptionEnum.None Then
        blRet = False
        If Len(slCommaSeparatedExceptionsList) > 0 Then
            ilPos = InStr(1, slCommaSeparatedExceptionsList, ",")
            If ilPos = 0 Then
                blRet = True
            'if comma, make sure not at start or end
            ElseIf Not (ilPos = 1 Or ilPos = Len(slCommaSeparatedExceptionsList)) Then
                blRet = True
            End If
        End If
    End If
    slList = mCleanReturnCodes()
    If blRet Then
        If Len(slList) > 0 Then
            SQLQuery = "Update ENT set entStatus = 'S',entErrorMsg = '' where entCode in (" & slList & ") "
            Select Case elExceptionsChoice
                Case UpdateExceptionEnum.StationCodes
                    SQLQuery = SQLQuery & " AND entShttCode " & slSecondIn & "(" & slCommaSeparatedExceptionsList & ")"
                Case UpdateExceptionEnum.AgreementCodes
                    SQLQuery = SQLQuery & " AND entAttCode " & slSecondIn & "(" & slCommaSeparatedExceptionsList & ")"
                Case UpdateExceptionEnum.VehicleCodes
                    SQLQuery = SQLQuery & " AND entVefCode " & slSecondIn & "(" & slCommaSeparatedExceptionsList & ")"
                Case UpdateExceptionEnum.EntCode
                    SQLQuery = SQLQuery & " AND entCode " & slSecondIn & "(" & slCommaSeparatedExceptionsList & ")"
            End Select
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand:
                ilPreviousValue = igShowMsgBox
                igShowMsgBox = 0
                gHandleError ErrorLog, "UpdateAsSuccessful"
                igShowMsgBox = ilPreviousValue
                smErrorMessage = smErrorMessage & " Error in 'UpdateAsSuccessful':" & SQLQuery & vbCrLf
                UpdateAsSuccessful = False
                Exit Function
            End If
        End If
    Else
        smErrorMessage = smErrorMessage & " Problem in UpdateAsSuccessful's exception list:" & slCommaSeparatedExceptionsList & vbCrLf
    End If
Cleanup:
    UpdateAsSuccessful = blRet
    Exit Function
ErrHand:
    ilPreviousValue = igShowMsgBox
    igShowMsgBox = 0
    gHandleError ErrorLog, "UpdateAsSuccessful"
    igShowMsgBox = ilPreviousValue
    smErrorMessage = smErrorMessage & " Error in 'UpdateAsSuccessful':" & SQLQuery & vbCrLf
    blRet = False
    GoTo Cleanup
End Function
'Public Function UpdateIncompleteByVehicle(ilVefCode As Integer, slFeedDateStart As String, slFeedDateEnd As String, elStatus As StatusEnum) As Boolean
' commented out because missing 7510, writing entErrorMsg
'    Dim blRet As Boolean
'    Dim ilPreviousValue As Integer
'    Dim slStatus As String
'    'update ent set entstatus = 'S' where entStatus = 'I' and entType = 'X' and ent3rdParty = 'E' and entUstCode = 1 and entVefCode = 3
'    ' and entFeedDate >= '2015-12-18' and entFeedDate <= '2015-12-19'
'    blRet = True
'    slFeedDateStart = Format$(slFeedDateStart, sgSQLDateForm)
'    slFeedDateEnd = Format$(slFeedDateEnd, sgSQLDateForm)
'    If elStatus = Successful Then
'        slStatus = "S"
'    ElseIf elStatus = NotSent Then
'        slStatus = "N"
'    Else
'        slStatus = "E"
'    End If
'    With tmENT
'        SQLQuery = "update ent set entstatus = '" & slStatus & "' where entStatus = 'I' and entType = '" & .sType & "' and ent3rdParty = '" & .s3rdParty & "' and entUstCode = " & .iUstCode & " and entVefCode = " & ilVefCode & " and entFeedDate >= '" & slFeedDateStart & "' and entFeedDate <= '" & slFeedDateEnd & "'"
'    End With
'    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'        GoSub ErrHand:
'    End If
'
'CLEANUP:
'    UpdateIncompleteByVehicle = blRet
'    Exit Function
'ErrHand:
'    ilPreviousValue = igShowMsgBox
'    igShowMsgBox = 0
'    gHandleError ErrorLog, "UpdateIncompleteByVehicle"
'    igShowMsgBox = ilPreviousValue
'    smErrorMessage = smErrorMessage & " Error in 'UpdateIncompleteByVehicle':" & SQLQuery & vbCrLf
'    blRet = False
'    GoTo CLEANUP
'End Function
Public Function UpdateIncompleteByFilename(elStatus As StatusEnum, Optional slUseThisFileName As String = "", Optional slUpdateToThisFileName As String = "", Optional slOnlyThese As String = "") As Boolean
    'the filename mut be unique to that export...that would be exportWeb.
    'slUseThisFileName... filename to be tested.  slUpdateToThisFileName...change to this filename
    'slOnlyThese.  entattcodes  "4,12,13,"
    Dim blRet As Boolean
    Dim slStatus As String
    Dim slFileName As String
    Dim ilPreviousValue As Integer
    Dim slAddToUpdate As String
    Dim slAddToWhere
'   update ent set entstatus = 'S' where entstatus = 'I' AND entustcode = 27 and entType = 'E' and entFilename = 'WebText.txt'

    blRet = True
    slAddToUpdate = ""
    If Len(slOnlyThese) > 1 Then
        slAddToWhere = mLoseLastLetterIfComma(slOnlyThese)
        slAddToWhere = " AND entAttCode in (" & slAddToWhere & ")"
    End If
    If Len(slUpdateToThisFileName) > 0 Then
        slAddToUpdate = "',entFileName = '" & slUpdateToThisFileName
    End If
    If elStatus = Successful Then
        slStatus = "S"
    ElseIf elStatus = NotSent Then
        slStatus = "N"
    Else
        slStatus = "E"
        '7510
        slAddToUpdate = slAddToUpdate & "',entErrorMsg = '" & mGetEntErrorMessage()
    End If
    If Len(slUseThisFileName) > 0 Then
        slFileName = slUseThisFileName
    Else
        If bmCopyForWeb Then
            slFileName = smUniqueWebName
            tmENT.sType = mDecodeType(senttoweb3rdpartyPosted)
        Else
            slFileName = smFileName
        End If
    End If
    If Len(slFileName) > 0 Then
        With tmENT
            SQLQuery = "update ent set entstatus = '" & slStatus & slAddToUpdate & "' where entStatus = 'I' and entType = '" & .sType & "' and entUstCode = " & .iUstCode & " and entFilename = '" & slFileName & "'" & slAddToWhere
        End With
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub ErrHand:
            ilPreviousValue = igShowMsgBox
            igShowMsgBox = 0
            gHandleError ErrorLog, "UpdateIncompleteByFilename"
            igShowMsgBox = ilPreviousValue
            smErrorMessage = smErrorMessage & " Error in 'UpdateIncompleteByFilename':" & SQLQuery & vbCrLf
            UpdateIncompleteByFilename = False
            If bmCreateUniqueFilenames Then
                mCreateUniqueFilename
            End If
            If bmCopyForWeb Then
                tmENT.sType = mDecodeType(Importposted3rdparty)
            End If
        End If
        'reverse the 'onlyThese'. Note that status E and S are reversed. I get rid of in, as above already done.
        If Len(slOnlyThese) > 1 Then
            slAddToUpdate = ""
'            If Len(slOnlyThese) > 1 Then
'                slAddToWhere = mLoseLastLetterIfComma(slOnlyThese)
'                slAddToWhere = " AND entAttCode not in (" & slAddToWhere & ")"
'            End If
            slAddToWhere = ""
            If Len(slUpdateToThisFileName) > 0 Then
                slAddToUpdate = "',entFileName = '" & slUpdateToThisFileName
            End If
            If elStatus = EntError Then
                slStatus = "S"
            ElseIf elStatus = NotSent Then
                slStatus = "N"
            Else
                slStatus = "E"
                '7510
                slAddToUpdate = slAddToUpdate & "',entErrorMsg = '" & mGetEntErrorMessage()
            End If
            If Len(slUseThisFileName) > 0 Then
                slFileName = slUseThisFileName
            Else
                If bmCopyForWeb Then
                    slFileName = smUniqueWebName
                    tmENT.sType = mDecodeType(senttoweb3rdpartyPosted)
                Else
                    slFileName = smFileName
                End If
            End If
            With tmENT
                SQLQuery = "update ent set entstatus = '" & slStatus & slAddToUpdate & "' where entStatus = 'I' and entType = '" & .sType & "' and entUstCode = " & .iUstCode & " and entFilename = '" & slFileName & "'" & slAddToWhere
            End With
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand:
                ilPreviousValue = igShowMsgBox
                igShowMsgBox = 0
                gHandleError ErrorLog, "UpdateIncompleteByFilename"
                igShowMsgBox = ilPreviousValue
                smErrorMessage = smErrorMessage & " Error in 'UpdateIncompleteByFilename':" & SQLQuery & vbCrLf
                UpdateIncompleteByFilename = False
                If bmCreateUniqueFilenames Then
                    mCreateUniqueFilename
                End If
                If bmCopyForWeb Then
                    tmENT.sType = mDecodeType(Importposted3rdparty)
                End If
            End If
        End If
    Else
        blRet = False
        smErrorMessage = smErrorMessage & "Problem in updateIncompleteByFilename:  no filename" & vbCrLf
    End If
Cleanup:
    UpdateIncompleteByFilename = blRet
    If bmCreateUniqueFilenames Then
        mCreateUniqueFilename
    End If
    If bmCopyForWeb Then
        tmENT.sType = mDecodeType(Importposted3rdparty)
    End If
    Exit Function
ErrHand:
    ilPreviousValue = igShowMsgBox
    igShowMsgBox = 0
    gHandleError ErrorLog, "UpdateIncompleteByFilename"
    igShowMsgBox = ilPreviousValue
    smErrorMessage = smErrorMessage & " Error in 'UpdateIncompleteByFilename':" & SQLQuery & vbCrLf
    blRet = False
    GoTo Cleanup
End Function
Public Function SetThirdPartyByHierarchy() As Boolean
    'for web imports and exports, where the web works as a 'middleman'
    'returns true if no error
    Dim blRet As Boolean
    Dim slChoice As String
    Dim slNewPart As Vendors
    Dim ilChoice As Integer
    Dim myRs As ADODB.Recordset
    
On Error GoTo ERRORBOX
    blRet = True
    slNewPart = Web
    
    '7701 get first one only.  Since I'm descending, gets audio before logs.  I left as do loop in case we want to do something more in future
    SQLQuery = "Select vatWvtVendorId as ID from VAT_Vendor_Agreement where vatattcode = " & tmENT.lAttCode & " ORDER BY VATWVTVendorID desc"
    Set myRs = gSQLSelectCall(SQLQuery)
    Do While Not myRs.EOF
        ilChoice = myRs!ID
        slNewPart = ilChoice
        Exit Do
    Loop
    
    
    
    
'    ilChoice = gIfNullInteger(myRs!vatWvtIdCodeAudio)
'    slNewPart = ilChoice
'    If ilChoice = 0 Then
'        ilChoice = gIfNullInteger(myRs!vatWvtIdCodeLog)
'        slNewPart = ilChoice
'    End If
'    slChoice = Trim$(myRs!attAudioDelivery)
'    If Len(slChoice) > 0 And slChoice <> "N" Then
'        Select Case slChoice
'            Case "X", "B"
'                slNewPart = Vendors.XDS_Break
'            Case "W"
'                slNewPart = Vendors.Wegener_Compel
'            Case "P"
'                slNewPart = Vendors.Wegener_IPump
'            Case "I"
'                slNewPart = Vendors.iDc
'       End Select
'    'AudioDelivery not set
'    Else
'        If Trim$(myRs!attExportToCBS) = "Y" Then
'            slNewPart = Vendors.cBs
'        ElseIf Trim$(myRs!attExportToClearCh) = "Y" Then
'            slNewPart = Vendors.ClearChannel
'        ElseIf Trim$(myRs!attWebInterface) = "C" Then
'            slNewPart = Vendors.cumulus
'        ElseIf Trim$(myRs!attExportToJelli) = "Y" Then
'            slNewPart = Vendors.Jelli
'        ElseIf Trim$(myRs!attExportToMarketron) = "Y" Then
'            slNewPart = Vendors.NetworkConnect
''        ElseIf Trim$(myRs!attExportToUnivision) = "Y" Then
''            slNewPart = Univision
'        End If
'    End If
Cleanup:
    tmENT.s3rdParty = mDecodeParty(slNewPart)
    SetThirdPartyByHierarchy = blRet
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
        End If
        Set myRs = Nothing
    End If
    Exit Function
ERRORBOX:
    smErrorMessage = smErrorMessage & " Problem in SetThirdPartyByHierarchy:" & Err.Description & vbCrLf
    blRet = False
    GoTo Cleanup
End Function

'Public Function UpdateEachByFeedDate(slFeedDate As String, llAttCode As Long, ilIngest As Integer, Optional llGsfCode As Long = 0) As Boolean
' commented out because missing 7510, writing entErrorMsg
'    Dim blRet As Boolean
'    Dim slList As String
'    Dim ilPreviousValue As Integer
'    Dim ilSpotCount As Integer
'    Dim llEnt As Long
'    Dim slNewValue As String
'
'    'select top 1 entcode,entspotcount from ent where entType = 'S'  AND entAttCode = 199 AND entFeeddate = '2015-02-15'
'    'AND entUstCode = 27  AND entGsfCode = 80 order by entcode desc
'
'    blRet = True
'    With tmENT
'        SQLQuery = "select top 1 entCode, entSpotCount from ent where entType = '" & .sType & "' and entAttCode = " & llAttCode & " AND entFeeddate = '" & Format(slFeedDate, sgSQLDateForm) & "' "
'        SQLQuery = SQLQuery & " AND entUstCode = " & .iUstCode & " AND entGsfCode = " & llGsfCode & " ORDER BY entCode desc"
'On Error GoTo ERRORSQL:
'        Set rst = gSQLSelectCall(SQLQuery)
'        If Not rst.EOF Then
'            llEnt = rst!EntCode
'            ilSpotCount = rst!entSpotCount
'            If ilSpotCount = ilIngest Then
'                slNewValue = "S"
'            Else
'                slNewValue = "E"
'            End If
'On Error GoTo 0
'            SQLQuery = "update ENT set entStatus = '" & slNewValue & "' , entIngestedCount = " & ilIngest & " WHERE entcode = " & llEnt
'            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                GoSub ErrHand:
'            End If
'        End If
'    End With
'CLEANUP:
'    UpdateEachByFeedDate = blRet
'    Exit Function
'ErrHand:
'    ilPreviousValue = igShowMsgBox
'    igShowMsgBox = 0
'    gHandleError ErrorLog, "UpdateEachByFeedDate"
'    igShowMsgBox = ilPreviousValue
'    smErrorMessage = smErrorMessage & " Error in 'UpdateEachByFeedDate':" & SQLQuery & vbCrLf
'    blRet = False
'    GoTo CLEANUP
'    Exit Function
'ERRORSQL:
'    ilPreviousValue = igShowMsgBox
'    igShowMsgBox = 0
'    gHandleError ErrorLog, "UpdateEachByFeedDate"
'    igShowMsgBox = ilPreviousValue
'    smErrorMessage = smErrorMessage & " Error in 'UpdateEachByFeedDate':" & Err.Description & vbCrLf
'    blRet = False
'    GoTo CLEANUP
'
'End Function
'Public Function UpdateEachByEntCode(llEntCode As Long, ilIngest As Integer) As Boolean
'    Dim blRet As Boolean
'    Dim slList As String
'    Dim ilPreviousValue As Integer
'
'    blRet = True
'    With tmENT
'        SQLQuery = "Update ENT set entStatus = 'S', entIngestedCount = " & ilIngest & " WHERE entCode = " & llEntCode & " AND " & ilIngest & " = entSpotCount"
'        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'            GoSub ErrHand:
'        End If
'        SQLQuery = "Update ENT set entStatus = 'E', entIngestedCount = " & ilIngest & " WHERE entCode = " & llEntCode & " AND " & ilIngest & " <> entSpotCount"
'        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'            GoSub ErrHand:
'        End If
'    End With
'Cleanup:
'    'smReturnCodes = ""
'    UpdateEachByEntCode = blRet
'    Exit Function
'ErrHand:
'    ilPreviousValue = igShowMsgBox
'    igShowMsgBox = 0
'    gHandleError ErrorLog, "UpdateEachByEntCode"
'    igShowMsgBox = ilPreviousValue
'    smErrorMessage = smErrorMessage & " Error in 'UpdateEachByEntCode':" & SQLQuery & vbCrLf
'    blRet = False
'    GoTo Cleanup
'End Function
Public Sub ClearWhenDontSend()
    mClearFeedBucket
End Sub
Public Function MulticastCopy(slAtts As String, slshtts As String) As Boolean
    Dim slArray() As String
    Dim c As Integer
    Dim ilUpper As Integer
    
    bmIsMulticast = False
On Error GoTo ERRORBOX
    slArray = Split(slAtts, ",")
    ilUpper = UBound(slArray)
    ReDim lmAtts(ilUpper)
    For c = 0 To ilUpper
        lmAtts(c) = slArray(c)
    Next c
    slArray = Split(slshtts, ",")
    If ilUpper = UBound(slArray) Then
        ReDim imShtts(ilUpper)
        For c = 0 To ilUpper
            imShtts(c) = slArray(c)
        Next c
        bmIsMulticast = True
    Else
        smErrorMessage = smErrorMessage & " Problem in 'MulticastCopy': upperbounds didn't match" & vbCrLf
    End If
Cleanup:
    MulticastCopy = bmIsMulticast
   Exit Function
ERRORBOX:
    bmIsMulticast = False
    smErrorMessage = smErrorMessage & " Error in 'MulticastCopy': " & Err.Description & vbCrLf
    GoTo Cleanup
End Function
Public Sub CleanEnts()
    Dim ilPreviousValue As Integer
    Dim dlDeleteBefore As Date
    Dim slDeleteBefore As String
     
    dlDeleteBefore = DateAdd("M", -DELETEBEFOREMONTHS, Now())
    slDeleteBefore = Format(dlDeleteBefore, sgShowDateForm)
    slDeleteBefore = gObtainStartStd(slDeleteBefore)
    slDeleteBefore = Format(slDeleteBefore, sgSQLDateForm)
    SQLQuery = "delete ent where entEnteredDate < '" & slDeleteBefore & "'"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        ilPreviousValue = igShowMsgBox
        igShowMsgBox = 0
        gHandleError ErrorLog, "CleanEnts"
        igShowMsgBox = ilPreviousValue
        smErrorMessage = smErrorMessage & " Error in 'CleanEnts':" & SQLQuery & vbCrLf
    End If
    Exit Sub
ErrHand:
    ilPreviousValue = igShowMsgBox
    igShowMsgBox = 0
    gHandleError ErrorLog, "CleanEnts"
    igShowMsgBox = ilPreviousValue
    smErrorMessage = smErrorMessage & " Error in 'CleanEnts':" & SQLQuery & vbCrLf
End Sub
Private Function mCreateEnts(slStatus As StatusEnum, blDontClearAfterSend As Boolean) As Boolean
    Dim ilPreviousMsgShow As Integer
    Dim blRet As Boolean
    Dim slOriginalStatus As String
    Dim llReturnEnt As Long
    Dim ilRealShtt As Integer
    Dim llRealAtt As Long
    Dim j As Integer
    
    
    ilPreviousMsgShow = igShowMsgBox
    igShowMsgBox = 0
    blRet = True
'    'temporarily block export to web
'    If tmENT.sType = mDecodeType(SentUnpostedToWeb) Then
'        GoTo cleanup
'    End If
    If Len(smReturnCodes) = 0 Then
        smReturnCodes = ","
    End If
On Error GoTo ERRORBOX
    tmENT.sFileName = smFileName
    slOriginalStatus = "S"
    If slStatus = Incomplete Then
        slOriginalStatus = "I"
    ElseIf slStatus = EntError Then
        slOriginalStatus = "E"
    ElseIf slStatus = NotSent Then
        slOriginalStatus = "N"
    End If
    'don't get returncode
    llReturnEnt = -1
    'this won't be used much anymore.  Assumes we write as "E", and then uses 'updateAsSuccessful' to change to 'S'
    If bmReturnEntCode Then
        llReturnEnt = 0
        slOriginalStatus = "E"
    End If
    myRsBuckets.Filter = adFilterNone
    If Not myRsBuckets.EOF Then
        myRsBuckets.MoveFirst
        Do While Not myRsBuckets.EOF
            '7509
           ' If myRsBuckets("SentOrReceived") > 0 Then
            If myRsBuckets("SentOrReceived") > 0 Or myRsBuckets("Deleted") > 0 Then
                With tmENT
                    .lgsfCode = myRsBuckets!Game
                    .iAstCount = myRsBuckets!Asts
                    .iIngestedCount = myRsBuckets!Ingested
                    .iSpotCount = myRsBuckets!SentOrReceived
                    '10099 I'm tracking regional or all if xds, so don't overwrite
                    If Not (.sType = "E" And .s3rdParty = "X") Then
                        .iBonusCount = myRsBuckets!Bonus
                    End If
                    .iMGCount = myRsBuckets!MakeGood
                    .iReplaceCount = myRsBuckets!Replacement
                    .sFeedDate = myRsBuckets!FeedDate
                    .iDeleteCount = myRsBuckets!Deleted
                    'in case changed previously by bmIngestDeterminesStatus
                    .sStatus = slOriginalStatus
                    If bmIngestDeterminesStatus Then
                        ' use spot count = ingest count, but if 0 something wrong. Also, status needs to be 'S'
                        If .iIngestedCount <> .iSpotCount And .iSpotCount > 0 And .sStatus = "S" Then
                            .sStatus = "E"
                        End If
                    End If
                    '7510
                    If .sStatus = "E" Then
                       .sErrorMsg = mGetEntErrorMessage()
                    Else
                        .sErrorMsg = ""
                    End If
                    If Not gENTAddNew(tmENT, ErrorLog, llReturnEnt) Then
                        blRet = False
                        smErrorMessage = smErrorMessage & " Error in 'CreateENTs': could not add for attcode " & tmENT.lAttCode & " and feed date " & tmENT.sFeedDate & vbCrLf
                    ElseIf bmReturnEntCode Then
                        If llReturnEnt > 0 Then
                            'keep return value for updating later!
                            If InStr(1, "," & llReturnEnt & ",") = 0 Then
                                smReturnCodes = smReturnCodes & llReturnEnt & ","
                            End If
                        Else
                            smErrorMessage = smErrorMessage & " Error in 'CreateEnts': did not return ent code" & vbCrLf
                        End If
                        'reset
                        llReturnEnt = 0
                    End If
                    If bmIsMulticast Then
                        llRealAtt = tmENT.lAttCode
                        ilRealShtt = tmENT.iShttCode
On Error GoTo ERRMULTI
                        For j = 0 To UBound(lmAtts) Step 1
                            tmENT.lAttCode = lmAtts(j)
                            tmENT.iShttCode = imShtts(j)
                            If Not gENTAddNew(tmENT, ErrorLog, llReturnEnt) Then
                                blRet = False
                                smErrorMessage = smErrorMessage & " Error in 'CreateENTs': could not add for attcode " & tmENT.lAttCode & " and feed date " & tmENT.sFeedDate & vbCrLf
                            ElseIf bmReturnEntCode Then
                                If llReturnEnt > 0 Then
                                    If InStr(1, "," & llReturnEnt & ",") = 0 Then
                                        smReturnCodes = smReturnCodes & llReturnEnt & ","
                                    End If
                                Else
                                    smErrorMessage = smErrorMessage & " Error in 'CreateEnts': did not return ent code" & vbCrLf
                                End If
                                'reset
                                llReturnEnt = 0
                            End If ' which way to add
                        Next j 'each multi #
                        tmENT.lAttCode = llRealAtt
                        tmENT.iShttCode = ilRealShtt
                    End If 'have multi
                End With
            End If
            myRsBuckets.MoveNext
        Loop
    End If
Cleanup:
    igShowMsgBox = ilPreviousMsgShow
    If Not blDontClearAfterSend Then
        mClearFeedBucket
    End If
    mCreateEnts = blRet
    Exit Function
ERRORBOX:
    'quit or resume next?
    smErrorMessage = smErrorMessage & " Error in 'CreateENTS': " & Err.Description & vbCrLf
    blRet = False
    GoTo Cleanup
    Exit Function
ERRMULTI:
    smErrorMessage = smErrorMessage & " Error in 'CreateENTS-creating multicasts': " & Err.Description & vbCrLf
    blRet = False
   ' Resume Next
End Function
Private Function mLoseLastLetterIfComma(slInput As String) As String
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
    mLoseLastLetterIfComma = slNewString
End Function
Private Sub mClearFeedBucket()
        
    Set myRsBuckets = mPrepRecordset()
    If bmIsMulticast Then
        Erase lmAtts
        Erase imShtts
        bmIsMulticast = False
    End If
End Sub
Private Sub mCreateUniqueFilename()
    Dim slTemp As String
    
    slTemp = gGetComputerName()
    If slTemp = "N/A" Then
        slTemp = "Unknown"
    End If
    slTemp = slTemp & "_" & sgUserName & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhmmss") & ".txt"
    'it's possible the filename is being set before type and 3rd party...but that won't really hurt any.
    If bmCopyForWeb Then
        smUniqueWebName = tmENT.sType & "_" & tmENT.s3rdParty & "_" & slTemp
    Else
        smFileName = tmENT.sType & "_" & tmENT.s3rdParty & "_" & slTemp
    End If
End Sub
Private Function mCleanReturnCodes(Optional blErase As Boolean = True) As String
    Dim slRet As String
    
    If Len(smReturnCodes) > 1 Then
        slRet = Mid(smReturnCodes, 2)
        slRet = mLoseLastLetterIfComma(slRet)
    End If
    If blErase Then
        smReturnCodes = ""
    End If
    mCleanReturnCodes = slRet
End Function
Private Function mDecodeProcessing(elProcess As ProcessingEnum) As String
    Dim slRet As String
    Select Case elProcess
        Case ProcessingEnum.Asts
            slRet = "Asts"
        Case ProcessingEnum.Bonus
            slRet = "Bonus"
        Case ProcessingEnum.Ingested
            slRet = "Ingested"
        Case ProcessingEnum.MakeGood
            slRet = "MakeGood"
        Case ProcessingEnum.Replacement
            slRet = "Replacement"
        Case ProcessingEnum.SentOrReceived
            slRet = "SentOrReceived"
        Case ProcessingEnum.Deleted
            slRet = "Deleted"
    End Select
    mDecodeProcessing = slRet
End Function
Private Function mDecodeType(elType As TypesEnum) As String
    Dim slRet As String
    
    Select Case elType
        Case TypesEnum.Sentunpostedtoweb
            slRet = "S"
        Case TypesEnum.Receivedpostedfromweb
            slRet = "R"
        Case TypesEnum.Exportunposted3rdparty
            slRet = "E"
        Case TypesEnum.Importposted3rdparty
            slRet = "I"
        Case TypesEnum.senttoweb3rdpartyPosted
            slRet = "P"
        Case TypesEnum.Forwardtoweb3rdparty
            slRet = "F"
        Case TypesEnum.importedinToweb3party
            slRet = "T"
    End Select
    mDecodeType = slRet
End Function
Private Function mEncodeType(slType As String) As TypesEnum
    Dim elRet As TypesEnum
    
    Select Case slType
        Case "S"
            elRet = TypesEnum.Sentunpostedtoweb
        Case "R"
            elRet = TypesEnum.Receivedpostedfromweb
        Case "E"
            elRet = TypesEnum.Exportunposted3rdparty
        Case "I"
            elRet = TypesEnum.Importposted3rdparty
        Case "P"
            elRet = TypesEnum.senttoweb3rdpartyPosted
        Case "F"
            elRet = TypesEnum.Forwardtoweb3rdparty
        Case "T"
            elRet = TypesEnum.importedinToweb3party
    End Select
    mEncodeType = elRet
End Function
Private Function mDecodeParty(elSource As Vendors) As String
    Dim slRet As String
    
    Select Case elSource
        Case Vendors.Web
            slRet = "W"
        Case Vendors.NetworkConnect
            slRet = "M"
        Case Vendors.XDS_Break, Vendors.XDS_ISCI
            slRet = "X"
        Case Vendors.Wegener_Compel
            slRet = "C"
        Case Vendors.Wegener_IPump
            slRet = "I"
        Case Vendors.cBs
            slRet = "B"
        Case Vendors.stratus
            slRet = "Q"
        Case Vendors.iHeart
            slRet = "H"
        Case Vendors.iDc
            slRet = "D"
        Case Vendors.Jelli
            slRet = "J"
        '7701
        Case Vendors.WideOrbit
            slRet = "O"
'        Case Vendors.Univision
'            slRet = "U"
        '8255
        Case Vendors.MrMaster
            slRet = "A"
        Case Vendors.RadioTraffic
            slRet = "T"
        Case Vendors.RCS
            slRet = "R"
        Case Vendors.Synchronicity
            slRet = "S"
        Case Vendors.BSI
            slRet = "L"
        Case Vendors.AnalyticOwl
            slRet = "Y"
        Case Vendors.LeadsRX
            slRet = "E"
        Case Vendors.RadioWorkflow
            slRet = "W"
    End Select
    mDecodeParty = slRet
End Function
Private Function mEncodeParty(slSource As String) As Vendors
    Dim elRet As Vendors
    'xds defaults to 'break'
    Select Case slSource
        Case "W"
            elRet = Vendors.Web
        Case "M"
            elRet = Vendors.NetworkConnect
        Case "X"
            elRet = Vendors.XDS_Break
        Case "C"
            elRet = Vendors.Wegener_Compel
        Case "I"
            elRet = Vendors.Wegener_IPump
        Case "B"
            elRet = Vendors.cBs
        Case "Q"
            elRet = Vendors.stratus
        Case "H"
            elRet = Vendors.iHeart
        Case "D"
            elRet = Vendors.iDc
        Case "J"
            elRet = Vendors.Jelli
        '7701
        Case "O"
            elRet = Vendors.WideOrbit
        '8255
        Case "A"
            elRet = Vendors.MrMaster
        Case "T"
            elRet = Vendors.RadioTraffic
        Case "R"
            elRet = Vendors.RCS
        Case "S"
            elRet = Vendors.Synchronicity
        Case "L"
            elRet = Vendors.BSI
        Case "Y"
            elRet = Vendors.AnalyticOwl
        Case "E"
            elRet = Vendors.LeadsRX
        Case "W"
            elRet = Vendors.RadioWorkflow
'        Case "U"
'            elRet = Univision
    End Select
    mEncodeParty = elRet
End Function

Private Sub mBasicStart()
    mSpecialRules
    Set myRsBuckets = mPrepRecordset()
    'EntErrorMsg = ""
End Sub
Private Sub mSpecialRules()
    ' all imports and web - ingestdeterminesstatus
    Dim elType As TypesEnum
  '  Dim elParty As ThirdPartyEnum
    
    bmIngestDeterminesStatus = False
    elType = mEncodeType(tmENT.sType)
   ' elParty = mEncodeParty(tmENT.s3rdParty)
    
    Select Case elType
        Case TypesEnum.Importposted3rdparty
            bmIngestDeterminesStatus = True
        Case TypesEnum.Receivedpostedfromweb
            bmIngestDeterminesStatus = True
    End Select
'    Select Case elParty
'        Case ThirdPartyEnum.Web
'
'    End Select
End Sub
Private Function mPrepRecordset() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "FeedDate", adChar, 10
            .Append "Game", adInteger
            .Append "SentOrReceived", adInteger
            .Append "Asts", adInteger
            .Append "Ingested", adInteger
            .Append "MakeGood", adInteger
            .Append "Bonus", adInteger
            .Append "Replacement", adInteger
            .Append "Deleted", adInteger
        End With
    myRs.Open
    Set mPrepRecordset = myRs
End Function
Private Function mGetEntErrorMessage() As String
    Dim slRet As String
    
    slRet = ""
    If Len(EntErrorMsg) > 0 Then
        slRet = EntErrorMsg
    ElseIf Len(ErrorLog) > 0 Then
        slRet = ErrorLog
    End If
    mGetEntErrorMessage = slRet
End Function
