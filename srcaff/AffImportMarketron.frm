VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmImportMarketron 
   Caption         =   "Import Marketron"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   6195
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5145
      Top             =   4215
   End
   Begin VB.PictureBox pbcTextWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5700
      ScaleHeight     =   225
      ScaleWidth      =   1005
      TabIndex        =   4
      Top             =   3795
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.ListBox lbcMsg 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   1395
      Width           =   5790
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5790
      Top             =   4305
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4890
      FormDesignWidth =   6195
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Import"
      Height          =   375
      Left            =   1125
      TabIndex        =   2
      Top             =   4380
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3150
      TabIndex        =   3
      Top             =   4380
      Width           =   1575
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5790
   End
   Begin VB.Menu mnuGuide 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuRemote 
         Caption         =   "Don't Download"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Don't Import"
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug log of file count"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmImportMarketron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private imExporting As Integer
Private imTerminate As Integer
Private oMyFileObj As FileSystemObject
Private lmCurrentRow As Long
Private smLines As String
Private smIniPath As String
Private tmImportSpot() As MARKETRONSPOTINFO
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private hmAst As Integer
Private Const MESSAGEBLACK As Long = 0
Private Const MESSAGERED As Long = 255
Private Const MESSAGEGREEN As Long = 39680
Private Const COMPLETEDFOLDER As String = "Completed\"
Private Const COMPLETEDDAYS As Integer = 365
Private Const ROWNAME As String = "<AffidavitSpot>"
Private lmMaxWidth As Long
'5217
'Private Const LOGFILE As String = "MarketronImportLog.txt"
Private Const FORMNAME As String = "FrmImportMarketron"
Private Const ERRORSQL As Integer = 9101
Private Const FILEFACTS As String = "MarketronFacts"
Private Const FILEERROR As String = "MarketronImport"
Private Const FILEDEBUG As String = "MarketronDebug"
Private Const FILEREIMPORT As String = "ReImportMarketron"
'4922.
Private lmUnMatchedSpots As Dictionary
'6158
Private rsMG As Recordset
Private rsMarket As Recordset
Private imMarket As Integer
Private Enum MGUpdates
    ADDMISSED = 0
    ADDMAKEGOOD = 1
    UPDATEMISSED = 2
    UPDATEMAKEGOOD = 3
    DELETEEITHER = 4
End Enum
Private Const NODATE As String = "1970-01-01"
'5602
Private myImport As CMarketron

'replaces logfile
Private smPathForgLogMsg As String
Private myErrors As CLogger
'7458
Private myEnt As CENThelper
'7266
Dim myAligner As cMarketronAligner
Dim smCurrentAgreementInfo As String

Private Property Get bmMoreRows() As Boolean
    'find next row name
    Dim llPos As Long
    
    If lmCurrentRow = 0 Then
        lmCurrentRow = 1
    End If
    llPos = InStr(lmCurrentRow + 1, smLines, ROWNAME)
    If llPos > 0 Then
        lmCurrentRow = llPos
        bmMoreRows = True
    Else
        lmCurrentRow = 0
        bmMoreRows = False
    End If
End Property
Private Function mTestFileName(slName As String) As Boolean
    mTestFileName = True
    '6602 remove
'    If InStr(1, slName, "-") < 2 Then
'        mTestFileName = False
'        Exit Function
'    End If
    If InStr(1, slName, ".txt") = 0 Then
        mTestFileName = False
    End If
End Function
Private Function mMarketronFolder() As String
    Dim slFolderPath As String
    Dim slNestedPath As String
    Dim myFile As file
    Dim dlDeleteDate As Date
    Dim slFolderName As String
    Dim blFirstTime As Boolean
    
    blFirstTime = False
    slFolderName = mSafeFileName(sgClientName)
    dlDeleteDate = gNow()
    slFolderPath = oMyFileObj.BuildPath(sgImportDirectory, slFolderName & "\")
    If Not oMyFileObj.FolderExists(slFolderPath) Then
        oMyFileObj.CreateFolder slFolderPath
    End If
    slNestedPath = slFolderPath & COMPLETEDFOLDER
    If Not oMyFileObj.FolderExists(slNestedPath) Then
        oMyFileObj.CreateFolder (slNestedPath)
    End If
    For Each myFile In oMyFileObj.GetFolder(slNestedPath).Files
    '4922 don't delete 'unread files'.  Write out problem files.
'        If DateDiff("d", myFile.DateCreated, dlDeleteDate) > COMPLETEDDAYS Then
'            myFile.Delete
'        End If
        If InStr(1, myFile.Name, "UNREAD", vbTextCompare) > 0 Then
'            '7808 don't show unread
'            If blFirstTime = False Then
'                blFirstTime = True
'                'mSetResults "There are unread files that need to be processed!  Please contact Counterpoint with " & LOGFILE & ".", MESSAGEBLACK
'                mSetResults "There are unread files that need to be processed!  Please contact Counterpoint with " & smPathForgLogMsg & ".", MESSAGEBLACK
'            End If
'            'gLogMsg "       " & myFile.Name & " dated: " & Format$(myFile.DateCreated, "dd/mm/yy"), smPathForgLogMsg, False
'            myErrors.WriteWarning "       " & myFile.Name & " dated: " & Format$(myFile.DateCreated, "dd/mm/yy")
        Else
            If DateDiff("d", myFile.DateCreated, dlDeleteDate) > COMPLETEDDAYS Then
                myFile.Delete
            End If
        End If
    Next myFile
    mMarketronFolder = slFolderPath
End Function
Private Function mSafeFileName(slOldName As String) As String
    Dim slTempName As String
    slTempName = Replace(slOldName, "?", "-")
    slTempName = Replace(slTempName, "/", "-")
    slTempName = Replace(slTempName, "\", "-")
    slTempName = Replace(slTempName, "%", "-")
    slTempName = Replace(slTempName, "*", "-")
    slTempName = Replace(slTempName, ":", "-")
    slTempName = Replace(slTempName, "|", "-")
    slTempName = Replace(slTempName, """", "-")
    slTempName = Replace(slTempName, ".", "-")
    slTempName = Replace(slTempName, "<", "-")
    slTempName = Replace(slTempName, ">", "-")
    mSafeFileName = slTempName
End Function
Private Sub cmdExport_Click()
    Dim iRet As Integer
    Dim sMsgFileName As String
    Dim slStatus As CSIRspGetXMLStatus
    Dim slErrorString As String
    Dim slName As String
    Dim slImportPath As String
    Dim myFile As file
    Dim blAtLeastOne As Boolean
    Dim slStartDate As String
    Dim slSignature As String
    Dim blImportFailed As Boolean
    Dim slVehicle As String
    Dim slStation As String
    Dim ilSpotCount As Integer
    Dim slFileName As String
    '4922
    Dim blNotAllImported As Boolean
    Dim blTest As Boolean
    '5602
    Dim ilOnMarketron As Integer
    '6815
    Dim ilPos As Integer
    
    blTest = False
    Screen.MousePointer = vbHourglass
'    Set lmUnMatchedSpots = New Dictionary
    imExporting = True
    lbcMsg.Clear
    lbcMsg.ForeColor = MESSAGEBLACK
    '7635
On Error Resume Next
    myErrors.CleanThisFolder = messages
    myErrors.CleanFolder
On Error GoTo 0
    slImportPath = mMarketronFolder()
    gOpenMKDFile hmAst, "Ast.Mkd"
    If Not myImport Is Nothing Then
        If myImport.isTest Then
            blTest = True
        End If
    End If
    If mnuImport.Checked Then
        mSetResults "   Guide chose to not import files", MESSAGEGREEN
    End If
    If Not blTest Then
        mSetResults "Contacting Marketron Server", MESSAGEBLACK
        If mnuRemote.Checked Then
            mSetResults "   Guide chose to not retrieve from Marketron", MESSAGEBLACK
        ElseIf Not myImport Is Nothing Then
            myImport.LogStart
            myImport.ImportPath = slImportPath
            ilOnMarketron = myImport.GetOrders(True)
            If ilOnMarketron > 0 Then
                mSetResults ilOnMarketron & " files were downloaded", MESSAGEGREEN
            ElseIf Len(myImport.ErrorMessage) > 0 Then
                mSetResults "could not download files.  See log", MESSAGERED
                myErrors.WriteError myImport.ErrorMessage
            End If
            myImport.LogEnd
        '7539  only use Jeff if have to.
        Else
            csiXMLStartRead smIniPath, "Marketron", slImportPath
            If Not csiXMLReadData() Then
                blImportFailed = True
                csiXMLStatus slStatus
                mSetResults "Could not read from Marketron--" & slStatus.sStatus, MESSAGERED
            End If
        End If
'        If mnuRemote.Checked Then
'            mSetResults "   Guide chose to not retrieve from Marketron", MESSAGEBLACK
'        Else
'            csiXMLStartRead smIniPath, "Marketron", slImportPath
'            If Not csiXMLReadData() Then
'                blImportFailed = True
'                csiXMLStatus slStatus
'                mSetResults "Could not read from Marketron--" & slStatus.sStatus, MESSAGERED
'            End If
'        End If
'        If Not myImport Is Nothing Then
'            ilOnMarketron = myImport.GetOrders()
'            If ilOnMarketron > 0 Then
'                mSetResults ilOnMarketron & " files were not downloaded", MESSAGERED
'            ElseIf Len(myImport.ErrorMessage) > 0 Then
'                mSetResults "could not read # of files available to download.  See log", MESSAGERED
'                myErrors.WriteError myImport.ErrorMessage
'            End If
'        End If
    Else
        mSetResults "not contacting Marketron server; test mode", MESSAGEBLACK
    End If
    DoEvents
    If imTerminate Then
        mSetResults "** User Terminated **", MESSAGERED
        imExporting = False
        GoTo Cleanup
    End If
    If Not mnuImport.Checked Then
        bgTaskBlocked = False
        sgTaskBlockedName = "Marketron Import"
         '7458
        Set myEnt = New CENThelper
        With myEnt
            .User = igUstCode
            .TypeEnt = Importposted3rdparty
            .ThirdParty = Vendors.NetworkConnect
            .ErrorLog = smPathForgLogMsg
        End With
        For Each myFile In oMyFileObj.GetFolder(slImportPath).Files
            '7266
            smCurrentAgreementInfo = ""
            If mTestFileName(myFile.Name) Then
                '7458
                myEnt.fileName = myFile.Name
                '6816
                Set lmUnMatchedSpots = New Dictionary
                If mProcessXml(slImportPath & myFile.Name, slStartDate, slSignature, slVehicle, slStation) Then
                    '7266
                    smCurrentAgreementInfo = slVehicle & "-" & slStation & "-" & slStartDate
                    ilSpotCount = mProcessSpots(slStartDate, slVehicle, slStation, slSignature, blNotAllImported)
                    If ilSpotCount > 0 And Not blNotAllImported Then
                        blAtLeastOne = True
                        mSetResults "Finished posting " & myFile.Name & " spots: " & ilSpotCount, MESSAGEBLACK
                        slFileName = myFile.Name
                        '6815 remove 'unread
                        If InStr(1, slFileName, "Unread") = 1 Then
                            'get rid of the 'unread' file. It worked!
                            If oMyFileObj.FILEEXISTS(slImportPath & COMPLETEDFOLDER & slFileName) Then
                                oMyFileObj.DeleteFile (slImportPath & COMPLETEDFOLDER & slFileName)
                            End If
                            ilPos = InStr(1, slFileName, "_")
                            If ilPos > 0 Then
                                slFileName = Mid(slFileName, ilPos + 1)
                            End If
                        End If
                    ElseIf ilSpotCount = 0 Then
                        '6815 don't add unread to unread
                        If InStr(1, myFile.Name, "Unread_") = 1 Then
                            slFileName = myFile.Name
                        Else
                            slFileName = "Unread_" & myFile.Name
                        End If
                        mSetResults myFile.Name & " could not be read.", MESSAGERED
                    '6815
                    ElseIf ilSpotCount < 0 Then
                        '6909 changed file name from 'station' to 'Match'
                        If InStr(1, myFile.Name, "UnreadNoMatch_") = 1 Then
                            slFileName = myFile.Name
                        Else
                            slFileName = "UnreadNoMatch_" & myFile.Name
                        End If
                        'slFileName = "UnreadNoStation_" & myFile.Name
                        mSetResults myFile.Name & " could not be read.", MESSAGERED
                    Else
                        If InStr(1, myFile.Name, "UnreadPartial_") = 1 Then
                            slFileName = myFile.Name
                        '7266 added elseif  stop "unreadPartial_Unread_"; get rid of 'unread_' in completed folder
                        ElseIf InStr(1, myFile.Name, "Unread_") = 1 Then
                            If oMyFileObj.FILEEXISTS(slImportPath & COMPLETEDFOLDER & myFile.Name) Then
                                oMyFileObj.DeleteFile (slImportPath & COMPLETEDFOLDER & myFile.Name)
                            End If
                            slFileName = Replace(myFile.Name, "Unread_", "UnreadPartial_")
                        Else
                            slFileName = "UnreadPartial_" & myFile.Name
                        End If
                       ' slFileName = "UnreadPartial_" & myFile.Name
                       'Dan 7/2/15 removed extra 'not'
                        mSetResults myFile.Name & " not all spots could be matched.", MESSAGERED
                    End If
                Else
                    If InStr(1, myFile.Name, "NoProcess_") = 1 Then
                        slFileName = myFile.Name
                    Else
                        slFileName = "NoProcess_" & myFile.Name
                    End If
                   ' slFileName = "NoProcess_" & myFile.Name
                    mSetResults myFile.Name & " had no posting information.", MESSAGEBLACK
                End If
                If imTerminate Then
                    mSetResults "** User Terminated **", MESSAGERED
                    imExporting = False
                    GoTo Cleanup
                End If
            Else
                If InStr(1, myFile.Name, "UnreadFailTest_") = 1 Then
                    slFileName = myFile.Name
                Else
                    slFileName = "UnreadFailTest_" & myFile.Name
                End If
               ' slFileName = "UnreadFailTest_" & myFile.Name
                mSetResults " A file was found which was not properly formatted: " & myFile.Name, MESSAGERED
            End If
    'testing don't move
    ''        'move file whether read or not
            If oMyFileObj.FILEEXISTS(slImportPath & COMPLETEDFOLDER & slFileName) Then
                oMyFileObj.DeleteFile (slImportPath & COMPLETEDFOLDER & slFileName)
            End If
            myFile.Move (slImportPath & COMPLETEDFOLDER & slFileName)
        Next myFile
        
        If bgTaskBlocked Then
             mSetResults "Some spots were blocked during Import.", MESSAGERED
             gMsgBox "Some spots were blocked during the Import." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
        End If
        bgTaskBlocked = False
        sgTaskBlockedName = ""
        'D.S. 08/20/13 make call here, but during updateCPTT build array of attcodes and start and end dates.
        
        
        If blAtLeastOne Then
            If blImportFailed Then
                mSetResults "Files in folder were read, but no new files were imported.", MESSAGERED
                mSetResults "**Markettron Posting completed**", MESSAGERED
            Else
                mSetResults "**Marketron Posting completed**", MESSAGEGREEN
            End If
        Else
            If Not blImportFailed Then
                mSetResults "**No files to read--posting ended**", MESSAGEBLACK
            End If
        End If
    Else
        mSetResults "**Completed**", MESSAGEGREEN
    End If  'guide will not import downloaded files
    'cmdExport.Enabled = False
Cleanup:
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    mCloseAst
    imExporting = False
    cmdCancel.Caption = "&Done"
    'cmdCancel.SetFocus
    Screen.MousePointer = vbDefault
    Set myFile = Nothing
     '7458
    Set myEnt = Nothing
    Exit Sub
ErrHand:
    'ttp 5217
    gHandleError smPathForgLogMsg, FORMNAME & "-cmdExport_Click"
    GoTo Cleanup
End Sub
Private Function mPrepMulticastChildren(tlMulticast() As OWNEDSTATIONS, slStartDate As String) As Boolean
    Dim mySqlQuery As String
    Dim slInStatement As String
    Dim c As Integer
    Dim slEndDate As String
    'for ents
    Dim slshtts As String
    
    slEndDate = DateAdd("d", 6, slStartDate)
    slEndDate = Format(slEndDate, sgSQLDateForm)
On Error GoTo ErrHand
    For c = 0 To UBound(tlMulticast) - 1 Step 1
        slInStatement = slInStatement & "," & tlMulticast(c).lAttCode
        slshtts = slshtts & "," & tlMulticast(c).iShttCode
    Next c
    'remove first ","
    slInStatement = Mid(slInStatement, 2)
    '7458
    slshtts = Mid(slshtts, 2)
    If Not myEnt.MulticastCopy(slInStatement, slshtts) Then
        myErrors.WriteWarning myEnt.ErrorMessage
    End If
    'delete child ast codes
    '12/13/13: replace Pledge with Feed
    'mySqlQuery = "DELETE from ast where astPledgeDate between '" & slStartDate & " ' and ' " & slEndDate & "' and astatfcode IN ( " & slInStatement & " )" '& slSafetyForTesting
    mySqlQuery = "DELETE from ast where astFeedDate between '" & slStartDate & " ' and ' " & slEndDate & "' and astatfcode IN ( " & slInStatement & " )" '& slSafetyForTesting
    cnn.BeginTrans
    If gSQLWaitNoMsgBox(mySqlQuery, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError smPathForgLogMsg, FORMNAME & "-mPrepMulticastChildren"
        cnn.RollbackTrans
        mPrepMulticastChildren = False
        Exit Function
    End If
    cnn.CommitTrans
    mPrepMulticastChildren = True
    Exit Function
ErrHand:
    'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, FORMNAME & "-mPrepMulticastChildren"
    mPrepMulticastChildren = False
  
End Function
Private Function mIsMulticast(ilStation As Integer, ilvehicle As Integer, slImportStartDate As String, tlChildStations() As OWNEDSTATIONS) As Boolean
    'O: tlChildStations
    Dim SQLQuery As String
    Dim mgt_rst As ADODB.Recordset
    Dim tmp_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim slName As String
    Dim blAdd As Boolean
    
On Error GoTo ErrHand
    'Dan change per Dick 11/15 exclude multicast = 0 to speed up
    blAdd = False
    mIsMulticast = False
    SQLQuery = "Select shttcode from  shtt where shttMulticastGroupid = (select shttMulticastGroupid from shtt where shttcode = " & ilStation & " AND shttMulticastGroupId > 0 ) and shttcode <> " & ilStation
    Set mgt_rst = gSQLSelectCall(SQLQuery)
    If Not (mgt_rst.EOF Or mgt_rst.BOF) Then
        mgt_rst.MoveFirst
        Do Until mgt_rst.EOF
            blAdd = False
            slName = gGetCallLettersByShttCode(mgt_rst!shttCode)
            tlChildStations(ilUpper).sCallLetters = Trim$(slName)
            tlChildStations(ilUpper).iShttCode = mgt_rst!shttCode
            tlChildStations(ilUpper).iSelected = 0
            'Find out if there is a current agreement for this station vehicle combination
            '7701
            SQLQuery = "SELECT attCode, attMulticast, vatWvtVendorId from att Left outer join VAT_Vendor_Agreement on attcode = vatAttCode WHERE (attShfCode = " & mgt_rst!shttCode & " And attVefCode = " & ilvehicle
           ' SQLQuery = "SELECT attCode, attMulticast,  attExportToMarketron from att WHERE (attShfCode = " & mgt_rst!shttCode & " And attVefCode = " & ilvehicle
            SQLQuery = SQLQuery & " AND attOffAir >= " & "'" & slImportStartDate & "'"
            SQLQuery = SQLQuery & " AND attOffAir >= attOnAir"
            SQLQuery = SQLQuery & " AND attDropDate >= " & "'" & slImportStartDate & "'" & ")"
            Set tmp_rst = gSQLSelectCall(SQLQuery)
            'not sure why this is a loop...if find one, that's good enough
            Do While Not tmp_rst.EOF
                'note: must both be marked as exporting to Marketron
                '7701
                If Trim$(tmp_rst!attMulticast) = "Y" And gIfNullInteger(tmp_rst!vatwvtvendorid) = Vendors.NetworkConnect Then
                'If Trim$(tmp_rst!attMulticast) = "Y" And Trim$(tmp_rst!attExportToMarketron) = "Y" Then
                    mIsMulticast = True
                    blAdd = True
                    tlChildStations(ilUpper).iSelected = 1
                    tlChildStations(ilUpper).lAttCode = tmp_rst!attCode
                    Exit Do
                End If
                tmp_rst.MoveNext
            Loop
            'Dan 11/15 per Dick, add if
            If blAdd Then
                ilUpper = ilUpper + 1
                ReDim Preserve tlChildStations(0 To ilUpper) As OWNEDSTATIONS
                blAdd = False
            End If
            mgt_rst.MoveNext
        Loop
    End If
Cleanup:
    If Not mgt_rst Is Nothing Then
        If (mgt_rst.State And adStateOpen) <> 0 Then
            mgt_rst.Close
            Set mgt_rst = Nothing
        End If
    End If
    Exit Function
ErrHand:
    'ttp 5217
    gHandleError smPathForgLogMsg, FORMNAME & "-mIsMulticast"
    mSetResults "Couldn't test for multicast stations! They were not posted against.", MESSAGERED
    mIsMulticast = False
    GoTo Cleanup
End Function
Private Sub mPrepAst(llAtt As Long, ilvehicle As Integer, slMondayFeedDate As String, ilStation As Integer)
    mLoadCpPosting llAtt, ilvehicle, slMondayFeedDate, ilStation '
    DoEvents
    igTimes = 1 'By Week
    'dan 6/6/11 multicasting needs tmAstInfo
    'gGetAstInfo hmAst, tmCPDat(), tmAstInfo(), -1, True, True, False
   ' gGetAstInfo hmAst, tmCPDat(), tmAstInfo(), -1, True, True, True
    '6158 change 'cpPosting' to false.  Rather small, but should've always been this way.
    gGetAstInfo hmAst, tmCPDat(), tmAstInfo(), -1, True, False, True
End Sub
Private Sub mCloseAst()
    gCloseMKDFile hmAst, "Ast.Mkd"
    Erase tmCPDat
    Erase tmAstInfo
End Sub
Private Sub mUpdateCpttToNotCreateAsts(llAtt As Long, slFeedDate As String)
    Dim SQLQuery As String
    '7895
    SQLQuery = "UPDATE cptt set cpttASTStatus = 'C' "
    SQLQuery = SQLQuery & " WHERE cpttAtfCode = " & llAtt
    SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slFeedDate, sgSQLDateForm) & "'"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError smPathForgLogMsg, FORMNAME & "-mUpdateCpttToNotCreateAsts"
        Exit Sub
    End If
    Exit Sub
ErrHand:
    gHandleError smPathForgLogMsg, FORMNAME & "-mUpdateCpttToNotCreateAsts"
End Sub
Private Sub mLoadCpPosting(llAtt As Long, ilvehicle As Integer, slFeedDate As String, ilStation As Integer)
    Dim cprst As ADODB.Recordset
    Dim SQLQuery As String
    Dim ilVpf As Integer
    
    SQLQuery = "SELECT cpttCode,cpttStatus,cpttPostingStatus,cpttAstStatus,attTimeType,shttTimeZone"
    SQLQuery = SQLQuery & " FROM cptt,att,shtt WHERE (shttCode = cpttShfCode AND attCode = cpttAtfCode " 'AND attExportToMarketron = 'Y'
    SQLQuery = SQLQuery & " AND cpttAtfCode = " & llAtt
    SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slFeedDate, sgSQLDateForm) & "')"
    Set cprst = gSQLSelectCall(SQLQuery)
    If Not cprst.EOF Then
        ReDim tgCPPosting(0 To 1) As CPPOSTING
        tgCPPosting(0).lCpttCode = cprst!cpttCode
        tgCPPosting(0).iStatus = cprst!cpttStatus
        tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
        tgCPPosting(0).lAttCode = llAtt
        tgCPPosting(0).iAttTimeType = cprst!attTimeType
        tgCPPosting(0).iVefCode = ilvehicle
        tgCPPosting(0).iShttCode = ilStation
        tgCPPosting(0).sZone = cprst!shttTimeZone
        tgCPPosting(0).sDate = Format$(slFeedDate, sgShowDateForm)
        tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
    End If
    cprst.Close
    Set cprst = Nothing
End Sub
Private Function mGetCodesFromFirstAst(slStartDate As String, slVehicle As String, slStation As String, ilvehicle As Integer, ilStation As Integer, slFeedDate As String) As Long
'return  AttCode  O-ilVehicle and ilStation and slFeedDate
    Dim rs As ADODB.Recordset
    Dim Sql As String
    Dim blAstFound As Boolean
    Dim c As Integer
    
On Error GoTo ErrHandle
    For c = 0 To UBound(tmImportSpot)
        Sql = "Select astAtfCode,astvefCode,astshfcode,astFeedDate from ast  where astcode = " & tmImportSpot(c).lAstCode
        Set rs = gSQLSelectCall(Sql)
        If Not (rs.EOF And rs.BOF) Then
            mGetCodesFromFirstAst = rs!astAtfCode
            ilvehicle = rs!astVefCode
            ilStation = rs!astShfCode
            slFeedDate = gAdjYear(gObtainPrevMonday(rs!astFeedDate))
            GoTo Cleanup
        End If
    Next c
    '7266 make sure this code is not commented out
    'use station/vehicle info and see if have a match.
    Sql = "Select AttCode,vefCode,shttcode from att join vef_Vehicles on attVefCode = vefcode join shtt on attShfCode = shttCode where vefname = '"
    '11/01/11 Dan M  use onair offair not agreestart agreeend
    'Sql = Sql & slVehicle & "' AND  shttcallletters = '" & slStation & "' AND attAgreeStart <= '" & slStartDate & "' AND attAgreeEnd >= '"
    Sql = Sql & slVehicle & "' AND  shttcallletters = '" & slStation & "' AND attOnAir <= '" & slStartDate & "' AND attOffAir >= '"
    Sql = Sql & slStartDate & "' AND attdropdate >= '" & slStartDate & "'"
    Set rs = gSQLSelectCall(Sql)
    If Not (rs.EOF And rs.BOF) Then
         mGetCodesFromFirstAst = rs!attCode
        ilvehicle = rs!vefCode
        ilStation = rs!shttCode
    End If
    'best can do:
    slFeedDate = gAdjYear(gObtainPrevMonday(slStartDate))
Cleanup:
    rs.Close
    Set rs = Nothing
    Exit Function
ErrHandle:
    mGetCodesFromFirstAst = 0
    GoTo Cleanup

End Function

Private Function mProcessXml(slPath As String, slStartDate As String, slSignature As String, slVehicle As String, slStation As String) As Boolean
'O- slStartDate
    Dim myFile As TextStream
    Dim c As Integer
    Dim ilChunk As Integer
    Dim ilMultiplier As Integer
    '7892
    Dim ilPos As Integer
    Dim ilHour As Integer
    '9538
    Dim myRemapper As cRemapper
    
    Set myRemapper = New cRemapper
    '9851
    'myRemapper.Start
    myRemapper.StartRemapping
    ilMultiplier = 1
    ilChunk = 20
    smLines = vbNullString
    mProcessXml = False
    lmCurrentRow = 0
On Error GoTo ERRNOREAD
    Set myFile = oMyFileObj.OpenTextFile(slPath)
    smLines = myFile.ReadAll
    If mTestFirstLine(smLines) Then
        slStartDate = mParseXml(smLines, "OrderStartDate")
        slSignature = Trim$(mParseXml(smLines, "Signature"))
        slVehicle = mParseXml(smLines, "ProgramName")
        '8704
        'slStation = mParseXml(smLines, "CallLetters") & "-" & mParseXml(smLines, "Band")
        slStation = mFixStation(mParseXml(smLines, "CallLetters"), mParseXml(smLines, "Band"))
        ReDim tmImportSpot(0 To ilChunk)
        Do While bmMoreRows
            If c > ilChunk * ilMultiplier Then
                ilMultiplier = ilMultiplier + 1
                ReDim Preserve tmImportSpot(0 To ilChunk * ilMultiplier)
            End If
            tmImportSpot(c).lAstCode = mParseXml(smLines, "SpotID")
            '9538
            tmImportSpot(c).lAstCode = myRemapper.Remap(tmImportSpot(c).lAstCode)
            tmImportSpot(c).sActualAirDate1 = mParseXml(smLines, "DateAired")
            tmImportSpot(c).sActualAirTime1 = mParseXml(smLines, "TimeAired")
            '7892
            If Not IsDate(tmImportSpot(c).sActualAirTime1) Then
               ' On Error Resume Next
                ilPos = InStr(1, tmImportSpot(c).sActualAirTime1, ":")
                If ilPos > 0 Then
                    ilHour = Mid(tmImportSpot(c).sActualAirTime1, 1, ilPos - 1)
                    If ilHour > 23 Then
                        If Trim(tmImportSpot(c).sActualAirTime1) = "24:00:00" Then
                            tmImportSpot(c).sActualAirTime1 = "23:59:59"
                        Else
                            ilHour = ilHour - 24
                            tmImportSpot(c).sActualAirTime1 = ilHour & Mid(tmImportSpot(c).sActualAirTime1, ilPos)
                            If IsDate(tmImportSpot(c).sActualAirTime1) And IsDate(tmImportSpot(c).sActualAirDate1) Then
                                tmImportSpot(c).sActualAirDate1 = DateAdd("d", 1, tmImportSpot(c).sActualAirDate1)
                            End If
                        End If
                    End If
                End If
              '  On Error GoTo ERRNOREAD
            End If
            tmImportSpot(c).iSpotLen = mParseXml(smLines, "Length")
            tmImportSpot(c).sISCI = mParseXml(smLines, "ISCICode")
            tmImportSpot(c).sStatusCode = mParseAndTranslateStatus()
            c = c + 1
        Loop
        If c > 0 Then
            ReDim Preserve tmImportSpot(0 To c - 1)
            mProcessXml = True
        End If
    End If
Cleanup:
    myFile.Close
    Set myFile = Nothing
    Exit Function
ERRNOREAD:
    GoTo Cleanup
End Function
Private Function mParseAndTranslateStatus() As String
    Dim slStatus As String
    
    slStatus = mParseXml(smLines, "AffidavitSpotStatus")
    Select Case slStatus
        Case "Aired"
            mParseAndTranslateStatus = "C"
        Case "Missed"
            mParseAndTranslateStatus = "N"
        Case Else
    End Select
End Function
Private Function mParseXml(sllines As String, slName As String) As String
    Dim slStartElement As String
    Dim slEndElement As String
    Dim slvalue As String
    Dim ilPos As Long
    Dim ilEndPos As Long
    Dim ilLength As Integer
    Dim ilStart As Long
    
    If lmCurrentRow = 0 Then
        lmCurrentRow = 1
    End If
    slStartElement = "<" & slName & ">"
    slEndElement = "</" & slName & ">"
    ilPos = InStr(lmCurrentRow, sllines, slStartElement)
    If ilPos > 0 Then
        ilStart = ilPos + Len(slStartElement)
        ilEndPos = InStr(lmCurrentRow, sllines, slEndElement)
        ilLength = ilEndPos - ilStart
        slvalue = Mid(sllines, ilStart, ilLength)
        mParseXml = mUnencodeXmlData(slvalue)
   Else
        mParseXml = vbNullString
   End If
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

Private Function mTestFirstLine(slLine As String) As Boolean
    '7539 allow <?xml version to be at start also
    If InStr(1, slLine, "<Affidavit>") >= 1 Then
    'If InStr(1, slLine, "<Affidavit>") = 1 Then
        mTestFirstLine = True
    Else
        mTestFirstLine = False
    End If
End Function
Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    Unload frmImportMarketron
End Sub



Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    If sgCommand = "/ReImport" Then
        Me.Left = -2 * Screen.Width
    End If
End Sub

Private Sub Form_Load()
    mInit
    If sgCommand = "/ReImport" Then
        mReImport
        Exit Sub
    End If
End Sub
Private Sub mInit()
    'csi internal guide-for testing help
    If (StrComp(sgUserName, "Guide", 1) = 0) And Not bgLimitedGuide Then
        mnuGuide.Visible = True
    End If
    frmImportMarketron.Caption = "Import Marketron - " & sgClientName
    imTerminate = False
    imExporting = False
    Set oMyFileObj = New FileSystemObject
    Set myErrors = New CLogger
    If sgCommand = "/ReImport" Then
        myErrors.LogPath = myErrors.CreateLogName(sgMsgDirectory & FILEREIMPORT)
    Else
        myErrors.LogPath = myErrors.CreateLogName(sgMsgDirectory & FILEERROR)
    End If
    smPathForgLogMsg = FILEERROR & "Log_" & Format(gNow(), "mm-dd-yy") & ".txt"
    smIniPath = gXmlIniPath(True)
    If LenB(smIniPath) = 0 Then
        cmdExport.Enabled = False
        mSetResults "Xml.ini doesn't exist.  This form cannot be activated.", MESSAGERED
        myErrors.WriteWarning "Xml.ini doesn't exit. Import Marketron cannot be activated."
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    '5602
    If Not mSetImportClass(smIniPath) Then
        cmdExport.Enabled = False
        mSetResults "Xml.ini has no values for Marketron, or values cannot be read, or proxy failed testing.  This form cannot be activated.", MESSAGERED
        myErrors.WriteWarning "Xml.ini has no values for Marketron or the values cannot be read."
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    lmMaxWidth = lbcMsg.Width
    'Dan M 7/11/13 for gGetLineParameters
    gPopDaypart
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmCPDat
    Erase tmAstInfo
    rsMG.Close
    rsMarket.Close
    Set myAligner = Nothing
    Set oMyFileObj = Nothing
    Set myErrors = Nothing
    Set myImport = Nothing
    Set frmImportMarketron = Nothing
End Sub
Private Sub mSetResults(Msg As String, FGC As Long)
    'add scroll bar as needed
    gAddMsgToListBox frmImportMarketron, lmMaxWidth, Msg, lbcMsg
    lbcMsg.ListIndex = lbcMsg.ListCount - 1
    'if ever got an error, remain red
    If lbcMsg.ForeColor <> MESSAGERED Then
        lbcMsg.ForeColor = FGC
    End If
    DoEvents
    'gLogMsg Msg, "MarketronImportLog.Txt", False
    myErrors.WriteFacts Msg
End Sub

Private Function mAttFromVehicleStation(slVehicle As String, slStation As String, slDate As String, ilVefCode As Integer, ilShttCode As Integer) As Long
    Dim rs As ADODB.Recordset
    Dim Sql As String
    
On Error GoTo ErrHandle
    Sql = "Select AttCode,vefCode,shttcode from att join vef_Vehicles on attVefCode = vefcode join shtt on attShfCode = shttCode where vefname = '"
    Sql = Sql & slVehicle & "' AND  shttcallletters = '" & slStation & "' AND attAgreeStart <= '" & Format(slDate, sgSQLDateForm) & "' AND attAgreeEnd >= '"
    Sql = Sql & Format(slDate, sgSQLDateForm) & "'"
    Set rs = gSQLSelectCall(Sql)
    If Not (rs.EOF And rs.BOF) Then
        mAttFromVehicleStation = rs(0)
        ilVefCode = rs(1)
        ilShttCode = rs(2)
    End If
Cleanup:
    rs.Close
    Set rs = Nothing
    Exit Function
ErrHandle:
    mAttFromVehicleStation = 0
    GoTo Cleanup
End Function
Private Function mProcessSpots(slStartDate As String, slVehicle As String, slStation As String, slSignature As String, blNotAllProcessed As Boolean) As Integer
    'Created by D.S. June 2007-adjusted for marketron Dan 10/26/10
    'O # of spots processed
    '0 blNotAllProcessed.  Were there spots in the import file that couldn't be fit with database spots?  Notice and mark file as such.
    Dim llIdx As Long
    Dim slNoAstExists As String
    Dim blRet As Boolean
    Dim blNoSpotAired As Boolean
    Dim llAtt As Long
    Dim ilvehicle As Integer
    Dim ilStation As Integer
    Dim slMondayFeedDate As String
    Dim ilSpotCount As Integer
    Dim tlMulticast() As OWNEDSTATIONS
    Dim slDate As String
    Dim blMulticast As Boolean
    Dim c As Integer
    Dim slMultiCastMessage As String
    
    blNotAllProcessed = False
    mProcessSpots = 0
On Error GoTo ErrHand
    llAtt = mGetCodesFromFirstAst(slStartDate, slVehicle, slStation, ilvehicle, ilStation, slMondayFeedDate)
    If llAtt = 0 Or ilvehicle = 0 Or ilStation = 0 Then
        '6815
        'mProcessSpots = 0
        mProcessSpots = -1
        Exit Function
    End If
    mSetResults "preparing spots to be posted against.", MESSAGEBLACK
    '7458
    With myEnt
        .Vehicle = ilvehicle
        .Station = ilStation
        .Agreement = llAtt
        .ProcessStart
    End With
    mPrepAst llAtt, ilvehicle, slMondayFeedDate, ilStation
    ReDim tlMulticast(0) As OWNEDSTATIONS
    slDate = Format(slStartDate, sgSQLDateForm)
    Set rsMG = mPrepRecordset()
    blMulticast = False
    imMarket = 0
    If mIsMulticast(ilStation, ilvehicle, slDate, tlMulticast) Then
        blMulticast = mPrepMulticastChildren(tlMulticast, slDate)
    '6158
    Else
        '6494 hide makegoods. Commenting this out blocks mg, although some coding still exits (see mMGResetBeforeImporting)
'        '6158  makegoods. As below, let error trip error catching in this function
'        mMGFindImported
'        SQLQuery = "Select  shttMktCode as market from shtt where shttcode = " & ilStation
'        Set rst = gSQLSelectCall(SQLQuery)
'        If Not rst.EOF Then
'            imMarket = rst!market
'        End If
    End If
    '6158 reset ast as not posted. delete missed from alt, delete makegoods from alt, ast, lst
    mMGResetBeforeImporting
   ' an error here will stop the processing and go to error handler below
    mImportByAstCode blNoSpotAired, llAtt, ilvehicle, ilStation, tlMulticast, blMulticast
    '4922
    '7266
    'mFixLostSpots llAtt, tlMulticast, blMulticast
    'returns if at least one spot was corrected
    If mFixLostSpots(ilvehicle, tlMulticast, blMulticast, slMondayFeedDate) Then
        blNoSpotAired = False
    End If
    If Not mUpdateCptt(blNoSpotAired, ilvehicle, llAtt, slMondayFeedDate, ilStation) Then
        mSetResults "Error with make goods in mUpdateCptt", MESSAGERED
        ilSpotCount = 0
        GoTo Cleanup
    End If
    mInsertToWebLog slVehicle, slStation, slSignature, llAtt, slStartDate
    If blMulticast Then
        For c = 0 To UBound(tlMulticast)
            If tlMulticast(c).iSelected = 1 Then
                'change cptt cpttaststatus = "C" 7895
                mUpdateCpttToNotCreateAsts tlMulticast(c).lAttCode, slMondayFeedDate
                mUpdateCptt blNoSpotAired, ilvehicle, tlMulticast(c).lAttCode, slMondayFeedDate, tlMulticast(c).iShttCode
                mInsertToWebLog slVehicle, slStation, slSignature, tlMulticast(c).lAttCode, slStartDate
                If tlMulticast(c).iGroupID = -9 Then
                    'what to do when an insert error happened? a spot is missing. Already wrote out a message.
                Else
                    slMultiCastMessage = slMultiCastMessage & tlMulticast(c).sCallLetters
                End If
            End If
        Next c
        If Len(slMultiCastMessage) > 0 Then
            slMultiCastMessage = " Posted to multicast stations: " & slMultiCastMessage
            mSetResults slMultiCastMessage, MESSAGEBLACK
        End If
    End If
    'Dan M removed this 4922. Import spots should be accounted for.
    '  If Not blNoImportError Then
    For llIdx = 0 To UBound(tmImportSpot) '- 1  Step 1
        If tmImportSpot(llIdx).iFound = False Then
            '7458 unmatched...use aired date
            If Not myEnt.Add(tmImportSpot(llIdx).sActualAirDate1, 0, SentOrReceived) Then
                myErrors.WriteWarning "using air date: " & myEnt.ErrorMessage
            End If
            '6714
            If tmImportSpot(llIdx).lAstCode > 0 Then
                blNotAllProcessed = True
                slNoAstExists = slVehicle & "," & slStation & ", attCode: "
                slNoAstExists = slNoAstExists & llAtt & ", Spot Length: "
                slNoAstExists = slNoAstExists & Trim$(tmImportSpot(llIdx).iSpotLen) & ", ISCI: "
                slNoAstExists = slNoAstExists & Trim$(tmImportSpot(llIdx).sISCI) & ", astCode:"
                'slNoAstExists = slNoAstExists & Trim$(tmImportSpot(llIdx).sCreativeTitle) & ", astCode: "
                slNoAstExists = slNoAstExists & Trim$(tmImportSpot(llIdx).lAstCode) & ", Air date: "
                slNoAstExists = slNoAstExists & Trim$(tmImportSpot(llIdx).sActualAirDate1) & ", Air time: "
                slNoAstExists = slNoAstExists & Trim$(tmImportSpot(llIdx).sActualAirTime1) & ", Status: "
                slNoAstExists = slNoAstExists & Trim$(tmImportSpot(llIdx).sStatusCode)
                mSetResults "Unable to process: " & slNoAstExists & " AST missing", MESSAGERED
            Else
                myErrors.WriteWarning "Unable to process a spot because ast code was 0"
            End If
        Else
            ilSpotCount = ilSpotCount + 1
        End If
    Next llIdx
    '7458 if ilSpotCount = 0 won't hurt us much
    If Not myEnt.CreateEnts() Then
        myErrors.WriteWarning myEnt.ErrorMessage
    End If
Cleanup:
    mProcessSpots = ilSpotCount
    If Not rsMG Is Nothing Then
        If (rsMG.State And adStateOpen) <> 0 Then
            rsMG.Close
        End If
        Set rsMG = Nothing
    End If
    If Not rsMarket Is Nothing Then
        If (rsMarket.State And adStateOpen) <> 0 Then
            rsMarket.Close
        End If
        Set rsMarket = Nothing
    End If
    Exit Function
ErrHand:
    'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, FORMNAME & "-mProcessSpots"
    ilSpotCount = 0
    GoTo Cleanup
End Function
Private Function mFixLostSpots(ilvehicle As Integer, tlMulticast() As OWNEDSTATIONS, blMulticast As Boolean, slMondayDate As String) As Boolean
    '7266 changed llatt to ilVehicle. changed to function; pass if at least one spot was correctly aligned
   ' Dim c As Long
    Dim slErrorMessage As String
    Dim blAtLeastOne As Boolean
    
    slErrorMessage = ""
    '6911 remove this
'    For c = 0 To UBound(tmImportSpot)
'        If tmImportSpot(c).iFound = False Then
'            mImportByFeedDate c, llAtt, tlMulticast, blMulticast
'        End If
'    Next c
    '7266
    blAtLeastOne = mImportAndAdjustAstCode(ilvehicle, tlMulticast, blMulticast, slMondayDate)
    If blMulticast Then
        If Not mInsertUnmatchedToSlaves(tlMulticast, slErrorMessage) Then
            mSetResults " One or more spots failed to be posted for the multicast station " & slErrorMessage, MESSAGERED
        End If
    End If
    mFixLostSpots = blAtLeastOne
End Function
Private Function mInsertUnmatchedToSlaves(tlMulticast() As OWNEDSTATIONS, slErrorMessage As String) As Boolean
    Dim vArray As Variant
    Dim c As Integer
    Dim blError As Boolean
    Dim llSpotLoop As Long
    Dim tlMyAst As AST
    
On Error GoTo ErrHandler
    blError = False
    vArray = lmUnMatchedSpots.Items
    For c = 0 To lmUnMatchedSpots.Count - 1
        llSpotLoop = vArray(c)
        With tlMyAst
            .lCode = 0
            .iStatus = tmAstInfo(llSpotLoop).iStatus
            .iVefCode = tmAstInfo(llSpotLoop).iVefCode
            .iCPStatus = tmAstInfo(llSpotLoop).iCPStatus
            .lSdfCode = tmAstInfo(llSpotLoop).lSdfCode
            .lLsfCode = tmAstInfo(llSpotLoop).lLstCode
            '12/9/13
            '.iPledgeStatus = tmAstInfo(llSpotLoop).iPledgeStatus
            gPackDate tmAstInfo(llSpotLoop).sFeedDate, .iFeedDate(0), .iFeedDate(1)
            gPackTime tmAstInfo(llSpotLoop).sFeedTime, .iFeedTime(0), .iFeedTime(1)
            '12/9/13
            'gPackDate tmAstInfo(llSpotLoop).sPledgeDate, .iPledgeDate(0), .iPledgeDate(1)
            'gPackTime tmAstInfo(llSpotLoop).sPledgeStartTime, .iPledgeStartTime(0), .iPledgeStartTime(1)
            'gPackTime tmAstInfo(llSpotLoop).sPledgeEndTime, .iPledgeEndTime(0), .iPledgeEndTime(1)
            gPackDate tmAstInfo(llSpotLoop).sAirDate, .iAirDate(0), .iAirDate(1)
            gPackTime tmAstInfo(llSpotLoop).sAirTime, .iAirTime(0), .iAirTime(1)
            .iAdfCode = tmAstInfo(llSpotLoop).iAdfCode
            .lDatCode = tmAstInfo(llSpotLoop).lDatCode
            .lCpfCode = tmAstInfo(llSpotLoop).lCpfCode
            .lRsfCode = tmAstInfo(llSpotLoop).lRRsfCode
'            .sStationCompliant = ""
'            .sAgencyCompliant = ""
'            .sAffidavitSource = ""
            .iUstCode = igUstCode
            '7894
            .iLen = tmAstInfo(llSpotLoop).iLen
            .lCntrNo = tmAstInfo(llSpotLoop).lCntrNo
            '7895
            .sAgencyCompliant = tmAstInfo(llSpotLoop).sAgencyCompliant
            .sStationCompliant = tmAstInfo(llSpotLoop).sStationCompliant
            .sAffidavitSource = tmAstInfo(llSpotLoop).sAffidavitSource
        End With
        If Not mInsertMulticast(tlMyAst, slErrorMessage, tlMulticast) Then
            blError = True
        End If
    Next c
    mInsertUnmatchedToSlaves = Not blError
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, FORMNAME & "-mInsertUnmatchedToSlaves"
    mInsertUnmatchedToSlaves = False
End Function
Private Function mImportAndAdjustAstCode(ilVefCode As Integer, tlMulticast() As OWNEDSTATIONS, blMulticast As Boolean, slMondayDate As String) As Boolean
    '7266 return true if at least one spot was fixed and it was 'aired'
    Dim c As Long
    Dim blNeedExports As Boolean
    Dim llAstReplacementCode As Long
    Dim tlMyAst As AST
    Dim blError As Boolean
    Dim llSpotLoop As Long
    Dim llMax As Long
    Dim slErrorMessage As String
    Dim llAlignCount As Long
    Dim blAtLeastOneAired As Boolean
    '7639
    Dim llNewCpf As Long
    Dim ilAddToStatus As Integer
    '9102
    Dim ilAdv As Integer
    Dim ilSpot As Integer
    Set myAligner = New cMarketronAligner
    blNeedExports = False
    llAlignCount = 0
    blAtLeastOneAired = False
On Error GoTo errbox
    For c = 0 To UBound(tmImportSpot)
        With tmImportSpot(c)
           '6714 don't bother aligning 0 astcodes!
           If .iFound = False And tmImportSpot(c).lAstCode <> 0 Then
                blNeedExports = True
                ilAdv = mReturnAdv(.sISCI, ilSpot, slMondayDate)
                If Not myAligner.AddImport(.sActualAirDate1, .sActualAirTime1, .lAstCode, .sISCI, ilAdv, ilSpot) Then
                    GoTo Cleanup
                End If
            End If
        End With
    Next c
    If blNeedExports Then
        If Not mGetExportsForAdjusting() Then
            GoTo Cleanup
        End If
        For c = 0 To UBound(tmImportSpot)
            If tmImportSpot(c).iFound Then
                myAligner.MarkExportAsFound tmImportSpot(c).lAstCode
            End If
        Next c
        If Not myAligner.Align() Then
            GoTo Cleanup
        Else
            llAlignCount = myAligner.AlignCount
            If llAlignCount = 1 Then
                myErrors.WriteWarning "1 spot was matched with a different astcode in " & smCurrentAgreementInfo
            ElseIf llAlignCount > 1 Then
                myErrors.WriteWarning llAlignCount & " spots were matched with different astcodes in " & smCurrentAgreementInfo
            End If
        End If
        llMax = UBound(tmAstInfo) - 1
        For c = 0 To UBound(tmImportSpot)
            If Not tmImportSpot(c).iFound Then
                llAstReplacementCode = myAligner.ReturnReplacementAst(tmImportSpot(c).lAstCode)
                If llAstReplacementCode > 0 Then
                    For llSpotLoop = 0 To llMax
                        If tmAstInfo(llSpotLoop).lCode = llAstReplacementCode Then
                            '7639
                             If tmImportSpot(c).sStatusCode = "C" Then
                                 llNewCpf = mAdjustISCIAsNeeded(llSpotLoop, tmImportSpot(c).sISCI)
                             Else
                                 llNewCpf = 0
                             End If
                            'update here
                             With tlMyAst
                                'update ast
                                .lCode = llAstReplacementCode
                                .iStatus = tmAstInfo(llSpotLoop).iStatus
                                ' update multicast
                                .iVefCode = ilVefCode
                                .iCPStatus = 1
                                .lSdfCode = tmAstInfo(llSpotLoop).lSdfCode
                                .lLsfCode = tmAstInfo(llSpotLoop).lLstCode
                                '12/9/13
                                '.iPledgeStatus = tmAstInfo(llSpotLoop).iPledgeStatus
                                gPackDate tmAstInfo(llSpotLoop).sFeedDate, .iFeedDate(0), .iFeedDate(1)
                                gPackTime tmAstInfo(llSpotLoop).sFeedTime, .iFeedTime(0), .iFeedTime(1)
                                '12/9/13
                                'gPackDate tmAstInfo(llSpotLoop).sPledgeDate, .iPledgeDate(0), .iPledgeDate(1)
                                'gPackTime tmAstInfo(llSpotLoop).sPledgeStartTime, .iPledgeStartTime(0), .iPledgeStartTime(1)
                                'gPackTime tmAstInfo(llSpotLoop).sPledgeEndTime, .iPledgeEndTime(0), .iPledgeEndTime(1)
                                .iAdfCode = tmAstInfo(llSpotLoop).iAdfCode
                                .lDatCode = tmAstInfo(llSpotLoop).lDatCode
                                '7639
                                '.lCpfCode = tmAstInfo(llSpotLoop).lCpfCode
                                '7639
                                 If llNewCpf > 0 Then
                                     .lCpfCode = llNewCpf
                                     ilAddToStatus = ASTEXTENDED_ISCICHGD
                                 Else
                                     .lCpfCode = tmAstInfo(llSpotLoop).lCpfCode
                                     ilAddToStatus = 0
                                 End If
                                .lRsfCode = tmAstInfo(llSpotLoop).lRRsfCode
                                .sStationCompliant = ""
                                .sAgencyCompliant = ""
                                .sAffidavitSource = ""
                                .iUstCode = igUstCode
                            End With
                            'returns new value for ilAstStatus for multicast
                            '7639 added lladdtostatus
                             If mUpdateAst(c, tlMyAst, ilAddToStatus) Then
                                blAtLeastOneAired = True
                             End If
                            'tlMyAst returned from above:
                            If tlMyAst.iStatus <> 4 Then
                                gPackDate tmImportSpot(c).sActualAirDate1, tlMyAst.iAirDate(0), tlMyAst.iAirDate(1)
                                gPackTime tmImportSpot(c).sActualAirTime1, tlMyAst.iAirTime(0), tlMyAst.iAirTime(1)
                            End If
                            'these will be added to multicast later, so remove if inserting here.
                            If lmUnMatchedSpots.Exists(tlMyAst.lCode) Then
                                lmUnMatchedSpots.Remove (tlMyAst.lCode)
                            End If
                            If blMulticast Then
                                '7639 if there's a new cpfcode, I need the old one to write the alt
    '                            'any time not true, keep record of
    '                            If Not mInsertMulticast(tlMyAst, slErrorMessage, tlMulticast) Then
    '                                blError = True
    '                            End If
                                If llNewCpf > 0 Then
                                    If Not mInsertMulticast(tlMyAst, slErrorMessage, tlMulticast, tmAstInfo(llSpotLoop).lCpfCode) Then
                                        blError = True
                                    End If
                                Else
                                    If Not mInsertMulticast(tlMyAst, slErrorMessage, tlMulticast, 0) Then
                                        blError = True
                                    End If
                                End If
                            End If
                            tmImportSpot(c).iFound = True
                            '7458
                            If Not myEnt.Add(tmAstInfo(llSpotLoop).sFeedDate, tmAstInfo(llSpotLoop).lgsfCode, Ingested) Then
                                myErrors.WriteWarning myEnt.ErrorMessage
                            End If
                            Exit For
                        End If
                    Next llSpotLoop
                End If  'have a replacement code
            End If 'not found previously--did we fix?
        Next c
        If blError Then
            mSetResults " One or more spots failed to be posted for the multicast station " & slErrorMessage, MESSAGERED
        End If
    End If
Cleanup:
    mImportAndAdjustAstCode = blAtLeastOneAired
    Set myAligner = Nothing
    Exit Function
errbox:
    blAtLeastOneAired = False
    GoTo Cleanup
End Function
Private Function mGetExportsForAdjusting() As Boolean
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilAllowedDays(6) As Integer
    Dim ilCompliant As Integer
    Dim blRet As Boolean
    Dim c As Long
    Dim slISCI As String
    
On Error GoTo errbox
    blRet = True
    For c = 0 To UBound(tmAstInfo) - 1
        If (tgStatusTypes(gGetAirStatus(tmAstInfo(c).iStatus)).iPledged <> 2) Then
            'dan, in v70 and up, use the lstLnStartTime and EndTime
            gGetLineParameters False, tmAstInfo(c), slStartDate, slEndDate, slStartTime, slEndTime, ilAllowedDays(), True
            '0 =none, 1=split copy, 2=blackout
            If tmAstInfo(c).iRegionType = 0 Then
                slISCI = tmAstInfo(c).sISCI
            Else
                slISCI = tmAstInfo(c).sRISCI
            End If
            If Not myAligner.AddExport(slStartDate, slEndDate, slStartTime, slEndTime, tmAstInfo(c).lCode, slISCI, tmAstInfo(c).iAdfCode, tmAstInfo(c).iLen) Then
                blRet = False
                Exit For
            End If
        End If
    Next c
    mGetExportsForAdjusting = blRet
    Exit Function
errbox:
    mGetExportsForAdjusting = False
End Function
'Private Function mImportByFeedDate(llImportLoop As Long, llAtfCode As Long, tlMulticast() As OWNEDSTATIONS, blMulticast As Boolean) As Boolean
'    ' O- #  resolved by feed date/time?
'    '4922 Dan most of this borrowed from mImportByFeedDate/DS
'    Dim myRst As ADODB.Recordset
'    Dim slSql As String
'    Dim slSavedTime As String
'    Dim slSavedDate As String
'    Dim blFound As Boolean
'    Dim llSpotLoop As Long
'    Dim llAstCode As Long
'    Dim slErrorMessage As String
'    Dim blError As Boolean
'    Dim blNoSpotAired As Boolean
'    Dim tlMyAst As AST
'    Dim llNewAst As Long
'    Dim ilRet As Integer
'    Dim tlDatPledgeInfo As DATPLEDGEINFO
'
'    llAstCode = tmImportSpot(llImportLoop).lAstCode
'    slSql = "select aetFeedDate, aetFeedTime from aet where aetStatus = 'M' AND aetastcode = " & llAstCode
'On Error GoTo ErrHand
'    Set myRst = gSQLSelectCall(slSql)
'    If Not myRst.EOF Then
'        slSavedTime = myRst("aetfeedtime").Value
'        slSavedTime = Format(slSavedTime, sgSQLTimeForm)
'        slSavedDate = Format(myRst("aetfeeddate").Value, sgSQLDateForm)
'        myRst.Close
'        '12/13/13: Obtain Pledge information from Dat.  See change below
'        slSql = "Select * FROM ast Where astFeedDate = '" & slSavedDate & "' "
'        slSql = slSql & "And  astCPStatus = 0 And astAtfCode = " & llAtfCode
'        Set myRst = gSQLSelectCall(slSql)
'        If Not myRst.EOF Then
'            'go through each one and see if astCodes match as a safety.
'            Do While Not myRst.EOF
'                DoEvents
'                '12/13/13: Obtain Pledge information from Dat
'                tlDatPledgeInfo.lAttCode = myRst!astAtfCode
'                tlDatPledgeInfo.lDatCode = myRst!astDatCode
'                tlDatPledgeInfo.iVefCode = myRst!astVefCode
'                tlDatPledgeInfo.sFeedDate = Format(myRst!astFeedDate, "m/d/yy")
'                tlDatPledgeInfo.sFeedTime = Format(myRst!astFeedTime, "hh:mm:ssam/pm")
'                ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
'                If tgStatusTypes(gGetAirStatus(tlDatPledgeInfo.iPledgeStatus)).iPledged <> 2 Then
'                    If gTimeToLong(Format(myRst!astFeedTime, "hh:mm:ssam/pm"), False) = gTimeToLong(slSavedTime, False) Then
'                        If myRst("astCode").Value = llAstCode Then
'                            blFound = True
'                            mImportByFeedDate = True
'                            Exit Do
'                        Else
'                            llNewAst = myRst!astCode
'                        End If
'                    End If
'                End If
'                myRst.MoveNext
'            Loop
'            'as expected, didn't find astcode in myrst.  Pick one record and use that
'            If Not blFound Then
'                myRst.Filter = " astcode = " & llNewAst
'                If Not myRst.EOF Then
'                    With tlMyAst
'                        'this is how we pass the astCode that is in the database, not the import file
'                        .lCode = myRst!astCode
'                        .iStatus = myRst!astStatus
'                        .iVefCode = myRst!astVefCode
'                        .iCPStatus = 1
'                        .lSdfCode = myRst!astSdfCode
'                        .lLsfCode = myRst!astLsfCode
'                        '12/9/13
'                        '.iPledgeStatus = tlDatPledgeInfo.iPledgeStatus
'                        gPackDate myRst!astFeedDate, .iFeedDate(0), .iFeedDate(1)
'                        gPackTime myRst!astFeedTime, .iFeedTime(0), .iFeedTime(1)
'                        '12/9/13
'                        'gPackDate myRst!astPledgeDate, .iPledgeDate(0), .iPledgeDate(1)
'                        'gPackTime myRst!astPledgeStartTime, .iPledgeStartTime(0), .iPledgeStartTime(1)
'                        'gPackTime myRst!astPledgeEndTime, .iPledgeEndTime(0), .iPledgeEndTime(1)
'                        .iAdfCode = myRst!astAdfCode
'                        .lDatCode = myRst!astDatCode
'                        .lCpfCode = myRst!astcpfcode
'                        .lRsfCode = myRst!astRsfCode
'                        .sStationCompliant = ""
'                        .sAgencyCompliant = ""
'                        .sAffidavitSource = ""
'                        .iUstCode = igUstCode
'                    End With
'                    blNoSpotAired = Not mUpdateAst(llImportLoop, tlMyAst)
'                    If blMulticast Then
'                        'returned from mUpdateAst above:
'                        If tlMyAst.iStatus <> 4 Then
'                            'note air date and time are from the imported file
'                            gPackDate tmImportSpot(llImportLoop).sActualAirDate1, tlMyAst.iAirDate(0), tlMyAst.iAirDate(1)
'                            gPackTime tmImportSpot(llImportLoop).sActualAirTime1, tlMyAst.iAirTime(0), tlMyAst.iAirTime(1)
'                        End If
'                        'these will be added to multicast later, so remove if inserting here.
'                        If lmUnMatchedSpots.Exists(tlMyAst.lCode) Then
'                            lmUnMatchedSpots.Remove (tlMyAst.lCode)
'                        End If
'                        'any time not true, keep record of
'                        If Not mInsertMulticast(tlMyAst, slErrorMessage, tlMulticast) Then
'                            blError = True
'                        End If
'                    End If
'                    mImportByFeedDate = True
'                End If
'            End If 'not blFound
'        End If 'not EOF ast table
'    End If 'not EOF aet table
'Cleanup:
'    If blError Then
'        mSetResults " One or more spots failed to be posted for the multicast station " & slErrorMessage, MESSAGERED
'    End If
'    If Not myRst Is Nothing Then
'        If (myRst.State And adStateOpen) <> 0 Then
'            myRst.Close
'        End If
'        Set myRst = Nothing
'    End If
'    Exit Function
'ErrHand:
'    Screen.MousePointer = vbDefault
'    gHandleError smPathForgLogMsg, FORMNAME & "-mImportByFeedDate"
'    mImportByFeedDate = False
'    GoTo Cleanup
'End Function



Private Sub mInsertToWebLog(slVehicleName As String, slStation As String, slSignature As String, llAttCode As Long, slDate As String)
    Dim SQLQuery As String
    Dim slCurrent As String
    
    slCurrent = gNow()
    If Len(slSignature) <= 30 Then
        slSignature = "marketron:" & slSignature
    Else
        slSignature = "m:" & slSignature
    End If
    SQLQuery = "Insert Into WebL (weblType, weblattCode, weblCallLetters, weblVehicleName, weblUserName, weblPostDay, weblDate, weblTime) " ' weblIP, weblCPUName,
    SQLQuery = SQLQuery & "Values (3," & llAttCode & ",'" & gFixQuote(slStation) & "','" & gFixQuote(slVehicleName) & "','" & gFixQuote(slSignature) & "', '" & slDate & "',"
    SQLQuery = SQLQuery & "'" & Format$(slCurrent, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(slCurrent, sgSQLTimeForm) & "'"
    SQLQuery = SQLQuery & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        GoTo ERRSQL
    End If
    Exit Sub
ERRSQL:
    'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, FORMNAME & "-mInsertToWebLog"
    mSetResults gMsg, MESSAGERED
End Sub
Private Function mImportByAstCode(blNoSpotAired As Boolean, llAtfCode As Long, ilVefCode As Integer, ilShfCode As Integer, tlMulticast() As OWNEDSTATIONS, blMulticast As Boolean) As Boolean
'    update the AST if it exists and mark it as found in the array. Return false only on error OR code matches but vehicle and station don't.
'   O-blNoSpotAired
'    6/7/2011 added multicasting
    Dim llImportLoop As Long
    Dim llAstCode As Long
    Dim llAstInfoKey As Long
    Dim llSpotLoop As Long
    Dim blError As Boolean
    Dim slErrorMessage As String
    Dim blFound As Boolean
    Dim tlMyAst As AST
    '6757
    Dim blThisSpotMissed As Boolean
    '7639
    Dim llNewCpf As Long
    Dim ilAddToStatus As Integer
On Error GoTo ErrHand
    mImportByAstCode = True
    blNoSpotAired = True
    '7458 get ast count
    For llSpotLoop = 0 To UBound(tmAstInfo) - 1
        DoEvents
        '7458
       If Not myEnt.Add(tmAstInfo(llSpotLoop).sFeedDate, tmAstInfo(llSpotLoop).lgsfCode, Asts) Then
            myErrors.WriteWarning myEnt.ErrorMessage
        End If
    Next llSpotLoop
    'reverse loop and lose recordset(replace with tmastinfo) for multicasting
    For llSpotLoop = 0 To UBound(tmAstInfo) - 1
        DoEvents
        llAstCode = tmAstInfo(llSpotLoop).lCode
        blFound = False
        For llImportLoop = LBound(tmImportSpot) To UBound(tmImportSpot) Step 1
            If llAstCode = tmImportSpot(llImportLoop).lAstCode Then
                'Dan M note:  Use attcode,station, and vehicle to test if same spot.
                'because the import doesn't pass most of this info, ilshfcode, and ilvefcode are set by looking at the first spot.
                ' All we are doing is making sure all further spots match this one.
                If tmAstInfo(llSpotLoop).iShttCode = ilShfCode And tmAstInfo(llSpotLoop).lAttCode = llAtfCode And tmAstInfo(llSpotLoop).iVefCode = ilVefCode Then
                    blFound = True
                    '7458
                    If Not myEnt.Add(tmAstInfo(llSpotLoop).sFeedDate, tmAstInfo(llSpotLoop).lgsfCode, Ingested) Then
                        myErrors.WriteWarning myEnt.ErrorMessage
                    End If
                   '6158
                    If Not mMGIsMakeGood(llSpotLoop, blMulticast) Then
                        '7639
                        If tmImportSpot(llImportLoop).sStatusCode = "C" Then
                            llNewCpf = mAdjustISCIAsNeeded(llSpotLoop, tmImportSpot(llImportLoop).sISCI)
                        Else
                            llNewCpf = 0
                        End If
                        With tlMyAst
                            'update ast
                            .lCode = llAstCode
                            .iStatus = tmAstInfo(llSpotLoop).iStatus
                            ' update multicast
                            .iVefCode = ilVefCode
                            .iCPStatus = 1
                            .lSdfCode = tmAstInfo(llSpotLoop).lSdfCode
                            .lLsfCode = tmAstInfo(llSpotLoop).lLstCode
                            '12/9/13
                            '.iPledgeStatus = tmAstInfo(llSpotLoop).iPledgeStatus
                            gPackDate tmAstInfo(llSpotLoop).sFeedDate, .iFeedDate(0), .iFeedDate(1)
                            gPackTime tmAstInfo(llSpotLoop).sFeedTime, .iFeedTime(0), .iFeedTime(1)
                            '12/9/13
                            'gPackDate tmAstInfo(llSpotLoop).sPledgeDate, .iPledgeDate(0), .iPledgeDate(1)
                            'gPackTime tmAstInfo(llSpotLoop).sPledgeStartTime, .iPledgeStartTime(0), .iPledgeStartTime(1)
                            'gPackTime tmAstInfo(llSpotLoop).sPledgeEndTime, .iPledgeEndTime(0), .iPledgeEndTime(1)
                            .iAdfCode = tmAstInfo(llSpotLoop).iAdfCode
                            .lDatCode = tmAstInfo(llSpotLoop).lDatCode
                            '7639
                            If llNewCpf > 0 Then
                                .lCpfCode = llNewCpf
                                ilAddToStatus = ASTEXTENDED_ISCICHGD
                            Else
                                .lCpfCode = tmAstInfo(llSpotLoop).lCpfCode
                                ilAddToStatus = 0
                            End If
                            .lRsfCode = tmAstInfo(llSpotLoop).lRRsfCode
                            .sStationCompliant = ""
                            .sAgencyCompliant = ""
                            .sAffidavitSource = ""
                            .iUstCode = igUstCode
                              '7894
                            .iLen = tmAstInfo(llSpotLoop).iLen
                            .lCntrNo = tmAstInfo(llSpotLoop).lCntrNo
                            gPackDate tmAstInfo(llSpotLoop).sAirDate, .iAirDate(0), .iAirDate(1)
                            gPackTime tmAstInfo(llSpotLoop).sAirTime, .iAirTime(0), .iAirTime(1)
                            '7895
                            .sAgencyCompliant = tmAstInfo(llSpotLoop).sAgencyCompliant
                            .sStationCompliant = tmAstInfo(llSpotLoop).sStationCompliant
                            .sAffidavitSource = tmAstInfo(llSpotLoop).sAffidavitSource
                        End With
                        'returns new value for ilAstStatus for multicast
                        '6757 if last one didn't air, thinking NONE aired
                        'blNoSpotAired = Not mUpdateAst(llImportLoop, tlMyAst)
                        '7639
                        blThisSpotMissed = Not mUpdateAst(llImportLoop, tlMyAst, ilAddToStatus)
                        If Not blThisSpotMissed Then
                            blNoSpotAired = False
                        End If
                        'tlMyAst returned from above:
                        If tlMyAst.iStatus <> 4 Then
                            gPackDate tmImportSpot(llImportLoop).sActualAirDate1, tlMyAst.iAirDate(0), tlMyAst.iAirDate(1)
                            gPackTime tmImportSpot(llImportLoop).sActualAirTime1, tlMyAst.iAirTime(0), tlMyAst.iAirTime(1)
                        End If
                        If blMulticast Then
                            '7639 if there's a new cpfcode, I need the old one to write the alt
'                            'any time not true, keep record of
'                            If Not mInsertMulticast(tlMyAst, slErrorMessage, tlMulticast) Then
'                                blError = True
'                            End If
                            If llNewCpf > 0 Then
                                If Not mInsertMulticast(tlMyAst, slErrorMessage, tlMulticast, tmAstInfo(llSpotLoop).lCpfCode) Then
                                    blError = True
                                End If
                            Else
                                If Not mInsertMulticast(tlMyAst, slErrorMessage, tlMulticast, 0) Then
                                    blError = True
                                End If
                            End If
                        End If
                    End If
                        'Dan M note:  for 4922, moved to mUpdateAst and mInsertMulticast 6/7/2012
                    'found match, done with this ast
                    Exit For
                ' wrong spot has same astcode! Skip this spot. Note that code continues, so this will be marked as missed.
                Else
                    mSetResults "ast code #" & llAstCode & " doesn't match information being imported: station code " & ilShfCode & " vehicle code " & ilVefCode & " or agreement code " & llAtfCode, MESSAGERED
                End If  'same spot?
            End If  'astcodes match
        Next llImportLoop
        ' Need to add these spots to multicast slaves. But first, see if there's a match with feedDate/Time; otherwise would be adding 2x
        If Not blFound And blMulticast Then
            lmUnMatchedSpots.Add tmAstInfo(llSpotLoop).lCode, llSpotLoop
        End If ' not found?
    Next llSpotLoop
    '6158
    rsMG.Filter = adFilterNone
    If Not blMulticast And rsMG.RecordCount > 0 Then
        If Not mMGImportAll() Then
            mSetResults " Problem working with MakeGoods - mMGImportAll.", MESSAGERED
            GoTo ErrHand
        End If
    End If
    If blError Then
        mSetResults " One or more spots failed to be posted for the multicast station " & slErrorMessage, MESSAGERED
    End If
    Exit Function
ErrHand:
    mImportByAstCode = False
    'Throw to mProcessSpots to stop all processing.
    Err.Raise ERRORSQL, "mImportByAstCode", Err.Description
End Function
Private Function mUpdateAst(llImportLoop As Long, myTlAst As AST, ilAddToStatus As Integer) As Boolean
    ' O- did spots air?
    '7639 added ilAddToStatus to handle adding '1000' (ASTEXTENDED_ISCICHGD)
    Dim llAstCode As Long
    Dim ilAstStatus As Integer
    
    tmImportSpot(llImportLoop).iFound = True
    llAstCode = myTlAst.lCode
    If tmImportSpot(llImportLoop).sStatusCode = "C" Then
        'C - Program and Commercial Aired Live.
        '   Marketron returns = "Aired"  Status = 0  Screen = 1-Aired Live
        'N - Neither the spot nor the program aired.
        '   Marketron returns = "Missed"  Status = 4   Screen = 5-Not Aired Other
        'unused:
        'D - Program and Commercial were both delayed.
        '   Marketron does not return  Status = 9   Screen = 10-Delay Cmml/Prg
        'S - Program did not air, but spot aired, either live or delayed.
        '   Marketron does not return  Status = 10  Screen = 11-Air Cmml Only
        'P - Program aired spot did not.
        '   Marketron does not return  Status = 4   Screen = 5-Not Aired Other
        'K - Delay B'cast
        '   Marketron does not return  Status = 1   Screen = 2-Delay B'cast
        mUpdateAst = True
        'update date/time aired
        If myTlAst.iStatus <= 1 Or myTlAst.iStatus = 9 Or myTlAst.iStatus = 10 Then
            ilAstStatus = myTlAst.iStatus
        Else
            ilAstStatus = 1
        End If
        '7639
        ilAstStatus = ilAstStatus + ilAddToStatus
        SQLQuery = "UPDATE ast SET astCPStatus = 1, astStatus = " & ilAstStatus & ", "
        SQLQuery = SQLQuery & "astAirDate = '" & Format$(tmImportSpot(llImportLoop).sActualAirDate1, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "astAirTime = '" & Format$(tmImportSpot(llImportLoop).sActualAirTime1, sgSQLTimeForm) & "'"
        '7639
        SQLQuery = SQLQuery & ", astcpfCode = " & myTlAst.lCpfCode
        '6158
        SQLQuery = SQLQuery & " WHERE (astCode = " & llAstCode & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            mUpdateAst = False
            Err.Raise ERRORSQL, "mUpdateAst", "Problem in mUpdateAst"
        End If
    Else
         'update status as not aired
         '6158 skip here and catch with the other 'missed'  7265 added it back in!
        ilAstStatus = 4
        SQLQuery = "UPDATE ast SET astCPStatus = 1, astStatus = " & ilAstStatus
        '7265 added here
        SQLQuery = SQLQuery & " WHERE (astCode = " & llAstCode & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            mUpdateAst = False
            Err.Raise ERRORSQL, "mUpdateAst", "Problem in mUpdateAst"
        End If
    End If
    'changes to status are passed along
    myTlAst.iStatus = ilAstStatus
'    SQLQuery = SQLQuery & " WHERE (astCode = " & llAstCode & ")"
'    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'        mUpdateAst = False
'        Err.Raise ERRORSQL, "mUpdateAst", "Problem in mUpdateAst"
'    End If
End Function
Private Function mInsertMulticast(tlAst As AST, slErrorMessage As String, tlMulticast() As OWNEDSTATIONS, Optional llOldCpf As Long = 0) As Boolean
    ' O- no error
    '7639 added blNewIsci
    Dim c As Integer
    Dim ilRet As Integer
    Dim blError As Boolean
    Dim ilRecLen As Integer
    
    blError = False
    ilRecLen = Len(tlAst)
    'for each multicast, run through insert the spot with minor changes
    For c = 0 To UBound(tlMulticast) - 1
        If tlMulticast(c).iSelected = 1 Then
            tlAst.lCode = 0
            tlAst.lAtfCode = tlMulticast(c).lAttCode
            tlAst.iShfCode = tlMulticast(c).iShttCode
            tlAst.lDatCode = mFindMatchingDat(tlAst.lDatCode, tlAst.lAtfCode)
            ilRet = btrInsert(hmAst, tlAst, ilRecLen, INDEXKEY0)
            If ilRet <> 0 Then
                'mark groupID so won't write 'posted' later. use blerror to write out error
                tlMulticast(c).iGroupID = -9
                blError = True
                If InStr(1, slErrorMessage, tlMulticast(c).sCallLetters) = 0 Then
                    slErrorMessage = slErrorMessage & " " & tlMulticast(c).sCallLetters
                End If
            '7639
            ElseIf llOldCpf > 0 Then
                mAddAltForIsci tlAst.lCode, tlAst.iAdfCode, llOldCpf
            End If
        End If
    Next c
    mInsertMulticast = Not blError
End Function

Private Function mUpdateCptt(blNoSpotsAired As Boolean, ilVefCode As Integer, llAtfCode As Long, slMondayFeedDate As String, ilStation As Integer) As Boolean
    'Created by D.S. June 2007  Modified Dan M 11/02/10 V70 new values in cptt added 2/25/2011
    'Set the CPTT week's value

    Dim slSuDate As String
    Dim ilStatus As Integer
    Dim llVeh As Long
    Dim ilAst As Integer
    'new values in cptt
    Dim ilSchdCount As Integer
    Dim ilAiredCount As Integer
    Dim ilPledgeCompliantCount As Integer
    Dim ilAgyCompliantCount As Integer
    Dim blRet As Boolean
    
    On Error GoTo ErrHand
    blRet = True
    slSuDate = DateAdd("d", 6, slMondayFeedDate)
    'Set any Not Aired to received as they are not exported
    For ilStatus = 0 To UBound(tgStatusTypes) Step 1
        If (tgStatusTypes(ilStatus).iPledged = 2) Then
            SQLQuery = "UPDATE ast SET "
            SQLQuery = SQLQuery & "astCPStatus = " & "1"    'Received
            SQLQuery = SQLQuery & " WHERE (astAtfCode = " & llAtfCode
            SQLQuery = SQLQuery & " AND astCPStatus = 0"
            SQLQuery = SQLQuery & " AND astStatus = " & tgStatusTypes(ilStatus).iStatus
            SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')" & ")"
            cnn.BeginTrans
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError smPathForgLogMsg, FORMNAME & "-mUpdateCptt"
                cnn.RollbackTrans
                mUpdateCptt = False
                Exit Function
            End If
            cnn.CommitTrans
        End If
    Next ilStatus
     '6158
    If Not mMGMissed(llAtfCode, slMondayFeedDate, slSuDate) Then
        blRet = False
    End If
    '7643 back in!
    '7265. This is no longer true!
    'ast's not found in xml from Marketron are marked as not aired 4
    SQLQuery = "UPDATE ast SET "
    SQLQuery = SQLQuery & "astCPStatus = 1, astStatus = 4"    'Received
    SQLQuery = SQLQuery & " WHERE (astAtfCode = " & llAtfCode
    SQLQuery = SQLQuery & " AND astCPStatus = 0"
    SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')" & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError smPathForgLogMsg, FORMNAME & "-mUpdateCptt"
        mUpdateCptt = False
        Exit Function
    End If
    'Determine if CPTTStatus should to set to 0=Partial or 1=Completed:  because of above code, will always be complete
    SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 0"
    SQLQuery = SQLQuery & " AND astAtfCode = " & llAtfCode
    SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF Then
        'Set CPTT as complete
        SQLQuery = "UPDATE cptt SET "
        llVeh = gBinarySearchVef(CLng(ilVefCode))
        If llVeh <> -1 Then
            If (tgVehicleInfo(llVeh).sVehType = "G") And (DateValue(slSuDate) > DateValue(Format$(gNow(), "m/d/yy"))) Then
                SQLQuery = SQLQuery & "cpttStatus = 0" & ", " 'Partial
                SQLQuery = SQLQuery & "cpttPostingStatus = 1" 'Partial
            Else
                If blNoSpotsAired Then
                    SQLQuery = SQLQuery & "cpttStatus = 2" & ", " 'Complete
                Else
                SQLQuery = SQLQuery & "cpttStatus = 1" & ", " 'Complete
                End If
                SQLQuery = SQLQuery & "cpttPostingStatus = 2"  'Complete
            End If
        Else
            If blNoSpotsAired Then
                SQLQuery = SQLQuery & "cpttStatus = 2" & ", " 'Complete
            Else
            SQLQuery = SQLQuery & "cpttStatus = 1" & ", " 'Complete
            End If
            SQLQuery = SQLQuery & "cpttPostingStatus = 2"  'Complete
        End If
        SQLQuery = SQLQuery & " WHERE cpttAtfCode = " & llAtfCode
        SQLQuery = SQLQuery & " AND (cpttStartDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError smPathForgLogMsg, FORMNAME & "-mUpdateCptt"
            mUpdateCptt = False
            Exit Function
        End If
    Else
        'Set CPTT as partial
        SQLQuery = "UPDATE cptt SET "
        SQLQuery = SQLQuery & "cpttStatus = 0" & ", " 'Partial
        SQLQuery = SQLQuery & "cpttPostingStatus = 1" 'Partial
        SQLQuery = SQLQuery & " WHERE cpttAtfCode = " & llAtfCode
        SQLQuery = SQLQuery & " AND (cpttStartDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError smPathForgLogMsg, FORMNAME & "-mUpdateCptt"
            mUpdateCptt = False
            Exit Function
        End If
    End If
    'Dan M V81 has new fields: how many spots to be aired? how many aired?  how many compliant? first step, get changes above into tmastInfo
    ilSchdCount = 0
    ilAiredCount = 0
    ilPledgeCompliantCount = 0
    ilAgyCompliantCount = 0
    gClearASTInfo False
    '7895
    mLoadCpPosting llAtfCode, ilVefCode, slMondayFeedDate, ilStation
    gGetAstInfo hmAst, tmCPDat(), tmAstInfo(), -1, False, False, True
    For ilAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
        'gIncSpotCounts tmAstInfo(ilAst).iPledgeStatus, tmAstInfo(ilAst).iStatus, tmAstInfo(ilAst).iCPStatus, tmAstInfo(ilAst).sTruePledgeDays, Format$(tmAstInfo(ilAst).sPledgeDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sAirDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sPledgeStartTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sTruePledgeEndTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
        gIncSpotCounts tmAstInfo(ilAst), ilSchdCount, ilAiredCount, ilPledgeCompliantCount, ilAgyCompliantCount
    Next ilAst
    'update dick's code.
'    SQLQuery = "Update cptt Set "
'    SQLQuery = SQLQuery & "cpttNoSpotsGen = " & ilSchdCount & ", "
'    SQLQuery = SQLQuery & "cpttNoSpotsAired = " & ilAiredCount & ", "
'    SQLQuery = SQLQuery & "cpttNoCompliant = " & ilCompliantCount & " "
'    SQLQuery = SQLQuery & " Where cpttCode = " & rst_Cptt!cpttCode
    SQLQuery = "Update cptt Set "
    SQLQuery = SQLQuery & "cpttNoSpotsGen = " & ilSchdCount & ", "
    SQLQuery = SQLQuery & "cpttNoSpotsAired = " & ilAiredCount & ", "
     'Dan M 8/11/14 Dick asked that this be added
    SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "cpttNoCompliant = " & ilPledgeCompliantCount & ", "
    SQLQuery = SQLQuery & "cpttAgyCompliant = " & ilAgyCompliantCount & " "
    SQLQuery = SQLQuery & " WHERE cpttAtfCode = " & llAtfCode
    SQLQuery = SQLQuery & " AND (cpttStartDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError smPathForgLogMsg, FORMNAME & "-mUpdateCptt"
        mUpdateCptt = False
        Exit Function
    End If
    gFileChgdUpdate "cptt.mkd", True
    mUpdateCptt = blRet
    Exit Function
ErrHand:
    'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, FORMNAME & "-mUpdateCptt"
    mUpdateCptt = False
End Function

Private Function mMGUpdateAlt(ilChoice As MGUpdates, llAst As Long, slDate As String, Optional llAlt As Long = 0, Optional ilAdv As Integer = 0) As Boolean
    Dim slSql As String
    
    slDate = Format$(slDate, sgSQLDateForm)
    Select Case ilChoice
        Case ADDMISSED
            slSql = "insert into alt (altAstCode,altMissedDate,altAdfCode,altMgDate) values (" & llAst & ",'" & slDate & "'," & ilAdv & ",'" & NODATE & "')"
        Case ADDMAKEGOOD
            slSql = "insert into alt (altLinkToAstCode,altMGDate,altAdfCode,altmisseddate) values (" & llAst & ",'" & slDate & "'," & ilAdv & ",'" & NODATE & "')"
        Case UPDATEMISSED
            slSql = "update alt set altMissedDate = '" & slDate & "', altAstCode = " & llAst & " where altcode = " & llAlt
        Case UPDATEMAKEGOOD
            slSql = "update alt set altMGDate = '" & slDate & "', altLinktoastcode = " & llAst & " where altcode = " & llAlt
        Case DELETEEITHER
            slSql = "delete from alt where altCode = " & llAlt
    End Select
    If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub ErrHand:
        gHandleError smPathForgLogMsg, FORMNAME & "-mMGUpdateAlt"
        mMGUpdateAlt = False
        Exit Function
    End If
    mMGUpdateAlt = True
    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, FORMNAME & "-mMGUpdateAlt"
    mMGUpdateAlt = False
End Function
Private Function mMGMissed(llAtfCode As Long, ByVal slMoDate As String, ByVal slSuDate As String) As Boolean
'returns false if error
    Dim blRet As Boolean
    Dim slSql As String
    Dim myRs As Recordset
    Dim llRet As Long
    Dim ilAdv As Integer
    
    blRet = True
    'this avoids mulitcast
    If imMarket > 0 Then
        ilAdv = 0
        slMoDate = Format$(slMoDate, sgSQLDateForm)
        slSuDate = Format$(slSuDate, sgSQLDateForm)
    On Error GoTo ERRORBOX
        slSql = "select astCode,astAirDate as myDate from ast where astatfCode = " & llAtfCode & " and astcpStatus = 0 and (astFeedDate >= '" & slMoDate & "' AND  astFeedDate <= '" & slSuDate & "') ORDER BY myDate"
        Set myRs = gSQLSelectCall(slSql)
        Do While Not myRs.EOF
            If Not IsNull(myRs!mydate) Then
                llRet = mMGMakeGoodExists(myRs!astCode, ilAdv)
                If llRet > 0 Then
                    If Not mMGUpdateAlt(UPDATEMISSED, myRs!astCode, myRs!mydate, llRet) Then
                        blRet = False
                    End If
                Else
                    If Not mMGUpdateAlt(ADDMISSED, myRs!astCode, myRs!mydate, , ilAdv) Then
                        blRet = False
                    End If
                End If
            Else
                blRet = False
                mSetResults "mMgMissed: air date missing for ast " & myRs!astCode, MESSAGERED
            End If
            myRs.MoveNext
        Loop
    End If
Cleanup:
    If Not myRs Is Nothing Then
        If (myRs.State And adStateOpen) <> 0 Then
            myRs.Close
        End If
        Set myRs = Nothing
    End If
    mMGMissed = blRet
    Exit Function
ERRORBOX:
    gHandleError smPathForgLogMsg, FORMNAME & "-mMGMissed"
    blRet = False
    GoTo Cleanup
End Function
Private Function mMGMissedExists(ilAdv As Integer) As Long
    'return altcode  For MakeGood, is there a missing to use?
    ' As opposed to mMGMakeGoodExists, we can pass adv in.
    Dim slSql As String
    Dim llMGCode As Long
    Dim ilMarket As Integer
    Dim ilStation As Integer
    Dim blFound As Boolean
    
    blFound = False
    llMGCode = 0
    slSql = "select astShfCode, altCode as mgCode from ast inner join alt on astCode = altAstCode where altmissedDate > '" & NODATE & "' and altmgDate = '" & NODATE & "' and altAdfCode = " & ilAdv & " ORDER BY ALTMISSEDDATE"
    Set rst = gSQLSelectCall(slSql)
    Do While Not rst.EOF
        llMGCode = rst!mgcode
        ilStation = rst!astShfCode
        If mMGIsMarket(ilStation) Then
            blFound = True
            Exit Do
        End If
        rst.MoveNext
    Loop
    If Not blFound Then
        llMGCode = 0
    End If
    mMGMissedExists = llMGCode
    
End Function
Private Function mMGMakeGoodExists(llAst As Long, ilAdv As Integer) As Long
    'return altcode, advcode.  For Missing, is there a makeGood ready to use?
    'get adv from lst.  Also, return alt if station is part of market.
    Dim slSql As String
    Dim llMGCode As Long
    Dim blFound As Boolean
    Dim ilStation As Integer
    
    blFound = False
    ilAdv = 0
    llMGCode = 0
    slSql = "select lstadfcode as adv,astShfCode from ast inner join lst on astlsfcode = lstcode where astcode = " & llAst
    Set rst = gSQLSelectCall(slSql)
    If Not rst.EOF Then
        ilAdv = rst!adv
        '6494 this is where program froze for Learfield
        slSql = "select astShfCode, altCode as mgCode from ast inner join alt on astCode = altLinkToAstCode where altMGDate > '" & NODATE & "' and altMissedDate = '" & NODATE & "' and altAdfCode = " & ilAdv & " ORDER BY ALTMISSEDDATE"
        'slSql = "SELECT AltCode as mgCode, altLinkToAst as  FROM alt WHERE altMGDate > '" & NODATE & "' AND altMissedDate = '" & NODATE & "' AND altAdfCode  = " & ilAdv & " ORDER BY ALTMGDATE "
        Set rst = gSQLSelectCall(slSql)
        Do While Not rst.EOF
            llMGCode = rst!mgcode
            ilStation = rst!astShfCode
            If mMGIsMarket(ilStation) Then
                blFound = True
                Exit Do
            End If
            rst.MoveNext
        Loop
        If Not blFound Then
            llMGCode = 0
        End If
   End If
    mMGMakeGoodExists = llMGCode
    
End Function
Private Function mMGIsMarket(ilStation As Integer) As Boolean
    Dim slSql As String
    Dim blRet As Boolean
    Dim ilMarket As Integer
    
    blRet = False
    If ilStation > 0 Then
        slSql = "Select  shttMktCode as market from shtt where shttcode = " & ilStation
        Set rsMarket = gSQLSelectCall(slSql)
        If Not rsMarket.EOF Then
            ilMarket = rsMarket!Market
            If ilMarket = imMarket Then
                blRet = True
            End If
        End If
    End If
    mMGIsMarket = blRet
    Exit Function
End Function
Private Function mMGIsMakeGood(llAstIndex As Long, blIsMulticast As Boolean) As Boolean
    Dim blRet As Boolean
    'Marketron is never makegood.
    blRet = False
    'this also assumes not multicast.
    If imMarket > 0 Then
        rsMG.Filter = "SourceAst = " & tmAstInfo(llAstIndex).lCode
        If Not rsMG.EOF Then
            blRet = True
        End If
    End If
    mMGIsMakeGood = blRet

End Function
Private Function mMGImportAll() As Boolean
    Dim blRet As Boolean

On Error GoTo errbox
    blRet = False
    If mMGImportSource() Then
        If mMGImportMakeGood() Then
            blRet = True
        End If
    End If
    mMGImportAll = blRet
    Exit Function
errbox:
    gHandleError smPathForgLogMsg, FORMNAME & "-mMGImportAll"
    blRet = False
    mMGImportAll = blRet
End Function
Private Function mMGImportSource() As Boolean
    Dim llPreviousAst As Long
    Dim myClone As Recordset
    Dim llSpot As Long
    Dim llImport As Long
    Dim blFound As Boolean
    Dim blRet As Boolean
    
On Error GoTo errbox
    blRet = True
    Set myClone = rsMG.Clone
   ' rsMG.Filter = adFilterNone
    'first time
    llPreviousAst = 0
    rsMG.MoveFirst
    Do While Not rsMG.EOF
        If rsMG!sourceast <> llPreviousAst Then
            llPreviousAst = rsMG!sourceast
            myClone.Filter = "SourceAst = " & llPreviousAst
            Do While Not myClone.EOF
                llSpot = myClone!ASTINDEX
                llImport = myClone!ImportIndex
                If mMGIsCompliant(llSpot, llImport) Then
                    myClone!found = True
                    blFound = True
                    If Not mMGUpdateAst(llSpot, llImport) Then
                        blRet = False
                    End If
                    Exit Do
                End If
                myClone.MoveNext
            Loop
            'no compliant?
            If Not blFound Then
                myClone.MoveFirst
                Do While Not myClone.EOF
                    llSpot = myClone!ASTINDEX
                    llImport = myClone!ImportIndex
                    If tmImportSpot(llImport).sStatusCode = "C" Then
                        myClone!found = True
                        blFound = True
                        If Not mMGUpdateAst(llSpot, llImport) Then
                            blRet = False
                        End If
                        Exit Do
                    End If
                    myClone.MoveNext
                Loop
                ' all are missed!
                If Not blFound Then
                    myClone.MoveFirst
                    myClone!found = True
                    llSpot = myClone!ASTINDEX
                    llImport = myClone!ImportIndex
                    If Not mMGUpdateAst(llSpot, llImport) Then
                        blRet = False
                    End If
                End If
            End If
        End If
        blFound = False
        rsMG.MoveNext
    Loop
Cleanup:
     If Not myClone Is Nothing Then
        If (myClone.State And adStateOpen) <> 0 Then
            myClone.Close
        End If
        Set myClone = Nothing
    End If
    mMGImportSource = blRet
    Exit Function
errbox:
    gHandleError smPathForgLogMsg, FORMNAME & "-mMGImportSource"
    blRet = False
    GoTo Cleanup
End Function
Private Function mMGImportMakeGood() As Boolean
    Dim blRet As Boolean
    Dim llSpotIndex As Long
    Dim llImportIndex As Long
    
On Error GoTo errbox
    rsMG.Filter = "Found = False"
    blRet = True
    Do While Not rsMG.EOF
        llSpotIndex = rsMG!ASTINDEX
        llImportIndex = rsMG!ImportIndex
        ' skip if 'missed'
        If tmImportSpot(llImportIndex).sStatusCode = "C" Then
            If Not mMGCreateMakeGood(llSpotIndex, llImportIndex) Then
                blRet = False
            End If
        End If
        tmImportSpot(llImportIndex).iFound = True
        rsMG.MoveNext
    Loop
    mMGImportMakeGood = blRet
    Exit Function
errbox:
    blRet = False
    mMGImportMakeGood = blRet
End Function
Private Function mMGCreateMakeGood(llSpotLoop As Long, llImportLoop As Long) As Boolean
    Dim blRet As Boolean
    Dim llLstCode As Long
    Dim llAlt As Long
    Dim llNewMG As Long
    Dim ilAdv As Integer
    Dim slDate As String
    
    blRet = True
On Error GoTo ERRORBOX
    ilAdv = tmAstInfo(llSpotLoop).iAdfCode
    slDate = tmImportSpot(llImportLoop).sActualAirDate1
    'have to create new lst.  This will have the new dates.
    llLstCode = mMGAddtoLst(tmAstInfo(llSpotLoop).lCode, slDate)
    llNewMG = mMGAddToAst(llSpotLoop, llImportLoop, llLstCode, tmAstInfo(llSpotLoop).lSdfCode)
    llAlt = mMGMissedExists(ilAdv)
    If llAlt > 0 Then
        If Not mMGUpdateAlt(UPDATEMAKEGOOD, llNewMG, slDate, llAlt) Then
            blRet = False
        End If
    Else
        If Not mMGUpdateAlt(ADDMAKEGOOD, llNewMG, slDate, , ilAdv) Then
            blRet = False
        End If
    End If
    mMGCreateMakeGood = blRet
    Exit Function
ERRORBOX:
    mMGCreateMakeGood = False
End Function
Private Function mMGIsCompliant(llSpotLoop As Long, llImportLoop As Long) As Boolean
    'return true if compliant
    Dim blRet As Boolean
    Dim slAllowedStartDate As String
    Dim slAllowedEndDate As String
    Dim slAllowedStartTime As String
    Dim slAllowedEndTime As String
    Dim ilAllowedDays(6) As Integer
    Dim slLineDays As String
    Dim ilGetLineParameters As Integer
    Dim ilCompliant As Integer
    Dim tlAst As ASTINFO
    
    blRet = False
    'can't be missed
    If tmImportSpot(llImportLoop).sStatusCode = "C" Then
        tlAst = tmAstInfo(llSpotLoop)
        tlAst.sAirDate = tmImportSpot(llImportLoop).sActualAirDate1
        tlAst.sAirTime = tmImportSpot(llImportLoop).sActualAirTime1
        slAllowedStartDate = ""
        slAllowedEndDate = ""
        slAllowedStartTime = ""
        slAllowedEndTime = ""
        slLineDays = ""
        ilGetLineParameters = 0
        ilCompliant = True
        'ilGetLineParameters = gGetLineParameters(True, tlAst, slAllowedStartDate, slAllowedEndDate, slAllowedStartTime, slAllowedEndTime, ilAllowedDays(), ilCompliant)
        ilGetLineParameters = gGetAgyCompliant(tlAst, slAllowedStartDate, slAllowedEndDate, slAllowedStartTime, slAllowedEndTime, ilAllowedDays(), ilCompliant)
        If ilCompliant = True Then
            blRet = True
        End If
    End If
    mMGIsCompliant = blRet
End Function
Private Function mMGAddtoLst(llAstCode As Long, slAirDate As String) As Long
  'copy lst connected to original ast. Change status (makeGood), date aired.
    Dim slProd As String
    Dim slCart As String
    Dim slISCI As String
    Dim ilAdf As Integer
    Dim ilAdfCode As Integer
    Dim llCntrNo As Long
    Dim llLineVef As Long
    Dim ilLineNo As Integer
    Dim ilAgfCode As Integer
    Dim ilPriceType As Integer
    Dim ilStatus As Integer
    Dim ilLoop As Integer
    Dim llLst As Long
    Dim llSdfCode As Long
    Dim llTemp As Long
    Dim rst_Temp As ADODB.Recordset
    Dim slSqlAirDate As String
    ReDim ilDay(0 To 6) As Integer
    Dim ilVefCode As Integer
    Dim ilLen As Integer
    
    mMGAddtoLst = 0
    slSqlAirDate = Format(slAirDate, sgSQLDateForm)
    On Error GoTo ErrHand
    
    SQLQuery = "SELECT astLsfCode FROM ast where astCode = " & llAstCode
    Set rst_Temp = gSQLSelectCall(SQLQuery)
    If Not rst_Temp.EOF Then
        llTemp = rst_Temp!astLsfCode
    End If
    
    SQLQuery = "SELECT * FROM Lst where lstCode = " & llTemp
    Set rst_Temp = gSQLSelectCall(SQLQuery)
    If rst_Temp.EOF Then
        'Error condition
    End If
    
    slProd = rst_Temp!lstProd
    slCart = rst_Temp!lstCart
    slISCI = rst_Temp!lstISCI
    ilAdfCode = rst_Temp!lstAdfCode
    llCntrNo = rst_Temp!lstCntrNo
    ilLineNo = rst_Temp!lstLineNo
    ilAgfCode = rst_Temp!lstAgfCode
    ilPriceType = rst_Temp!lstPriceType
    llLineVef = rst_Temp!lstLnVefCode
    llSdfCode = rst_Temp!lstSdfCode
    ilVefCode = rst_Temp!lstLogVefCode
    ilLen = rst_Temp!lstLen
    For ilLoop = 0 To 6 Step 1
        ilDay(ilLoop) = 0
    Next ilLoop
    ilStatus = ASTEXTENDED_MG
    ilDay(gWeekDayLong(gDateValue(slAirDate))) = 1
    SQLQuery = "INSERT INTO lst (lstCode, lstType, lstSdfCode, lstCntrNo, "
    SQLQuery = SQLQuery & "lstAdfCode, lstAgfCode, lstProd, "
    SQLQuery = SQLQuery & "lstLineNo, lstLnVefCode, lstStartDate, "
    SQLQuery = SQLQuery & "lstEndDate, lstMon, lstTue, "
    SQLQuery = SQLQuery & "lstWed, lstThu, lstFri, "
    SQLQuery = SQLQuery & "lstSat, lstSun, lstSpotsWk, "
    SQLQuery = SQLQuery & "lstPriceType, lstPrice, lstSpotType, "
    SQLQuery = SQLQuery & "lstLogVefCode, lstLogDate, lstLogTime, "
    SQLQuery = SQLQuery & "lstDemo, lstAud, lstISCI, "
    SQLQuery = SQLQuery & "lstWkNo, lstBreakNo, lstPositionNo, "
    SQLQuery = SQLQuery & "lstSeqNo, lstZone, lstCart, "
    SQLQuery = SQLQuery & "lstCpfCode, lstCrfCsfCode, lstStatus, "
    SQLQuery = SQLQuery & "lstLen, lstUnits, lstCifCode, "
    SQLQuery = SQLQuery & "lstAnfCode, lstEvtIDCefCode, lstSplitNetwork, "
    'SQLQuery = SQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, lstUnused)"
    SQLQuery = SQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, "
    SQLQuery = SQLQuery & "lstLnStartTime, lstLnEndTime, lstUnused)"
    SQLQuery = SQLQuery & " VALUES (" & "Replace" & ", " & 2 & ", " & llSdfCode & ", " & llCntrNo & ", "
    SQLQuery = SQLQuery & ilAdfCode & ", " & ilAgfCode & ", '" & slProd & "', "
    SQLQuery = SQLQuery & ilLineNo & ", " & llLineVef & ", '" & slSqlAirDate & "', "
    SQLQuery = SQLQuery & "'" & slSqlAirDate & "', " & ilDay(0) & ", " & ilDay(1) & ", "
    SQLQuery = SQLQuery & ilDay(2) & ", " & ilDay(3) & ", " & ilDay(4) & ", "
    SQLQuery = SQLQuery & ilDay(5) & ", " & ilDay(6) & ", " & 0 & ", "
    SQLQuery = SQLQuery & ilPriceType & ", " & 0 & ", " & 5 & ", "
    SQLQuery = SQLQuery & ilVefCode & ", '" & slSqlAirDate & "', '" & slSqlAirDate & "', "
    SQLQuery = SQLQuery & "'" & 0 & "', " & 0 & ", '" & slISCI & "', "
    SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", '" & "" & "', '" & slCart & "', "
    SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & ilStatus & ", "
    SQLQuery = SQLQuery & ilLen & ", " & 0 & ", " & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", " & 0 & ", '" & "N" & "', "
    'SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", '" & "N" & "', " & 0 & ", '" & "" & "'" & ")"
    SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", '" & "N" & "', " & 0 & ", "
    SQLQuery = SQLQuery & "'" & Format("12am", sgSQLTimeForm) & "', '" & Format("12am", sgSQLTimeForm) & "', '" & "" & "'" & ")"
    llLst = gInsertAndReturnCode(SQLQuery, "lst", "lstCode", "Replace")
    rst_Temp.Close
    mMGAddtoLst = llLst
    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, FORMNAME & "-mMGAddToLst"
    mMGAddtoLst = 0
End Function
Private Function mMGAddToAst(llSpotLoop As Long, llImportLoop As Long, llLstCode As Long, llSdfCode As Long) As Long
    Dim llAttCode As Long
    Dim ilShttCode As Integer
    Dim ilVefCode As Integer
    Dim slAirDate As String
    Dim slAirTime As String
    Dim ilStatus As Integer
    Dim ilAdfCode As Integer
    Dim llDATCode As Long
    Dim llCpfCode As Long
    Dim llRsfCode As Long
    Dim slStationCompliant As String
    Dim slAgencyCompliant As String
    Dim slAffidavitSource As String
    Dim llAst As Long
    
On Error GoTo ErrHand
    ilStatus = ASTEXTENDED_MG
    With tmAstInfo(llSpotLoop)
        llAttCode = .lAttCode
        ilShttCode = .iShttCode
        ilVefCode = .iVefCode
        ilAdfCode = .iAdfCode
        llDATCode = .lDatCode
        llCpfCode = .lCpfCode
        llRsfCode = .lRRsfCode
    End With
    With tmImportSpot(llImportLoop)
        slAirDate = .sActualAirDate1
        slAirTime = .sActualAirTime1
    End With
    slStationCompliant = ""
    slAgencyCompliant = ""
    slAffidavitSource = ""
    SQLQuery = "INSERT INTO ast"
    SQLQuery = SQLQuery + "(astcode,astAtfCode, astShfCode, astVefCode, "
    SQLQuery = SQLQuery + "astSdfCode, astLsfCode, astAirDate, astAirTime, "
    '12/13/13: Support New AST layout
    'SQLQuery = SQLQuery + "astStatus, astCPStatus, astFeedDate, astFeedTime, astPledgeDate, astPledgeStartTime, astPledgeEndTime)"
    SQLQuery = SQLQuery + "astStatus, astCPStatus, astFeedDate, astFeedTime, "
    SQLQuery = SQLQuery + "astAdfCode, astDatCode, astCpfCode, astRsfCode, astStationCompliant, astAgencyCompliant, astAffidavitSource, astUstCode)"
    SQLQuery = SQLQuery + " VALUES "
    SQLQuery = SQLQuery + "( Replace , " & llAttCode & ", " & ilShttCode & ", "
    SQLQuery = SQLQuery & ilVefCode & ", " & llSdfCode & ", " & llLstCode & ", "
    SQLQuery = SQLQuery + "'" & Format$(slAirDate, sgSQLDateForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "', "
    'SQLQuery = SQLQuery & ilStatus & ", " & "1" & ", '" & Format$(slAirDate, sgSQLDateForm) & "', "
    'SQLQuery = SQLQuery & "'" & Format$(slAirTime, sgSQLTimeForm) & "', '" & Format$(slAirDate, sgSQLDateForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "')"
    SQLQuery = SQLQuery & ilStatus & ", " & "1" & ", '" & Format$(slAirDate, sgSQLDateForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & ilAdfCode & ", " & llDATCode & ", " & llCpfCode & ", " & llRsfCode & ", "
    SQLQuery = SQLQuery & "'" & slStationCompliant & "', '" & slAgencyCompliant & "', '" & slAffidavitSource & "', " & igUstCode & ")"
    llAst = gInsertAndReturnCode(SQLQuery, "ast", "astCode", "Replace")
    mMGAddToAst = llAst
    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, FORMNAME & "-mMGAddToAst"
    mMGAddToAst = 0
End Function
Private Function mMGUpdateAst(llSpotLoop As Long, llImportLoop As Long) As Boolean
    Dim tlMyAst As AST
    
    'in the code I copied, lcode and ivefcode were set by local variables.
    ' virtually unneeded. Only so can use mUpdateAst.  All we care about is air dates, which are tied to llImportLoop
    With tlMyAst
        .lCode = tmAstInfo(llSpotLoop).lCode
        .iStatus = tmAstInfo(llSpotLoop).iStatus
        .iVefCode = tmAstInfo(llSpotLoop).iVefCode
        .iCPStatus = 1
        .lSdfCode = tmAstInfo(llSpotLoop).lSdfCode
        .lLsfCode = tmAstInfo(llSpotLoop).lLstCode
        '12/9/13
        '.iPledgeStatus = tmAstInfo(llSpotLoop).iPledgeStatus
        gPackDate tmAstInfo(llSpotLoop).sFeedDate, .iFeedDate(0), .iFeedDate(1)
        gPackTime tmAstInfo(llSpotLoop).sFeedTime, .iFeedTime(0), .iFeedTime(1)
        '12/9/13
        'gPackDate tmAstInfo(llSpotLoop).sPledgeDate, .iPledgeDate(0), .iPledgeDate(1)
        'gPackTime tmAstInfo(llSpotLoop).sPledgeStartTime, .iPledgeStartTime(0), .iPledgeStartTime(1)
        'gPackTime tmAstInfo(llSpotLoop).sPledgeEndTime, .iPledgeEndTime(0), .iPledgeEndTime(1)
        .iAdfCode = tmAstInfo(llSpotLoop).iAdfCode
        .lDatCode = tmAstInfo(llSpotLoop).lDatCode
        .lCpfCode = tmAstInfo(llSpotLoop).lCpfCode
        .lRsfCode = tmAstInfo(llSpotLoop).lRRsfCode
        .sStationCompliant = ""
        .sAgencyCompliant = ""
        .sAffidavitSource = ""
        .iUstCode = igUstCode
    End With
    mUpdateAst llImportLoop, tlMyAst, 0
    'this may mean I never get error!
    mMGUpdateAst = True
    Exit Function
ErrHand:
    gHandleError smPathForgLogMsg, FORMNAME & "-mMGUpdateAst"
    mMGUpdateAst = False
End Function
Private Sub mMGFindImported()
    'the rsMG only contains asts that are duplicated in import file
    ' note that if ast doesn't exist in tmAst, I have an issue!
    Dim c As Integer
    Dim ilUpper As Integer
    Dim llAst As Long
    Dim myAsts As Dictionary
    Dim llPreviousAst As Long
    Dim ilIndex As Integer
    Dim ilAstIndex As Integer
    
    Set myAsts = New Dictionary
    llPreviousAst = 0
    ilUpper = UBound(tmImportSpot)
    'this will get us duplicate asts for recordset.  But doesn't add the first ast that was duplicated later.
    For c = 0 To ilUpper
        llAst = tmImportSpot(c).lAstCode
        If myAsts.Exists(llAst) Then
            For ilAstIndex = 0 To UBound(tmAstInfo) - 1
                If tmAstInfo(ilAstIndex).lCode = llAst Then
                    rsMG.AddNew Array("Found", "SourceAst", "ImportIndex", "astIndex"), Array("False", llAst, c, ilAstIndex)
                    Exit For
                End If
            Next ilAstIndex
        Else
            myAsts.Add llAst, c
        End If
    Next c
    'now we have to add the records that got put into dictionary the first time and subsequently were found to be duplicated
    If Not rsMG.EOF Then
        rsMG.MoveFirst
        Do While Not rsMG.EOF
            If llPreviousAst <> rsMG!sourceast Then
                llPreviousAst = rsMG!sourceast
                'should always exist.
                If myAsts.Exists(llPreviousAst) Then
                    ilIndex = myAsts(llPreviousAst)
                    For ilAstIndex = 0 To UBound(tmAstInfo) - 1
                        If tmAstInfo(ilAstIndex).lCode = llPreviousAst Then
                            rsMG.AddNew Array("Found", "SourceAst", "ImportIndex", "astIndex"), Array("False", llPreviousAst, ilIndex, ilAstIndex)
                            Exit For
                        End If
                    Next ilAstIndex
                End If
            End If
            rsMG.MoveNext
        Loop
    End If
End Sub
Private Function mPrepRecordset() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "SourceAst", adInteger
            .Append "ImportIndex", adInteger
            .Append "AstIndex", adInteger
            .Append "Found", adBoolean
        End With
    myRs.Open
    myRs("SourceAst").Properties("optimize") = True
    myRs.Sort = "SourceAst"
    Set mPrepRecordset = myRs
End Function
Private Sub mMGResetBeforeImporting()
    Dim llSpotLoop As Long
    Dim llAstCode As Long
    Dim llLstCode As Long
    Dim llMissed As Long
    Dim llAlt As Long
    Dim slSql As String
    
    For llSpotLoop = 0 To UBound(tmAstInfo) - 1
        DoEvents
        With tmAstInfo(llSpotLoop)
            llAstCode = .lCode
            If .iCPStatus = 1 Then
                'missed? find if in alt and delete modify, then change posting status in ast
                If .iStatus = 4 Then
                    If Not mMGDeleteMissed(llAstCode) Then
                        Err.Raise 1001, "mMGResetBeforeImporting", Err.Description
                    End If
                    .iCPStatus = 0
                    .iStatus = .iPledgeStatus
                    slSql = "update ast set astCpStatus = 0 , astStatus = " & .iStatus & " where astcode = " & llAstCode
                    If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                        Err.Raise 1001, "mMGResetBeforeImporting", Err.Description
                    End If
                ' if makegood, we have a lot to do: delete ast, lst, modify or delete alt. Otherwise, just update posting status
                Else
                    slSql = "Select altCode, altAstCode as missed from alt where altlinkToAstCode = " & llAstCode
                    Set rst = gSQLSelectCall(slSql)
                    If Not rst.EOF Then
                        llAlt = rst!altCode
                        llMissed = rst!missed
                        llLstCode = .lLstCode
                        If Not mMGDeleteMakeGood(llAstCode, llLstCode, llMissed, llAlt) Then
                            Err.Raise 1001, "mMGResetBeforeImporting", Err.Description
                        End If
                    Else
                        .iCPStatus = 0
                        .iStatus = .iPledgeStatus
                        slSql = "update ast set astCpStatus = 0 , astStatus = " & .iStatus & " where astcode = " & llAstCode
                        If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                            Err.Raise 1001, "mMGResetBeforeImporting", Err.Description
                        End If
                    End If
                End If
            End If
        End With
    Next llSpotLoop
End Sub
Private Function mMGDeleteMissed(llAstCode As Long) As Boolean
    Dim slSql As String
    Dim llMGCode As Long
    Dim llAltCode As Long
    Dim blRet As Boolean
    
    blRet = True
On Error GoTo ERRORBOX
    slSql = "SELECT AltCode , altLinkToAstcode as MGCode  FROM alt WHERE altAstCode = " & llAstCode
    Set rst = gSQLSelectCall(slSql)
    If Not rst.EOF Then
        llAltCode = rst!altCode
        llMGCode = rst!mgcode
        If llMGCode > 0 Then
            If Not mMGUpdateAlt(UPDATEMISSED, 0, NODATE, llAltCode) Then
                blRet = False
            End If
        Else
            If Not mMGUpdateAlt(DELETEEITHER, 0, NODATE, llAltCode) Then
                blRet = False
            End If
        End If
    End If
    mMGDeleteMissed = blRet
    Exit Function
ERRORBOX:
    mMGDeleteMissed = False
End Function
Private Function mMGDeleteMakeGood(llAst As Long, llLst As Long, llMissed As Long, llAlt As Long) As Boolean
    'delete previous asts and lsts. Delete or update alt as needed (if missed already found, update)
    Dim slSql As String
    Dim blRet As Boolean
    
    blRet = True
    slSql = "Delete from ast where astcode = " & llAst
    If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
        GoTo ERRSQL
    End If
    slSql = "Delete from lst where lstcode = " & llLst
    If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
        GoTo ERRSQL
    End If
    If llMissed > 0 Then
        If Not mMGUpdateAlt(UPDATEMAKEGOOD, 0, NODATE, llAlt) Then
            blRet = False
        End If
    Else
        If Not mMGUpdateAlt(DELETEEITHER, 0, "", llAlt) Then
            blRet = False
        End If
    End If
    mMGDeleteMakeGood = blRet
    Exit Function
ERRSQL:
    gHandleError smPathForgLogMsg, FORMNAME & "-mMGDeleteMakeGood"
    mMGDeleteMakeGood = False
End Function

Private Function mSetImportClass(slIniPath As String) As Boolean
    'return true if values exist in ini file, not if created myExport
    Dim slRet As String
    Dim blRet As Boolean
    Dim slServicePage As String
    Dim slHost As String
    Dim slPassword As String
    Dim slUserName As String
    '7878
'    Dim myXml As MSXML2.DOMDocument
'    Dim myElem As MSXML2.IXMLDOMElement
    Dim blReturnAll As Boolean
    '7539
    Dim slProxyUrl As String
    Dim slProxyPort As String
    Dim slProxyTestUrl As String
    Dim blUseSecure As Boolean
    Dim blUseProxySecure As Boolean
    '8925
    Dim slPort As String
    
    blRet = False
    slUserName = ""
    slPassword = ""
    slProxyUrl = ""
    slProxyPort = ""
    slProxyTestUrl = ""
    blUseSecure = False
    blUseProxySecure = False
    slPort = ""
On Error GoTo ERRORBOX
    '7539 this ini value will stop using this class; in case doesn't work for client
    gLoadFromIni "MARKETRON", "UseOldMethod", slIniPath, slRet
    If UCase(slRet) = "TRUE" Then
        mSetImportClass = True
        Exit Function
    End If
'7539 no longer true
'    'not sure if this code will work with proxy; so if proxy set, don't run.
'    gLoadFromIni "MARKETRON", "ProxyServer", slIniPath, slRet
'    If slRet <> "Not Found" Then
'        mSetImportClass = True
'        Exit Function
'    End If
    'treat as if not needed
    gLoadFromIni "MARKETRON", "ProxyServer", slIniPath, slRet
    If slRet <> "Not Found" Then
        slProxyUrl = slRet
        'must have port defined also
        gLoadFromIni "MARKETRON", "ProxyPort", slIniPath, slRet
        If slRet <> "Not Found" Then
            slProxyPort = slRet
            gLoadFromIni "MARKETRON", "ProxyTestURL", slIniPath, slRet
            If slRet <> "Not Found" Then
                slProxyTestUrl = slRet
            End If
            gLoadFromIni "MARKETRON", "UseSecureProxy", slIniPath, slRet
            If UCase(slRet) = "TRUE" Then
                blUseProxySecure = True
            End If
        Else
            slProxyUrl = ""
        End If
    End If
    gLoadFromIni "MARKETRON", "UseSecure", slIniPath, slRet
    If UCase(slRet) = "TRUE" Then
        blUseSecure = True
    End If
    'will assume 'available', which is the normal download.
    gLoadFromIni "MARKETRON", "OrderStatus", slIniPath, slRet
    If slRet = "Not Found" Then
        slRet = "Available"
    End If
    If slRet = "Received" Then
        blReturnAll = True
    Else
        blReturnAll = False
    End If
    '8925 port was missing.
    gLoadFromIni "MARKETRON", "Port", slIniPath, slRet
    If slRet <> "Not Found" Then
        slPort = slRet
    End If
    'here on out is needed to continue
    gLoadFromIni "MARKETRON", "Host", slIniPath, slRet
    If slRet = "Not Found" Then
        slRet = ""
    End If
    If Len(slRet) = 0 Then
        mSetImportClass = blRet
        Exit Function
    End If
    slHost = slRet
    gLoadFromIni "MARKETRON", "WebServiceRcvURL", slIniPath, slRet
    If slRet = "Not Found" Then
        slRet = ""
    End If
    If Len(slRet) = 0 Then
        mSetImportClass = blRet
        Exit Function
    End If
    slServicePage = slRet
    gLoadFromIni "MARKETRON", "Authentication", slIniPath, slRet
    If slRet = "Not Found" Then
        slRet = ""
    End If
    If Len(slRet) = 0 Then
        mSetImportClass = blRet
        Exit Function
    End If
    blRet = True
'    Set myXml = New MSXML2.DOMDocument
'    If Not myXml.loadXML(slRet) Then
'        mSetImportClass = blRet
'        Exit Function
'    End If
'    Set myElem = myXml.selectSingleNode("//Username")
'    If Not myElem Is Nothing Then
'        slUserName = myElem.Text
'    End If
'    Set myElem = myXml.selectSingleNode("//Password")
'    If Not myElem Is Nothing Then
'        slPassword = myElem.Text
'    End If
    '7878
    slUserName = gParseXml(slRet, "Username", 0)
    slPassword = gParseXml(slRet, "Password", 0)
    If Len(slPassword) > 0 And Len(slUserName) > 0 Then
        Set myImport = New CMarketron
        With myImport
            If StrComp(slHost, "Test", vbTextCompare) = 0 Then
                .isTest = True
            End If
            .SoapUrl = slHost
            .ImportPage = slServicePage
            'couldn't set address
            If Len(.ErrorMessage) > 0 Then
               ' mSetResults "Couldn't set secondary calls to Marketron", MESSAGEBLACK
                '7539 changed this and set blret = false
                myErrors.WriteWarning "Couldn't set Marketron importer: " & myImport.ErrorMessage & "."
                Set myImport = Nothing
                blRet = False
                GoTo Cleanup
            End If
            .Password = slPassword
            .UserName = slUserName
            .ReturnAll = blReturnAll
            .UseSecure = blUseSecure
            '8925
            If Len(slPort) > 0 Then
                .Port = slPort
            End If
            If Len(slProxyUrl) > 0 Then
                If Not .Proxy(slProxyUrl, slProxyPort, blUseProxySecure, slProxyTestUrl) Then
                    myErrors.WriteWarning "Couldn't use defined proxy in xml.ini"
                    blRet = False
                    Set myImport = Nothing
                    GoTo Cleanup
                End If
            End If
            .LogPath = .CreateLogName(sgMsgDirectory & FILEDEBUG)
        End With
    End If
Cleanup:
'    Set myElem = Nothing
'    Set myXml = Nothing
    mSetImportClass = blRet
    Exit Function
ERRORBOX:
    blRet = False
    myErrors.WriteError "mSetImportClass-" & Err.Description
    GoTo Cleanup
End Function

Private Sub mnuDebug_Click()
     mnuDebug.Checked = True
    If Not myImport Is Nothing Then
        myImport.LogPath = myImport.CreateLogName(sgMsgDirectory & FILEDEBUG)
    End If
End Sub

Private Sub mnuImport_Click()
    mnuImport.Checked = Not mnuImport.Checked
End Sub

Private Sub mnuRemote_Click()
    mnuRemote.Checked = Not mnuRemote.Checked
End Sub

Private Function mFindMatchingDat(llFromDatCode As Long, llToAttCode) As Long
    Dim blMatch As Boolean
    Dim Dat1_rst As ADODB.Recordset
    Dim Dat2_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    mFindMatchingDat = 0
    If llFromDatCode <= 0 Then
        Exit Function
    End If
    SQLQuery = "Select * from Dat where datCode = " & llFromDatCode
    Set Dat1_rst = gSQLSelectCall(SQLQuery)
    If Not Dat1_rst.EOF Then
        blMatch = True
        SQLQuery = "Select * from Dat where datAtfCode = " & llToAttCode
        Set Dat2_rst = gSQLSelectCall(SQLQuery)
        Do While Not Dat2_rst.EOF
            blMatch = True
            If Dat1_rst!datFdMon <> Dat2_rst!datFdMon Then
                blMatch = False
            End If
            If Dat1_rst!datFdTue <> Dat2_rst!datFdTue Then
                blMatch = False
            End If
            If Dat1_rst!datFdWed <> Dat2_rst!datFdWed Then
                blMatch = False
            End If
            If Dat1_rst!datFdThu <> Dat2_rst!datFdThu Then
                blMatch = False
            End If
            If Dat1_rst!datFdFri <> Dat2_rst!datFdFri Then
                blMatch = False
            End If
            If Dat1_rst!datFdSat <> Dat2_rst!datFdSat Then
                blMatch = False
            End If
            If Dat1_rst!datFdSun <> Dat2_rst!datFdSun Then
                blMatch = False
            End If
            If gTimeToLong(Format$(Dat1_rst!datFdStTime, sgShowTimeWSecForm), False) <> gTimeToLong(Format$(Dat2_rst!datFdStTime, sgShowTimeWSecForm), False) Then
                blMatch = False
            End If
            If gTimeToLong(Format$(Dat1_rst!datFdEdTime, sgShowTimeWSecForm), True) <> gTimeToLong(Format$(Dat2_rst!datFdEdTime, sgShowTimeWSecForm), True) Then
                blMatch = False
            End If
            If Dat1_rst!datFdStatus <> Dat2_rst!datFdStatus Then
                blMatch = False
            End If
            If Dat1_rst!datPdMon <> Dat2_rst!datPdMon Then
                blMatch = False
            End If
            If Dat1_rst!datPdTue <> Dat2_rst!datPdTue Then
                blMatch = False
            End If
            If Dat1_rst!datPdWed <> Dat2_rst!datPdWed Then
                blMatch = False
            End If
            If Dat1_rst!datPdThu <> Dat2_rst!datPdThu Then
                blMatch = False
            End If
            If Dat1_rst!datPdFri <> Dat2_rst!datPdFri Then
                blMatch = False
            End If
            If Dat1_rst!datPdSat <> Dat2_rst!datPdSat Then
                blMatch = False
            End If
            If Dat1_rst!datPdSun <> Dat2_rst!datPdSun Then
                blMatch = False
            End If
            If gTimeToLong(Format$(Dat1_rst!datPdStTime, sgShowTimeWSecForm), False) <> gTimeToLong(Format$(Dat2_rst!datPdStTime, sgShowTimeWSecForm), False) Then
                blMatch = False
            End If
            If gTimeToLong(Format$(Dat1_rst!datPdEdTime, sgShowTimeWSecForm), True) <> gTimeToLong(Format$(Dat2_rst!datPdEdTime, sgShowTimeWSecForm), True) Then
                blMatch = False
            End If
            If Dat1_rst!datPdDayFed <> Dat2_rst!datPdDayFed Then
                blMatch = False
            End If
            If Dat1_rst!datAirPlayNo <> Dat2_rst!datAirPlayNo Then
                blMatch = False
            End If
            If blMatch Then
                mFindMatchingDat = Dat2_rst!datCode
                Exit Function
            End If
            Dat2_rst.MoveNext
        Loop
    End If
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "Import Marketron-mFindMatchingDat"
    Exit Function
End Function
Private Function mAdjustISCIAsNeeded(llSpotLoop As Long, slImportISCI As String) As Long
    'return CpfCode if import isci is different; otherwise 0
    'add or update cpf with new isci; then create alt for old isci
    Dim slAstISCI As String
    Dim llRet As Long
    
    llRet = 0
    'Dan M 7/29/15 don't do 7639 until later: reports need to be fixed. To restore: lose goto line below
    '8018 comment out
   ' GoTo Cleanup
On Error GoTo ERRORBOX
    If llSpotLoop <= UBound(tmAstInfo) Then
        If tmAstInfo(llSpotLoop).iRegionType = 0 Then
            slAstISCI = tmAstInfo(llSpotLoop).sISCI
        Else
            slAstISCI = tmAstInfo(llSpotLoop).sRISCI
        End If
        If slAstISCI <> slImportISCI Then
            SQLQuery = "select cpfCode from cpf_Copy_Prodct_ISCI where cpfisci = '" & slImportISCI & "'"
            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                llRet = rst!cpfCode
            Else
                SQLQuery = "INSERT into cpf_Copy_Prodct_ISCI (cpfCode,cpfName,cpfIsci,cpfCreative,cpfRotEndDate,cpfsifCode) VALUES (Replace,'','" & slImportISCI & "','','" & NODATE & "',0)"
                llRet = gInsertAndReturnCode(SQLQuery, "cpf_Copy_Prodct_ISCI", "cpfCode", "Replace")
            End If
            'now create alt
            If llRet <> 0 Then
                mAddAltForIsci tmAstInfo(llSpotLoop).lCode, tmAstInfo(llSpotLoop).iAdfCode, tmAstInfo(llSpotLoop).lCpfCode
            Else
                mSetResults "warning!  Issue in mAdjustISCIAsNeeded. Please see log", MESSAGERED
                myErrors.WriteWarning "sql call invalid:" & SQLQuery, False
            End If

        End If
    Else
        mSetResults "warning!  Issue in mAdjustISCIAsNeeded. Please see log", MESSAGERED
        myErrors.WriteWarning "index passed: " & llSpotLoop & " is not valid.  Max: " & UBound(tmAstInfo), False
    End If
Cleanup:
    mAdjustISCIAsNeeded = llRet
    Exit Function
ERRORBOX:
    mSetResults "warning!  Issue in mAdjustISCIAsNeeded. Please see log", MESSAGERED
    gHandleError smPathForgLogMsg, FORMNAME & "-mAdjustISCIAsNeeded"
    llRet = 0
    GoTo Cleanup
End Function
Private Sub mAddAltForIsci(llAstCode As Long, ilAdfCode As Integer, llCpfCode As Long)
    'note that the tmastInfo has never been updated with any new isci data
    SQLQuery = "insert into alt (altAstCode,altMissedDate,altAdfCode,altMgDate,altCpfCode) values (" & llAstCode & ",'" & NODATE & "'," & ilAdfCode & ",'" & NODATE & "'," & llCpfCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub Errbox:
        mSetResults "warning!  Issue in mAddAltForIsci. Please see log", MESSAGERED
        gHandleError smPathForgLogMsg, FORMNAME & "-mAddAltForIsci"
        Exit Sub
    End If
    Exit Sub
errbox:
    mSetResults "warning!  Issue in mAddAltForIsci. Please see log", MESSAGERED
    gHandleError smPathForgLogMsg, FORMNAME & "-mAddAltForIsci"

End Sub
Private Function mReImport() As Integer
    Dim ilSpotCount As Integer
    Dim slImportPath As String
    Dim blNotAllImported As Boolean
    Dim slStartDate As String
    Dim slSignature As String
    Dim slVehicle As String
    Dim slStation As String
    
    Screen.MousePointer = vbHourglass
    imExporting = True
    lbcMsg.Clear
    gOpenMKDFile hmAst, "Ast.Mkd"
    Set myEnt = New CENThelper
    With myEnt
        .User = igUstCode
        .TypeEnt = Importposted3rdparty
        .ThirdParty = Vendors.NetworkConnect
        .ErrorLog = smPathForgLogMsg
    End With
    smCurrentAgreementInfo = ""
    myEnt.fileName = "ReImportMarketron.Txt"
    slImportPath = sgImportDirectory
    Set lmUnMatchedSpots = New Dictionary
    If mProcessXml(slImportPath & "ReImportMarketron.Txt", slStartDate, slSignature, slVehicle, slStation) Then
        '7266
        smCurrentAgreementInfo = slVehicle & "-" & slStation & "-" & slStartDate
        ilSpotCount = mProcessSpots(slStartDate, slVehicle, slStation, slSignature, blNotAllImported)
        If ilSpotCount > 0 And Not blNotAllImported Then
            sgReImportStatus = "Marketron Import: Successful"
        ElseIf ilSpotCount = 0 Then
            sgReImportStatus = "Marketron Import: No Spots Matched"
        ElseIf ilSpotCount < 0 Then
            sgReImportStatus = "Marketron Import: Unable to find Agreement"
        Else
            sgReImportStatus = "Marketron Import: Partially processed as not all spots matched"
        End If
    Else
        sgReImportStatus = "Marketron Import Failed"
    End If
    mCloseAst
    imExporting = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Set myEnt = Nothing
    tmcTerminate.Enabled = True
End Function
Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Unload frmImportMarketron
End Sub
Public Function mFixStation(slCall As String, slBand As String) As String
    '8704
    Dim ilPos As Integer
    Dim slRet As String
    Dim slTemp As String
    
    slRet = slCall & "-" & slBand
On Error GoTo errbox
    ilPos = InStr(slCall, "-")
    If ilPos > 0 Then
        ilPos = InStr(slCall, "_")
        If ilPos > 0 Then
            slTemp = Mid(slCall, 1, ilPos - 1)
            If InStr(slTemp, "-") > 0 Then
                slRet = slTemp
            End If
        End If
    End If
    mFixStation = slRet
    Exit Function
errbox:
    mFixStation = slRet
End Function
Private Function mReturnAdv(slISCI As String, ilSpotLength As Integer, slMondayDate As String) As Integer
    Dim ilCount As Integer
    Dim ilRet As Integer
    Dim slSql As String
    Dim slSundayDate As String
    Dim myRst As ADODB.Recordset
    Dim llDate As Long
    
    llDate = gDateValue(slMondayDate) + 6
    slSundayDate = Format$(llDate, sgSQLDateForm)
    slMondayDate = Format(slMondayDate, sgSQLDateForm)
    ilSpotLength = 0
    slSql = "Select cifLen,adfCode from CIF_Copy_Inventory inner join  ADF_Advertisers  on cifAdfCode = adfCode inner join CPF_Copy_Prodct_ISCI on cifcpfcode = cpfcode where cpfISCI = '" & Trim(slISCI) & "'  AND cifRotStartDate <= '" & slSundayDate & "' and cifRotEnddate >= '" & slMondayDate & "'"
    Set myRst = gSQLSelectCall(slSql)
    Do While (Not myRst.EOF)
        ilCount = ilCount + 1
        ilRet = myRst("adfCode")
        ilSpotLength = myRst("cifLen")
        myRst.MoveNext
    Loop
    'almost impossible to have 2 different advertisers or 2 lengths at this point (when testing against one week)
    If ilCount > 1 Then
        ilRet = 0
        ilSpotLength = 0
    End If
    mReturnAdv = ilRet
End Function
