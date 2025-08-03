VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmImportiPump 
   Caption         =   "Import iPump"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3210
      TabIndex        =   3
      Top             =   4305
      Width           =   1575
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   1185
      TabIndex        =   2
      Top             =   4305
      Width           =   1890
   End
   Begin VB.ListBox lbcMsg 
      Height          =   2205
      Left            =   180
      TabIndex        =   1
      Top             =   1320
      Width           =   5790
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
      Left            =   5760
      ScaleHeight     =   225
      ScaleWidth      =   1005
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   1035
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5850
      Top             =   4230
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
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   165
      TabIndex        =   4
      Top             =   1035
      Width           =   5790
   End
End
Attribute VB_Name = "frmImportiPump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bmExporting As Boolean
Private bmTerminate As Boolean
Private oMyFileObj As FileSystemObject
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private hmAst As Integer
Private Const MESSAGEBLACK As Long = 0
Private Const MESSAGERED As Long = 255
Private Const MESSAGEGREEN As Long = 39680
Private Const COMPLETEDFOLDER As String = "Completed\"
Private Const COMPLETEDDAYS As Integer = 365
'time zone subtract client from 'site ipump' to get time difference. Reverse of export.
Private Const NOTIMEDIFFERENCE As Integer = 0
Private Const EASTERN As Integer = 8
Private Const CENTRAL As Integer = 7
Private Const MOUNTAIN As Integer = 6
Private Const PACIFIC As Integer = 5
Private Const ALASKAN As Integer = 4
Private Const HAWAIIAN As Integer = 3
Private Const TIMENONE  As String = "1/1/1970"
Private Const ZONENODAYLIGHTCHANGE As Integer = 0
Private Const ZONEDAYLIGHTATSTATION As Integer = 1
Private Const ZONEDAYLIGHTATTIME As Integer = 2

Private lmMaxWidth As Long
'Private Const LOGFILE As String = "iPumpImportLog.txt"
Private Const FORMNAME As String = "FrmImportiPump"
Private Const ERRORSQL As Integer = 9101
Private Const FILEERROR As String = "iPumpImport"
Private Const myForm As String = "iPump"
Private smPathForgLogMsg As String
Private myErrors As CLogger
Private rsImported As ADODB.Recordset
'time zone
Private imDaylight As Integer
Private imTimeZone As Integer

Private Sub cmdCancel_Click()
    If bmExporting Then
        bmTerminate = True
        Exit Sub
    End If
    Unload frmImportiPump
End Sub

Private Sub cmdImport_Click()
    mCleanFolders
    bmTerminate = False
    bgTaskBlocked = False
    sgTaskBlockedName = "iPump Import"
    mImport
    If bgTaskBlocked Then
          mSetResults "Some spots were blocked during Import.", MESSAGERED
          gMsgBox "Some spots were blocked during the Import." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
     End If
     bgTaskBlocked = False
     sgTaskBlockedName = ""
End Sub
Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    mInit
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmCPDat
    Erase tmAstInfo
    rsImported.Close
    Set oMyFileObj = Nothing
    Set myErrors = Nothing
    Set frmImportiPump = Nothing
End Sub
Private Sub mInit()
    frmImportiPump.Caption = "Import iPump - " & sgClientName
    bmTerminate = False
    bmExporting = False
    Set oMyFileObj = New FileSystemObject
    lmMaxWidth = lbcMsg.Width
    Set myErrors = New CLogger
    myErrors.LogPath = myErrors.CreateLogName(sgMsgDirectory & FILEERROR)
    smPathForgLogMsg = FILEERROR & "Log_" & Format(gNow(), "mm-dd-yy") & ".txt"
    'gPopDaypart
End Sub
Private Sub mImport()
    Dim iRet As Integer
    Dim sMsgFileName As String
    Dim slStatus As CSIRspGetXMLStatus
    Dim slErrorString As String
    Dim slName As String
    Dim slImportPath As String
    Dim myFile As file
    Dim blAtLeastOne As Boolean
    Dim blImportFailed As Boolean
    Dim ilSpotCount As Integer
    Dim slFileName As String
    Dim blNotAllImported As Boolean
    Dim slIPumpID As String
    'time zone stuff
    Dim ilSiteZone As Integer
    
    Screen.MousePointer = vbHourglass
    bmExporting = True
    lbcMsg.Clear
    slImportPath = mIPumpFolder()
    'files available to import
    '8886
    'If Dir(slImportPath & "*.weg") <> "" Then
    DoEvents
    If bmTerminate Then
        mSetResults "** User Terminated **", MESSAGERED
        bmExporting = False
        GoTo Cleanup
    End If
    myErrors.WriteFacts "Importing begun: " & sgUserName, True
    gOpenMKDFile hmAst, "Ast.Mkd"
    ilSiteZone = mZoneGetSite()
    For Each myFile In oMyFileObj.GetFolder(slImportPath).Files
        If mTestFileName(myFile.Name) Then
            If mProcessFile(slImportPath & myFile.Name, slIPumpID) Then
                If rsImported.RecordCount > 0 Then
                    rsImported.MoveFirst
                    ilSpotCount = mProcessSpots(slIPumpID, ilSiteZone, blNotAllImported)
                    If ilSpotCount > 0 And Not blNotAllImported Then
                        blAtLeastOne = True
                        mSetResults "Finished posting " & myFile.Name & " spots: " & ilSpotCount, MESSAGEBLACK
                        slFileName = myFile.Name
                    ElseIf ilSpotCount = 0 Then
                        slFileName = "Unread_" & myFile.Name
                        mSetResults myFile.Name & " could not be read.", MESSAGERED
                    Else
                        slFileName = "UnreadMatching_" & myFile.Name
                        mSetResults myFile.Name & " could not be matched.", MESSAGERED
                    End If
                Else
                    slFileName = "UnreadNoProcess_" & myFile.Name
                    mSetResults myFile.Name & " could not be read in mProcessFile.", MESSAGEBLACK
                End If
            Else
                slFileName = "Unread_" & myFile.Name
                mSetResults myFile.Name & " had no posting information.", MESSAGEBLACK
            End If
            If bmTerminate Then
                mSetResults "** User Terminated **", MESSAGEBLACK
                bmExporting = False
                GoTo Cleanup
            End If
        Else
            slFileName = "UnreadFailTest_" & myFile.Name
            mSetResults " A file was found which was not properly formatted: " & myFile.Name, MESSAGERED
        End If
'testing don't move
        If oMyFileObj.FILEEXISTS(slImportPath & COMPLETEDFOLDER & slFileName) Then
            oMyFileObj.DeleteFile (slImportPath & COMPLETEDFOLDER & slFileName)
        End If
        myFile.Move (slImportPath & COMPLETEDFOLDER & slFileName)
    Next myFile
   ' End If
    If blAtLeastOne Then
        If blImportFailed Then
            mSetResults "Files in folder were read, but no new files were imported.", MESSAGERED
            mSetResults "**iPump Posting completed**", MESSAGERED
        Else
            mSetResults "**iPump Posting completed**", MESSAGEGREEN
            myErrors.WriteFacts "**iPump Posting completed**"
        End If
    Else
        If Not blImportFailed Then
            mSetResults "**No files to read--posting ended**", MESSAGEBLACK
        End If
    End If
    cmdImport.Enabled = False
Cleanup:
    If Not rsImported Is Nothing Then
        If (rsImported.State And adStateOpen) <> 0 Then
            rsImported.Close
        End If
        Set rsImported = Nothing
    End If
    mCloseAst
    bmExporting = False
    cmdCancel.Caption = "&Done"
    cmdCancel.SetFocus
    Screen.MousePointer = vbDefault
    Set myFile = Nothing
    Exit Sub
ErrHand:
    myErrors.WriteError "-mImport: " & Err.Description, False, True
    GoTo Cleanup
End Sub
Private Function mIPumpFolder() As String
    Dim slFolderPath As String
    Dim slNestedPath As String
    Dim myFile As file
    Dim dlDeleteDate As Date
    Dim slFolderName As String
    Dim blFirstTime As Boolean
    Dim slNewFolder As String
    
    blFirstTime = False
    slFolderName = mSafeFileName(sgClientName)
    dlDeleteDate = gNow()
    slFolderPath = oMyFileObj.BuildPath(sgImportDirectory, myForm & "\" & slFolderName & "\")
    If Not oMyFileObj.FolderExists(slFolderPath) Then
        slNewFolder = oMyFileObj.BuildPath(sgImportDirectory, myForm)
        If Not oMyFileObj.FolderExists(slNewFolder) Then
            oMyFileObj.CreateFolder (slNewFolder)
        End If
        oMyFileObj.CreateFolder slFolderPath
    End If
    slNestedPath = slFolderPath & COMPLETEDFOLDER
    If Not oMyFileObj.FolderExists(slNestedPath) Then
        oMyFileObj.CreateFolder (slNestedPath)
    End If
    For Each myFile In oMyFileObj.GetFolder(slNestedPath).Files
        If InStr(1, myFile.Name, "UNREAD", vbTextCompare) > 0 Then
            If blFirstTime = False Then
                blFirstTime = True
                mSetResults "There are unread files that need to be processed!  Please contact Counterpoint with " & smPathForgLogMsg & ".", MESSAGEBLACK
            End If
            myErrors.WriteWarning "       " & myFile.Name & " dated: " & Format$(myFile.DateCreated, "dd/mm/yy")
        Else
            If DateDiff("d", myFile.DateCreated, dlDeleteDate) > COMPLETEDDAYS Then
                myFile.Delete
            End If
        End If
    Next myFile
    mIPumpFolder = slFolderPath
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
Private Sub mSetResults(Msg As String, FGC As Long)
    'add scroll bar as needed
    gAddMsgToListBox frmImportiPump, lmMaxWidth, Msg, lbcMsg
    lbcMsg.ListIndex = lbcMsg.ListCount - 1
    'if ever got an error, remain red
    If lbcMsg.ForeColor <> MESSAGERED Then
        lbcMsg.ForeColor = FGC
    End If
    DoEvents
    If FGC = MESSAGERED Then
        myErrors.WriteWarning Msg
    Else
        myErrors.WriteFacts Msg
    End If
End Sub
Private Sub mCleanFolders()
    If Not myErrors Is Nothing Then
        With myErrors
            .CleanThisFolder = messages
            '8886 lose passing 'iPump'
            .CleanFolder
            If Len(.ErrorMessage) > 0 Then
                .WriteWarning "Couldn't delete old files from 'messages': " & .ErrorMessage
            End If
        End With
    End If
End Sub
Private Function mTestFileName(slName As String) As Boolean
    mTestFileName = True
    If InStr(1, UCase(slName), ".WEG") = 0 Then
        mTestFileName = False
    End If
End Function
Private Function mProcessFile(slFullPath As String, slIPumpID As String) As Boolean
'returns false for errors reading text; if no facts returns true.
' O- slIPumpID
    Dim blReturn As Boolean
    Dim slLine As String
    Dim ilVersion As Integer
    Dim oMyFileObj As FileSystemObject
    Dim myFile As TextStream
    Dim slName As String
    Dim blTextExists As Boolean
    
    Set oMyFileObj = New FileSystemObject
    If Not rsImported Is Nothing Then
        If (rsImported.State And adStateOpen) <> 0 Then
            rsImported.Close
        End If
    End If
    Set rsImported = mPrepRecordset()
    blReturn = True
    blTextExists = oMyFileObj.FILEEXISTS(slFullPath)
    If blTextExists Then
        Set myFile = oMyFileObj.OpenTextFile(slFullPath, ForReading, False)
        'header?
        slLine = myFile.ReadLine
        Do While Not myFile.AtEndOfStream
            slLine = myFile.ReadLine
            If Not mFillRs(slLine, slIPumpID) Then
                blReturn = False
                GoTo Cleanup
            End If
        Loop
       myFile.Close
       Set myFile = Nothing
    End If
Cleanup:
    Set oMyFileObj = Nothing
    If Not myFile Is Nothing Then
    On Error Resume Next
        myFile.Close
    On Error GoTo 0
        Set myFile = Nothing
    End If
    mProcessFile = blReturn
End Function

Private Function mProcessSpots(sliPumpCode As String, ilSiteZone As Integer, blNotAllProcessed As Boolean) As Integer
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
    Dim slLastDate As String
    Dim ilSpotCount As Integer
    Dim c As Integer
    Dim slStation As String
    Dim slVehicle As String
    Dim rsClone As ADODB.Recordset
    Dim llVefCode As Long
    Dim ilStnCode As Integer
    Dim llPreviousVehicle As Long
    Dim slStartDate As String
    
    blNotAllProcessed = False
    mProcessSpots = 0
    ilvehicle = 0
    llAtt = 0
On Error GoTo ErrHand
    ilStation = mGetStation(sliPumpCode, slStation)
    If ilStation = 0 Then
        mProcessSpots = 0
        Exit Function
    End If
    mGetVehicles ilStation
    rsImported.Sort = "attcode"
    rsImported.MoveFirst
    slMondayFeedDate = gAdjYear(gObtainPrevMonday(Trim$(rsImported!Date)))
    Set rsClone = rsImported.Clone
    rsClone.Sort = "date desc"
    Do While Not rsImported.EOF
        If llAtt <> rsImported!attCode Then
            llAtt = rsImported!attCode
            ilvehicle = rsImported!Vehicle
            rsClone.Filter = "attCode = " & llAtt
            mPrepAst llAtt, ilvehicle, slMondayFeedDate, ilStation
            slLastDate = Format(rsClone!Date, sgSQLDateForm)
            If ilSiteZone > 0 Then
                If Not mAdjustTime(rsClone, ilSiteZone, slMondayFeedDate, slLastDate) Then
                    mProcessSpots = 0
                    Exit Function
                End If
            End If
            mResetBeforeImporting slMondayFeedDate, slLastDate
            ' an error here will stop the processing and go to error handler below
            mImportByAstCode blNoSpotAired, ilStation, rsClone
            If Not mUpdateCptt(blNoSpotAired, ilvehicle, llAtt, slMondayFeedDate, slLastDate) Then
                mSetResults "Error in mUpdateCptt", MESSAGERED
                ilSpotCount = 0
                GoTo Cleanup
            End If
            llVefCode = gBinarySearchVef(CLng(ilvehicle))
            If llVefCode <> -1 Then
                slVehicle = Trim$(tgVehicleInfo(llVefCode).sVehicle)
            Else
                slVehicle = ""
            End If
            ilStnCode = gBinarySearchStationInfoByCode(ilStation)
            If ilStnCode <> -1 Then
                slStation = Trim$(tgStationInfo(ilStnCode).sCallLetters)
            Else
                slStation = ""
            End If
             slStartDate = slMondayFeedDate
            If Len(slVehicle) > 0 And Len(slVehicle) > 0 Then
                mInsertToWebLog slVehicle, slStation, "iPump", llAtt, slStartDate
            End If
        End If
        rsImported.MoveNext
    Loop
    llPreviousVehicle = 0
    rsImported.Filter = "Found = false"
    Do While Not rsImported.EOF
        blNotAllProcessed = True
        If rsImported!Vehicle <> llPreviousVehicle Then
            llPreviousVehicle = rsImported!Vehicle
            llVefCode = gBinarySearchVef(llPreviousVehicle)
            If llVefCode <> -1 Then
                slVehicle = Trim$(tgVehicleInfo(llVefCode).sVehicle)
            Else
                slVehicle = ""
            End If
        End If
        slNoAstExists = slVehicle & "," & slStation & ", attCode: "
        slNoAstExists = slNoAstExists & rsImported!attCode & ", ISCI: " & Trim$(rsImported!ISCI) & ", astCode:" & rsImported!code & ", Air date: "
        slNoAstExists = slNoAstExists & rsImported!Date & ", Air time: "
        slNoAstExists = slNoAstExists & Trim$(rsImported!TIME)
        mSetResults "Unable to process: " & slNoAstExists & " AST missing", MESSAGERED
        rsImported.MoveNext
    Loop
    rsImported.Filter = "Found = true"
    ilSpotCount = rsImported.RecordCount
    'mFinishPrevious ilStation, slStation
Cleanup:
    mProcessSpots = ilSpotCount
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, FORMNAME & "-mProcessSpots"
    ilSpotCount = 0
    GoTo Cleanup
End Function
Private Sub mPrepAst(llAtt As Long, ilvehicle As Integer, slMondayFeedDate As String, ilStation As Integer)

    mLoadCpPosting llAtt, ilvehicle, slMondayFeedDate, ilStation '
    DoEvents
    igTimes = 1 'By Week
    gGetAstInfo hmAst, tmCPDat(), tmAstInfo(), -1, True, False, True
End Sub
Private Sub mCloseAst()
    gCloseMKDFile hmAst, "Ast.Mkd"
    Erase tmCPDat
    Erase tmAstInfo
End Sub
Private Sub mLoadCpPosting(llAtt As Long, ilvehicle As Integer, slFeedDate As String, ilStation As Integer)
    Dim cprst As ADODB.Recordset
    Dim SQLQuery As String
    Dim ilVpf As Integer
    
    SQLQuery = "SELECT cpttCode,cpttStatus,cpttPostingStatus,cpttAstStatus,attTimeType,shttTimeZone,ShttackDaylight as Daylight, shttTztCode as TimeZone"
    SQLQuery = SQLQuery & " FROM cptt,att,shtt WHERE (shttCode = cpttShfCode AND attCode = cpttAtfCode "
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
    imTimeZone = cprst!TimeZone
    imDaylight = cprst!DAYLIGHT
    cprst.Close
    Set cprst = Nothing
End Sub
Private Function mImportByAstCode(blNoSpotAired As Boolean, ilShfCode As Integer, rsClone As ADODB.Recordset) As Boolean
'    update the AST if it exists and mark it as found in the array. Return false only on error OR code matches but vehicle and station don't.
'   O-blNoSpotAired
'    6/7/2011 added multicasting
    Dim llAstCode As Long
    Dim llSpotLoop As Long
 '   Dim tlMyAst As AST
    Dim llImportAttCode As Long
    Dim ilVefCode As Integer
    
On Error GoTo ErrHand
    mImportByAstCode = True
    blNoSpotAired = True
    llImportAttCode = rsClone!attCode
    ilVefCode = rsClone!Vehicle
    'reverse loop and lose recordset(replace with tmastinfo) for multicasting
    For llSpotLoop = 0 To UBound(tmAstInfo) - 1
        DoEvents
        llAstCode = tmAstInfo(llSpotLoop).lCode
        rsClone.Filter = "attcode = " & llImportAttCode & " AND code = " & llAstCode
        If Not rsClone.EOF Then
            rsClone!found = True
            blNoSpotAired = Not mUpdateAst(llAstCode, tmAstInfo(llSpotLoop).iStatus, rsClone!Date, rsClone!TIME)
        End If
    Next llSpotLoop
    Exit Function
ErrHand:
    mImportByAstCode = False
    'Throw to mProcessSpots to stop all processing.
    Err.Raise ERRORSQL, "mImportByAstCode", Err.Description
End Function

Private Function mUpdateAst(llAstCode As Long, ilStatus As Integer, slDate As String, slTime As String) As Boolean
    ' O- did spots air?
    Dim ilAstStatus As Integer
    mUpdateAst = True
    If ilStatus <= 1 Or ilStatus = 9 Or ilStatus = 10 Then
        ilAstStatus = ilStatus
    Else
        ilAstStatus = 1
    End If
    SQLQuery = "UPDATE ast SET astCPStatus = 1, astStatus = " & ilAstStatus & ", "
    SQLQuery = SQLQuery & "astAirDate = '" & Format$(slDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "astAirTime = '" & Format$(slTime, sgSQLTimeForm) & "'"
    '6158
    SQLQuery = SQLQuery & " WHERE (astCode = " & llAstCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        mUpdateAst = False
        Err.Raise ERRORSQL, "mUpdateAst", "Problem in mUpdateAst"
    End If
End Function
Private Function mUpdateCptt(blNoSpotsAired As Boolean, ilVefCode As Integer, llAtfCode As Long, slMondayFeedDate As String, slLastDate As String) As Boolean
    'Created by D.S. June 2007  Modified Dan M 11/02/10 V81 new values in cptt added 2/25/2011
    'Set the CPTT value, but only for days between monday and 'last date'
    Dim slSuDate As String
    Dim ilStatus As Integer
    Dim llVeh As Long
    Dim ilAst As Integer
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
            SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slLastDate, sgSQLDateForm) & "')" & ")"
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
    'ast's not found  are marked as not aired 4
    SQLQuery = "UPDATE ast SET "
    SQLQuery = SQLQuery & "astCPStatus = 1, astStatus = 4"    'Received
    SQLQuery = SQLQuery & " WHERE (astAtfCode = " & llAtfCode
    SQLQuery = SQLQuery & " AND astCPStatus = 0"
    SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slLastDate, sgSQLDateForm) & "')" & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError smPathForgLogMsg, FORMNAME & "-mUpdateCptt"
        mUpdateCptt = False
        Exit Function
    End If
    'Determine if CPTTStatus should to set to 0=Partial or 1=Completed
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
                SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & "cpttPostingStatus = 1" 'Partial
            Else
                If blNoSpotsAired Then
                    SQLQuery = SQLQuery & "cpttStatus = 2" & ", " 'Complete
                Else
                    SQLQuery = SQLQuery & "cpttStatus = 1" & ", " 'Complete
                End If
                SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & "cpttPostingStatus = 2"  'Complete
            End If
        Else
            If blNoSpotsAired Then
                SQLQuery = SQLQuery & "cpttStatus = 2" & ", " 'Complete
            Else
                SQLQuery = SQLQuery & "cpttStatus = 1" & ", " 'Complete
            End If
            SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
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
        SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
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
    gGetAstInfo hmAst, tmCPDat(), tmAstInfo(), -1, False, False, True
    For ilAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
        gIncSpotCounts tmAstInfo(ilAst), ilSchdCount, ilAiredCount, ilPledgeCompliantCount, ilAgyCompliantCount
    Next ilAst
    SQLQuery = "Update cptt Set "
    SQLQuery = SQLQuery & "cpttNoSpotsGen = " & ilSchdCount & ", "
    SQLQuery = SQLQuery & "cpttNoSpotsAired = " & ilAiredCount & ", "
    SQLQuery = SQLQuery & "cpttNoCompliant = " & ilPledgeCompliantCount & ", "
    SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
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
Private Sub mInsertToWebLog(slVehicleName As String, slStation As String, slSignature As String, llAttCode As Long, slDate As String)
    Dim SQLQuery As String
    Dim slCurrent As String
    
    slCurrent = gNow()
    SQLQuery = "Insert Into WebL (weblType, weblattCode, weblCallLetters, weblVehicleName, weblUserName, weblPostDay, weblDate, weblTime) "
    SQLQuery = SQLQuery & "Values (3," & llAttCode & ",'" & gFixQuote(slStation) & "','" & gFixQuote(slVehicleName) & "','" & gFixQuote(slSignature) & "', '" & Format(slDate, sgSQLDateForm) & "',"
    SQLQuery = SQLQuery & "'" & Format$(slCurrent, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(slCurrent, sgSQLTimeForm) & "'"
    SQLQuery = SQLQuery & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        GoTo ERRSQL
    End If
    Exit Sub
ERRSQL:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, FORMNAME & "-mInsertToWebLog"
    mSetResults gMsg, MESSAGERED
End Sub
   Private Function mFillRs(slLine As String, slIPumpID As String) As Boolean
        Dim slNames() As String
        
On Error GoTo ErrCatch
        slNames = Split(slLine, ",")
        If UBound(slNames) <> 4 Then
            mFillRs = False
            GoTo Cleanup
        End If
        With rsImported
            .AddNew Array("code", "date", "time", "isci", "Found"), _
            Array(slNames(0), slNames(1), slNames(2), slNames(3), False)
        End With
        slIPumpID = slNames(4)
        mFillRs = True
Cleanup:
        Erase slNames
        Exit Function
ErrCatch:
    mFillRs = False
    gHandleError smPathForgLogMsg, FORMNAME & "-mFillRs"
    mSetResults gMsg, MESSAGERED
    GoTo Cleanup
   End Function
Private Function mPrepRecordset() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
    With myRs.Fields
        .Append "Code", adInteger
        '.Append "date", adChar, 10
        .Append "date", adDate
        .Append "time", adChar, 8
        .Append "isci", adChar, 20
        .Append "AttCode", adInteger
        .Append "Vehicle", adInteger
        .Append "Found", adBoolean
    End With
    myRs.Open
    myRs("AttCode").Properties("optimize") = True
    Set mPrepRecordset = myRs
End Function
Private Function mGetStation(sliPumpCode As String, slStation As String) As Integer
    'O ilshtt and slStation(call letters)
    Dim Sql As String
    Dim ilRet As Integer
    
On Error GoTo ERRSQL
    Sql = "select shttcode , shttCallLetters from shtt where shttipumpid = '" & sliPumpCode & "'"
    Set rst = gSQLSelectCall(Sql)
    If Not rst.EOF Then
        ilRet = rst!shttCode
    Else
        ilRet = 0
    End If
    mGetStation = ilRet
    Exit Function
ERRSQL:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, FORMNAME & "-mGetStation"
    mSetResults gMsg, MESSAGERED
    
End Function

Private Sub mGetVehicles(ilStation As Integer)
    Dim Sql As String
    Dim blNoMatch As Boolean
    Dim slErrorString As String
    
    slErrorString = ""
    blNoMatch = False
On Error GoTo ERRSQL
    Do While Not rsImported.EOF
        Sql = "select astatfCode as Code, astvefCode as vehicle from ast where  astCode = " & rsImported!code & " AND ASTSHFCODE = " & ilStation
        Set rst = gSQLSelectCall(Sql)
        If Not rst.EOF Then
            rsImported!attCode = rst!code
            rsImported!Vehicle = rst!Vehicle
        Else
            blNoMatch = True
            slErrorString = slErrorString & CStr(rsImported!code) & ","
        End If
        rsImported.MoveNext
    Loop
    If blNoMatch Then
        slErrorString = mLoseLastLetter(slErrorString)
        mSetResults "Problem with some spots: The station's reference do not match.", MESSAGERED
        myErrors.WriteWarning "These ast codes do match the file's station : " & slErrorString
    End If
    Exit Sub
ERRSQL:
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, FORMNAME & "-mGetVehicles"
    mSetResults gMsg, MESSAGERED

End Sub
Private Function mLoseLastLetter(slInput As String) As String
    Dim llLength As Long
    Dim slNewString As String

    llLength = Len(slInput)
    If llLength > 0 Then
        slNewString = Mid(slInput, 1, llLength - 1)
    End If
    mLoseLastLetter = slNewString
End Function
Private Sub mResetBeforeImporting(slFirst As String, slLast As String)
    Dim llSpotLoop As Long
    Dim llAstCode As Long
    Dim llAlt As Long
    Dim slSql As String
    
    For llSpotLoop = 0 To UBound(tmAstInfo) - 1
        DoEvents
        With tmAstInfo(llSpotLoop)
            'between first and last days
            If DateDiff("d", slFirst, .sFeedDate) >= 0 And DateDiff("d", .sFeedDate, slLast) >= 0 Then
                llAstCode = .lCode
                .iCPStatus = 0
                .iStatus = .iPledgeStatus
                slSql = "update ast set astCpStatus = 0 , astStatus = " & .iStatus & " where astcode = " & llAstCode
                If gSQLWaitNoMsgBox(slSql, False) <> 0 Then
                    Err.Raise 1001, "mResetBeforeImporting", Err.Description
                End If
            End If
        End With
    Next llSpotLoop
End Sub
'time zone
Private Function mZoneGetSite() As Integer
    Dim ilRet As Integer
    Dim slZone As String
    Dim rsZone As ADODB.Recordset
    Dim slSql As String
    
On Error GoTo ERRORBOX
    ilRet = NOTIMEDIFFERENCE
    slSql = "SELECT safIPumpZone as SiteZone From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set rsZone = gSQLSelectCall(slSql)
    If Not rsZone.EOF Then
        slZone = UCase(rsZone!SiteZone)
        Select Case slZone
            Case "E"
                ilRet = EASTERN
            Case "C"
                ilRet = CENTRAL
            Case "M"
                ilRet = MOUNTAIN
            Case "P"
                ilRet = PACIFIC
'            Case "A"
'                ilRet = ALASKAN
'            Case "H"
'                ilRet = HAWAIIAN
        End Select
    End If
Cleanup:
    mZoneGetSite = ilRet
    If Not rsZone Is Nothing Then
        If (rsZone.State And adStateOpen) <> 0 Then
            rsZone.Close
        End If
        Set rsZone = Nothing
    End If
    Exit Function
ERRORBOX:
    gHandleError smPathForgLogMsg, FORMNAME & "-mZoneGetSite"
    ilRet = NOTIMEDIFFERENCE
    GoTo Cleanup
End Function
Private Function mAdjustTime(rsClone As Recordset, ilSiteZone As Integer, slMondayFeedDate As String, slLastDate As String) As Boolean
    Dim blRet As Boolean
    Dim ilDaylightAdjust As Integer
    Dim ilZoneAdjust As Integer
    Dim dlStartSavingTime As Date
    Dim dlEndSavingTime As Date
    Dim slAdjustedDate As String
    Dim slAdjustedTime As String

    blRet = True
On Error GoTo ERRORBOX
    ilDaylightAdjust = mZoneForSavings(dlStartSavingTime, dlEndSavingTime, slMondayFeedDate, slLastDate)
    ilZoneAdjust = mZoneGetAdjustForStation(ilSiteZone, ilDaylightAdjust)
    If ilZoneAdjust > 0 Or ilDaylightAdjust = ZONEDAYLIGHTATTIME Then
        Do While Not rsClone.EOF
            slAdjustedDate = rsClone!Date
            slAdjustedTime = rsClone!TIME
            slAdjustedDate = mZoneAdjustTime(slAdjustedDate, slAdjustedTime, ilZoneAdjust, ilDaylightAdjust, dlStartSavingTime, dlEndSavingTime)
                rsClone!TIME = Format(slAdjustedDate, "hh:nn") 'ss?
                rsClone!Date = Format(slAdjustedDate, sgShowDateForm)
            rsClone.MoveNext
        Loop
        rsClone.MoveFirst
    End If
    mAdjustTime = blRet
    Exit Function
ERRORBOX:
    myErrors.WriteError "Error in mAdjustTime: " & Err.Description
    mAdjustTime = False
End Function
Private Function mZoneForSavings(dlStartSavingTime As Date, dlEndSavingTime As Date, slImportStart As String, slImportEnd As String) As Integer
    'ZONENODAYLIGHTCHANGE ZONEDAYLIGHTATSTATION ZONEDAYLIGHTATTIME
    Dim ilRet As Integer
    
    ilRet = ZONENODAYLIGHTCHANGE
    dlStartSavingTime = mDaylightSavings(True, slImportStart)
    dlEndSavingTime = mDaylightSavings(False, slImportEnd)
    ' this week's export is within daylight savings.  The station's choice to follow dst matters
    If DateDiff("d", dlStartSavingTime, slImportStart) > 0 And DateDiff("d", slImportEnd, dlEndSavingTime) > 0 Then
        ilRet = ZONEDAYLIGHTATSTATION
    ' this export crosses into daylight savings time! must check at date time for when this happens
    ElseIf (DateDiff("d", slImportStart, dlStartSavingTime) > 0 And DateDiff("d", dlStartSavingTime, slImportEnd) > 0) Or (DateDiff("d", slImportEnd, dlEndSavingTime) > 0 And DateDiff("d", dlEndSavingTime, slImportStart) > 0) Then
        ilRet = ZONEDAYLIGHTATTIME
    End If
    mZoneForSavings = ilRet
End Function
Private Function mDaylightSavings(blStart As Boolean, slStartDate As String) As Date
    ' start: 2nd sunday of march end: 1st sunday in November  both 2:00 am
    Dim slDate As String
    Dim ilDay As Integer
  
 On Error GoTo ERRORBOX
    If blStart Then
        slDate = "3/01/" & DatePart("yyyy", slStartDate)
        ilDay = DatePart("w", slDate)
        If ilDay <> 1 Then
            slDate = gObtainNextSunday(slDate)
        End If
        slDate = DateAdd("ww", 1, slDate)
    Else
        slDate = "11/01/" & DatePart("yyyy", slStartDate)
        ilDay = DatePart("w", slDate)
        If ilDay <> 1 Then
            slDate = gObtainNextSunday(slDate)
        End If
    End If
    mDaylightSavings = CDate(slDate & " 2:00:00 am")
    Exit Function
ERRORBOX:
    myErrors.WriteError Err.Description, True, True
    mSetResults "Error in mDaylightSavings", MESSAGERED
End Function
Private Function mZoneGetAdjustForStation(ilSiteZone As Integer, ilDaylightAdjust As Integer) As Integer
    'ZONENODAYLIGHTCHANGE ZONEDAYLIGHTATSTATION ZONEDAYLIGHTATTIME
    'out if number to add to adjust for difference from station's zone to 'master zone'.  could be negative
    'opposite of export:
    'a pacific export to Eastern station?  Get a plus number.  When used later with dateAdd, will add 3 hours (not incl. daylight issues)
    Dim ilRet As Integer
    Dim ilStation As Integer
    Dim blNeedDaylightAdjust As Boolean
    Dim slTimeZone As String
    Dim ilIndex As Integer
    
    ilStation = NOTIMEDIFFERENCE
    For ilIndex = 0 To UBound(tgTimeZoneInfo) - 1 Step 1
        If tgTimeZoneInfo(ilIndex).iCode = imTimeZone Then
            slTimeZone = tgTimeZoneInfo(ilIndex).sCSIName
            Exit For
        End If
    Next ilIndex
    If Len(slTimeZone) > 0 Then
        slTimeZone = Mid(slTimeZone, 1, 1)
        Select Case slTimeZone
            Case "E"
                ilStation = EASTERN
            Case "C"
                ilStation = CENTRAL
            Case "M"
                ilStation = MOUNTAIN
            Case "P"
                ilStation = PACIFIC
            Case "A"
                ilStation = ALASKAN
            Case Else
                ilStation = HAWAIIAN
        End Select
        ' opposite of export
       ' ilRet = ilSiteZone - ilStation
        ilRet = ilStation - ilSiteZone
    Else
        ilRet = 0
    End If
    ' 0 yes or 1 no aknowledge daylight
    If imDaylight = 1 Then
        blNeedDaylightAdjust = True
    Else
        blNeedDaylightAdjust = False
    End If
    'station doesn't aknowledge daylight, and this export is during daylight (but not across daylight change!)
    If blNeedDaylightAdjust And ilDaylightAdjust = ZONEDAYLIGHTATSTATION Then
        'opposite of export
        'ilRet = ilRet + 1
        ilRet = ilRet - 1
    End If
     mZoneGetAdjustForStation = ilRet
End Function
Private Function mZoneAdjustTime(slDate As String, slTime As String, ilAdjust As Integer, ilDaylightAdjust As Integer, dlStartSavingTime As Date, dlEndSavingTime As Date) As String
    ' return adjusted date/time if site says to always send as specific zone
    Dim slRet As String
    
    slRet = DateAdd("h", ilAdjust, slDate & " " & slTime)
    If ilDaylightAdjust = ZONEDAYLIGHTATTIME Then
        ' date is on daylight savings time, is the time after 2:00 am?
        If DateDiff("d", slRet, dlStartSavingTime) = 0 Then
            If DateDiff("h", dlStartSavingTime, slRet) > 0 Then
                 slRet = DateAdd("h", 1, slRet)
            End If
        ' date is when dst ends. Is time before 2:00 am?
        ElseIf DateDiff("d", slRet, dlEndSavingTime) = 0 Then
            If DateDiff("h", slRet, dlEndSavingTime) > 0 Then
                 slRet = DateAdd("h", 1, slRet)
            End If
        ' date is during daylight savings time, but not the start or end date of dst?
        ElseIf DateDiff("d", dlStartSavingTime, slRet) > 0 And DateDiff("d", slRet, dlEndSavingTime) > 0 Then
            slRet = DateAdd("h", 1, slRet)
        End If
    End If
    mZoneAdjustTime = slRet
End Function
'end time zone
Private Function mFinishPrevious(ilStation As Integer, slStation As String) As Boolean
    Dim llAtt As Long
    Dim ilvehicle As Integer
    Dim slMondayOld As String
    Dim slSundayOld As String
    Dim blNoSpotsAired As Boolean
    Dim slVehicle As String
    Dim llVefCode As Long
    
    'what do I set this to?
    blNoSpotsAired = False
    rsImported.Filter = adFilterNone
    If rsImported.RecordCount > 0 Then
        rsImported.MoveFirst
        Do While Not rsImported.EOF
            If llAtt <> rsImported!attCode Then
                llAtt = rsImported!attCode
                ilvehicle = rsImported!Vehicle
                slMondayOld = gAdjYear(gObtainPrevMonday(Trim$(rsImported!Date)))
                slSundayOld = DateAdd("d", -1, slMondayOld)
                slMondayOld = DateAdd("d", -7, slMondayOld)
                'mPrepAst llAtt, ilvehicle, slMondayFeedDate, ilStation
                llVefCode = gBinarySearchVef(CLng(ilvehicle))
                If llVefCode <> -1 Then
                    slVehicle = Trim$(tgVehicleInfo(llVefCode).sVehicle)
                Else
                    slVehicle = ""
                End If
                If mUpdateCptt(blNoSpotsAired, ilvehicle, llAtt, slMondayOld, slSundayOld) Then
                    If Len(slVehicle) > 0 And Len(slVehicle) > 0 Then
                        mInsertToWebLog slVehicle, slStation, "iPump", llAtt, slMondayOld
                    End If
                Else
                    mSetResults "could not set previous week's posting to complete.", MESSAGEBLACK
                    myErrors.WriteWarning slStation & " " & slVehicle & " agrement # " & llAtt & " for week of " & slMondayOld & " couldn't complete posting."
                End If
            End If
        Loop
    End If
End Function
