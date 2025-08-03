VERSION 5.00
Begin VB.Form frmImportBIA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Aired Spots"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   ControlBox      =   0   'False
   Icon            =   "AffImportBIA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkReportStationsNotUpdated 
      Caption         =   "Report stations not in BIA file."
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CheckBox chkReportOnly 
      Caption         =   "Show report only. No database changes will occur."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   4215
   End
   Begin VB.CommandButton cmdViewReport 
      Caption         =   "View Report"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtFile 
      Height          =   300
      Left            =   990
      TabIndex        =   5
      Top             =   480
      Width           =   3600
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "Browse"
      Height          =   300
      Left            =   4845
      TabIndex        =   4
      Top             =   480
      Width           =   1065
   End
   Begin VB.ListBox lbcMsg 
      Height          =   2205
      ItemData        =   "AffImportBIA.frx":030A
      Left            =   120
      List            =   "AffImportBIA.frx":030C
      TabIndex        =   1
      Top             =   1410
      Width           =   5790
   End
   Begin VB.PictureBox ReSize1 
      Height          =   480
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   11
      Top             =   4425
      Width           =   1200
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   1125
      TabIndex        =   2
      Top             =   4500
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3150
      TabIndex        =   3
      Top             =   4500
      Width           =   1575
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   480
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   12
      Top             =   4440
      Width           =   1200
   End
   Begin VB.PictureBox plcGauge 
      Height          =   210
      Left            =   1335
      ScaleHeight     =   150
      ScaleWidth      =   3285
      TabIndex        =   10
      Top             =   1125
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Label lbcFile 
      Caption         =   "Import File"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   495
      Width           =   780
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   5790
   End
End
Attribute VB_Name = "frmImportBIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmImportBIA
'*
'*  Created August 14, 2006 by Jeff Dutschke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text

Private imImporting As Integer
Private imTerminate As Integer
Private hmFrom As Integer
Private lmTotalBIARecords As Long
Private lmProcessedBIARecords As Long
Private lmPercent As Long
Private tmBIAInfo() As BIASTATIONINFO
Private tmBIAReportInfo() As BIAREPORTINFO
Private tmUsedMarkets() As MARKETINFO
Private tmUsedOwners() As OWNERINFO
Private tmUsedFormats() As FORMATINFO
Private tmBIARegionSet() As BIAREGIONSET
Private smCallLetters As String
Private smBand As String
Private smMarketName As String
Private smRank As String
Private smOwnerName As String
Private smFormat As String
Private blUpdateDatabase As Boolean
Private smReportPathFileName As String

'***************************************************************************
'
'***************************************************************************
Private Sub cmdImport_Click()
    Dim iLoop As Integer
    Dim sFileName As String
    Dim iRet As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim sToFile As String
    Dim sDateTime As String
    Dim sMsgFileName As String
    Dim sMoDate As String

    On Error GoTo ErrHand
    lbcMsg.Clear
    blUpdateDatabase = True
    If chkReportOnly.Value Then
        ' User wants to view the report without actually making the changes to the database.
        blUpdateDatabase = False
        gLogMsg "SHOW REPORT ONLY. NO CHANGES ARE BEING MADE TO THE DATABASE.", "BIAImportLog.Txt", False
    End If
    Kill smReportPathFileName
    ReDim tmBIAReportInfo(0 To 0) As BIAREPORTINFO
    ReDim tmUsedMarkets(0 To 0) As MARKETINFO
    ReDim tmUsedOwners(0 To 0) As OWNERINFO
    ReDim tmUsedFormats(0 To 0) As FORMATINFO
    ReDim tmBIARegionSet(0 To 0) As BIAREGIONSET

    Screen.MousePointer = vbHourglass
    
    If Not gPopMarkets() Then
        Screen.MousePointer = vbDefault
        Call FailBIAImport("Unable to Load Existing Market Names.")
        Exit Sub
    End If
    
    If Not gPopOwnerNames() Then
        Screen.MousePointer = vbDefault
        Call FailBIAImport("Unable to Load Existing Owner Names.")
        Exit Sub
    End If

    If Not gPopFormats() Then
        Screen.MousePointer = vbDefault
        Call FailBIAImport("Unable to Load Existing Format Names.")
        Exit Sub
    End If

    If Not gPopStations() Then
        Screen.MousePointer = vbDefault
        Call FailBIAImport("Unable to Load Existing Station Names.")
        Exit Sub
    End If
    iRet = mCheckFile()
    If iRet = 1 Then
        Screen.MousePointer = vbDefault
        txtFile.SetFocus
        Exit Sub
    End If
    If iRet = 2 Then
        cmdImport.Enabled = False
        cmdCancel.Caption = "&Done"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imImporting = True
    cmcBrowse.Enabled = False
    cmdImport.Enabled = False
    cmdViewReport.Enabled = False
    txtFile.Enabled = False
    chkReportOnly.Enabled = False
    chkReportStationsNotUpdated.Enabled = False
    cmdCancel.Caption = "Abort"
    SetResults "Importing BIA Information", RGB(0, 0, 0)

    iRet = mImportBIAInfo()
    If iRet = False Then
        Screen.MousePointer = vbDefault
        Call FailBIAImport("Terminated - mImportBIAInfo returned False")
        Exit Sub
    End If

    
    If blUpdateDatabase Then
        If Not mRemoveDuplicateMarketNames() Then
            Screen.MousePointer = vbDefault
            Call FailBIAImport("Unable to Remove Dupliacte Market Names.")
            Exit Sub
        End If
        
        If Not gPopMarkets() Then
            Screen.MousePointer = vbDefault
            Call FailBIAImport("Unable to Load Existing Market Names.")
            Exit Sub
        End If
    End If
    'If UBound(tgMarketInfo) > 0 Then
    '    ArraySortTyp fnAV(tgMarketInfo(), 0), UBound(tgMarketInfo), 0, LenB(tgMarketInfo(0)), 2, LenB(tgMarketInfo(0).sName), 1
    'End If

    
    If lmTotalBIARecords > 0 Then
        plcGauge.Value = 0
        plcGauge.Visible = True
    End If
    lmPercent = 0
    lmProcessedBIARecords = 0

    iRet = mProcessBIAInfo()
    plcGauge.Visible = False
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Call FailBIAImport("User Terminated")
        Exit Sub
    End If
    If iRet = False Then
        Screen.MousePointer = vbDefault
        Call FailBIAImport("Terminated - mProcessBIAInfo returned False")
        Exit Sub
    End If

    ' Call TEST_FixBadStationNames
    iRet = mSetUsedArrays()
    
    Call ScanForBrokenMarketLinks
    Call ScanForBrokenOwnerLinks
    Call ScanForBrokenFormatLinks

    If UBound(tmBIAReportInfo) < 1 Then
        SetResults "All station information was 100% up to date. No records were changed.", RGB(0, 0, 0)
    End If

    If chkReportStationsNotUpdated.Value Then
        Call ScanStationsNotUpdated     ' Not found in the BIA file.
    End If

    Call CreateBIAStatusReport

    If imTerminate Then
        Screen.MousePointer = vbDefault
        Call FailBIAImport("User Terminated")
        Exit Sub
    End If
    imImporting = False
    cmdViewReport.Enabled = True
    SetResults "Operation Completed Successfully.", RGB(0, 200, 0)
    If Not blUpdateDatabase Then
        SetResults "*** SHOW REPORT ONLY WAS CHECKED. NO CHANGES HAVE OCCURRED.", RGB(0, 0, 200)
    End If
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Call ShowReport(False)
    Exit Sub

ErrHand:
    Resume Next
End Sub

'***************************************************************************
'
'***************************************************************************
Private Function mImportBIAInfo() As Integer
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim slCallLetters As String
    'Dim slFields(1 To 6) As String
    Dim slFields(0 To 5) As String
    Dim ilMaxRecords As Long
    
    On Error GoTo mImportBIAInfoErr:
    mImportBIAInfo = False
    lbcMsg.Clear
    
    SetResults "Importing BIA station information...", RGB(0, 0, 0)
    slFromFile = txtFile.Text
    'ilRet = 0
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        SetResults "Unable to open file. Error = " & Trim$(Str$(ilRet)), RGB(255, 0, 0)
        Exit Function
    End If
    
    lmTotalBIARecords = 0
    ilMaxRecords = 10000
    ReDim tmBIAInfo(0 To ilMaxRecords) As BIASTATIONINFO
    
    Do While Not EOF(hmFrom)
        ilRet = 0
        Line Input #hmFrom, slLine
        gParseCDFields slLine, False, slFields()
        
        If ilRet <> 0 Then
            SetResults "Unable to read from file. Error = " & Trim$(Str$(ilRet)), RGB(255, 0, 0)
            Exit Function
        End If
'        tmBIAInfo(lmTotalBIARecords).sCallLetters = mMakeStationName(slFields(1), slFields(2))
'        tmBIAInfo(lmTotalBIARecords).sMarketName = Trim(slFields(3))
'        tmBIAInfo(lmTotalBIARecords).iRank = slFields(4)
'        tmBIAInfo(lmTotalBIARecords).sOwnerName = Trim(slFields(5))
'        'tmBIAInfo(lmTotalBIARecords).sFormat = gFixQuote(Trim(slFields(6)))
'        tmBIAInfo(lmTotalBIARecords).sFormat = Trim(slFields(6))
        tmBIAInfo(lmTotalBIARecords).sCallLetters = mMakeStationName(slFields(0), slFields(1))
        tmBIAInfo(lmTotalBIARecords).sMarketName = Trim(slFields(2))
        tmBIAInfo(lmTotalBIARecords).iRank = slFields(3)
        tmBIAInfo(lmTotalBIARecords).sOwnerName = Trim(slFields(4))
        'tmBIAInfo(lmTotalBIARecords).sFormat = gFixQuote(Trim(slFields(5)))
        tmBIAInfo(lmTotalBIARecords).sFormat = Trim(slFields(5))

        
        lmTotalBIARecords = lmTotalBIARecords + 1
        If lmTotalBIARecords >= ilMaxRecords Then
            ' Allocate another 1000 entries.
            ilMaxRecords = ilMaxRecords + 5000
            ReDim Preserve tmBIAInfo(0 To ilMaxRecords) As BIASTATIONINFO
        End If
        
    Loop
    Close hmFrom
    ReDim Preserve tmBIAInfo(0 To lmTotalBIARecords - 1) As BIASTATIONINFO

    ' Sort the array.
    If UBound(tmBIAInfo) > 0 Then
        ArraySortTyp fnAV(tmBIAInfo(), 0), UBound(tmBIAInfo) - 1, 0, LenB(tmBIAInfo(0)), 0, LenB(tmBIAInfo(0).sCallLetters), 0
    End If

    SetResults Trim(Str(lmTotalBIARecords)) & " records imported.", RGB(0, 0, 0)
    mImportBIAInfo = True
    Exit Function
mImportBIAInfoErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmImportBIA-mImportBIA"
    mImportBIAInfo = False
    Exit Function

End Function

'***************************************************************************
'
'***************************************************************************
Private Function mProcessBIAInfo() As Integer
    Dim llBIAIdx As Long
    Dim llStationIdx As Long
    Dim llMktIdx As Long
    Dim llOwnerIdx As Long
    Dim llFormatIdx As Long
    Dim llIdx As Long
    Dim slCallLetters As String
    Dim ilRet As Integer
    Dim ilOldRank As Integer
    Dim sOriginalOwnerName As String
    Dim llTempIdx As Long
    
    On Error GoTo ErrHandler:
    mProcessBIAInfo = False
    
    ' Process all records loaded in the BIA list
    SetResults "Processing BIA information.", RGB(0, 0, 0)
    For llBIAIdx = 0 To lmTotalBIARecords - 1
        DoEvents
        If imTerminate Then
            Exit Function
        End If
'        If llBIAIdx > 0 And llBIAIdx Mod 1000 = 0 Then
'            SetResults Trim(Str(llBIAIdx)) & " records processed.", RGB(0, 0, 0)
'
''            mProcessBIAInfo = True
''            Exit Function
'        End If
         If lmTotalBIARecords > 0 Then
            lmProcessedBIARecords = lmProcessedBIARecords + 1
            lmPercent = (lmProcessedBIARecords * CSng(100)) / lmTotalBIARecords
            If lmPercent >= 100 Then
                If lmProcessedBIARecords + 1 < lmTotalBIARecords Then
                    lmPercent = 99
                Else
                    lmPercent = 100
                End If
            End If
            If plcGauge.Value <> lmPercent Then
                plcGauge.Value = lmPercent
                DoEvents
            End If
        End If
       slCallLetters = Trim(tmBIAInfo(llBIAIdx).sCallLetters)
        ' Find this station in the stations array. If not found then ignore this BIA record.
        llStationIdx = LookupStation(slCallLetters)
        If llStationIdx <> -1 Then
            ' We found the station we're looking for. Find the current market name.
            If Len(Trim(tmBIAInfo(llBIAIdx).sMarketName)) > 0 Then  ' Don't look this up if blank.
                'llTempIdx = LookupMarketByName(tmBIAInfo(llBIAIdx).sMarketName)
                'If llTempIdx = 38 Then
                '    llMktIdx = 1
                'End If
                'llMktIdx = mBinarySearchMarketName(Trim(tmBIAInfo(llBIAIdx).sMarketName))
                llMktIdx = LookupMarketByName(tmBIAInfo(llBIAIdx).sMarketName)
                If UCase(Trim(tmBIAInfo(llBIAIdx).sMarketName)) = "ALBUQUERQUE-SANTA FE, NM" Then
                    If llMktIdx = 6 Then
                        imTerminate = False
                    End If
                End If
                'If llTempIdx <> llMktIdx Then
                '    gMsgBox "error"
                'End If
                If llMktIdx <> -1 Then
                    ' Since this market name is valid because it exists in the BIA data, mark it here as
                    ' used, even though it might not have a reference from stations.
                    'Call AddMarketToUsedArray(llMktIdx)    'Moved to mSetUsedArrays
                    ' Yes it does exist. Verify the station market code is pointing to it.
                    If tgStationInfo(llStationIdx).iMktCode <> tgMarketInfo(llMktIdx).iCode Then
                        ' Update this stations market code since it is not pointing to the correct market name.
                        If Not UpdateStationsMarket(llStationIdx, llMktIdx, llBIAIdx) Then
                            Exit Function
                        End If
                    End If
                Else
                    ' The market name does not exist. Add it and point this stations market code to it.
                    If Not AddMarket(llStationIdx, llMktIdx, llBIAIdx) Then
                        Exit Function
                    End If
                End If
    
                ' Check the Rank
                If tgMarketInfo(llMktIdx).iRank <> tmBIAInfo(llBIAIdx).iRank Then
                    ' The rank does not match.
                    If Not UpdateRank(llMktIdx, llBIAIdx) Then
                        Exit Function
                    End If
                End If
            End If
            
            ' Check Owner Information
            If Len(Trim(tmBIAInfo(llBIAIdx).sOwnerName)) > 0 Then

                'If Trim(tmBIAInfo(llBIAIdx).sOwnerName) = "New World Broadcasting Company Incorporated" Then
                '    gMsgBox "dd"
                'End If
                
                ' WORKAROUND CODE FOR THE OWNER LENGTH NAME ISSUE
                ' The last name field in the artt table is only 40 and max owner name from BIA is unknown.
                ' Some of them have been identified to be larger than 40 characters. So this code was
                ' added to temporarily handle this situation.
                sOriginalOwnerName = ""
                If Len(Trim(tmBIAInfo(llBIAIdx).sOwnerName)) > 60 Then
                    ' Save the original name so we can report on this if the station is updated.
                    sOriginalOwnerName = Trim(tmBIAInfo(llBIAIdx).sOwnerName)
                    tmBIAInfo(llBIAIdx).sOwnerName = Left(tmBIAInfo(llBIAIdx).sOwnerName, 40)
                End If
                ' END OF WORKAROUND CODE

                llOwnerIdx = LookupOwnerByName(tmBIAInfo(llBIAIdx).sOwnerName)
                If llOwnerIdx <> -1 Then
                    'Call AddOwnerToUsedArray(llOwnerIdx)    'Moved to mSetUsedArrays
                    ' This owner name DOES exist. Verify the station owner code is pointing to it.
                    If tgStationInfo(llStationIdx).iOwnerCode <> tgOwnerInfo(llOwnerIdx).iCode Then
                        ' Update this stations owner code since it is not pointing to the correct owner name.
                        If Not UpdateStationsOwner(llStationIdx, llOwnerIdx, llBIAIdx) Then
                            Exit Function
                        End If
                        If Len(Trim(sOriginalOwnerName)) > 0 Then
                            ' Only report this if the station had to be updated.
                            Call UpdateReport(llBIAIdx, "WARNING: Owner name from BIA has been truncated. (" & sOriginalOwnerName & ") was truncated to (" & Left(tmBIAInfo(llBIAIdx).sOwnerName, 60) & ")")
                        End If
                    End If
                Else
                    ' The owner name does not exist. Add it and point this stations owner code to it.
                    If Not AddOwner(llStationIdx, llOwnerIdx, llBIAIdx) Then
                        Exit Function
                    End If
                    If Len(Trim(sOriginalOwnerName)) > 0 Then
                        ' Only report this if the station had to be updated.
                        Call UpdateReport(llBIAIdx, "WARNING: Owner name from BIA has been truncated. (" & sOriginalOwnerName & ") was truncated to (" & Left(tmBIAInfo(llBIAIdx).sOwnerName, 40) & ")")
                    End If
                End If
            End If

            ' Check Station Format Information
            If Len(Trim(tmBIAInfo(llBIAIdx).sFormat)) > 0 Then
                llFormatIdx = LookupFormatByName(tmBIAInfo(llBIAIdx).sFormat)
                If llFormatIdx <> -1 Then
                    'Call AddFormatToUsedArray(llFormatIdx)    'Moved to mSetUsedArrays
                    ' This format name DOES exist. Verify the station format code is pointing to it.
                    If tgStationInfo(llStationIdx).iFormatCode <> tgFormatInfo(llFormatIdx).iCode Then
                        ' Update this stations format code since it is not pointing to the correct format name.
                        If Not UpdateStationsFormat(llStationIdx, llFormatIdx, llBIAIdx) Then
                            Exit Function
                        End If
                    End If
                Else
                    ' This format is not currently in the array so add it.
                    If Not AddFormat(llStationIdx, llFormatIdx, llBIAIdx) Then
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    SetResults Trim(Str(lmTotalBIARecords)) & " Records Processed.", RGB(0, 0, 0)

    mProcessBIAInfo = True
    Exit Function
    
ErrHandler:
    Screen.MousePointer = vbDefault
    gMsg = "A general error has occured in mProcessBIAInfo: "
    gLogMsg "A general error has occured in mProcessBIAInfo: ", "BIAImportLog.Txt", False
    gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
End Function

'***************************************************************************
'
'***************************************************************************
Private Function AddMarket(StationIDX As Long, MktIdx As Long, BIAIdx As Long) As Integer
    Dim NewIdx As Long
    
    On Error GoTo ErrHandler:
    AddMarket = False
    SQLQuery = "Insert Into MKT (mktName, mktRank, mktGroupName) Values( "
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim(tmBIAInfo(BIAIdx).sMarketName)) & "', "
    SQLQuery = SQLQuery & tmBIAInfo(BIAIdx).iRank & ", " & "''" & ")"

    If blUpdateDatabase Then
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHandler:
        End If
    End If
    SQLQuery = "Select MAX(mktCode) from MKT"
    Set rst = cnn.Execute(SQLQuery)

    NewIdx = UBound(tgMarketInfo)
    ReDim Preserve tgMarketInfo(0 To NewIdx + 1) As MARKETINFO
    tgMarketInfo(NewIdx).iCode = rst(0).Value
    tgMarketInfo(NewIdx).sName = tmBIAInfo(BIAIdx).sMarketName
    tgMarketInfo(NewIdx).iRank = tmBIAInfo(BIAIdx).iRank
    tgMarketInfo(NewIdx).sARB = 0
    tgMarketInfo(NewIdx).sBIA = 0

    mAddToRegionSet "M", tgStationInfo(StationIDX).iMktCode, tgMarketInfo(NewIdx).iCode
    tgStationInfo(StationIDX).iMktCode = tgMarketInfo(NewIdx).iCode
    SQLQuery = "Update shtt Set shttMktCode = " & tgStationInfo(StationIDX).iMktCode & " Where shttCode = " & tgStationInfo(StationIDX).iCode
    If blUpdateDatabase Then
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHandler:
        End If
    End If

    MktIdx = NewIdx ' Make sure we update this.
    Call UpdateReport(BIAIdx, "New market was added (" & Trim(tmBIAInfo(BIAIdx).sMarketName) & ")")
    AddMarket = True
    Exit Function

ErrHandler:
    gMsg = ""
    gHandleError "AffErrorLog.txt", "frmImportBIA-AddMarket"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function AddOwner(StationIDX As Long, OwnerIdx As Long, BIAIdx As Long) As Integer
    Dim NewIdx As Long

    On Error GoTo ErrHandler:
    AddOwner = False

    SQLQuery = "INSERT INTO artt(arttLastName, arttType) "
    SQLQuery = SQLQuery & "VALUES ( '" & gFixQuote(Trim(tmBIAInfo(BIAIdx).sOwnerName)) & "', 'O')"

    If blUpdateDatabase Then
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHandler:
        End If
    End If
    SQLQuery = "Select MAX(arttCode) from artt"
    Set rst = cnn.Execute(SQLQuery)

    NewIdx = UBound(tgOwnerInfo)
    ReDim Preserve tgOwnerInfo(0 To NewIdx + 1) As OWNERINFO
    tgOwnerInfo(NewIdx).iCode = rst(0).Value
    tgOwnerInfo(NewIdx).sName = tmBIAInfo(BIAIdx).sOwnerName

    mAddToRegionSet "O", tgStationInfo(StationIDX).iOwnerCode, tgOwnerInfo(NewIdx).iCode
    tgStationInfo(StationIDX).iOwnerCode = tgOwnerInfo(NewIdx).iCode
    SQLQuery = "Update shtt Set shttOwnerArttCode = " & tgOwnerInfo(NewIdx).iCode & " Where shttCode = " & tgStationInfo(StationIDX).iCode
    If blUpdateDatabase Then
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHandler:
        End If
    End If

    OwnerIdx = NewIdx
    Call UpdateReport(BIAIdx, "New owner was added (" & Trim(tmBIAInfo(BIAIdx).sOwnerName) & ")")
    AddOwner = True
    Exit Function

ErrHandler:
    gHandleError "AffErrorLog.txt", "frmImportBIA-AddOwner"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function AddFormat(StationIDX As Long, FormatIdx As Long, BIAIdx As Long) As Integer
    Dim NewIdx As Long
    
    On Error GoTo ErrHandler:
    AddFormat = False
    SQLQuery = "Insert Into FMT_Station_Format (fmtName, fmtGroupName, fmtDftCode, fmtUstCode) Values( '" & gFixQuote(Trim(tmBIAInfo(BIAIdx).sFormat)) & "', " & "''," & 0 & "," & igUstCode & ")"

    If blUpdateDatabase Then
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHandler:
        End If
    End If
    SQLQuery = "Select MAX(fmtCode) from FMT_Station_Format"
    Set rst = cnn.Execute(SQLQuery)

    NewIdx = UBound(tgFormatInfo)
    On Error GoTo ErrHandler:
    ReDim Preserve tgFormatInfo(0 To NewIdx + 1) As FORMATINFO
    tgFormatInfo(NewIdx).iCode = 1
    If Not IsNull(rst(0).Value) Then
        tgFormatInfo(NewIdx).iCode = rst(0).Value
    End If
    tgFormatInfo(NewIdx).sName = tmBIAInfo(BIAIdx).sFormat
    tgFormatInfo(NewIdx).iUstCode = 0

    SQLQuery = "Update shtt Set shttfmtCode = " & tgFormatInfo(NewIdx).iCode & " Where shttCode = " & tgStationInfo(StationIDX).iCode
    If blUpdateDatabase Then
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHandler:
        End If
    End If

    FormatIdx = NewIdx
    Call UpdateReport(BIAIdx, "Format was added and assigned (" & Trim(tmBIAInfo(BIAIdx).sFormat) & ")")
    mAddToRegionSet "F", tgStationInfo(StationIDX).iFormatCode, tgFormatInfo(NewIdx).iCode
    tgStationInfo(StationIDX).iFormatCode = tgFormatInfo(NewIdx).iCode
    AddFormat = True
    Exit Function

ErrHandler:
    gHandleError "AffErrorLog.txt", "frmImportBIA-AddFormat"
End Function

'***************************************************************************
'
'***************************************************************************
Private Function UpdateStationsMarket(StationIDX As Long, MktIdx As Long, BIAIdx As Long) As Integer
    Dim NewIdx As Long
    
    On Error GoTo ErrHandler:
    UpdateStationsMarket = False
    mAddToRegionSet "M", tgStationInfo(StationIDX).iMktCode, tgMarketInfo(MktIdx).iCode
    tgStationInfo(StationIDX).iMktCode = tgMarketInfo(MktIdx).iCode
    SQLQuery = "Update shtt Set shttMktCode = " & tgStationInfo(StationIDX).iMktCode & " Where shttCode = " & tgStationInfo(StationIDX).iCode
    If blUpdateDatabase Then
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHandler:
        End If
    End If

    Call UpdateReport(BIAIdx, "Market was assigned (" & Trim(tmBIAInfo(BIAIdx).sMarketName) & ")")
    UpdateStationsMarket = True
    Exit Function

ErrHandler:
    gHandleError "AffErrorLog.txt", "frmImportBIA-UpdateStationsMarket"
End Function

'***************************************************************************
'
'***************************************************************************
Private Function UpdateStationsOwner(StationIDX As Long, OwnerIdx As Long, BIAIdx As Long) As Integer
    Dim NewIdx As Long

    On Error GoTo ErrHandler:
    UpdateStationsOwner = False
    mAddToRegionSet "O", tgStationInfo(StationIDX).iOwnerCode, tgOwnerInfo(OwnerIdx).iCode
    tgStationInfo(StationIDX).iOwnerCode = tgOwnerInfo(OwnerIdx).iCode
    SQLQuery = "Update shtt Set shttOwnerArttCode = " & tgStationInfo(StationIDX).iOwnerCode & " Where shttCode = " & tgStationInfo(StationIDX).iCode
    If blUpdateDatabase Then
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHandler:
        End If
    End If

    Call UpdateReport(BIAIdx, "Owner was assigned (" & Trim(tmBIAInfo(BIAIdx).sOwnerName) & ")")
    UpdateStationsOwner = True
    Exit Function

ErrHandler:
    gHandleError "AffErrorLog.txt", "frmImportBIA-UpdateStationsOwner"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function UpdateStationsFormat(StationIDX As Long, FormatIdx As Long, BIAIdx As Long) As Integer
    Dim NewIdx As Long

    On Error GoTo ErrHandler:
    UpdateStationsFormat = False
    mAddToRegionSet "F", tgStationInfo(StationIDX).iFormatCode, tgFormatInfo(FormatIdx).iCode
    tgStationInfo(StationIDX).iFormatCode = tgFormatInfo(FormatIdx).iCode
    SQLQuery = "Update shtt Set shttfmtCode = " & tgStationInfo(StationIDX).iFormatCode & " Where shttCode = " & tgStationInfo(StationIDX).iCode
    If blUpdateDatabase Then
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHandler:
        End If
    End If

    Call UpdateReport(BIAIdx, "Format was assigned (" & Trim(tmBIAInfo(BIAIdx).sFormat) & ")")
    UpdateStationsFormat = True
    Exit Function

ErrHandler:
    gHandleError "AffErrorLog.txt", "frmImportBIA-UpdateStationsFormat"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function UpdateRank(MktIdx As Long, BIAIdx As Long) As Integer
    On Error GoTo ErrHandler:
    UpdateRank = False
    SQLQuery = "Update MKT Set mktRank = '" & tmBIAInfo(BIAIdx).iRank & "' Where mktCode = " & tgMarketInfo(MktIdx).iCode
    If blUpdateDatabase Then
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoSub ErrHandler:
        End If
    End If
    
    Call UpdateReport(BIAIdx, "Market rank was changed from " & Trim(Str(tgMarketInfo(MktIdx).iRank)) & " to " & Trim(Str(tmBIAInfo(BIAIdx).iRank)))
    tgMarketInfo(MktIdx).iRank = tmBIAInfo(BIAIdx).iRank
    UpdateRank = True
    Exit Function

ErrHandler:
    gHandleError "AffErrorLog.txt", "frmImportBIA-UpdateRank"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Sub AddMarketToUsedArray(MktIdx As Long)
    Dim NewIdx As Long
    Dim llLoop As Long
    
    For llLoop = 0 To UBound(tmUsedMarkets) - 1 Step 1
        If tgMarketInfo(MktIdx).iCode = tmUsedMarkets(llLoop).iCode Then
            Exit Sub
        End If
    Next llLoop
    NewIdx = UBound(tmUsedMarkets)
    ReDim Preserve tmUsedMarkets(0 To NewIdx + 1) As MARKETINFO
    tmUsedMarkets(NewIdx).iCode = tgMarketInfo(MktIdx).iCode
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub AddOwnerToUsedArray(OwnerIdx As Long)
    Dim NewIdx As Long
    Dim llLoop As Long
    
    For llLoop = 0 To UBound(tmUsedOwners) - 1 Step 1
        If tgOwnerInfo(OwnerIdx).iCode = tmUsedOwners(llLoop).iCode Then
            Exit Sub
        End If
    Next llLoop
    NewIdx = UBound(tmUsedOwners)
    ReDim Preserve tmUsedOwners(0 To NewIdx + 1) As OWNERINFO
    tmUsedOwners(NewIdx).iCode = tgOwnerInfo(OwnerIdx).iCode
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub AddFormatToUsedArray(FormatIdx As Long)
    Dim NewIdx As Long
    Dim llLoop As Long
    
    For llLoop = 0 To UBound(tmUsedFormats) - 1 Step 1
        If tgFormatInfo(FormatIdx).iCode = tmUsedFormats(llLoop).iCode Then
            Exit Sub
        End If
    Next llLoop
    NewIdx = UBound(tmUsedFormats)
    ReDim Preserve tmUsedFormats(0 To NewIdx + 1) As FORMATINFO
    tmUsedFormats(NewIdx).iCode = tgFormatInfo(FormatIdx).iCode
End Sub

'***************************************************************************
'
'***************************************************************************
Private Function ScanForBrokenMarketLinks() As Integer
    Dim llStationIdx As Long
    Dim llMktIdx As Long
    Dim LinkIsOk As Boolean
    Dim llIdx As Long
    Dim llSef As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHandler
    SetResults "Verifying station market information...", RGB(0, 0, 0)

    If UBound(tgStationInfo) < 1 Then
        Exit Function
    End If
    If UBound(tgMarketInfo) < 1 Then
        Exit Function
    End If
    
'Moved to mSetUsedArrays
'    For llStationIdx = 0 To UBound(tgStationInfo)
'        DoEvents
'        If imTerminate Then
'            Exit Function
'        End If
'        LinkIsOk = False
'        If tgStationInfo(llStationIdx).iMktCode > 0 Then
'            For llMktIdx = 0 To UBound(tgMarketInfo) - 1
'                If tgStationInfo(llStationIdx).iMktCode = tgMarketInfo(llMktIdx).iCode Then
'                    ' If this market name is blank, don't add it. It will get deleted later.
'                    If Len(Trim(tgMarketInfo(llMktIdx).sName)) > 0 Then
'                        llIdx = UBound(tmUsedMarkets)
'                        ReDim Preserve tmUsedMarkets(0 To llIdx + 1) As MARKETINFO
'                        tmUsedMarkets(llIdx).iCode = tgStationInfo(llStationIdx).iMktCode
'                        llMktIdx = UBound(tgMarketInfo) ' Exit out of this loop.
'                        LinkIsOk = True
'                        Exit For
'                    End If
'                End If
'            Next
'            If Not LinkIsOk Then
'                ' The station is not pointing to a valid market name.
'                SQLQuery = "Update shtt Set shttMktCode = 0 Where shttCode = " & tgStationInfo(llStationIdx).iCode
'                If blUpdateDatabase Then
'                    'cnn.Execute SQLQuery, rdExecDirect
'                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                        GoSub ErrHandler:
'                    End If
'                End If
'                Call UpdateReport(-1, Trim(tgStationInfo(llStationIdx).sCallLetters) & " Bad market pointer was removed. Station now has no market.")
'                ScanForBrokenMarketLinks = ScanForBrokenMarketLinks + 1
'            End If
'        End If
'    Next
    
    ' At this point the arrary tmUsedMarkets contains a list of valid markets. Any not found in this
    ' list will now be deleted.
    If UBound(tmUsedMarkets) < 1 Then
        Exit Function
    End If
    For llMktIdx = 0 To UBound(tgMarketInfo) - 1 Step 1
        LinkIsOk = False
        For llIdx = 0 To UBound(tmUsedMarkets) - 1 Step 1
            If tmUsedMarkets(llIdx).iCode = tgMarketInfo(llMktIdx).iCode Then
                llIdx = UBound(tmUsedMarkets) ' Exit out of this loop.
                LinkIsOk = True
            End If
        Next
        If Not LinkIsOk Then
            ilRet = True
            If blUpdateDatabase Then
                For llSef = 0 To UBound(tmBIARegionSet) - 1 Step 1
                    If tmBIARegionSet(llSef).sCategory = "M" Then
                        If tmBIARegionSet(llSef).iFromCode = tgMarketInfo(llMktIdx).iCode Then
                            ilRet = gUpdateRegions("M", tmBIARegionSet(llSef).iFromCode, tmBIARegionSet(llSef).iToCode, "BIAImportLog.Txt")
                            Exit For
                        End If
                    End If
                Next llSef
            End If
            If ilRet Then
                ' This market name is not being used. Remove it.
                SQLQuery = "Delete From MKT Where mktCode = " & tgMarketInfo(llMktIdx).iCode
                If blUpdateDatabase And ilRet Then
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHandler:
                    End If
                End If
                Call UpdateReport(-1, "Bad market entry (" & tgMarketInfo(llMktIdx).sName & ") was removed.")
            End If
        End If
    Next
    Erase tmUsedMarkets
    Exit Function
    
ErrHandler:
    gHandleError "AffErrorLog.txt", "frmImportBIA-ScanForBrokenMarketLinks"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function ScanForBrokenOwnerLinks() As Integer
    Dim llStationIdx As Long
    Dim llOwnerIdx As Long
    Dim LinkIsOk As Boolean
    Dim llIdx As Long
    Dim llSef As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHandler
    SetResults "Verifying station owner information...", RGB(0, 0, 0)

    If UBound(tgStationInfo) < 1 Then
        Exit Function
    End If
    If UBound(tgOwnerInfo) < 1 Then
        Exit Function
    End If
    
'Moved to mSetUsedArrays
'    For llStationIdx = 0 To UBound(tgStationInfo)
'        DoEvents
'        If imTerminate Then
'            Exit Function
'        End If
'        LinkIsOk = False
'        If tgStationInfo(llStationIdx).iOwnerCode > 0 Then  ' Look only when a link exist.
'            For llOwnerIdx = 0 To UBound(tgOwnerInfo)
'                If tgStationInfo(llStationIdx).iOwnerCode = tgOwnerInfo(llOwnerIdx).iCode Then
'                    llIdx = UBound(tmUsedOwners)
'                    ReDim Preserve tmUsedOwners(0 To llIdx + 1) As OWNERINFO
'                    tmUsedOwners(llIdx).iCode = tgStationInfo(llStationIdx).iOwnerCode
'                    llOwnerIdx = UBound(tgOwnerInfo) ' Exit out of this loop.
'                    LinkIsOk = True
'                End If
'            Next
'            If Not LinkIsOk Then
'                ' The station is not pointing to a valid owner name.
'                SQLQuery = "Update shtt Set shttOwnerArttCode = 0 Where shttCode = " & tgStationInfo(llStationIdx).iCode
'                If blUpdateDatabase Then
'                    'cnn.Execute SQLQuery, rdExecDirect
'                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                        GoSub ErrHandler:
'                    End If
'                End If
'                Call UpdateReport(-1, Trim(tgStationInfo(llStationIdx).sCallLetters) & " Bad owner pointer was removed. Station now has no owner.")
'            End If
'        End If
'    Next

    ' At this point the arrary tmUsedOwners contains a list of valid owners. Any not found in this
    ' list will now be deleted.
    If UBound(tmUsedOwners) < 1 Then
        Exit Function
    End If
    For llOwnerIdx = 0 To UBound(tgOwnerInfo) - 1 Step 1
        LinkIsOk = False
        For llIdx = 0 To UBound(tmUsedOwners) - 1 Step 1
            If tmUsedOwners(llIdx).iCode = tgOwnerInfo(llOwnerIdx).iCode Then
                llIdx = UBound(tmUsedOwners) ' Exit out of this loop.
                LinkIsOk = True
            End If
        Next
        If Not LinkIsOk Then
            ilRet = True
            If blUpdateDatabase Then
                For llSef = 0 To UBound(tmBIARegionSet) - 1 Step 1
                    If tmBIARegionSet(llSef).sCategory = "O" Then
                        If tmBIARegionSet(llSef).iFromCode = tgOwnerInfo(llOwnerIdx).iCode Then
                            ilRet = gUpdateRegions("O", tmBIARegionSet(llSef).iFromCode, tmBIARegionSet(llSef).iToCode, "BIAImportLog.Txt")
                            Exit For
                        End If
                    End If
                Next llSef
            End If
            If ilRet Then
                ' This market name is not being used. Remove it.
                SQLQuery = "Delete From artt Where arttCode = " & tgOwnerInfo(llOwnerIdx).iCode & " And arttType = 'O'"
                If blUpdateDatabase Then
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHandler:
                    End If
                End If
                Call UpdateReport(-1, "Bad owner entry (" & tgOwnerInfo(llOwnerIdx).sName & ") was removed.")
            End If
        End If
    Next
    Erase tmUsedOwners
    Exit Function

ErrHandler:
    gHandleError "AffErrorLog.txt", "frmImportBIA-ScanForBrokenOwnerLinks"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function ScanForBrokenFormatLinks() As Integer
    Dim llStationIdx As Long
    Dim llFmtIdx As Long
    Dim LinkIsOk As Boolean
    Dim llIdx As Long
    Dim llSef As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHandler
    SetResults "Verifying station format information...", RGB(0, 0, 0)

    If UBound(tgStationInfo) < 1 Then
        Exit Function
    End If
    If UBound(tgFormatInfo) < 1 Then
        Exit Function
    End If
    
'Moved to mSetUsedArrays
'    For llStationIdx = 0 To UBound(tgStationInfo)
'        DoEvents
'        If imTerminate Then
'            Exit Function
'        End If
'        LinkIsOk = False
'        If tgStationInfo(llStationIdx).iFormatCode > 0 Then  ' Look only when a link exist.
'            For llFmtIdx = 0 To UBound(tgFormatInfo)
'                If tgStationInfo(llStationIdx).iFormatCode = tgFormatInfo(llFmtIdx).iCode Then
'                    llIdx = UBound(tmUsedFormats)
'                    ReDim Preserve tmUsedFormats(0 To llIdx + 1) As FORMATINFO
'                    tmUsedFormats(llIdx).iCode = tgStationInfo(llStationIdx).iFormatCode
'                    llFmtIdx = UBound(tgFormatInfo) ' Exit out of this loop.
'                    LinkIsOk = True
'                End If
'            Next
'            If Not LinkIsOk Then
'                ' The station is not pointing to a valid format name.
'                SQLQuery = "Update shtt Set shttFmtCode = 0 Where shttCode = " & tgStationInfo(llStationIdx).iCode
'                If blUpdateDatabase Then
'                    'cnn.Execute SQLQuery, rdExecDirect
'                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                        GoSub ErrHandler:
'                    End If
'                End If
'                Call UpdateReport(-1, Trim(tgStationInfo(llStationIdx).sCallLetters) & " Bad format pointer was removed. Station now has no format.")
'            End If
'        End If
'    Next

    ' At this point the arrary tmUsedFormats contains a list of valid formats. Any not found in this
    ' list will now be deleted.
    If UBound(tmUsedFormats) < 1 Then
        Exit Function
    End If
    For llFmtIdx = 0 To UBound(tgFormatInfo) - 1 Step 1
        LinkIsOk = False
        For llIdx = 0 To UBound(tmUsedFormats) - 1 Step 1
            If tmUsedFormats(llIdx).iCode = tgFormatInfo(llFmtIdx).iCode Then
                llIdx = UBound(tmUsedFormats) ' Exit out of this loop.
                LinkIsOk = True
            End If
        Next
        If Not LinkIsOk Then
            ilRet = True
            If blUpdateDatabase Then
                For llSef = 0 To UBound(tmBIARegionSet) - 1 Step 1
                    If tmBIARegionSet(llSef).sCategory = "F" Then
                        If tmBIARegionSet(llSef).iFromCode = tgFormatInfo(llFmtIdx).iCode Then
                            ilRet = gUpdateRegions("F", tmBIARegionSet(llSef).iFromCode, tmBIARegionSet(llSef).iToCode, "BIAImportLog.Txt")
                            Exit For
                        End If
                    End If
                Next llSef
            End If
            If ilRet Then
                ' This format name is not being used. Remove it.
                SQLQuery = "Delete From FMT_Station_Format Where fmtCode = " & tgFormatInfo(llFmtIdx).iCode & ""
                If blUpdateDatabase Then
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHandler:
                    End If
                End If
                Call UpdateReport(-1, "Bad format entry (" & tgFormatInfo(llFmtIdx).sName & ") was removed.")
            End If
        End If
    Next
    Erase tmUsedFormats
    Exit Function

ErrHandler:
    gHandleError "AffErrorLog.txt", "frmImportBIA-ScanForBrokenFormatLinks"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function ScanStationsNotUpdated() As Long
    Dim llLoop As Long
    Dim llBIAIdx As Long
    Dim slCallLetters As String

    SetResults "Reporting stations not in BIA file...", RGB(0, 0, 0)
    Call UpdateReport(-1, "Stations not updated report.")
    Call UpdateReport(-1, "This list shows stations not in the BIA file and therefore could not be validated")
    Call UpdateReport(-1, "---------------------------------------------------------------------------------")
    For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        DoEvents
        If imTerminate Then
            Exit Function
        End If
        slCallLetters = Trim(tgStationInfo(llLoop).sCallLetters)
        llBIAIdx = mBinarySearchBIAStation(slCallLetters)
        If llBIAIdx = -1 Then
            Call UpdateReport(-1, slCallLetters)
        End If
    Next
    Call UpdateReport(-1, "---------------------------------------------------------------------------------")
End Function

'***************************************************************************
'
'***************************************************************************
'Private Function TEST_FixBadStationNames() As Long
'    Dim llLoop As Long
'    Dim llBIAIdx As Long
'    Dim slCallLetters As String
'
'    On Error GoTo ErrHand
'    SetResults "Deleting stations not in BIA file...", RGB(0, 0, 0)
'    Call UpdateReport(-1, "---------------------------------------------------------------------------------")
'    For llLoop = 0 To UBound(tgStationInfo)
'        DoEvents
'        If imTerminate Then
'            Exit Function
'        End If
'        slCallLetters = Trim(tgStationInfo(llLoop).sCallLetters)
'        llBIAIdx = mBinarySearchBIAStation(slCallLetters)
'        If llBIAIdx = -1 Then
'            ' This station does not exist. Lets delete it.
'            SQLQuery = "Delete From shtt Where shttCallLetters = '" & slCallLetters & "'"
'            cnn.Execute SQLQuery, rdExecDirect
'            Call UpdateReport(-1, slCallLetters)
'        End If
'    Next
'    Call UpdateReport(-1, "---------------------------------------------------------------------------------")
'ErrHand:
'Resume Next
'End Function


'***************************************************************************
'
'***************************************************************************
Private Function mBinarySearchBIAStation(sCallLetters As String) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim ilResult As Integer
    
    On Error GoTo ErrHand
    
    mBinarySearchBIAStation = -1    ' Start out as not found.
    llMin = LBound(tmBIAInfo)
    llMax = UBound(tmBIAInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        ilResult = StrComp(Trim(tmBIAInfo(llMiddle).sCallLetters), sCallLetters, vbTextCompare)
        Select Case ilResult
            Case 0:
                mBinarySearchBIAStation = llMiddle  ' Found it !
                Exit Function
            Case 1:
                llMax = llMiddle - 1
            Case -1:
                llMin = llMiddle + 1
        End Select
    Loop
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in mBinarySearchBIAStation: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "BIAImportLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    Exit Function
    
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mBinarySearchMarketName(sMarketName As String) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim ilResult As Integer
    
    On Error GoTo ErrHand
    
    mBinarySearchMarketName = -1    ' Start out as not found.
    llMin = LBound(tgMarketInfo)
    llMax = UBound(tgMarketInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        ilResult = StrComp(Trim(tgMarketInfo(llMiddle).sName), sMarketName, vbTextCompare)
        Select Case ilResult
            Case 0:
                mBinarySearchMarketName = llMiddle  ' Found it !
                Exit Function
            Case 1:
                llMax = llMiddle - 1
            Case -1:
                llMin = llMiddle + 1
        End Select
    Loop
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in mBinarySearchMarketName: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number, "BIAImportLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    Exit Function
    
End Function

'***************************************************************************
'
'***************************************************************************
Private Function LookupStation(slCallLetters As String) As Long
    Dim llLoop As Long
    
    LookupStation = -1
    On Error GoTo ErrHandler
    For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        If StrComp(Trim(tgStationInfo(llLoop).sCallLetters), Trim(slCallLetters), vbTextCompare) = 0 Then
            LookupStation = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function LookupBIAStation(CallLetters As String) As Long
    Dim llLoop As Long
    Dim slCallLetters As String
    
    LookupBIAStation = -1
    On Error GoTo ErrHandler
    For llLoop = 0 To lmTotalBIARecords
        slCallLetters = Trim(tmBIAInfo(llLoop).sCallLetters)
        If StrComp(slCallLetters, Trim(CallLetters), vbTextCompare) = 0 Then
            LookupBIAStation = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function LookupMarket(mktCode As Integer) As Long
    Dim llLoop As Long
    
    LookupMarket = -1
    On Error GoTo ErrHandler
    If UBound(tgMarketInfo) < 1 Then
        Exit Function
    End If
    For llLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
        If tgMarketInfo(llLoop).iCode = mktCode Then
            LookupMarket = llLoop
            Exit Function
        End If
    Next
    Exit Function
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function LookupOwner(OwnerCode As Integer) As Long
    Dim llLoop As Long
    
    LookupOwner = -1
    On Error GoTo ErrHandler
    If UBound(tgOwnerInfo) < 1 Then
        Exit Function
    End If
    For llLoop = 0 To UBound(tgOwnerInfo) - 1 Step 1
        If tgOwnerInfo(llLoop).iCode = OwnerCode Then
            LookupOwner = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function LookupFormat(fmtCode As Integer) As Long
    Dim llLoop As Long
    
    LookupFormat = -1
    On Error GoTo ErrHandler
    If UBound(tgFormatInfo) < 1 Then
        Exit Function
    End If
    For llLoop = 0 To UBound(tgFormatInfo) - 1 Step 1
        If tgFormatInfo(llLoop).iCode = fmtCode Then
            LookupFormat = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function LookupMarketByName(sMarketName As String) As Long
    Dim llLoop As Long
    Dim llIdx As Long

    LookupMarketByName = -1
    On Error GoTo ErrHandler
    llIdx = UBound(tgMarketInfo)
    If llIdx < 1 Then
        Exit Function
    End If
    For llLoop = 0 To llIdx
        If StrComp(Trim(tgMarketInfo(llLoop).sName), Trim(sMarketName), vbTextCompare) = 0 Then
            LookupMarketByName = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function LookupOwnerByName(sOwnerName As String) As Long
    Dim llLoop As Long
    Dim llIdx As Long

    LookupOwnerByName = -1
    On Error GoTo ErrHandler
    llIdx = UBound(tgOwnerInfo)
    If llIdx < 1 Then
        Exit Function
    End If
    For llLoop = 0 To llIdx
        If StrComp(Trim(tgOwnerInfo(llLoop).sName), Trim(sOwnerName), vbTextCompare) = 0 Then
            LookupOwnerByName = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function LookupFormatByName(sFormatName As String) As Long
    Dim llLoop As Long
    Dim llIdx As Long

    LookupFormatByName = -1
    On Error GoTo ErrHandler
    llIdx = UBound(tgFormatInfo)
    If llIdx < 1 Then
        Exit Function
    End If
    For llLoop = 0 To llIdx
        If StrComp(Trim(tgFormatInfo(llLoop).sName), Trim(sFormatName), vbTextCompare) = 0 Then
            LookupFormatByName = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mMakeStationName(StationName As String, Band As String) As String
    Dim slName As String

    slName = Trim(Replace(StationName, "-", "")) & "-" & Trim(Replace(Band, "-", ""))
    mMakeStationName = slName
End Function

'***************************************************************************
'
'***************************************************************************
Private Function CreateBIAStatusReport()
    Dim llBIAReportIdx As Long
    Dim ilLoop As Long
    Dim hlFile As Integer
    Dim slStatus As String
    Dim slStartCallLetters As String
    Dim slCurrentCallLetters As String
    
    CreateBIAStatusReport = False
    SetResults "Creating Status Report...", RGB(0, 0, 0)
    llBIAReportIdx = UBound(tmBIAReportInfo)
    If llBIAReportIdx < 1 Then
        CreateBIAStatusReport = True
        Exit Function
    End If
    On Error GoTo IgnoreError
    Kill smReportPathFileName
    On Error GoTo ErrHandler
    'hlFile = FreeFile
    'Open smReportPathFileName For Append As hlFile
    ilRet = gFileOpen(smReportPathFileName, "Append", hlFile)
    If Not blUpdateDatabase Then
        Print #hlFile, "*** SHOW REPORT ONLY WAS CHECKED."
        Print #hlFile, "*** NO CHANGES WERE MADE TO THE DATABASE"
        Print #hlFile, "*** THE FOLLOWING INFORMATION SHOWS WHAT WOULD HAVE OCCURRED"
        Print #hlFile, "***"
    End If
    slStartCallLetters = ""
    For ilLoop = 0 To llBIAReportIdx
'        slCurrentCallLetters = tmBIAReportInfo(ilLoop).sCallLetters
'        If slStartCallLetters <> slCurrentCallLetters Then
'            slStartCallLetters = slCurrentCallLetters
'            Print #hlFile, slStartCallLetters
'        End If
'        slStatus = "    " & Trim(tmBIAReportInfo(ilLoop).sReportInfo)
'        Print #hlFile, slStatus

        slStatus = Trim(tmBIAReportInfo(ilLoop).sReportInfo)
        If Len(Trim(tmBIAReportInfo(ilLoop).sCallLetters)) > 0 Then
            Print #hlFile, Trim(tmBIAReportInfo(ilLoop).sCallLetters) & ", ", slStatus
        Else
            Print #hlFile, slStatus
        End If
    Next
    Close hlFile
    CreateBIAStatusReport = True
    Exit Function

IgnoreError:
    Resume Next
ErrHandler:
    SetResults "Error creating status report.", RGB(255, 0, 0)
End Function

'***************************************************************************
'
'***************************************************************************
Private Sub UpdateReport(BIAIdx As Long, sMsg As String)
    Dim llIdx As Long
    
    llIdx = UBound(tmBIAReportInfo)
    ReDim Preserve tmBIAReportInfo(0 To llIdx + 1) As BIAREPORTINFO
    If BIAIdx <> -1 Then
        tmBIAReportInfo(llIdx).sCallLetters = Trim(tmBIAInfo(BIAIdx).sCallLetters)
    Else
        tmBIAReportInfo(llIdx).sCallLetters = ""
    End If
    tmBIAReportInfo(llIdx).sReportInfo = sMsg
End Sub

'***************************************************************************
'
'***************************************************************************
Private Function mCheckFile()
    sgBIAFileName = txtFile.Text
    frmBIACheck.Show vbModal
    mCheckFile = igBIARetStatus
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Sub SetResults(Msg As String, FGC As Long)
    gLogMsg Msg, "BIAImportLog.Txt", False
    lbcMsg.AddItem Msg
    lbcMsg.ListIndex = lbcMsg.ListCount - 1
    lbcMsg.ForeColor = FGC
    DoEvents
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub cmcBrowse_Click()

    Dim slCurDir As String
    
    slCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    'Set default Folder
    CommonDialog1.InitDir = sgImportDirectory
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 3
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    txtFile.Text = Trim$(CommonDialog1.FileName)
    ChDir slCurDir
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub cmdCancel_Click()
    Dim ilResp As Integer
    
    If imImporting Then
        ilResp = gMsgBox("Are you sure you want to abort?", vbYesNo)
        If ilResp = vbYes Then
            imTerminate = True
        End If
        Exit Sub
    End If
    Unload frmImportBIA
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub cmdViewReport_Click()
    Call ShowReport(True)
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub ShowReport(PromptIfMissing As Boolean)
    Dim sCmd As String
    Dim slDateTime As String
    Dim ilRet As Integer
    
    On Error GoTo ErrHandler
    ilRet = 0
    slDateTime = FileDateTime(smReportPathFileName)
    If ilRet = -1 Then
        If PromptIfMissing Then
            gMsgBox "There is no report file to view"
        End If
        Exit Sub
    End If

    sCmd = "Notepad.exe " & smReportPathFileName
    Call Shell(sCmd, vbNormalFocus)
    Exit Sub
    
ErrHandler:
    ilRet = -1
    Resume Next
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = vbNormal
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imImporting Then
        Screen.MousePointer = vbHourglass
    End If
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    
    Screen.MousePointer = vbHourglass
    frmImportBIA.Caption = "BIA Station Information - " & sgClientName
    imTerminate = False
    imImporting = False
    chkReportStationsNotUpdated.Value = 1
    
    txtFile.Text = sgImportDirectory & "BIA_Data.txt"
    smReportPathFileName = sgMsgDirectory & "BIAStatus.txt"
    Screen.MousePointer = vbDefault
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    If Not blUpdateDatabase Then
        ' If the database was not updated, then reload these arrays.
        If Not gPopMarkets() Then
            Call FailBIAImport("Unable to Load Existing Market Names.")
            Exit Sub
        End If
        If Not gPopOwnerNames() Then
            Call FailBIAImport("Unable to Load Existing Owner Names.")
            Exit Sub
        End If
        If Not gPopFormats() Then
            Call FailBIAImport("Unable to Load Existing Format Names.")
            Exit Sub
        End If
    End If
    Erase tmBIAReportInfo
    Erase tmBIARegionSet
    Erase tmBIAInfo
    Set frmImportBIA = Nothing
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub FailBIAImport(sMsg As String)
    SetResults sMsg, RGB(255, 0, 0)
    imImporting = False
    cmdViewReport.Enabled = True
    cmdCancel.Caption = "&Done"
    cmdCancel.SetFocus
    Screen.MousePointer = vbDefault
End Sub


Private Function mRemoveDuplicateMarketNames() As Integer
    Dim llOutsideLoop As Long
    Dim llInsideLoop As Long
    Dim ilRet As Integer
    
    mRemoveDuplicateMarketNames = True
    On Error GoTo ErrHand:
    SetResults "Checking and removing duplicated Market Names...", RGB(0, 0, 0)
    For llOutsideLoop = LBound(tgMarketInfo) To UBound(tgMarketInfo) - 1 Step 1
        If tgMarketInfo(llOutsideLoop).iCode > 0 Then
            For llInsideLoop = llOutsideLoop + 1 To UBound(tgMarketInfo) - 1 Step 1
                If StrComp(tgMarketInfo(llInsideLoop).sName, tgMarketInfo(llOutsideLoop).sName, vbTextCompare) = 0 Then
                    SQLQuery = "Update shtt Set shttMktCode = " & tgMarketInfo(llOutsideLoop).iCode & " Where shttMktCode = " & tgMarketInfo(llInsideLoop).iCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHand:
                    End If
                    SQLQuery = "Update mat Set matMktCode = " & tgMarketInfo(llOutsideLoop).iCode & " Where matMktCode = " & tgMarketInfo(llInsideLoop).iCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHand:
                    End If
                    SQLQuery = "Update mgt Set mgtMktCode = " & tgMarketInfo(llOutsideLoop).iCode & " Where mgtMktCode = " & tgMarketInfo(llInsideLoop).iCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHand:
                    End If
                    'SQLQuery = "SELECT * FROM raf_region_area WHERE ((rafCategory = 'M') and (rafType = 'C' OR rafType = 'N'))"
                    'Set rst_Raf = cnn.Execute(SQLQuery)
                    'Do While Not rst_Raf.EOF
                    '    SQLQuery = "SELECT * FROM sef_Split_Entity WHERE sefIntCode = " & tgMarketInfo(llOutsideLoop).iCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
                    '    Set rst_Sef = cnn.Execute(SQLQuery)
                    '    If rst_Sef.EOF Then
                    '        SQLQuery = "Update sef_Split_Entity Set sefIntCode = " & tgMarketInfo(llOutsideLoop).iCode & " Where sefIntCode = " & tgMarketInfo(llInsideLoop).iCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
                    '        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '            GoSub ErrHand:
                    '        End If
                    '    Else
                    '        SQLQuery = "DELETE FROM sef_Split_Entity WHERE sefIntCode = " & tgMarketInfo(llInsideLoop).iCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
                    '        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '            GoSub ErrHand:
                    '        End If
                    '    End If
                    '    rst_Raf.MoveNext
                    'Loop
                    ilRet = gUpdateRegions("M", tgMarketInfo(llInsideLoop).iCode, tgMarketInfo(llOutsideLoop).iCode, "BIAImportLog.Txt")
                    If Not ilRet Then
                        mRemoveDuplicateMarketNames = False
                        Exit Function
                    End If
                    SQLQuery = "DELETE FROM mkt WHERE mktCode = " & tgMarketInfo(llInsideLoop).iCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHand:
                    End If
                    tgMarketInfo(llInsideLoop).iCode = -1
                End If
            Next llInsideLoop
        End If
    Next llOutsideLoop
    
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmImportBIA-mRemoveDuplicateMarket-Names"
    mRemoveDuplicateMarketNames = False
    Exit Function
End Function

Private Function mSetUsedArrays() As Integer
    Dim llStationIdx As Long
    Dim llMktIdx As Long
    Dim llOwnerIdx As Long
    Dim llFmtIdx As Long
    Dim ilLinkIsOk As Integer
    
    mSetUsedArrays = True
    On Error GoTo ErrHandler
    
    If UBound(tgStationInfo) < 1 Then
        Exit Function
    End If

    For llStationIdx = 0 To UBound(tgStationInfo) - 1 Step 1
        DoEvents
        If imTerminate Then
            Exit Function
        End If
        ilLinkIsOk = False
        If (tgStationInfo(llStationIdx).iMktCode > 0) And (UBound(tgMarketInfo) > 0) Then
            For llMktIdx = 0 To UBound(tgMarketInfo) - 1 Step 1
                If tgStationInfo(llStationIdx).iMktCode = tgMarketInfo(llMktIdx).iCode Then
                    ' If this market name is blank, don't add it. It will get deleted later.
                    If Len(Trim(tgMarketInfo(llMktIdx).sName)) > 0 Then
                        AddMarketToUsedArray llMktIdx
                        ilLinkIsOk = True
                        Exit For
                    End If
                End If
            Next
            If Not ilLinkIsOk Then
                ' The station is not pointing to a valid market name.
                SQLQuery = "Update shtt Set shttMktCode = 0 Where shttCode = " & tgStationInfo(llStationIdx).iCode
                If blUpdateDatabase Then
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHandler:
                    End If
                End If
                Call UpdateReport(-1, Trim(tgStationInfo(llStationIdx).sCallLetters) & " Bad market pointer was removed. Station now has no market.")
            End If
        End If
    
        ilLinkIsOk = False
        If (tgStationInfo(llStationIdx).iOwnerCode > 0) And (UBound(tgOwnerInfo) > 0) Then ' Look only when a link exist.
            For llOwnerIdx = 0 To UBound(tgOwnerInfo) - 1 Step 1
                If tgStationInfo(llStationIdx).iOwnerCode = tgOwnerInfo(llOwnerIdx).iCode Then
                    AddOwnerToUsedArray llOwnerIdx
                    ilLinkIsOk = True
                End If
            Next
            If Not ilLinkIsOk Then
                ' The station is not pointing to a valid owner name.
                SQLQuery = "Update shtt Set shttOwnerArttCode = 0 Where shttCode = " & tgStationInfo(llStationIdx).iCode
                If blUpdateDatabase Then
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHandler:
                    End If
                End If
                Call UpdateReport(-1, Trim(tgStationInfo(llStationIdx).sCallLetters) & " Bad owner pointer was removed. Station now has no owner.")
            End If
        End If
    
        ilLinkIsOk = False
        If (tgStationInfo(llStationIdx).iFormatCode > 0) And (UBound(tgFormatInfo) > 0) Then ' Look only when a link exist.
            For llFmtIdx = 0 To UBound(tgFormatInfo) - 1 Step 1
                If tgStationInfo(llStationIdx).iFormatCode = tgFormatInfo(llFmtIdx).iCode Then
                    AddFormatToUsedArray llFmtIdx
                    ilLinkIsOk = True
                End If
            Next
            If Not ilLinkIsOk Then
                ' The station is not pointing to a valid format name.
                SQLQuery = "Update shtt Set shttFmtCode = 0 Where shttCode = " & tgStationInfo(llStationIdx).iCode
                If blUpdateDatabase Then
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        GoSub ErrHandler:
                    End If
                End If
                Call UpdateReport(-1, Trim(tgStationInfo(llStationIdx).sCallLetters) & " Bad format pointer was removed. Station now has no format.")
            End If
        End If
    Next
    Exit Function

ErrHandler:
    gHandleError "AffErrorLog.txt", "frmImportBIA-mSetUserArrays"
    mSetUsedArrays = False
    Exit Function
End Function

Private Function mUpdateRegions(slCategory As String, ilFromCode As Integer, ilToCode As Integer) As Integer
    Dim rst_Raf As ADODB.Recordset
    Dim rst_Sef As ADODB.Recordset

    mUpdateRegions = True
    On Error GoTo ErrHand:
    SQLQuery = "SELECT * FROM raf_region_area WHERE ((rafCategory = '" & slCategory & "') and (rafType = 'C' OR rafType = 'N'))"
    Set rst_Raf = cnn.Execute(SQLQuery)
    Do While Not rst_Raf.EOF
        SQLQuery = "SELECT * FROM sef_Split_Entity WHERE sefIntCode = " & ilToCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
        Set rst_Sef = cnn.Execute(SQLQuery)
        If rst_Sef.EOF Then
            SQLQuery = "Update sef_Split_Entity Set sefIntCode = " & ilToCode & " Where sefIntCode = " & ilFromCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                GoSub ErrHand:
            End If
        Else
            SQLQuery = "DELETE FROM sef_Split_Entity WHERE sefIntCode = " & ilFromCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                GoSub ErrHand:
            End If
        End If
        rst_Raf.MoveNext
    Loop
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmImportBIA-mUpdateRegions"
    mUpdateRegions = False
    Exit Function
End Function

Private Sub mAddToRegionSet(slCategory As String, ilFromCode As Integer, ilToCode As Integer)
    Dim ilUpper As Integer
    ilUpper = UBound(tmBIARegionSet)
    tmBIARegionSet(ilUpper).sCategory = slCategory
    tmBIARegionSet(ilUpper).iFromCode = ilFromCode
    tmBIARegionSet(ilUpper).iToCode = ilToCode
    ReDim Preserve tmBIARegionSet(0 To ilUpper + 1) As BIAREGIONSET
End Sub
