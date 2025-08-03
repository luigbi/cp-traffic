VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportAiredSpot 
   Caption         =   "Import Aired Spots"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   Icon            =   "AffImportAiredSpot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   6195
   Begin VB.TextBox txtFile 
      Height          =   300
      Left            =   990
      TabIndex        =   6
      Top             =   480
      Width           =   3600
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "Browse"
      Height          =   300
      Left            =   4845
      TabIndex        =   5
      Top             =   480
      Width           =   1065
   End
   Begin VB.ListBox lbcMsg 
      Enabled         =   0   'False
      Height          =   2205
      ItemData        =   "AffImportAiredSpot.frx":08CA
      Left            =   120
      List            =   "AffImportAiredSpot.frx":08CC
      TabIndex        =   1
      Top             =   1410
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5895
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbcFile 
      Caption         =   "Import File"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   495
      Width           =   780
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   150
      TabIndex        =   4
      Top             =   3765
      Width           =   5490
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
End
Attribute VB_Name = "frmImportAiredSpot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmContact - allows for selection of station/vehicle/advertiser for contact information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text

Private smDate As String     'Export Date
Private imNumberDays As Integer
Private imVefCode As Integer
Private imAdfCode As Integer
Private smVefName As String
Private imAllClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
'Private hmMsg As Integer
Private hmTo As Integer
Private hmFrom As Integer
Private cprst As ADODB.Recordset
Private drst As ADODB.Recordset
Private tmCPDat() As DAT
Private tmAirSpotInfo() As AIRSPOTINFO
Private hmAst As Integer
Private tmAstInfo() As ASTINFO








'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
'Private Function mOpenMsgFile(sMsgFileName As String) As Integer
'    Dim slToFile As String
'    Dim slDateTime As String
'    Dim slFileDate As String
'    Dim slNowDate As String
'    Dim ilRet As Integer
'
'    On Error GoTo mOpenMsgFileErr:
'    ilRet = 0
'    slNowDate = Format$(gNow(), sgShowDateForm)
'    slToFile = sgMsgDirectory & "ImptAiredSpots.Txt"
'    slDateTime = FileDateTime(slToFile)
'    If ilRet = 0 Then
'        slFileDate = Format$(slDateTime, sgShowDateForm)
'        If DateValue(gAdjYear(slFileDate)) = DateValue(gAdjYear(slNowDate)) Then  'Append
'            On Error GoTo 0
'            ilRet = 0
'            On Error GoTo mOpenMsgFileErr:
'            hmMsg = FreeFile
'            Open slToFile For Append As hmMsg
'            If ilRet <> 0 Then
'                Close hmMsg
'                hmMsg = -1
'                gMsgBox "Open File " & slToFile & " error #" & Str$(Err.Number), vbOKOnly
'                mOpenMsgFile = False
'                Exit Function
'            End If
'        Else
'            Kill slToFile
'            On Error GoTo 0
'            ilRet = 0
'            On Error GoTo mOpenMsgFileErr:
'            hmMsg = FreeFile
'            Open slToFile For Output As hmMsg
'            If ilRet <> 0 Then
'                Close hmMsg
'                hmMsg = -1
'                gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
'                mOpenMsgFile = False
'                Exit Function
'            End If
'        End If
'    Else
'        On Error GoTo 0
'        ilRet = 0
'        On Error GoTo mOpenMsgFileErr:
'        hmMsg = FreeFile
'        Open slToFile For Output As hmMsg
'        If ilRet <> 0 Then
'            Close hmMsg
'            hmMsg = -1
'            gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
'            mOpenMsgFile = False
'            Exit Function
'        End If
'    End If
'    On Error GoTo 0
'    'Print #hmMsg, "** Import Aired Station Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
'    'Print #hmMsg, ""
'    sMsgFileName = slToFile
'    mOpenMsgFile = True
'    Exit Function
'mOpenMsgFileErr:
'    ilRet = 1
'    Resume Next
'End Function

Private Sub cmcBrowse_Click()

    Dim slCurDir As String
    
    slCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 3
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    txtFile.Text = Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub cmdExport_Click()
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
    Screen.MousePointer = vbHourglass
    If Not mCheckFile() Then
        Screen.MousePointer = vbDefault
        txtFile.SetFocus
        Exit Sub
    End If
'    If Not mOpenMsgFile(sMsgFileName) Then
'        cmdCancel.SetFocus
'        Exit Sub
'    End If
    imExporting = True
    On Error GoTo 0
    lbcMsg.Enabled = True
    lacResult.Caption = ""
    bgTaskBlocked = False
    sgTaskBlockedName = "Aired Station Spots Import"
    iRet = mImportSpots()
    If (iRet = False) Then
        bgTaskBlocked = False
        sgTaskBlockedName = ""
        gLogMsg "** Terminated - mImportSpots returned False**", "UnivisionImportLog.Txt", False
        'Print #hmMsg, "** Terminated **"
        'Close #hmMsg
        Close #hmTo
        imExporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    If imTerminate Then
        bgTaskBlocked = False
        sgTaskBlockedName = ""
        gLogMsg "** User Terminated **", "UnivisionImportLog.Txt", False
        'Print #hmMsg, "** User Terminated **"
        'Close #hmMsg
        Close #hmTo
        imExporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    'Clear old aet records out
    On Error GoTo ErrHand:
    If bgTaskBlocked Then
        gMsgBox "Some spots were blocked during the Import." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
    End If
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    imExporting = False
    gLogMsg "** Completed Import Aired Station Spots" & " **", "UnivisionImportLog.Txt", False
    gLogMsg "", "UnivisionImportLog.Txt", False
    'Print #hmMsg, "** Completed Import Aired Station Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    'Close #hmMsg
    lacResult.Caption = "Results: " & sMsgFileName
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Exit Sub
cmdExportErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportAireSpot-cmdExport"
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    Unload frmImportAiredSpot
End Sub


Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    
    Screen.MousePointer = vbHourglass
    frmImportAiredSpot.Caption = "Aired Station Spots - " & sgClientName
    imAllClick = False
    imTerminate = False
    imExporting = False
    gOpenMKDFile hmAst, "Ast.Mkd"
    
    txtFile.Text = sgImportDirectory & "CSISpots.txt"
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gCloseMKDFile hmAst, "Ast.Mkd"
    Erase tmAirSpotInfo
    Erase tmCPDat
    Erase tmAstInfo
    cprst.Close
    drst.Close
    Set frmImportAiredSpot = Nothing
End Sub

Private Function mImportSpots() As Integer
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim slDate As String
    Dim ilLoop As Integer
    Dim ilStatus As Integer
    Dim llPrevAttCode As Long
    Dim ilFound As Integer
    Dim llAstCode As Long
    Dim slSDate As String
    Dim slEDate As String
    Dim slMoDate As String
    Dim slSuDate As String
    Dim slELine As String
    Dim slFLine As String
    Dim slInDate As String
    Dim slInTime As String
    Dim ilPos As Integer
    Dim ilAdf As Integer
    Dim ilAdfCode As Integer
    Dim slZone As String
    Dim ilStation As Integer
    Dim ilSetAet As Integer
    Dim llAetCode As Long
    Dim ilAddBonus As Integer
    Dim slTime As String
    Dim ilSpotsAired As Integer
    'Dim slFields(1 To 16) As String
    Dim slFields(0 To 15) As String
    Dim ilAst As Integer
    Dim ilAnyAstExist As Integer
    Dim ilSchdCount As Integer
    Dim ilAiredCount As Integer
    Dim ilPledgeCompliantCount As Integer
    Dim ilAgyCompliantCount As Integer
    Dim ilAnyNotCompliant As Integer
    
    slFromFile = txtFile.Text
    'ilRet = 0
    'On Error GoTo mImportSpotsErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        Exit Function
    End If
    slSDate = ""
    ReDim tmAirSpotInfo(0 To 0) As AIRSPOTINFO
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mImportSpotsErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, False, slFields()
                For ilLoop = LBound(slFields) To UBound(slFields) Step 1
                    slFields(ilLoop) = Trim$(slFields(ilLoop))
                Next ilLoop
                'If slFields(1) = "S" Then
                If slFields(0) = "S" Then
                    On Error GoTo ErrHand
                    'llAstCode = Val(slFields(13))
                    llAstCode = Val(slFields(12))
                    '12/13/13: astPledgeDate is not used
                    'SQLQuery = "Select astShfCode, astAtfCode, astVefCode, astPledgeDate, astFeedDate, astStatus FROM ast WHERE (astCode =" & llAstCode & ")"
                    SQLQuery = "Select astShfCode, astAtfCode, astVefCode, astFeedDate, astStatus FROM ast WHERE (astCode =" & llAstCode & ")"
                    Set rst = gSQLSelectCall(SQLQuery)
                    'If (Not rst.EOF) And (Val(slFields(16)) <> 9) Then
                    If (Not rst.EOF) And (Val(slFields(15)) <> 9) Then
                        ilFound = False
                        For ilLoop = 0 To UBound(tmAirSpotInfo) - 1 Step 1
                            If (rst!astShfCode = tmAirSpotInfo(ilLoop).iShfCode) And (rst!astAtfCode = tmAirSpotInfo(ilLoop).lAtfCode) And (rst!astVefCode = tmAirSpotInfo(ilLoop).iVefCode) Then
                                ilFound = True
                                If DateValue(gAdjYear(rst!astFeedDate)) < DateValue(gAdjYear(tmAirSpotInfo(ilLoop).sStartDate)) Then
                                    tmAirSpotInfo(ilLoop).sStartDate = rst!astFeedDate
                                End If
                                If DateValue(gAdjYear(rst!astFeedDate)) > DateValue(gAdjYear(tmAirSpotInfo(ilLoop).sEndDate)) Then
                                    tmAirSpotInfo(ilLoop).sEndDate = rst!astFeedDate
                                End If
                                Exit For
                            End If
                        Next ilLoop
                        If Not ilFound Then
                            tmAirSpotInfo(UBound(tmAirSpotInfo)).iShfCode = rst!astShfCode
                            tmAirSpotInfo(UBound(tmAirSpotInfo)).lAtfCode = rst!astAtfCode
                            tmAirSpotInfo(UBound(tmAirSpotInfo)).iVefCode = rst!astVefCode
                            tmAirSpotInfo(UBound(tmAirSpotInfo)).sStartDate = rst!astFeedDate
                            tmAirSpotInfo(UBound(tmAirSpotInfo)).sEndDate = rst!astFeedDate
                            ReDim Preserve tmAirSpotInfo(0 To UBound(tmAirSpotInfo) + 1) As AIRSPOTINFO
                        End If
                        ilSetAet = True
                        'If Val(slFields(16)) = 0 Then
                        If Val(slFields(15)) = 0 Then
                            'update date/time aired
                            SQLQuery = "UPDATE ast SET "
                            SQLQuery = SQLQuery + "astCPStatus = 1" & ", " 'Received
                            If gGetAirStatus(rst!astStatus) <= 1 Then
                                SQLQuery = SQLQuery + "astStatus = " & rst!astStatus & ", "       'Aired
                            Else
                                SQLQuery = SQLQuery + "astStatus = " & 1 & ", "       'Delayed
                            End If
                            'slInDate = slFields(14)
                            slInDate = slFields(13)
                            'slInTime = slFields(15)
                            slInTime = slFields(14)
                            If (gIsDate(slInDate) = False) Or (Len(Trim$(slInDate)) = 0) Or (gIsTime(slInTime) = False) Or (Len(Trim$(slInTime)) = 0) Then
                                'Write error as date or time in error
                                gLogMsg "Invalid Aired Date or Time: Vehicle- " & slELine & " Call Letters- " & slFLine & " " & slLine & " line bypassed", "UnivisionImportLog.Txt", False
                                'Print #hmMsg, "Invalid Aired Date or Time: Vehicle- " & slELine & " Call Letters- " & slFLine & " " & slLine & " line bypassed"
                                lbcMsg.AddItem "Invalid Aired Date or Time: " & slELine & " " & slFLine & " " & slLine & " line bypassed"
                                ilSetAet = False
                            Else
                            If (InStr(1, slInTime, "PM", vbTextCompare) = 0) And (InStr(1, slInTime, "AM", vbTextCompare) = 0) Then
                                ilPos = InStr(1, slInTime, "N", vbTextCompare)
                                If ilPos > 0 Then
                                    slInTime = Left(slInTime, ilPos - 1) & "PM"
                                End If
                                ilPos = InStr(1, slInTime, "M", vbTextCompare)
                                If ilPos > 0 Then
                                    slInTime = Left(slInTime, ilPos - 1) & "AM"
                                End If
                                ilPos = InStr(1, slInTime, "P", vbTextCompare)
                                If ilPos > 0 Then
                                    slInTime = Left(slInTime, ilPos - 1) & "PM"
                                End If
                                ilPos = InStr(1, slInTime, "A", vbTextCompare)
                                If ilPos > 0 Then
                                    slInTime = Left(slInTime, ilPos - 1) & "AM"
                                End If
                            End If
                            SQLQuery = SQLQuery & "astAirDate = '" & Format$(slInDate, sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "astAirTime = '" & Format$(slInTime, sgSQLTimeForm) & "'"
                            SQLQuery = SQLQuery + " WHERE (astCode = " & llAstCode & ")"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "UnivisionImportLog.Txt", "ImportAiredSpot-mImportSpots"
                                cnn.RollbackTrans
                                mImportSpots = False
                                Exit Function
                            End If
                            cnn.CommitTrans
                            End If
'                        ElseIf Val(slFields(16)) = 9 Then
''                        ElseIf Val(slFields(15)) = 9 Then
'                            'Spot moved but cancel part aired, create bonus spot
'
'                            'Print #hmMsg, "Canceled Spot Aired: Vehicle- " & slELine & " Call Letters- " & slFLine & " " & slLine
'                            'lbcMsg.AddItem "Canceled Spot Aired: " & slELine & " " & slFLine & " " & slLine
                        'ElseIf (Val(slFields(16)) = 2) Or (Val(slFields(16)) = 4) Then
                        ElseIf (Val(slFields(15)) = 2) Or (Val(slFields(15)) = 4) Then
                            'update status as not aired
                            SQLQuery = "UPDATE ast SET "
                            SQLQuery = SQLQuery + "astCPStatus = 1" & ", " 'Received
                            SQLQuery = SQLQuery + "astStatus = 3"       'NA-Blackout
                            SQLQuery = SQLQuery + " WHERE (astCode = " & llAstCode & ")"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "UnivisionImportLog.Txt", "ImportAiredSpot-mImportSpots"
                                cnn.RollbackTrans
                                mImportSpots = False
                                Exit Function
                            End If
                            cnn.CommitTrans
                        Else
                            'update status as not aired
                            SQLQuery = "UPDATE ast SET "
                            SQLQuery = SQLQuery + "astCPStatus = 1" & ", " 'Received
                            SQLQuery = SQLQuery + "astStatus = 4"       'NA-Other
                            SQLQuery = SQLQuery + " WHERE (astCode = " & llAstCode & ")"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "UnivisionImportLog.Txt", "ImportAiredSpot-mImportSpots"
                                cnn.RollbackTrans
                                mImportSpots = False
                                Exit Function
                            End If
                            cnn.CommitTrans
                        End If
                        If ilSetAet Then
                            SQLQuery = "UPDATE aet SET "
                            SQLQuery = SQLQuery & "aetStatus = 'I'"
                            SQLQuery = SQLQuery & " WHERE (aetStatus <> 'D' And aetAstCode = " & llAstCode & ")"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "UnivisionImportLog.Txt", "ImportAiredSpot-mImportSpots"
                                cnn.RollbackTrans
                                mImportSpots = False
                                Exit Function
                            End If
                            cnn.CommitTrans
                        End If
                    Else
                        'If Val(slFields(16)) = 9 Then   'Deleted spot aired
                        If Val(slFields(15)) = 9 Then   'Deleted spot aired
                            'Create bonus spot
                            SQLQuery = "Select * FROM aet WHERE (aetStatus = 'D' and aetAstCode =" & llAstCode & ")"
                            Set rst = gSQLSelectCall(SQLQuery)
                            If Not rst.EOF Then
                                ilSetAet = True
                            Else
                                ilSetAet = False
                                'The Delete status did not exist, try to find aet that we can model from
                                'This would have happen for spots deleted before the status code set to 'D' 10/6/03
                                SQLQuery = "Select * FROM aet WHERE (aetStatus = 'I' and aetAstCode =" & llAstCode & ")"
                                Set rst = gSQLSelectCall(SQLQuery)
                                If Not rst.EOF Then
                                    'ilSetAet = True
                                Else
                                    'ilSetAet = False
                                    SQLQuery = "Select * FROM aet WHERE (aetAstCode =" & llAstCode & ")"
                                    Set rst = gSQLSelectCall(SQLQuery)
                                End If
                            End If
                            If Not rst.EOF Then
                                'slInDate = slFields(14)
                                slInDate = slFields(13)
                                'slInTime = slFields(15)
                                slInTime = slFields(14)
                                If (gIsDate(slInDate) = False) Or (Len(Trim$(slInDate)) = 0) Or (gIsTime(slInTime) = False) Or (Len(Trim$(slInTime)) = 0) Then
                                    'Write error as date or time in error
                                    gLogMsg "Invalid Aired Date or Time: Vehicle- " & slELine & " Call Letters- " & slFLine & " " & slLine & " line bypassed", "UnivisionImportLog.Txt", False
                                    'Print #hmMsg, "Invalid Aired Date or Time: Vehicle- " & slELine & " Call Letters- " & slFLine & " " & slLine & " line bypassed"
                                    lbcMsg.AddItem "Invalid Aired Date or Time: " & slELine & " " & slFLine & " " & slLine & " line bypassed"
                                    ilSetAet = False
                                Else
                                    If (InStr(1, slInTime, "PM", vbTextCompare) = 0) And (InStr(1, slInTime, "AM", vbTextCompare) = 0) Then
                                        ilPos = InStr(1, slInTime, "N", vbTextCompare)
                                        If ilPos > 0 Then
                                            slInTime = Left(slInTime, ilPos - 1) & "PM"
                                        End If
                                        ilPos = InStr(1, slInTime, "M", vbTextCompare)
                                        If ilPos > 0 Then
                                            slInTime = Left(slInTime, ilPos - 1) & "AM"
                                        End If
                                        ilPos = InStr(1, slInTime, "P", vbTextCompare)
                                        If ilPos > 0 Then
                                            slInTime = Left(slInTime, ilPos - 1) & "PM"
                                        End If
                                        ilPos = InStr(1, slInTime, "A", vbTextCompare)
                                        If ilPos > 0 Then
                                            slInTime = Left(slInTime, ilPos - 1) & "AM"
                                        End If
                                    End If
                                    ilAdfCode = 0
                                    For ilAdf = LBound(tgAdvtInfo) To UBound(tgAdvtInfo) - 1 Step 1
                                        If StrComp(Trim$(rst!aetAdvt), Trim$(tgAdvtInfo(ilAdf).sAdvtName), vbTextCompare) = 0 Then
                                            ilAdfCode = tgAdvtInfo(ilAdf).iCode
                                            Exit For
                                        End If
                                    Next ilAdf
                                    If ilAdfCode = 0 Then
                                        gLogMsg "Unable to process: Vehicle- " & slELine & " Call Letters- " & slFLine & " " & slLine & " as Advertiser missing", "UnivisionImportLog.Txt", False
                                        'Print #hmMsg, "Unable to process: Vehicle- " & slELine & " Call Letters- " & slFLine & " " & slLine & " as Advertiser missing"
                                        lbcMsg.AddItem "Unable to process: " & slELine & " " & slFLine & " " & slLine & " as Advertiser missing"
                                        ilSetAet = False
                                    Else
                                        slZone = ""
                                        For ilStation = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
                                            If rst!aetShfCode = tgStationInfo(ilStation).iCode Then
                                                slZone = tgStationInfo(ilStation).sZone
                                                Exit For
                                            End If
                                        Next ilStation
                                        llAetCode = rst!aetCode
                                        llAstCode = rst!aetAstCode
                                        'Test if previously added
                                        ilAddBonus = True
                                        SQLQuery = "Select astAirDate, astAirTime, astStatus FROM ast WHERE (astAtfCode =" & rst!aetAtfCode & " And astShfCode =" & rst!aetShfCode & " And astVefCode =" & rst!aetVefCode & " And astStatus = " & ASTEXTENDED_BONUS & ")"
                                        Set drst = gSQLSelectCall(SQLQuery)
                                        While Not drst.EOF
                                            'If aired date and time matched, then record previously processed
                                            If DateValue(gAdjYear(drst!astAirDate)) = DateValue(gAdjYear(slInDate)) Then
                                                slTime = Format$(drst!astAirTime, "h:mm:ssa/p")
                                                If TimeValue(slTime) = TimeValue(slInTime) Then
                                                    ilAddBonus = False
                                                    'Exit While
                                                End If
                                            End If
                                            drst.MoveNext
                                        Wend
                                        If ilAddBonus Then
                                            ilRet = gAddBonusSpot(rst!aetCntrNo, ilAdfCode, rst!aetVefCode, slInDate, slInTime, slZone, rst!aetAtfCode, rst!aetShfCode, rst!aetProd, rst!aetCart, rst!aetISCI, rst!aetLen)
                                        End If
                                    End If
                                End If
                                If ilSetAet Then
                                    SQLQuery = "UPDATE aet SET "
                                    SQLQuery = SQLQuery & "aetStatus = 'I'"
                                    SQLQuery = SQLQuery & " WHERE (aetCode = " & llAetCode & ")"
                                    cnn.BeginTrans
                                    'cnn.Execute SQLQuery, rdExecDirect
                                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                        '6/10/16: Replaced GoSub
                                        'GoSub ErrHand:
                                        Screen.MousePointer = vbDefault
                                        gHandleError "UnivisionImportLog.Txt", "ImportAiredSpot-mImportSpots"
                                        cnn.RollbackTrans
                                        mImportSpots = False
                                        Exit Function
                                    End If
                                    cnn.CommitTrans
                                End If
                            Else
                        'Write error as AST missing
                                gLogMsg "Unable to process: Vehicle- " & slELine & " Call Letters- " & slFLine & " " & slLine & " probable previously import 28 days ago", "UnivisionImportLog.Txt", False
                                'Print #hmMsg, "Unable to process: Vehicle- " & slELine & " Call Letters- " & slFLine & " " & slLine & " probable previously import 28 days ago"
                                lbcMsg.AddItem "Unable to process: " & slELine & " " & slFLine & " " & slLine & " probable previously import 28 days ago"
                            End If
                        Else
                            'Write error as AST missing
                            gLogMsg "Unable to process: Vehicle- " & slELine & " Call Letters- " & slFLine & " " & slLine & " as AST missing", "UnivisionImportLog.Txt", False
                            'Print #hmMsg, "Unable to process: Vehicle- " & slELine & " Call Letters- " & slFLine & " " & slLine & " as AST missing"
                            lbcMsg.AddItem "Unable to process: " & slELine & " " & slFLine & " " & slLine & " as AST missing"
                        End If
                    End If
                'ElseIf slFields(1) = "E" Then
                ElseIf slFields(0) = "E" Then
                    'slELine = slFields(2)
                    slELine = slFields(1)
                'ElseIf slFields(1) = "F" Then
                ElseIf slFields(0) = "F" Then
                    'slFLine = slFields(2)
                    slFLine = slFields(1)
                End If
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    'Set any Not Aired to received as they are not exported
    llPrevAttCode = -1
    For ilLoop = 0 To UBound(tmAirSpotInfo) - 1 Step 1
        If llPrevAttCode <> tmAirSpotInfo(ilLoop).lAtfCode Then
            slSDate = tmAirSpotInfo(ilLoop).sStartDate
            slMoDate = gAdjYear(gObtainPrevMonday(slSDate))
            Do
                slSuDate = DateAdd("d", 6, slMoDate)
                For ilStatus = 0 To UBound(tgStatusTypes) Step 1
                    If (tgStatusTypes(ilStatus).iPledged = 2) Then
                        SQLQuery = "UPDATE ast SET "
                        SQLQuery = SQLQuery + "astCPStatus = " & "1"    'Received
                        SQLQuery = SQLQuery + " WHERE (astAtfCode = " & tmAirSpotInfo(ilLoop).lAtfCode
                        SQLQuery = SQLQuery + " AND astCPStatus = 0"
                        SQLQuery = SQLQuery + " AND astStatus = " & tgStatusTypes(ilStatus).iStatus
                        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')" & ")"
                        cnn.BeginTrans
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "UnivisionImportLog.Txt", "ImportAiredSpot-mImportSpots"
                            cnn.RollbackTrans
                            mImportSpots = False
                            Exit Function
                        End If
                        cnn.CommitTrans
                    End If
                Next ilStatus
                slMoDate = DateAdd("d", 7, slMoDate)
            Loop While DateValue(gAdjYear(slMoDate)) < DateValue(gAdjYear(tmAirSpotInfo(ilLoop).sEndDate))
        End If
        llPrevAttCode = tmAirSpotInfo(ilLoop).lAtfCode
    Next ilLoop
    

    'Determine if CPTTStatus should to set to 0=Partial or 1=Completed
    llPrevAttCode = -1
    For ilLoop = 0 To UBound(tmAirSpotInfo) - 1 Step 1
        If llPrevAttCode <> tmAirSpotInfo(ilLoop).lAtfCode Then
            slSDate = tmAirSpotInfo(ilLoop).sStartDate
            slMoDate = gAdjYear(gObtainPrevMonday(slSDate))
            Do
                slSuDate = DateAdd("d", 6, slMoDate)
                'Test to see if any spots aired or were they all not aired
                ilSpotsAired = gDidAnySpotsAir(tmAirSpotInfo(ilLoop).lAtfCode, slMoDate, slSuDate)
                If ilSpotsAired Then
                    'We know at least one spot aired
                   ilSpotsAired = True
                Else
                    'no aired spots were found
                    ilSpotsAired = False
                End If
                
                'Check for any spots that have not aired - astCPStatus = 0 = not aired
                SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 0"
                SQLQuery = SQLQuery + " AND astAtfCode = " & tmAirSpotInfo(ilLoop).lAtfCode
                'SQLQuery = SQLQuery + " AND astShfCode = " & tmAirSpotInfo(ilLoop).iShfCode
                'SQLQuery = SQLQuery + " AND astVefCode = " & tmAirSpotInfo(ilLoop).iVefCode
                SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                Set rst = gSQLSelectCall(SQLQuery)
                If rst.EOF Then
                    'Set CPTT as complete
                    SQLQuery = "UPDATE cptt SET "
                    If ilSpotsAired Then
                        SQLQuery = SQLQuery + "cpttStatus = 1" & ", " 'Complete spots aired
                    Else
                        SQLQuery = SQLQuery + "cpttStatus = 2" & ", " 'Complete NO spots aired
                    End If
                    SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery + "cpttPostingStatus = 2"  'Complete
                    SQLQuery = SQLQuery + " WHERE cpttAtfCode = " & tmAirSpotInfo(ilLoop).lAtfCode
                    'SQLQuery = SQLQuery + " AND cpttShfCode = " & tmAirSpotInfo(ilLoop).iShfCode
                    'SQLQuery = SQLQuery + " AND cpttVefCode = " & tmAirSpotInfo(ilLoop).iVefCode
                    SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                    cnn.BeginTrans
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "UnivisionImportLog.Txt", "ImportAiredSpot-mImportSpots"
                        cnn.RollbackTrans
                        mImportSpots = False
                        Exit Function
                    End If
                    cnn.CommitTrans
                Else
                    'Set CPTT as Partial
                    SQLQuery = "UPDATE cptt SET "
                    SQLQuery = SQLQuery + "cpttStatus = 0" & ", " 'Partial
                    SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery + "cpttPostingStatus = 1" 'Partial
                    SQLQuery = SQLQuery + " WHERE cpttAtfCode = " & tmAirSpotInfo(ilLoop).lAtfCode
                    'SQLQuery = SQLQuery + " AND cpttShfCode = " & tmAirSpotInfo(ilLoop).iShfCode
                    'SQLQuery = SQLQuery + " AND cpttVefCode = " & tmAirSpotInfo(ilLoop).iVefCode
                    SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                    cnn.BeginTrans
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "UnivisionImportLog.Txt", "ImportAiredSpot-mImportSpots"
                        cnn.RollbackTrans
                        mImportSpots = False
                        Exit Function
                    End If
                    cnn.CommitTrans
                End If
                'D.S. 02/25/11 Start new compliant code
                SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attCode, attTimeType, attGenCP, attStartTime, attLogType, attPostType, shttWebEmail, shttWebPW, attWebEmail, attSendLogEmail, attWebPW, attAgreeStart, attAgreeEnd, attDropDate, attOnAir, attOffAir, attMulticast, attWebInterface"
                SQLQuery = SQLQuery & " FROM shtt, cptt, att"
                SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
                SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
                SQLQuery = SQLQuery & " AND attExportType = 2"
                'D.S. 11/27/13
                'SQLQuery = SQLQuery & " AND cpttVefCode = " & ilVefCode
                SQLQuery = SQLQuery & " AND cpttAtfCode = " & tmAirSpotInfo(ilLoop).lAtfCode
                SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slMoDate, sgSQLDateForm) & "')"
                SQLQuery = SQLQuery & " Order by shttcallletters"
                Set cprst = gSQLSelectCall(SQLQuery)
                If Not cprst.EOF Then
                    'Create AST records - gGetAstInfo requires tgCPPosting to be initialized
                    ReDim tgCPPosting(0 To 1) As CPPOSTING
                    tgCPPosting(0).lCpttCode = cprst!cpttCode
                    tgCPPosting(0).iStatus = cprst!cpttStatus
                    tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                    tgCPPosting(0).lAttCode = cprst!cpttatfCode
                    tgCPPosting(0).iAttTimeType = cprst!attTimeType
                    tgCPPosting(0).iVefCode = rst!astVefCode
                    tgCPPosting(0).iShttCode = cprst!shttCode
                    tgCPPosting(0).sZone = cprst!shttTimeZone
                    tgCPPosting(0).sDate = Format$(slMoDate, sgShowDateForm)
                    tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                    ilSchdCount = 0
                    ilAiredCount = 0
                    ilPledgeCompliantCount = 0
                    ilAgyCompliantCount = 0
                    igTimes = 1 'By Week
                    ilAdfCode = -1
                    'Dan M 9/26/13 6442
                    ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, True, False, True)
                    'ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, False, False, True)
                    For ilAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                        ilAnyAstExist = True
                        'gIncSpotCounts tmAstInfo(ilAst).iPledgeStatus, tmAstInfo(ilAst).iStatus, tmAstInfo(ilAst).iCPStatus, tmAstInfo(ilAst).sTruePledgeDays, Format$(tmAstInfo(ilAst).sPledgeDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sAirDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sPledgeStartTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sTruePledgeEndTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
                        gIncSpotCounts tmAstInfo(ilAst), ilSchdCount, ilAiredCount, ilPledgeCompliantCount, ilAgyCompliantCount
                    Next ilAst
                    If ilAiredCount <> ilPledgeCompliantCount Then
                        ilAnyNotCompliant = True
                    End If
                    SQLQuery = "Update cptt Set "
                    SQLQuery = SQLQuery & "cpttNoSpotsGen = " & ilSchdCount & ", "
                    SQLQuery = SQLQuery & "cpttNoSpotsAired = " & ilAiredCount & ", "
                    SQLQuery = SQLQuery & "cpttNoCompliant = " & ilPledgeCompliantCount & ", "
                    SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery & "cpttAgyCompliant = " & ilAgyCompliantCount & " "
                    SQLQuery = SQLQuery & " Where cpttCode = " & cprst!cpttCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "UnivisionImportLog.Txt", "ImportAiredSpot-mImportSpots"
                        mImportSpots = False
                        Exit Function
                    End If
                End If
                'D.S. 02/25/11 End new code
                slMoDate = DateAdd("d", 7, slMoDate)
            Loop While DateValue(gAdjYear(slMoDate)) < DateValue(gAdjYear(tmAirSpotInfo(ilLoop).sEndDate))
        End If
        llPrevAttCode = tmAirSpotInfo(ilLoop).lAtfCode
    Next ilLoop
    gFileChgdUpdate "cptt.mkd", True
    mImportSpots = True
    Exit Function
mImportSpotsErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportAiredSpot-mImportSpots"
    mImportSpots = False
    Exit Function

End Function


Private Function mCheckFile()
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim ilLoop As Integer
    Dim ilSFound As Integer
    'Dim slFields(1 To 16) As String
    Dim slFields(0 To 15) As String
    
    slFromFile = txtFile.Text
    'ilRet = 0
    'On Error GoTo mImportSpotsErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Beep
        gMsgBox "Unable to open the Import file error: " & Trim$(Str$(ilRet)), vbCritical
        mCheckFile = False
        Close hmFrom
        Exit Function
    End If
    mCheckFile = True
    ilSFound = False
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mImportSpotsErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, False, slFields()
                For ilLoop = LBound(slFields) To UBound(slFields) Step 1
                    slFields(ilLoop) = Trim$(slFields(ilLoop))
                Next ilLoop
                'If slFields(1) = "S" Then
                If slFields(0) = "S" Then
                    ilSFound = True
                    'If slFields(16) = "" Then
                    If slFields(15) = "" Then
                        Beep
                        gMsgBox "Import file missing Status Column on Spot record", vbCritical
                        Close hmFrom
                        mCheckFile = False
                        Exit Function
                    End If
                End If
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    If Not ilSFound Then
        Beep
        gMsgBox "No Spot records found in the Import file", vbCritical
        mCheckFile = False
    End If
    Exit Function
mImportSpotsErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportAiredSpot-mCheckFile"
End Function
