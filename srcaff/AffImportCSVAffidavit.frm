VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportCSVAffidavit 
   Caption         =   "Import CSV Affidavit"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   Icon            =   "AffImportCSVAffidavit.frx":0000
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
      ItemData        =   "AffImportCSVAffidavit.frx":08CA
      Left            =   120
      List            =   "AffImportCSVAffidavit.frx":08CC
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
Attribute VB_Name = "frmImportCSVAffidavit"
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
Private hmFrom As Integer
Private hmCSV As Integer
Private smLogPathFileName As String
Private cprst As ADODB.Recordset
Private drst As ADODB.Recordset
Private tmCPDat() As DAT
Private Type CPTTCSVINFO
    lAttCode As Long
    lShttCode As Long
    iVefCode As Integer
    sDate As String * 10
End Type

Private tmCpttCSVInfo() As CPTTCSVINFO
Private hmAst As Integer
Private tmAstInfo() As ASTINFO
Const STATIONINDEX = 0
Const VEHICLEINDEX = 1
Const DATEAIREDINDEX = 2
Const LENGTHINDEX = 3
Const TIMEAIREDINDEX = 4
Const ISCIINDEX = 5









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
    CommonDialog1.FilterIndex = 2
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
    smLogPathFileName = "CSVAffidavitImportLog.Txt"
    gLogMsgWODT "ON", hmCSV, sgMsgDirectory & smLogPathFileName
    gLogMsgWODT "W", hmCSV, "Started Import CSV Affidavit Spots " & Now & " **"

    iRet = mImportSpots()
    If (iRet = False) Then
        gLogMsgWODT "W", hmCSV, "** Terminated - mImportSpots returned False " & Now & " **"
        gLogMsgWODT "C", hmCSV, ""
        imExporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    If imTerminate Then
        gLogMsgWODT "W", hmCSV, "** User Terminated " & Now & " **"
        gLogMsgWODT "C", hmCSV, ""
        imExporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.SetFocus
        Exit Sub
    End If
    gLogMsgWODT "W", hmCSV, "** Completed Import CSV Affidavit Spots " & Now & " **"
    gLogMsgWODT "C", hmCSV, ""
    On Error GoTo ErrHand:
    mShowReport
    imExporting = False
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
    Unload frmImportCSVAffidavit
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
    frmImportCSVAffidavit.Caption = "Aired Station Spots - " & sgClientName
    imAllClick = False
    imTerminate = False
    imExporting = False
    gOpenMKDFile hmAst, "Ast.Mkd"
    
    'txtFile.Text = sgImportDirectory & "CSISpots.txt"
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gCloseMKDFile hmAst, "Ast.Mkd"
    Erase tmCpttCSVInfo
    Erase tmCPDat
    Erase tmAstInfo
    cprst.Close
    drst.Close
    Set frmImportCSVAffidavit = Nothing
End Sub

Private Function mImportSpots() As Integer
        
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim ilLoop As Integer
    Dim llShttCode As Long
    Dim ilVefCode As Integer
    Dim llPrevShttCode As Long
    Dim ilPrevVefCode As Integer
    Dim llPrevMoDate As Long
    Dim slFields(0 To 15) As String
    Dim blHeaderFd As Boolean
    Dim slStation As String
    Dim slVehicle As String
    Dim slDateAired As String
    Dim slLength As String
    Dim slTimeAired As String
    Dim slISCI As String
    Dim slMoDate As String
    Dim slSDate As String
    Dim slSuDate As String
    Dim ilFound As Integer
    
    Dim ilSpotsAired As Integer
    Dim ilAnyAstExist As Integer
    Dim ilSchdCount As Integer
    Dim ilAiredCount As Integer
    Dim ilPledgeCompliantCount As Integer
    Dim ilAgyCompliantCount As Integer
    Dim ilAnyNotCompliant As Integer
    Dim ilAdfCode As Integer
    Dim ilAst As Integer
    Dim blPostingCompleted As Boolean
    Dim llNoMatch As Long
    Dim llWebPosted As Long
    Dim llMatches As Long
    Dim llNoLines As Long
    
    slFromFile = txtFile.Text
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        Exit Function
    End If
    blHeaderFd = False
    llPrevShttCode = 0
    ilPrevVefCode = 0
    llPrevMoDate = 0
    slSDate = ""
    blPostingCompleted = False
    llNoMatch = 0
    llWebPosted = 0
    llMatches = 0
    llNoLines = 0
    ReDim tmAstInfo(0 To 0) As ASTINFO
    ReDim tmCpttCSVInfo(0 To 0) As CPTTCSVINFO
    Do While Not EOF(hmFrom)
        If imTerminate Then
            mImportSpots = False
            Exit Function
        End If
            
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
                If Not blHeaderFd Then
                    If (slFields(STATIONINDEX) = "STATION") And (slFields(DATEAIREDINDEX) = "DATE AIRED") And (slFields(TIMEAIREDINDEX) = "TIME AIRED") And (slFields(ISCIINDEX) = "ISCI") Then
                        blHeaderFd = True
                    End If
                Else
                    If slFields(0) <> "" Then
                        llNoLines = llNoLines + 1
                        On Error GoTo ErrHand
                        slStation = slFields(STATIONINDEX)
                        slVehicle = UCase$(slFields(VEHICLEINDEX))
                        slDateAired = slFields(DATEAIREDINDEX)
                        slMoDate = gObtainPrevMonday(slDateAired)
                        slLength = slFields(LENGTHINDEX)
                        slTimeAired = slFields(TIMEAIREDINDEX)
                        slISCI = slFields(ISCIINDEX)
                        llShttCode = gBinarySearchStation(Trim$(slStation))
                        If llShttCode <> -1 Then
                            llShttCode = tgStationInfo(llShttCode).iCode
                        End If
                        ilVefCode = -1
                        For ilLoop = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                            If Trim$(tgVehicleInfo(ilLoop).sVehicle) = slVehicle Then
                                ilVefCode = tgVehicleInfo(ilLoop).iCode
                                Exit For
                            End If
                        Next ilLoop
                        If (llShttCode = -1) Or (ilVefCode = -1) Then
                            If (llShttCode = -1) And (ilVefCode = -1) Then
                                gLogMsgWODT "W", hmCSV, "Station and Vehicle not found, line bypassed- " & slLine
                            ElseIf (llShttCode = -1) Then
                                gLogMsgWODT "W", hmCSV, "Station not found, line bypassed- " & slLine
                            Else
                                gLogMsgWODT "W", hmCSV, "Vehicle not found, line bypassed- " & slLine
                            End If
                        Else
                            If (llShttCode <> llPrevShttCode) Or (ilVefCode <> ilPrevVefCode) Or (gDateValue(slMoDate) <> llPrevMoDate) Then
                                'Set any spot not matched as cdStatus = 0
                                For ilLoop = 0 To UBound(tmAstInfo) - 1 Step 1
                                    If (tmAstInfo(ilLoop).iCPStatus = -1) Then
                                        tmAstInfo(ilLoop).iCPStatus = 0
                                        If Trim$(tmAstInfo(ilLoop).sAffidavitSource) = "" Then
                                            SQLQuery = "UPDATE ast SET "
                                            SQLQuery = SQLQuery + "astCPStatus = 0"
                                            SQLQuery = SQLQuery + " WHERE (astCode = " & tmAstInfo(ilLoop).lCode & ")"
                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                Screen.MousePointer = vbDefault
                                                gHandleError "CSVAffidavitImportError.Txt", "ImportCSVAffidavit-mImportSpots"
                                                mImportSpots = False
                                                Exit Function
                                            End If
                                        End If
                                   End If
                                Next ilLoop
                                blPostingCompleted = False
                                llPrevShttCode = llShttCode
                                ilPrevVefCode = ilVefCode
                                llPrevMoDate = gDateValue(slMoDate)
                                SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attCode, attTimeType, attGenCP, attStartTime, attLogType, attPostType, shttWebEmail, shttWebPW, attWebEmail, attSendLogEmail, attWebPW, attAgreeStart, attAgreeEnd, attDropDate, attOnAir, attOffAir, attMulticast, attWebInterface"
                                SQLQuery = SQLQuery & " FROM shtt, cptt, att"
                                SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
                                SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
                                'SQLQuery = SQLQuery & " AND attExportType = 2"
                                SQLQuery = SQLQuery & " AND cpttVefCode = " & ilVefCode
                                SQLQuery = SQLQuery & " AND cpttShfCode = " & llShttCode
                                SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slMoDate, sgSQLDateForm) & "')"
                                SQLQuery = SQLQuery & " Order by shttcallletters"
                                Set cprst = gSQLSelectCall(SQLQuery)
                                If Not cprst.EOF Then
                                    If cprst!cpttPostingStatus <> 2 Then
                                        'Create AST records - gGetAstInfo requires tgCPPosting to be initialized
                                        tmCpttCSVInfo(UBound(tmCpttCSVInfo)).lAttCode = cprst!cpttatfCode
                                        tmCpttCSVInfo(UBound(tmCpttCSVInfo)).sDate = slMoDate
                                        tmCpttCSVInfo(UBound(tmCpttCSVInfo)).lShttCode = llShttCode
                                        tmCpttCSVInfo(UBound(tmCpttCSVInfo)).iVefCode = ilVefCode
                                        ReDim Preserve tmCpttCSVInfo(0 To UBound(tmCpttCSVInfo) + 1) As CPTTCSVINFO
                                        ReDim tgCPPosting(0 To 1) As CPPOSTING
                                        tgCPPosting(0).lCpttCode = cprst!cpttCode
                                        tgCPPosting(0).iStatus = cprst!cpttStatus
                                        tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                                        tgCPPosting(0).lAttCode = cprst!cpttatfCode
                                        tgCPPosting(0).iAttTimeType = cprst!attTimeType
                                        tgCPPosting(0).iVefCode = ilVefCode
                                        tgCPPosting(0).iShttCode = cprst!shttCode
                                        tgCPPosting(0).sZone = cprst!shttTimeZone
                                        tgCPPosting(0).sDate = Format$(slMoDate, sgShowDateForm)
                                        tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                                        igTimes = 1 'By Week
                                        ilAdfCode = -1
                                        ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, True, False, True)
                                        For ilLoop = 0 To UBound(tmAstInfo) - 1 Step 1
                                            tmAstInfo(ilLoop).iCPStatus = -1
                                        Next ilLoop
                                    Else
                                        blPostingCompleted = True
                                        ReDim tmAstInfo(0 To 0) As ASTINFO
                                    End If
                                Else
                                    ReDim tmAstInfo(0 To 0) As ASTINFO
                                End If
                            End If
                            If blPostingCompleted Then
                                ilFound = True
                                llWebPosted = llWebPosted + 1
                            Else
                                ilFound = False
                            End If
                            For ilLoop = 0 To UBound(tmAstInfo) - 1 Step 1
                                If (tmAstInfo(ilLoop).iCPStatus = -1) And (gDateValue(tmAstInfo(ilLoop).sAirDate) = gDateValue(slDateAired)) Then
                                    If tmAstInfo(ilLoop).iRegionType = 0 Then
                                        If slISCI = Trim$(tmAstInfo(ilLoop).sISCI) Then
                                            ilFound = True
                                        End If
                                    Else
                                        If slISCI = Trim$(tmAstInfo(ilLoop).sRISCI) Then
                                            ilFound = True
                                        End If
                                    End If
                                    If ilFound Then
                                        tmAstInfo(ilLoop).iCPStatus = 1
                                        If Trim$(tmAstInfo(ilLoop).sAffidavitSource) = "" Then  'Not posted
                                            llMatches = llMatches + 1
                                            tmAstInfo(ilLoop).sAirDate = slDateAired
                                            tmAstInfo(ilLoop).sAirTime = slTimeAired
                                            SQLQuery = "UPDATE ast SET "
                                            SQLQuery = SQLQuery + "astCPStatus = 1" & ", " 'Received
                                            SQLQuery = SQLQuery & "astAirDate = '" & Format$(slDateAired, sgSQLDateForm) & "', "
                                            SQLQuery = SQLQuery & "astAirTime = '" & Format$(slTimeAired, sgSQLTimeForm) & "'"
                                            SQLQuery = SQLQuery + " WHERE (astCode = " & tmAstInfo(ilLoop).lCode & ")"
                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                Screen.MousePointer = vbDefault
                                                gHandleError "CSVAffidavitImportError.Txt", "ImportCSVAffidavit-mImportSpots"
                                                mImportSpots = False
                                                Exit Function
                                            End If
                                        Else
                                            llWebPosted = llWebPosted + 1
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                llNoMatch = llNoMatch + 1
                                gLogMsgWODT "W", hmCSV, "Matching Spot not found, line bypassed- " & slLine
                            End If
                        End If
                    End If
                End If
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    gLogMsgWODT "W", hmCSV, "Matches Found Count = " & llMatches
    gLogMsgWODT "W", hmCSV, "No Match Found Count = " & llNoMatch
    gLogMsgWODT "W", hmCSV, "Web Posted/Bypassed Count = " & llWebPosted
    gLogMsgWODT "W", hmCSV, "Totals Lines Processed = " & llNoLines

    'Determine if CPTTStatus should to set to 0=Partial or 1=Completed
    For ilLoop = 0 To UBound(tmCpttCSVInfo) - 1 Step 1
        If imTerminate Then
            mImportSpots = False
            Exit Function
        End If
        slSDate = tmCpttCSVInfo(ilLoop).sDate
        slMoDate = gAdjYear(gObtainPrevMonday(slSDate))
        slSuDate = DateAdd("d", 6, slMoDate)
        'Test to see if any spots aired or were they all not aired
        ilSpotsAired = gDidAnySpotsAir(tmCpttCSVInfo(ilLoop).lAttCode, slMoDate, slSuDate)
        If ilSpotsAired Then
            'We know at least one spot aired
           ilSpotsAired = True
        Else
            'no aired spots were found
            ilSpotsAired = False
        End If
        
        'Check for any spots that have not aired - astCPStatus = 0 = not aired
        SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 0"
        SQLQuery = SQLQuery + " AND astAtfCode = " & tmCpttCSVInfo(ilLoop).lAttCode
        'SQLQuery = SQLQuery + " AND astShfCode = " & tmCpttCSVInfo(ilLoop).iShfCode
        'SQLQuery = SQLQuery + " AND astVefCode = " & tmCpttCSVInfo(ilLoop).iVefCode
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
            SQLQuery = SQLQuery + " WHERE cpttAtfCode = " & tmCpttCSVInfo(ilLoop).lAttCode
            SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "CSVAffidavitImportError.Txt", "ImportCSVAffidavit-mImportSpots"
                mImportSpots = False
                Exit Function
            End If
        Else
            'Set CPTT as Partial
            SQLQuery = "UPDATE cptt SET "
            SQLQuery = SQLQuery + "cpttStatus = 0" & ", " 'Partial
            SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
            SQLQuery = SQLQuery + "cpttPostingStatus = 1" 'Partial
            SQLQuery = SQLQuery + " WHERE cpttAtfCode = " & tmCpttCSVInfo(ilLoop).lAttCode
            SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "CSVAffidavitImportError.Txt", "ImportCSVAffidavit-mImportSpots"
                mImportSpots = False
                Exit Function
            End If
        End If
        'D.S. 02/25/11 Start new compliant code
        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attCode, attTimeType, attGenCP, attStartTime, attLogType, attPostType, shttWebEmail, shttWebPW, attWebEmail, attSendLogEmail, attWebPW, attAgreeStart, attAgreeEnd, attDropDate, attOnAir, attOffAir, attMulticast, attWebInterface"
        SQLQuery = SQLQuery & " FROM shtt, cptt, att"
        SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
        SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
        SQLQuery = SQLQuery & " AND attExportType = 2"
        'D.S. 11/27/13
        'SQLQuery = SQLQuery & " AND cpttVefCode = " & ilVefCode
        SQLQuery = SQLQuery & " AND cpttAtfCode = " & tmCpttCSVInfo(ilLoop).lAttCode
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
                gHandleError "CSVAffidavitImportError.Txt", "ImportCSVAffidavit-mImportSpots"
                mImportSpots = False
                Exit Function
            End If
        End If
    Next ilLoop
    gFileChgdUpdate "cptt.mkd", True
    mImportSpots = True
    Exit Function
mImportSpotsErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportCSVAffidavit-mImportSpots"
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
                If (slFields(STATIONINDEX) = "STATION") And (slFields(DATEAIREDINDEX) = "DATE AIRED") And (slFields(TIMEAIREDINDEX) = "TIME AIRED") And (slFields(ISCIINDEX) = "ISCI") Then
                    ilSFound = True
                    Exit Do
                End If
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    If Not ilSFound Then
        Beep
        gMsgBox "Header record not found in the Import file", vbCritical
        mCheckFile = False
    End If
    Exit Function
mImportSpotsErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportCSVAffidavit-mCheckFile"
End Function

Private Sub mShowReport()
    Dim slCmd As String
    Dim slDateTime As String
    Dim ilRet As Integer
    
    'On Error GoTo ErrHandler
    ilRet = 0
    'slDateTime = FileDateTime(smReportPathFileName)
    ilRet = gFileExist(sgMsgDirectory & smLogPathFileName)
    If ilRet <> 0 Then
        Exit Sub
    End If
    ilRet = MsgBox("View Result File?", vbApplicationModal + vbInformation + vbYesNo, "Question")
    If ilRet = vbNo Then
        Exit Sub
    End If
    slCmd = "Notepad.exe " & sgMsgDirectory & smLogPathFileName
    Call Shell(slCmd, vbNormalFocus)
    Exit Sub
    
'ErrHandler:
'    ilRet = -1
'    Resume Next
End Sub
