VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportCSV 
   Caption         =   "Import Station Information"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   Icon            =   "AffImptCSV.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmcSplit 
      Caption         =   "Compress Import File"
      Height          =   315
      Left            =   1695
      TabIndex        =   12
      Top             =   3585
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1665
      TabIndex        =   4
      Top             =   750
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   3315
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.ListBox lbcNames 
      Height          =   2010
      ItemData        =   "AffImptCSV.frx":08CA
      Left            =   255
      List            =   "AffImptCSV.frx":08D1
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.ListBox lbcError 
      Height          =   2010
      ItemData        =   "AffImptCSV.frx":08DF
      Left            =   240
      List            =   "AffImptCSV.frx":08E6
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   7245
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6225
      TabIndex        =   8
      Top             =   3585
      Width           =   1245
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "Browse"
      Height          =   330
      Left            =   6075
      TabIndex        =   2
      Top             =   360
      Width           =   1395
   End
   Begin VB.CommandButton cmcImport 
      Caption         =   "Import"
      Height          =   315
      Left            =   4650
      TabIndex        =   7
      Top             =   3585
      Width           =   1245
   End
   Begin VB.TextBox txtFile 
      Height          =   300
      Left            =   945
      TabIndex        =   1
      Top             =   375
      Width           =   4770
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7530
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lacDate 
      Caption         =   "Last Date to Import"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lbcPercent 
      Height          =   210
      Left            =   225
      TabIndex        =   11
      Top             =   3630
      Width           =   1605
   End
   Begin VB.Label lbcMsg 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   1065
      TabIndex        =   10
      Top             =   3300
      Width           =   5625
   End
   Begin VB.Label lbcFile 
      Caption         =   "Import File"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   390
      Width           =   780
   End
End
Attribute VB_Name = "frmImportCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmImportCSV - displays import csv information
'*
'*  Created Aug,1998 by Dick LeVine
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imTerminate As Integer
Private smVehName As String
Private imAllClick As Integer
Private imImporting As Integer
Private hmFrom As Integer
Private hmMsg As Integer
Private smMsgFile As String
Private lmTotalNoBytes As Long
Private lmProcessedNoBytes As Long
Private lmStartDate As Long
Private lmLastDate As Long
Private smCurDir As String
Private lmFloodPercent As Long
Private hmSplit As Integer
Private smSplitFileName As String
'Private smFields(1 To 60) As String
Private smFields(0 To 59) As String
Private smMissingVef() As String
Private smMissingShtt() As String
Private smStationNotMatching() As String
Private smVehicleNotMatching() As String
Private smStationNameError() As String
Private smVehicleNameError() As String
Private smZoneError() As String
Private imVefCombo As Integer
Private imVefCode As Integer
Private tmMissingAtt() As MISSINGATT
Private tmPledgeCount() As MISSINGATT
Private tmAgreeID() As AGREEID
Private tmLstMYLInfo() As LSTMYLINFO
Private smMissingTime() As String
Private tmPledgeInfo() As PledgeInfo
Private tmLogSpotInfo() As LOGSPOTINFO
'array of 2-char state codes with its time zone :0 = est, 1 = cst , 2 = mst, 3 = pst.   Washington DC and Puerto Rico at end (both assumed to be EST)
Private Const smZone As String * 156 = "AL1AK3AZ2AR1CA3C02CT0DE0FL1GA0HI3ID3IL1IN0IA1KS1KY0LA1ME0MD0MA0MI0MN1MS1MO1MT2NE1NV3NH0NJ0NM2NY0NC0ND1OH0OK1OR3PA0RI0SC0SD1TN0TX1UT2VT0VA0WA3WV0WI1WY2PR0DC0"
'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenSplitFile                  *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenSplitFile(slDate As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slInDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim ilRet As Integer
    Dim slLetter As String

    'On Error GoTo mOpenSplitFileErr:
    If Len(smSplitFileName) = 0 Then
        slInDate = slDate
        'slInDate = gObtainPrevMonday(slInDate)
        gObtainYearMonthDayStr slInDate, True, slYear, slMonth, slDay
        slToFile = "S" & right$(slYear, 2) & slMonth & slDay & "A" & ".csv"
        'slToFile = "PS" & Right$(slYear, 2) & slMonth & slDay & ".csv"
    Else
        slLetter = Chr$(Asc(Mid$(smSplitFileName, 8, 1)) + 1)
        slToFile = Left$(smSplitFileName, 7) & slLetter & ".csv"
    End If
    smSplitFileName = slToFile
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        Kill slToFile
        On Error GoTo 0
        'ilRet = 0
        'On Error GoTo mOpenSplitFileErr:
        'hmSplit = FreeFile
        'Open slToFile For Output As hmSplit
        ilRet = gFileOpen(slToFile, "Output", hmSplit)
        If ilRet <> 0 Then
            Close hmSplit
            hmSplit = -1
            gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenSplitFile = False
            Exit Function
        End If
    Else
        On Error GoTo 0
        'ilRet = 0
        'On Error GoTo mOpenSplitFileErr:
        'hmSplit = FreeFile
        'Open slToFile For Output As hmSplit
        ilRet = gFileOpen(slToFile, "Output", hmSplit)
        If ilRet <> 0 Then
            Close hmSplit
            hmSplit = -1
            gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenSplitFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, "** Split File Name: " & slToFile & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    mOpenSplitFile = True
    Exit Function
'mOpenSplitFileErr:
'    ilRet = 1
'    Resume Next
End Function



'*******************************************************
'*                                                     *
'*      Procedure Name:mReadAndSplit                   *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File and split into two   *
'*                                                     *
'*******************************************************
Private Function mReadAndSplit(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilLoop As Integer
    Dim ilStatus As Integer
    Dim ilPos As Integer
    Dim slMsg As String
    Dim llPercent As Long
    Dim ilUpper As Integer
    Dim llTstDate As Long
    Dim slCntrNo As String
    Dim llCntrNo As Long
    Dim slAirDate As String
    Dim slAirTime As String
    Dim llSpotCount As Long
    Dim slChar As String
    Dim slStr As String
    Dim ilAirDateMsg As Integer
    Dim slFdDate As String
    Dim slFdTime As String
    Dim slPdDate As String
    Dim slPdSTime As String
    Dim slPdETime As String
    Dim slCallLetters As String
    Dim llNextPercent As Long
    Dim slFileDate As String

    ilAirDateMsg = False
    llSpotCount = -1
    llNextPercent = 35
    ilRet = 0
    smSplitFileName = ""
    'On Error GoTo mReadAndSplitErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        lbcError.AddItem "Open " & slFromFile & " error#" & Str$(ilRet)
        Print #hmMsg, "Open " & slFromFile & " error#" & Str$(ilRet)
        Close hmFrom
        mReadAndSplit = False
        Exit Function
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    lmProcessedNoBytes = 0
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mReadAndSplitErr:
        'Line Input #hmFrom, slLine
        slLine = ""
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                Exit Do
            ElseIf slChar <> sgCR Then
                slLine = slLine & slChar
            End If
        Loop
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Print #hmMsg, "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Close hmFrom
            mReadAndSplit = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, True, smFields()
                'slCallLetters = smFields(1)
                slCallLetters = smFields(0)
                'If (Len(Trim$(smFields(7))) > 0) Or (Len(Trim$(smFields(10))) > 0) Then
                If (Len(Trim$(smFields(6))) > 0) Or (Len(Trim$(smFields(9))) > 0) Then
                    'If (Len(Trim$(smFields(7))) > 0) Then
                    If (Len(Trim$(smFields(6))) > 0) Then
                        'slAirDate = Format$(smFields(7), sgShowDateForm)
                        slAirDate = Format$(smFields(6), sgShowDateForm)
                    Else
                        'slAirDate = Format$(smFields(10), sgShowDateForm)
                        slAirDate = Format$(smFields(9), sgShowDateForm)
                    End If
                    'If (Len(Trim$(smFields(8))) > 0) Then
                    If (Len(Trim$(smFields(7))) > 0) Then
                        'slAirTime = Format$(smFields(8), sgShowTimeWSecForm)
                        slAirTime = Format$(smFields(7), sgShowTimeWSecForm)
                    Else
                        'slAirTime = Format$(smFields(11), sgShowTimeWSecForm)
                        slAirTime = Format$(smFields(10), sgShowTimeWSecForm)
                    End If
                    If (DateValue(gAdjYear(slAirDate)) >= lmStartDate) And (DateValue(gAdjYear(slAirDate)) <= lmLastDate) Then
                        'slFdDate = Format$(smFields(10), sgShowDateForm)
                        slFdDate = Format$(smFields(9), sgShowDateForm)
                        'slFdTime = Format$(Format$(smFields(11), "h:mmam/pm"), sgShowTimeWSecForm)
                        slFdTime = Format$(Format$(smFields(10), "h:mmam/pm"), sgShowTimeWSecForm)
                        'If smFields(12) <> "" Then
                        If smFields(11) <> "" Then
                            'slPdDate = Format$(smFields(12), sgShowDateForm)
                            slPdDate = Format$(smFields(11), sgShowDateForm)
                        Else
                            slPdDate = slFdDate
                        End If
                        'If smFields(13) <> "" Then
                        If smFields(12) <> "" Then
                            'If gTimeToLong(Format$(smFields(11), "h:mm:ssam/pm"), False) = gTimeToLong(Format$(smFields(13), "h:mm:ssam/pm"), False) Then
                            If gTimeToLong(Format$(smFields(10), "h:mm:ssam/pm"), False) = gTimeToLong(Format$(smFields(11), "h:mm:ssam/pm"), False) Then
                                slPdSTime = slFdTime
                            Else
                                'slPdSTime = Format$(smFields(13), sgShowTimeWSecForm)
                                slPdSTime = Format$(smFields(12), sgShowTimeWSecForm)
                            End If
                        Else
                            slPdSTime = slFdTime
                        End If
                        'slPdETime = Format$(gTimeToLong(slPdSTime, False) + 60, "hh:mm:ss")
                        If (DateValue(gAdjYear(slAirDate)) <> DateValue(gAdjYear(slPdDate))) Or (gTimeToLong(slAirTime, False) <> gTimeToLong(slPdSTime, False)) Then
                            If llSpotCount = -1 Then
                                slFileDate = gObtainStartStd(slPdDate)
                                llSpotCount = 0
                                ilRet = mOpenSplitFile(slFileDate)
                                If Not ilRet Then
                                    mReadAndSplit = False
                                    Exit Function
                                End If
                            ElseIf llSpotCount = 40000 Then
                                llSpotCount = 0
                                ilRet = mOpenSplitFile(slFileDate)
                                If Not ilRet Then
                                    mReadAndSplit = False
                                    Exit Function
                                End If
                            End If
                            llSpotCount = llSpotCount + 1
                            Print #hmSplit, slLine
                        End If
                    End If
                Else
                    'slStr = smFields(4)
                    slStr = smFields(3)
                    ilPos = InStr(slStr, "-")
                    If ilPos > 0 Then
                        slCntrNo = Left$(slStr, ilPos - 1) & Mid$(slStr, ilPos + 1)
                    Else
                        slCntrNo = slStr
                    End If
                    llCntrNo = Val(slCntrNo)
                    If Not ilAirDateMsg Then
                        lbcError.AddItem "Air Date Missing: See Output Text File"
                        ilAirDateMsg = True
                    End If
                    'Print #hmMsg, "Air Date Missing: " & Trim$(smFields(1)) & " " & Trim$(smFields(3)) & Str$(llCntrNo) & " " & smFields(10) & " " & smFields(11) & " " & slLine
                    Print #hmMsg, "Air Date Missing: " & Trim$(smFields(0)) & " " & Trim$(smFields(2)) & Str$(llCntrNo) & " " & smFields(9) & " " & smFields(10) & " " & slLine
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
            llPercent = (lmProcessedNoBytes / lmTotalNoBytes) * 100
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmFloodPercent <> llPercent Then
                lmFloodPercent = llPercent
                lbcPercent.Caption = Str$(llPercent) & "%"
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    Close hmSplit
    If ilRet <> 0 Then
        mReadAndSplit = False
    Else
        mReadAndSplit = True
        lmFloodPercent = 100
        lbcPercent.Caption = "100%"
    End If
    Print #hmMsg, "** Creating Affiliate Spots Split Files Finished: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    Exit Function
mReadAndSplitErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportCSV-mReadAndSplit"
    Exit Function
End Function



'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFileMYLSpots               *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Private Function mReadFileMYLSpots(ilVefCode As Integer, slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim ilMaxCol As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilMatch As Integer
    Dim slCurDate As String
    Dim ilFound As Integer
    'Dim ilVefCode As Integer
    Dim ilAdfCode As Integer
    Dim slLogDate As String
    Dim slLogTime As String
    Dim slProd As String
    Dim slCart As String
    Dim ilLen As Integer
    Dim iUpper As Integer
    Dim slStr As String
    Dim slMsg As String
    Dim llPercent As Long
    Dim llDate As Long
    Dim ilHour As Integer
    Dim ilTest As Integer
    Dim ilBreakNo As Integer
    Dim ilPosition As Integer
        
    slCurDate = Format(gNow(), "mm/dd/yyyy")
    ReDim tmLstMYLInfo(0 To 0) As LSTMYLINFO
    'ilRet = 0
    'On Error GoTo mReadFileMYLSpotsErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        lbcError.AddItem "Open " & slFromFile & " error#" & Str$(ilRet)
        Print #hmMsg, "Open " & slFromFile & " error#" & Str$(ilRet)
        Close hmFrom
        mReadFileMYLSpots = False
        Exit Function
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    'ilVefCode = -1
    'For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
    '    If (StrComp(Trim$(tgVehicleInfo(ilLoop).sVehicle), "Music of Your Life", vbTextCompare) = 0) Or (StrComp(Trim$(tgVehicleInfo(ilLoop).sVehicle), "MYL", vbTextCompare) = 0) Or (StrComp(Trim$(tgVehicleInfo(ilLoop).sCodeStn), "MYL", vbTextCompare) = 0) Then
    '        ilVefCode = tgVehicleInfo(ilLoop).iCode
    '        Exit For
    '    End If
    'Next ilLoop
    'If ilVefCode = -1 Then
    '    lbcError.AddItem "Unable to Find 'Music of Your Life' as a Vehicle" & " error#" & Str$(ilRet)
    '    Print #hmMsg, "Unable to Find 'Music of Your Life' as a Vehicle" & " error#" & Str$(ilRet)
    '    Close hmFrom
    '    mReadFileMYLSpots = False
    '    Exit Function
    'End If
    ilAdfCode = -1
    For ilLoop = 0 To UBound(tgAdvtInfo) - 1 Step 1
        If (StrComp(Trim$(tgAdvtInfo(ilLoop).sAdvtName), "Music of Your Life", vbTextCompare) = 0) Or (StrComp(Trim$(tgAdvtInfo(ilLoop).sAdvtName), "MYL", vbTextCompare) = 0) Or (StrComp(Trim$(tgAdvtInfo(ilLoop).sAdvtAbbr), "MYL", vbTextCompare) = 0) Then
            ilAdfCode = tgAdvtInfo(ilLoop).iCode
            Exit For
        End If
    Next ilLoop
    If ilAdfCode = -1 Then
        lbcError.AddItem "Unable to Find 'Music of Your Life' as an Advertiser" & " error#" & Str$(ilRet)
        Print #hmMsg, "Unable to Find 'Music of Your Life' as a Advertiser" & " error#" & Str$(ilRet)
        Close hmFrom
        mReadFileMYLSpots = False
        Exit Function
    End If
    lmProcessedNoBytes = 0
    slLogDate = "1/1/1970"
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mReadFileMYLSpotsErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Print #hmMsg, "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Close hmFrom
            mReadFileMYLSpots = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, True, smFields()
                For ilLoop = LBound(smFields) To UBound(smFields) Step 1
                    smFields(ilLoop) = Trim$(smFields(ilLoop))
                Next ilLoop
                On Error GoTo ErrHand
                'Get LST for date if different
                'If DateValue(gAdjYear(slLogDate)) <> DateValue(gAdjYear(smFields(1))) Then
                If DateValue(gAdjYear(slLogDate)) <> DateValue(gAdjYear(smFields(0))) Then
                    ilHour = -1
                    ReDim tmLstMYLInfo(0 To 0) As LSTMYLINFO
                    'slLogDate = Format$(smFields(1), sgShowDateForm)
                    slLogDate = Format$(smFields(0), sgShowDateForm)
                    SQLQuery = "SELECT lstType, lstLogTime, lstLen, lstWkNo, lstBreakNo, lstPositionNo, lstZone, lstAnfCode, lstCode FROM lst "
                    SQLQuery = SQLQuery + " WHERE (lstLogVefCode = " & ilVefCode
                    SQLQuery = SQLQuery + " AND lstBkoutLstCode = 0"
                    SQLQuery = SQLQuery + " AND lstLogDate = '" & Format$(slLogDate, sgSQLDateForm) & "')"
                    SQLQuery = SQLQuery + " ORDER BY lstZone, lstLogDate, lstLogTime, lstBreakNo, lstPositionNo"
        
                    Set rst = gSQLSelectCall(SQLQuery)
                    While Not rst.EOF
                        iUpper = UBound(tmLstMYLInfo)
                        tmLstMYLInfo(iUpper).sLogTime = Format$(rst!lstLogTime, sgShowTimeWSecForm)
                        tmLstMYLInfo(iUpper).lLogTime = gTimeToLong(rst!lstLogTime, False)
                        tmLstMYLInfo(iUpper).iHour = tmLstMYLInfo(iUpper).lLogTime \ 3600
                        tmLstMYLInfo(iUpper).iType = rst!lstType
                        tmLstMYLInfo(iUpper).iLen = rst!lstLen
                        tmLstMYLInfo(iUpper).iWkNo = rst!lstWkNo
                        tmLstMYLInfo(iUpper).iBreakNo = rst!lstBreakNo
                        tmLstMYLInfo(iUpper).iPositionNo = rst!lstPositionNo
                        If IsNull(rst!lstZone) Then
                            tmLstMYLInfo(iUpper).sZone = ""
                        Else
                            tmLstMYLInfo(iUpper).sZone = rst!lstZone
                        End If
                        tmLstMYLInfo(iUpper).iAnfCode = rst!lstAnfCode
                        tmLstMYLInfo(iUpper).lCode = rst!lstCode
                        If iUpper = LBound(tmLstMYLInfo) Then
                            tmLstMYLInfo(iUpper).iHourBreakNo = 1
                            ilHour = tmLstMYLInfo(iUpper).iHour
                        Else
                            If ilHour <> tmLstMYLInfo(iUpper).iHour Then
                                tmLstMYLInfo(iUpper).iHourBreakNo = 1
                                ilHour = tmLstMYLInfo(iUpper).iHour
                            Else
                                If tmLstMYLInfo(iUpper - 1).lLogTime = tmLstMYLInfo(iUpper).lLogTime Then
                                    tmLstMYLInfo(iUpper).iHourBreakNo = tmLstMYLInfo(iUpper - 1).iHourBreakNo
                                Else
                                    tmLstMYLInfo(iUpper).iHourBreakNo = tmLstMYLInfo(iUpper - 1).iHourBreakNo + 1
                                End If
                            End If
                        End If
                        iUpper = iUpper + 1
                        ReDim Preserve tmLstMYLInfo(0 To iUpper) As LSTMYLINFO
                        rst.MoveNext
                    Wend
                End If
                ilFound = False
                'llDate = DateValue(gAdjYear(smFields(1)))
                llDate = DateValue(gAdjYear(smFields(0)))
                'ilHour = Hour(smFields(2) & ":00")
                ilHour = Hour(smFields(1) & ":00")
                'ilLen = Val(smFields(4))
                ilLen = Val(smFields(3))
                'slProd = gFixQuote(smFields(5))
                slProd = gFixQuote(smFields(4))
                'slCart = smFields(6)
                slCart = smFields(5)
                For ilLoop = 0 To UBound(tmLstMYLInfo) - 1 Step 1
                    'First the first avail or spot in matching hour
                    If ilHour = tmLstMYLInfo(ilLoop).iHour Then
                        'ilBreakNo = (Asc(smFields(3)) - Asc("A")) \ 2 + 1
                        ilBreakNo = (Asc(smFields(2)) - Asc("A")) \ 2 + 1
                        'ilPosition = ((Asc(smFields(3)) - Asc("A")) Mod 2) + 1
                        ilPosition = ((Asc(smFields(2)) - Asc("A")) Mod 2) + 1
                        If ilBreakNo = tmLstMYLInfo(ilLoop).iHourBreakNo Then
                            ilFound = True
                            cnn.BeginTrans
                            ilRet = 0
                            If ilPosition = 1 Then
                                'Remove first postion
                                'Remove second position event if at same time
                                'Note: first avail in hour is always a 2/60
                                '      and 'A' can be a 60sec replacement or
                                '      30sec.  If 30sec, then a 'B' spots will
                                '      also to sent.
                                If ilLoop < UBound(tmLstMYLInfo) - 1 Then
                                    SQLQuery = "DELETE FROM lst WHERE (lstCode = " & tmLstMYLInfo(ilLoop).lCode & ")"
                                    'cnn.Execute SQLQuery, rdExecDirect
                                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                        '6/11/16: Replaced GoSub
                                        'GoSub ErrHand:
                                        Screen.MousePointer = vbDefault
                                        gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileMYLSpots"
                                        cnn.RollbackTrans
                                        ilRet = 1
                                    End If
                                    If ilRet = 0 Then
                                        If tmLstMYLInfo(ilLoop).lLogTime = tmLstMYLInfo(ilLoop + 1).lLogTime Then
                                            SQLQuery = "DELETE FROM lst WHERE (lstCode = " & tmLstMYLInfo(ilLoop + 1).lCode & ")"
                                            'cnn.Execute SQLQuery, rdExecDirect
                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                '6/11/16: Replaced GoSub
                                                'GoSub ErrHand:
                                                Screen.MousePointer = vbDefault
                                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileMYLSpots"
                                                cnn.RollbackTrans
                                                ilRet = 1
                                            End If
                                        End If
                                    End If
                                End If
                                tmLstMYLInfo(ilLoop).iPositionNo = 1
                            Else
                                tmLstMYLInfo(ilLoop).iPositionNo = 2
                            End If
                            If ilRet = 0 Then
                                SQLQuery = "INSERT INTO lst (lstType, lstSdfCode, lstCntrNo, "
                                SQLQuery = SQLQuery & "lstAdfCode, lstAgfCode, lstProd, "
                                SQLQuery = SQLQuery & "lstLineNo, lstLnVefCode, lstStartDate,"
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
                                '12/28/06
                                'SQLQuery = SQLQuery & "lstAnfCode)"
                                SQLQuery = SQLQuery & "lstAnfCode, lstEvtIDCefCode, lstSplitNetwork, "
                                SQLQuery = SQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, "
                                SQLQuery = SQLQuery & "lstLnStartTime, lstLnEndTime, lstUnused)"
                                SQLQuery = SQLQuery & " VALUES (" & 0 & ", " & 0 & ", " & 0 & ", "
                                SQLQuery = SQLQuery & ilAdfCode & ", " & 0 & ", '" & slProd & "', "
                                SQLQuery = SQLQuery & 0 & ", " & 0 & ", '" & Format$("1/1/1970", sgSQLDateForm) & "', "
                                SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', " & 0 & ", " & 0 & ", "
                                SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", "
                                SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", "
                                SQLQuery = SQLQuery & 1 & ", " & 0 & ", " & 5 & ", "
                                SQLQuery = SQLQuery & ilVefCode & ", '" & Format$(slLogDate, sgSQLDateForm) & "', '" & Format$(tmLstMYLInfo(ilLoop).sLogTime, sgSQLTimeForm) & "', "
                                SQLQuery = SQLQuery & "'" & "" & "', " & 0 & ", '" & "" & "', "
                                SQLQuery = SQLQuery & tmLstMYLInfo(ilLoop).iWkNo & ", " & tmLstMYLInfo(ilLoop).iBreakNo & ", " & tmLstMYLInfo(ilLoop).iPositionNo & ", "
                                SQLQuery = SQLQuery & 0 & ", '" & tmLstMYLInfo(ilLoop).sZone & "', '" & slCart & "', "
                                SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", "
                                SQLQuery = SQLQuery & ilLen & ", " & 0 & ", " & 0 & ", "
                                '12/28/06
                                'SQLQuery = SQLQuery & tmLstMYLInfo(ilLoop).iAnfCode & ")"
                                SQLQuery = SQLQuery & tmLstMYLInfo(ilLoop).iAnfCode & ", " & 0 & ", '" & "N" & "', "
                                SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", '" & "N" & "', " & 0 & ", "
                                SQLQuery = SQLQuery & "'" & Format("12am", sgSQLTimeForm) & "', '" & Format("12am", sgSQLTimeForm) & "', '" & "" & "'" & ")"
                                'cnn.Execute SQLQuery, rdExecDirect
                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    '6/11/16: Replaced GoSub
                                    'GoSub ErrHand:
                                    Screen.MousePointer = vbDefault
                                    gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileMYLSpots"
                                    cnn.RollbackTrans
                                    ilRet = 1
                                End If
                                If ilRet = 0 Then
                                    cnn.CommitTrans
                                End If
                            End If
                            Exit For
                        End If
                    End If
                Next ilLoop
                If Not ilFound Then
                    lbcError.AddItem "Suppress Spot Not Found for: " & Trim$(slLine)
                    Print #hmMsg, "Suppress Spot Not Found for: " & Trim$(slLine)
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
            llPercent = (lmProcessedNoBytes / lmTotalNoBytes) * 100
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmFloodPercent <> llPercent Then
                lmFloodPercent = llPercent
                lbcPercent.Caption = Str$(llPercent) & "%"
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    If ilRet <> 0 Then
        mReadFileMYLSpots = False
    Else
        mReadFileMYLSpots = True
        lmFloodPercent = 100
        lbcPercent.Caption = "100%"
    End If
    Print #hmMsg, "** Import MYL Spots Info Finished: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    Exit Function
mReadFileMYLSpotsErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportCSV-mReadFileMYLSpots"
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFileCP                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Private Function mReadFileCP(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilLoop As Integer
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilCode As Integer
    Dim slCallLetters As String
    Dim ilShttCode As Integer
    Dim slVehicle As String
    Dim ilVefCode As Integer
    Dim llAttCode As Long
    Dim ilStatus As Integer
    Dim ilSelected As Integer
    Dim slCurDate As String
    Dim slCurTime As String
    Dim ilPos As Integer
    Dim slMsg As String
    Dim llPercent As Long
    Dim ilUpper As Integer
    Dim slSDate As String
    Dim slRDate As String
    Dim slStr As String
    Dim llStationID As Long
    Dim slTime As String
    Dim ilCycle As Integer
    Dim llSDate As Long
    Dim ilPostingStatus As Integer
    Dim slChar As String
    Dim ilVef As Integer
        
    slCurDate = Format(gNow(), sgShowDateForm)
    slCurTime = Format(gNow(), sgShowTimeWSecForm)
    slTime = Format("12:00AM", "hh:mm:ss")
    ilCycle = 7
    ReDim smMissingVef(0 To 0) As String
    ReDim smMissingShtt(0 To 0) As String
    ilUpper = 0
    ReDim tmAgreeID(0 To 0) As AGREEID
    ReDim tmMissingAtt(0 To 0) As MISSINGATT
    ilRet = 0
    On Error GoTo ErrHand
    SQLQuery = "SELECT attShfCode, attVefCode, attOnAir, attOffAir, attDropDate, attCode"
    SQLQuery = SQLQuery + " FROM att"
    Set rst = gSQLSelectCall(SQLQuery)
    If ilRet <> 0 Then
        mReadFileCP = False
        Exit Function
    End If
    While Not rst.EOF
        ilUpper = UBound(tmAgreeID)
        tmAgreeID(ilUpper).lCode = rst!attCode
        tmAgreeID(ilUpper).iShttCode = rst!attshfCode
        tmAgreeID(ilUpper).iVefCode = rst!attvefCode
        tmAgreeID(ilUpper).lOnAir = DateValue(gAdjYear(rst!attOnAir))
        tmAgreeID(ilUpper).lOffAir = DateValue(gAdjYear(rst!attOffAir))
        tmAgreeID(ilUpper).lDropDate = DateValue(gAdjYear(rst!attDropDate))
        If tmAgreeID(ilUpper).lDropDate < tmAgreeID(ilUpper).lOffAir Then
            tmAgreeID(ilUpper).lEndDate = tmAgreeID(ilUpper).lDropDate
        Else
            tmAgreeID(ilUpper).lEndDate = tmAgreeID(ilUpper).lOffAir
        End If
        ilUpper = ilUpper + 1
        ReDim Preserve tmAgreeID(0 To ilUpper) As AGREEID
        rst.MoveNext
    Wend

    'ilRet = 0
    'On Error GoTo mReadFileCPErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        lbcError.AddItem "Open " & slFromFile & " error#" & Str$(ilRet)
        Print #hmMsg, "Open " & slFromFile & " error#" & Str$(ilRet)
        Close hmFrom
        mReadFileCP = False
        Exit Function
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    lmProcessedNoBytes = 0
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mReadFileCPErr:
        'Line Input #hmFrom, slLine
        slLine = ""
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                Exit Do
            ElseIf slChar <> sgCR Then
                slLine = slLine & slChar
            End If
        Loop
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Print #hmMsg, "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Close hmFrom
            mReadFileCP = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, True, smFields()
                For ilLoop = LBound(smFields) To UBound(smFields) Step 1
                    smFields(ilLoop) = Trim$(smFields(ilLoop))
                Next ilLoop
                'smFields(1) = UCase$(Trim$(smFields(1)))
                smFields(1) = UCase$(Trim$(smFields(0)))
                'slCallLetters = smFields(1)
                slCallLetters = smFields(0)
                ilShttCode = -1
                If Len(slCallLetters) > 40 Then
                    lbcError.AddItem slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                    Print #hmMsg, slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                    slCallLetters = Left$(slCallLetters, 40)
                End If
                'llStationID = Val(smFields(2))
                llStationID = Val(smFields(1))
                For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                    'If StrComp(Trim$(tgStationInfo(ilLoop).sCallLetters), slCallLetters, vbTextCompare) = 0 Then
                    If tgStationInfo(ilLoop).lID = llStationID Then
                        ilShttCode = tgStationInfo(ilLoop).iCode
                        Exit For
                    End If
                Next ilLoop
                
                If ilShttCode > 0 Then
                    'slVehicle = Trim$(smFields(3))
                    slVehicle = Trim$(smFields(2))
                    ilVefCode = -1
                    ilCycle = 7
                    ilSelected = False
                    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                        If sgShowByVehType = "Y" Then
                            smVehName = Mid$(tgVehicleInfo(ilLoop).sVehicle, 3)
                        Else
                            smVehName = tgVehicleInfo(ilLoop).sVehicle
                        End If
                        'If StrComp(Trim$(smVehName), smFields(3), vbTextCompare) = 0 Then
                        If StrComp(Trim$(smVehName), smFields(2), vbTextCompare) = 0 Then
                            'ilVefCode = tgVehicleInfo(ilLoop).icode
                            'If lbcNames.Selected(ilLoop) Then
                            '    ilSelected = True
                            'End If
                            ilCycle = tgVehicleInfo(ilLoop).iNoDaysCycle
                            'Exit For
                            ilFound = False
                            ilVefCode = tgVehicleInfo(ilLoop).iCode
                            For ilVef = 0 To lbcNames.ListCount - 1 Step 1
                                If lbcNames.ItemData(ilVef) = ilVefCode Then
                                    If lbcNames.Selected(ilLoop) Then
                                        ilSelected = True
                                    End If
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilVef
                            If ilFound Then
                                Exit For
                            End If
                        End If
                    Next ilLoop
                    If (ilVefCode > 0) And (ilSelected) Then
                        llAttCode = -1
                        'slSDate = Format$(smFields(4), sgShowDateForm)
                        slSDate = Format$(smFields(3), sgShowDateForm)
                        llSDate = DateValue(gAdjYear(slSDate))
                        For ilLoop = LBound(tmAgreeID) To UBound(tmAgreeID) - 1 Step 1
                            If (tmAgreeID(ilLoop).iShttCode = ilShttCode) And (tmAgreeID(ilLoop).iVefCode = ilVefCode) Then
                                If (llSDate >= tmAgreeID(ilLoop).lOnAir) And (llSDate <= tmAgreeID(ilLoop).lEndDate) Then
                                    llAttCode = tmAgreeID(ilLoop).lCode
                                    Exit For
                                End If
                            End If
                        Next ilLoop
                        If llAttCode > 0 Then
                            If (DateValue(gAdjYear(slSDate)) >= lmStartDate) And (DateValue(gAdjYear(slSDate)) <= lmLastDate) Then
                                'If smFields(5) <> "" Then
                                If smFields(4) <> "" Then
                                    'slRDate = Format$(smFields(5), sgShowDateForm)
                                    slRDate = Format$(smFields(4), sgShowDateForm)
                                Else
                                    slRDate = ""
                                End If
                                'ilStatus = Val(smFields(6))
                                ilStatus = Val(smFields(5))
                                If ilStatus > 0 Then
                                    ilPostingStatus = 2
                                Else
                                    ilPostingStatus = 0
                                End If
                                'Add CP
                                Do
                                    If slRDate <> "" Then
                                        SQLQuery = "INSERT INTO cptt(cpttAtfCode, cpttShfCode, cpttVefCode, "
                                        SQLQuery = SQLQuery & "cpttCreateDate, cpttStartDate, cpttReturnDate, "
                                        SQLQuery = SQLQuery & "cpttStatus, cpttUsfCode, cpttPostingStatus)"
                                        SQLQuery = SQLQuery & " VALUES "
                                        SQLQuery = SQLQuery & "(" & llAttCode & ", " & ilShttCode & ", " & ilVefCode & ", "
                                        SQLQuery = SQLQuery & "'" & Format$(slCurDate, sgSQLDateForm) & "', '" & Format$(slSDate, sgSQLDateForm) & "', '" & Format$(slRDate, sgSQLDateForm) & "', "
                                        SQLQuery = SQLQuery & "" & ilStatus & ", " & igUstCode & ", " & ilPostingStatus & ")"
                                    Else
                                        SQLQuery = "INSERT INTO cptt(cpttAtfCode, cpttShfCode, cpttVefCode, "
                                        SQLQuery = SQLQuery & "cpttCreateDate, cpttStartDate, "
                                        SQLQuery = SQLQuery & "cpttStatus, cpttUsfCode, cpttPostingStatus)"
                                        SQLQuery = SQLQuery & " VALUES "
                                        SQLQuery = SQLQuery & "(" & llAttCode & ", " & ilShttCode & ", " & ilVefCode & ", "
                                        SQLQuery = SQLQuery & "'" & Format$(slCurDate, sgSQLDateForm) & "', '" & Format$(slSDate, sgSQLDateForm) & "', "
                                        SQLQuery = SQLQuery & "" & ilStatus & ", " & igUstCode & ", " & ilPostingStatus & ")"
                                    End If
                                    cnn.BeginTrans
                                    'cnn.Execute SQLQuery, rdExecDirect
                                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                        '6/11/16: Replaced GoSub
                                        'GoSub ErrHand:
                                        Screen.MousePointer = vbDefault
                                        gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileCP"
                                        cnn.RollbackTrans
                                        ilRet = 1
                                    End If
                                    If ilRet = 0 Then
                                        cnn.CommitTrans
                                    End If
                                    ilCycle = ilCycle - 7
                                    slSDate = Format$(DateValue(gAdjYear(slSDate)) + 7, sgShowDateForm)
                                Loop While ilCycle > 0
                                gFileChgdUpdate "cptt.mkd", True
                            End If
                        Else
                            'If Not llAttMissingMsg Then
                            '    lbcError.AddItem "Agreement(s) Missing: See Output Text File"
                            '    llAttMissingMsg = True
                            'End If
                            'Print #hmMsg, "Agreement Missing: " & Trim$(smFields(1)) & " " & Trim$(smFields(3))
                            'Print #hmMsg, "Agreement Missing: " & Trim$(smFields(0)) & " " & Trim$(smFields(2))
                            ilFound = False
                            For ilLoop = 0 To UBound(tmMissingAtt) - 1 Step 1
                                If (ilShttCode = tmMissingAtt(ilLoop).iShttCode) And (ilVefCode = tmMissingAtt(ilLoop).iVefCode) Then
                                    ilFound = True
                                    tmMissingAtt(ilLoop).lCount = tmMissingAtt(ilLoop).lCount + 1
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                'lbcError.AddItem "Agreement Missing: " & Trim$(smFields(1)) & " " & Trim$(smFields(3))
                                lbcError.AddItem "Agreement Missing: " & Trim$(smFields(0)) & " " & Trim$(smFields(2))
                                tmMissingAtt(UBound(tmMissingAtt)).iShttCode = ilShttCode
                                tmMissingAtt(UBound(tmMissingAtt)).iVefCode = ilVefCode
                                tmMissingAtt(UBound(tmMissingAtt)).lCount = 1
                                ReDim Preserve tmMissingAtt(0 To UBound(tmMissingAtt) + 1) As MISSINGATT
                            End If
                        End If
                    Else
                        If ilVefCode < 0 Then
                            ilFound = False
                            For ilLoop = 0 To UBound(smMissingVef) - 1 Step 1
                                'If StrComp(smFields(3), smMissingVef(ilLoop), 1) = 0 Then
                                If StrComp(smFields(2), smMissingVef(ilLoop), 1) = 0 Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                'lbcError.AddItem "Vehicle Missing: " & Trim$(smFields(3))
                                lbcError.AddItem "Vehicle Missing: " & Trim$(smFields(2))
                                'Print #hmMsg, "Vehicle Missing: " & Trim$(smFields(3))
                                Print #hmMsg, "Vehicle Missing: " & Trim$(smFields(2))
                                'smMissingVef(UBound(smMissingVef)) = smFields(3)
                                smMissingVef(UBound(smMissingVef)) = smFields(2)
                                ReDim Preserve smMissingVef(0 To UBound(smMissingVef) + 1) As String
                            End If
                        End If
                    End If
                Else
                    ilFound = False
                    For ilLoop = 0 To UBound(smMissingShtt) - 1 Step 1
                        'If StrComp(smFields(1), smMissingShtt(ilLoop), 1) = 0 Then
                        If StrComp(smFields(0), smMissingShtt(ilLoop), 1) = 0 Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        'lbcError.AddItem "Station Missing: " & Trim$(smFields(1))
                        lbcError.AddItem "Station Missing: " & Trim$(smFields(0))
                        'Print #hmMsg, "Station Missing: " & Trim$(smFields(1))
                        Print #hmMsg, "Station Missing: " & Trim$(smFields(0))
                        'smMissingShtt(UBound(smMissingShtt)) = smFields(1)
                        smMissingShtt(UBound(smMissingShtt)) = smFields(0)
                        ReDim Preserve smMissingShtt(0 To UBound(smMissingShtt) + 1) As String
                    End If
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
            llPercent = (lmProcessedNoBytes / lmTotalNoBytes) * 100
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmFloodPercent <> llPercent Then
                lmFloodPercent = llPercent
                lbcPercent.Caption = Str$(llPercent) & "%"
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    If ilRet <> 0 Then
        mReadFileCP = False
    Else
        mReadFileCP = True
        lmFloodPercent = 100
        lbcPercent.Caption = "100%"
    End If
    Print #hmMsg, "** Import CP's Finished: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    Exit Function
mReadFileCPErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportCSV-mReadFileCP"
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFilePledge_Sv                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Private Function mReadFilePledge_Sv(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim ilMaxCol As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilMatch As Integer
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilCode As Integer
    Dim slCallLetters As String
    Dim ilShttCode As Integer
    Dim slVehicle As String
    Dim ilVefCode As Integer
    Dim llAttCode As Long
    Dim llLstCode As Long
    Dim ilStatus As Integer
    Dim ilSelected As Integer
    Dim slCurDate As String
    Dim slCurTime As String
    Dim ilPos As Integer
    Dim slMsg As String
    Dim llPercent As Long
    Dim ilUpper As Integer
    Dim llEffSDate As Long
    Dim llEffEDate As Long
    Dim slName As String
    Dim ilAdfCode As Integer
    Dim slCntrNo As String
    Dim llCntrNo As Long
    Dim slProd As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim slFdDate As String
    Dim slFdTime As String
    Dim slPdDate As String
    Dim slPdTime As String
    Dim llAttMissingMsg As Long
    Dim ilLstMissingMsg As Integer
    Dim ilOverlapMsg As Integer
    Dim slStr As String
    Dim llStationID As Long
    Dim slZone As String
    Dim slFdSTime As String
    Dim slFdETime As String
    Dim slPdSTime As String
    Dim slPdETime As String
    Dim slOnAir As String
    Dim slOffAir As String
    Dim llExcludeDate As Long
    Dim llTFNDate As Long
    Dim ilPledge As Integer
    Dim llFdTime As Long
    Dim llTstFdTime As Long
    Dim ilUpdate As Integer
    Dim ilMerged As Integer
    Dim ilPd As Integer
    Dim ilDupl As Integer
    Dim ilDuplMsg As Integer
    Dim ilFoundType As Integer
    Dim llAttFound As Long
    Dim ilFdSHour As Integer
    Dim ilFdEHour As Integer
    Dim llAtt As Long
    Dim slChar As String
    Dim llTemp As Long
    Dim ilVef As Integer
    Dim VehCombo_rst As ADODB.Recordset
    ReDim ilLDays(0 To 6) As Integer
    ReDim ilPdDays(0 To 6) As Integer
    ReDim ilFdDays(0 To 6) As Integer

        
    slCurDate = Format(gNow(), sgShowDateForm)
    slCurTime = Format(gNow(), sgShowTimeWSecForm)
    ilOverlapMsg = False
    llExcludeDate = DateValue("1/1/2000")
    llTFNDate = DateValue("12/31/1999")
    ReDim smMissingVef(0 To 0) As String
    ReDim smMissingShtt(0 To 0) As String
    ReDim smMissingTime(0 To 0) As String
    ilUpper = 0
    ReDim tmAgreeID(0 To 0) As AGREEID
    ReDim tmPledgeCount(0 To 0) As MISSINGATT
    ReDim tmMissingAtt(0 To 0) As MISSINGATT
    ReDim tmPledgeInfo(0 To 0) As PledgeInfo
    ilLstMissingMsg = False
    ilDuplMsg = False
    ilRet = 0
    On Error GoTo ErrHand
    SQLQuery = "SELECT attShfCode, attVefCode, attOnAir, attOffAir, attDropDate, attCode"
    SQLQuery = SQLQuery + " FROM att"
    Set rst = gSQLSelectCall(SQLQuery)
    If ilRet <> 0 Then
        mReadFilePledge_Sv = False
        Exit Function
    End If
    While Not rst.EOF
        ilUpper = UBound(tmAgreeID)
        tmAgreeID(ilUpper).lCode = rst!attCode
        tmAgreeID(ilUpper).iShttCode = rst!attshfCode
        tmAgreeID(ilUpper).iVefCode = rst!attvefCode
        tmAgreeID(ilUpper).lOnAir = DateValue(gAdjYear(rst!attOnAir))
        tmAgreeID(ilUpper).lOffAir = DateValue(gAdjYear(rst!attOffAir))
        tmAgreeID(ilUpper).lDropDate = DateValue(gAdjYear(rst!attDropDate))
        If tmAgreeID(ilUpper).lDropDate < tmAgreeID(ilUpper).lOffAir Then
            tmAgreeID(ilUpper).lEndDate = tmAgreeID(ilUpper).lDropDate
        Else
            tmAgreeID(ilUpper).lEndDate = tmAgreeID(ilUpper).lOffAir
        End If
        ilUpper = ilUpper + 1
        ReDim Preserve tmAgreeID(0 To ilUpper) As AGREEID
        rst.MoveNext
    Wend

    'ilRet = 0
    'On Error GoTo mReadFilePledge_SvErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        lbcError.AddItem "Open " & slFromFile & " error#" & Str$(ilRet)
        Print #hmMsg, "Open " & slFromFile & " error#" & Str$(ilRet)
        Close hmFrom
        mReadFilePledge_Sv = False
        Exit Function
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    lmProcessedNoBytes = 0
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mReadFilePledge_SvErr:
        'Line Input #hmFrom, slLine
        slLine = ""
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                Exit Do
            ElseIf slChar <> sgCR Then
                slLine = slLine & slChar
            End If
        Loop
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Print #hmMsg, "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Close hmFrom
            mReadFilePledge_Sv = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, True, smFields()
                For ilLoop = LBound(smFields) To UBound(smFields) Step 1
                    smFields(ilLoop) = Trim$(smFields(ilLoop))
                Next ilLoop
                'smFields(1) = UCase$(Trim$(smFields(1)))
                smFields(0) = UCase$(Trim$(smFields(0)))
                'slCallLetters = smFields(1)
                slCallLetters = smFields(0)
                ilShttCode = -1
                If Len(slCallLetters) > 40 Then
                    lbcError.AddItem slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                    Print #hmMsg, slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                    slCallLetters = Left$(slCallLetters, 40)
                End If
                'llStationID = Val(smFields(2))
                llStationID = Val(smFields(1))
                For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                    'If StrComp(Trim$(tgStationInfo(ilLoop).sCallLetters), slCallLetters, vbTextCompare) = 0 Then
                    If tgStationInfo(ilLoop).lID = llStationID Then
                        ilShttCode = tgStationInfo(ilLoop).iCode
                        slZone = UCase$(Left$(tgStationInfo(ilLoop).sZone, 1))
                        Exit For
                    End If
                Next ilLoop
                If ilShttCode > 0 Then
                    'slVehicle = Trim$(smFields(3))
                    slVehicle = Trim$(smFields(2))
                    ilVefCode = -1
                    ilSelected = False
                    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                        If sgShowByVehType = "Y" Then
                            smVehName = Mid$(tgVehicleInfo(ilLoop).sVehicle, 3)
                        Else
                            smVehName = tgVehicleInfo(ilLoop).sVehicle
                        End If
                        'If StrComp(Trim$(smVehName), smFields(3), vbTextCompare) = 0 Then
                        If StrComp(Trim$(smVehName), smFields(2), vbTextCompare) = 0 Then
                            'ilVefCode = tgVehicleInfo(ilLoop).icode
                            'If lbcNames.Selected(ilLoop) Then
                            '    ilSelected = True
                            'End If
                            'Exit For
                            ilFound = False
                            ilVefCode = tgVehicleInfo(ilLoop).iCode
                            For ilVef = 0 To lbcNames.ListCount - 1 Step 1
                                If lbcNames.ItemData(ilVef) = ilVefCode Then
                                    If lbcNames.Selected(ilLoop) Then
                                        ilSelected = True
                                    End If
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilVef
                            If ilFound Then
                                Exit For
                            End If
                        End If
                    Next ilLoop
                    If (ilVefCode > 0) And (ilSelected) Then
                        'If Len(smFields(11)) <> 0 Then
                        If Len(smFields(10)) <> 0 Then
                            'Truncate seconds
                            'slFdSTime = Format$(smFields(11), "hh:mm") & ":00" '"hh:mm:ss")
                            slFdSTime = Format$(smFields(10), "hh:mm") & ":00" '"hh:mm:ss")
                        Else
                            slFdSTime = ""
                        End If
                        'If Len(smFields(12)) <> 0 Then
                        If Len(smFields(11)) <> 0 Then
                            'slFdETime = Format$(smFields(12), "h:mmam/pm")
                            slFdETime = Format$(smFields(11), "h:mmam/pm")
                        Else
                            slFdETime = Format$(slFdSTime, "h:mmam/pm")
                        End If
                        slFdETime = Format$(gLongToTime(gTimeToLong(slFdETime, False) + 60), "hh:mm;ss")
                        '***********************************************
                        'Adjust Feed times because import data is wrong
                        'CST- Add One Hour
                        'MST- Add Two Hours
                        'PST- Add Three hours
                        If (slZone = "C") Or (slZone = "M") Or (slZone = "P") Then
                            ilFdSHour = Val(Left$(slFdSTime, 2))
                            ilFdEHour = Val(Left$(slFdETime, 2))
                            Select Case slZone
                                Case "C"
                                    ilFdSHour = ilFdSHour + 1
                                    ilFdEHour = ilFdEHour + 1
                                Case "M"
                                    ilFdSHour = ilFdSHour + 2
                                    ilFdEHour = ilFdEHour + 2
                                Case "P"
                                    ilFdSHour = ilFdSHour + 3
                                    ilFdEHour = ilFdEHour + 3
                            End Select
                            If ilFdSHour > 23 Then
                                ilFdSHour = ilFdSHour - 24
                            End If
                            If ilFdEHour > 23 Then
                                ilFdEHour = ilFdEHour - 24
                            End If
                            If ilFdSHour < 10 Then
                                slFdSTime = "0" & Trim$(Str$(ilFdSHour)) & Mid$(slFdSTime, 3)
                            Else
                                slFdSTime = Trim$(Str$(ilFdSHour)) & Mid$(slFdSTime, 3)
                            End If
                            If ilFdEHour < 10 Then
                                slFdETime = "0" & Trim$(Str$(ilFdEHour)) & Mid$(slFdETime, 3)
                            Else
                                slFdETime = Trim$(Str$(ilFdEHour)) & Mid$(slFdETime, 3)
                            End If
                        End If
                        '***********************************************
        
                        'ilStatus = Val(smFields(13))
                        ilStatus = Val(smFields(12))
                        If ilStatus > 0 Then
                            ilStatus = ilStatus - 1
                        End If
                        If ilStatus = 6 Then 'Change status 7 to 2
                            ilStatus = 1
                        End If
                        'If Len(smFields(21)) <> 0 Then
                        If Len(smFields(20)) <> 0 Then
                            'slPdSTime = Format$(smFields(21), "hh:mm:ss")
                            slPdSTime = Format$(smFields(20), "hh:mm:ss")
                        Else
                            slPdSTime = slFdSTime
                        End If
                        'If Len(smFields(22)) <> 0 Then
                        If Len(smFields(21)) <> 0 Then
                            'slPdETime = Format$(smFields(22), "hh:mm:ss")
                            slPdETime = Format$(smFields(21), "hh:mm:ss")
                        Else
                            slPdETime = slPdSTime
                        End If
                        If gTimeToLong(slPdSTime, False) + 60 >= gTimeToLong(slPdETime, False) Then
                            'If Len(smFields(21)) <> 0 Then
                            If Len(smFields(20)) <> 0 Then
                                'Truncate seconds
                                'slPdSTime = Format$(smFields(21), "hh:mm") & ":00" '"hh:mm:ss")
                                slPdSTime = Format$(smFields(20), "hh:mm") & ":00" '"hh:mm:ss")
                            Else
                                slPdSTime = slFdSTime
                            End If
                            'If Len(smFields(22)) <> 0 Then
                            If Len(smFields(21)) <> 0 Then
                                'slPdETime = Format$(smFields(22), "h:mmam/pm")
                                slPdETime = Format$(smFields(21), "h:mmam/pm")
                            Else
                                slPdETime = Format$(slPdSTime, "h:mmam/pm")
                            End If
                            slPdETime = Format$(gLongToTime(gTimeToLong(slPdETime, False) + 60), "hh:mm:ss")
                        End If
                        llAttFound = False
                        For llAtt = LBound(tmAgreeID) To UBound(tmAgreeID) - 1 Step 1
                            If (tmAgreeID(llAtt).iShttCode = ilShttCode) And (tmAgreeID(llAtt).iVefCode = ilVefCode) Then
                                llAttFound = True
                                ilFoundType = -1
                                llAttCode = tmAgreeID(llAtt).lCode
                                'If smFields(23) <> "" Then
                                If smFields(22) <> "" Then
                                    'llEffSDate = DateValue(gAdjYear(smFields(23)))
                                    llEffSDate = DateValue(gAdjYear(smFields(22)))
                                Else
                                    llEffSDate = tmAgreeID(llAtt).lOnAir
                                End If
                                'If smFields(24) <> "" Then
                                If smFields(23) <> "" Then
                                    'llEffEDate = DateValue(gAdjYear(smFields(24)))
                                    llEffEDate = DateValue(gAdjYear(smFields(23)))
                                Else
                                    llEffEDate = DateValue("12/31/2069")
                                End If
                                'If llEffSDate >= llExcludeDate Then
                                If llEffEDate < lmStartDate Then
                                    'ilFoundType = 0
                                    Exit For
                                Else
                                    'If llEffEDate >= llTFNDate Then
                                    '    llEffEDate = DateValue("12/31/2069")
                                    'End If
                                    'If (tmAgreeID(llAtt).lOnAir <= llEffSDate) And (tmAgreeID(llAtt).lOffAir >= llEffEDate) Then
                                    '    ilFoundType = 1 'Add Pledge to agreement
                                    'Else
                                    '    If (llEffEDate < tmAgreeID(llAtt).lOnAir) Or (llEffSDate > tmAgreeID(llAtt).lOffAir) Then
                                    '        ilFoundType = 2 'Create new Agreement and add Pledge
                                    '    Else
                                    '        ilFoundType = 3 'Overlap dates
                                    '    End If
                                    'End If
                                    If (llEffEDate < tmAgreeID(llAtt).lOnAir) Or (llEffSDate > tmAgreeID(llAtt).lEndDate) Then
                                        ilFoundType = 0
                                    Else
                                        ilFoundType = 1
                                    End If
                                End If
                                On Error GoTo ErrHand
                                If (ilFoundType = 1) Or (ilFoundType = 2) Then
                                    'Test if avails defined for Feed Times
                                    ilRet = 0
                                    If ilFoundType = 2 Then
                                        'Add Agreement
                                        SQLQuery = "SELECT *"
                                        SQLQuery = SQLQuery + " FROM att"
                                        SQLQuery = SQLQuery & " WHERE (attCode = " & llAttCode & ")"
                                        Set rst = gSQLSelectCall(SQLQuery)
                                        If (ilRet = 0) And (Not rst.EOF) Then
                                            'Insert agreement with new Dates
                                            slOnAir = Format$(llEffSDate, sgShowDateForm)
                                            slOffAir = Format$(llEffEDate, sgShowDateForm)
                                            'D.S. 8/2/05
                                            llTemp = gFindAttHole()
                                            If llTemp = -1 Then
                                                mReadFilePledge_Sv = False
                                                Screen.MousePointer = vbDefault
                                                Exit Function
                                            End If
                                            SQLQuery = "INSERT INTO att (attCode, attShfCode, attVefCode, attAgreeStart, "
                                            SQLQuery = SQLQuery & "attAgreeEnd, attOnAir, attOffAir, "
                                            SQLQuery = SQLQuery & "attSigned, attSignDate, attLoad, "
                                            SQLQuery = SQLQuery & "attTimeType, attComp, attBarCode, "
                                            SQLQuery = SQLQuery & "attDropDate, attUsfCode, "
                                            SQLQuery = SQLQuery & "attEnterDate, attEnterTime, attNotice, "
                                            SQLQuery = SQLQuery & "attCarryCmml, attNoCDs, attSendTape, "
                                            SQLQuery = SQLQuery & "attACName, attACPhone, attGenLog, "
                                            SQLQuery = SQLQuery & "attGenCP, attPostingType, attPrintCP, "
                                            SQLQuery = SQLQuery & "attComments, attGenOther, attAgreementID, attPledgeType)"
                                            SQLQuery = SQLQuery & " VALUES (" & llTemp & ", " & ilShttCode & "," & ilVefCode & ",'" & Format$(slOnAir, sgSQLDateForm) & "',"
                                            SQLQuery = SQLQuery & "'" & Format$(slOffAir, sgSQLDateForm) & "','" & Format$(slOnAir, sgSQLDateForm) & "','" & Format$(slOffAir, sgSQLDateForm) & "',"
                                            SQLQuery = SQLQuery & rst!attSigned & ",'" & Format$(rst!attSignDate, sgSQLDateForm) & "'," & rst!attLoad & ","
                                            SQLQuery = SQLQuery & rst!attTimeType & "," & rst!attComp & "," & rst!attBarCode & ","
                                            SQLQuery = SQLQuery & "'" & Format$("12/31/2069", sgSQLDateForm) & "'," & 1 & ","
                                            SQLQuery = SQLQuery & "'" & Format$(slCurDate, sgSQLDateForm) & "','" & Format$(slCurTime, sgSQLTimeForm) & "','" & rst!attNotice & "',"
                                            SQLQuery = SQLQuery & rst!attCarryCmml & "," & rst!attNoCDs & "," & rst!attSendTape & ","
                                            SQLQuery = SQLQuery & "'" & rst!attACName & "','" & rst!attACPhone & "','" & rst!attGenLog & "',"
                                            SQLQuery = SQLQuery & "'" & rst!attGenCP & "'," & rst!attPostingType & "," & rst!attPrintCP & ","
                                            SQLQuery = SQLQuery & "'" & "" & "','" & rst!attGenOther & "'," & 0 & ",'" & "A" & "'" & ")"
                                            cnn.BeginTrans
                                            'cnn.Execute SQLQuery, rdExecDirect
                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                '6/11/16: Replaced GoSub
                                                'GoSub ErrHand:
                                                Screen.MousePointer = vbDefault
                                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFilePledge_Sv"
                                                cnn.RollbackTrans
                                                ilRet = 1
                                            End If
                                            If ilRet = 0 Then
                                                cnn.CommitTrans
                                                If llTemp = 0 Then
                                                    SQLQuery = "Select MAX(attCode) from att"
                                                    Set rst = gSQLSelectCall(SQLQuery)
                                                    ilUpper = UBound(tmAgreeID)
                                                    tmAgreeID(ilUpper).lCode = rst(0).Value
                                                Else
                                                    ilUpper = UBound(tmAgreeID)
                                                    tmAgreeID(ilUpper).lCode = llTemp
                                                End If
                                                tmAgreeID(ilUpper).iShttCode = ilShttCode
                                                tmAgreeID(ilUpper).iVefCode = ilVefCode
                                                tmAgreeID(ilUpper).lOnAir = llEffSDate
                                                tmAgreeID(ilUpper).lOffAir = llEffEDate
                                                tmAgreeID(ilUpper).lEndDate = llEffEDate
                                                ilUpper = ilUpper + 1
                                                ReDim Preserve tmAgreeID(0 To ilUpper) As AGREEID
                                                If llTemp = 0 Then
                                                    llAttCode = rst(0).Value
                                                Else
                                                    llAttCode = llTemp
                                                End If
                                                ReDim tgDat(0 To 0) As DAT
                                                SQLQuery = "Select vefCombineVefCode from VEF_Vehicles Where vefCode = " & imVefCode
                                                Set VehCombo_rst = gSQLSelectCall(SQLQuery)
                                                If Not VehCombo_rst.EOF Then
                                                    imVefCombo = VehCombo_rst!vefCombineVefCode
                                                End If
                                                gGetAvails llAttCode, ilShttCode, ilVefCode, imVefCombo, slOnAir, True
                                                'Write the avail out
                                                
                                            End If
                                        Else
                                            ilRet = 1
                                        End If
                                    End If
                                    If ilRet = 0 Then
                                        For ilLoop = 0 To 6 Step 1
                                            ilFdDays(ilLoop) = Val(smFields(ilLoop + 4))
                                        Next ilLoop
                                        For ilLoop = 0 To 6 Step 1
                                            ilPdDays(ilLoop) = Val(smFields(ilLoop + 14))
                                        Next ilLoop
                                        'Find Pledge to merge exception into
                                        ilMerged = False
                                        ReDim tmPledgeInfo(0 To 0) As PledgeInfo
                                        SQLQuery = "SELECT * "
                                        SQLQuery = SQLQuery + " FROM dat"
                                        'SQLQuery = SQLQuery + " WHERE (dat.datShfCode= " & ilShttCode & ""
                                        'SQLQuery = SQLQuery + " AND dat.datVefCode = " & ilVefCode
                                        SQLQuery = SQLQuery + " WHERE (datatfCode = " & llAttCode
                                        SQLQuery = SQLQuery & " AND datFdStTime = '" & Format$(slFdSTime, sgSQLTimeForm) & "')"
                                        Set rst = gSQLSelectCall(SQLQuery)
                                        If Not rst.EOF Then
                                            While Not rst.EOF
                                                ilUpper = UBound(tmPledgeInfo)
                                                tmPledgeInfo(ilUpper).ilDays(0) = Val(rst!datFdMon)
                                                tmPledgeInfo(ilUpper).ilDays(1) = Val(rst!datFdTue)
                                                tmPledgeInfo(ilUpper).ilDays(2) = Val(rst!datFdWed)
                                                tmPledgeInfo(ilUpper).ilDays(3) = Val(rst!datFdThu)
                                                tmPledgeInfo(ilUpper).ilDays(4) = Val(rst!datFdFri)
                                                tmPledgeInfo(ilUpper).ilDays(5) = Val(rst!datFdSat)
                                                tmPledgeInfo(ilUpper).ilDays(6) = Val(rst!datFdSun)
                                                tmPledgeInfo(ilUpper).iFdStatus = rst!datFdStatus
                                                If IsNull(rst!datPdStTime) Then
                                                    tmPledgeInfo(ilUpper).lPdSTime = 0
                                                Else
                                                    tmPledgeInfo(ilUpper).lPdSTime = gTimeToLong(rst!datPdStTime, False)
                                                End If
                                                tmPledgeInfo(ilUpper).lCode = rst!datCode
                                                ilUpper = ilUpper + 1
                                                ReDim Preserve tmPledgeInfo(0 To ilUpper) As PledgeInfo
                                                rst.MoveNext
                                            Wend
                                            For ilPd = 0 To UBound(tmPledgeInfo) - 1 Step 1
                                                For ilLoop = 0 To 6 Step 1
                                                    ilLDays(ilLoop) = tmPledgeInfo(ilPd).ilDays(ilLoop)
                                                Next ilLoop
                                                If tmPledgeInfo(ilPd).iFdStatus = 0 Then
                                                    'Remove Days
                                                    For ilLoop = 0 To 6 Step 1
                                                        If ilFdDays(ilLoop) = 1 Then
                                                            ilLDays(ilLoop) = 0
                                                        End If
                                                    Next ilLoop
                                                    ilUpdate = False
                                                    For ilLoop = 0 To 6 Step 1
                                                        If ilLDays(ilLoop) <> 0 Then
                                                            ilUpdate = True
                                                        End If
                                                    Next ilLoop
                                                    If ilUpdate Then
                                                        SQLQuery = "UPDATE dat"
                                                        SQLQuery = SQLQuery & " SET datFdMon = " & ilLDays(0) & ","
                                                        SQLQuery = SQLQuery & "datFdTue = " & ilLDays(1) & ","
                                                        SQLQuery = SQLQuery & "datFdWed = " & ilLDays(2) & ","
                                                        SQLQuery = SQLQuery & "datFdThu = " & ilLDays(3) & ","
                                                        SQLQuery = SQLQuery & "datFdFri = " & ilLDays(4) & ","
                                                        SQLQuery = SQLQuery & "datFdSat = " & ilLDays(5) & ","
                                                        SQLQuery = SQLQuery & "datFdSun = " & ilLDays(6) & ","
                                                        SQLQuery = SQLQuery & "datPdMon = " & ilLDays(0) & ","
                                                        SQLQuery = SQLQuery & "datPdTue = " & ilLDays(1) & ","
                                                        SQLQuery = SQLQuery & "datPdWed = " & ilLDays(2) & ","
                                                        SQLQuery = SQLQuery & "datPdThu = " & ilLDays(3) & ","
                                                        SQLQuery = SQLQuery & "datPdFri = " & ilLDays(4) & ","
                                                        SQLQuery = SQLQuery & "datPdSat = " & ilLDays(5) & ","
                                                        SQLQuery = SQLQuery & "datPdSun = " & ilLDays(6)
                                                        SQLQuery = SQLQuery & " WHERE (datCode = " & tmPledgeInfo(ilPd).lCode & ")"
                                                        cnn.BeginTrans
                                                        'cnn.Execute SQLQuery, rdExecDirect
                                                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                            '6/11/16: Replaced GoSub
                                                            'GoSub ErrHand:
                                                            Screen.MousePointer = vbDefault
                                                            gHandleError "AffErrorLog.txt", "ImportCSV-mReadFilePledge_Sv"
                                                            cnn.RollbackTrans
                                                            ilRet = 1
                                                        End If
                                                        If ilRet = 0 Then
                                                            cnn.CommitTrans
                                                        End If
                                                    Else
                                                        SQLQuery = "DELETE FROM dat WHERE (datCode = " & tmPledgeInfo(ilPd).lCode & ")"
                                                        'cnn.Execute SQLQuery, rdExecDirect
                                                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                            '6/11/16: Replaced GoSub
                                                            'GoSub ErrHand:
                                                            Screen.MousePointer = vbDefault
                                                            gHandleError "AffErrorLog.txt", "ImportCSV-mReadFilePledge_Sv"
                                                            ilRet = 1
                                                        End If
                                                    End If
                                                Else
                                                    If tmPledgeInfo(ilPd).lPdSTime = gTimeToLong(slPdSTime, False) Then
                                                        ilMerged = True
                                                        'Check if any day has been previously set on- then assume that this
                                                        'is a dulpicate type record
                                                        ilDupl = False
                                                        For ilLoop = 0 To 6 Step 1
                                                            If (ilFdDays(ilLoop) = 1) And (ilLDays(ilLoop) = 1) Then
                                                                ilDupl = True
                                                                Exit For
                                                            End If
                                                        Next ilLoop
                                                        If Not ilDupl Then
                                                            'Merge Days
                                                            SQLQuery = "UPDATE dat SET "
                                                            If ilFdDays(0) = 1 Then
                                                                SQLQuery = SQLQuery & "datFdMon = " & 1 & ","
                                                            End If
                                                            If ilFdDays(1) = 1 Then
                                                                SQLQuery = SQLQuery & "datFdTue = " & 1 & ","
                                                            End If
                                                            If ilFdDays(2) = 1 Then
                                                                SQLQuery = SQLQuery & "datFdWed = " & 1 & ","
                                                            End If
                                                            If ilFdDays(3) = 1 Then
                                                                SQLQuery = SQLQuery & "datFdThu = " & 1 & ","
                                                            End If
                                                            If ilFdDays(4) = 1 Then
                                                                SQLQuery = SQLQuery & "datFdFri = " & 1 & ","
                                                            End If
                                                            If ilFdDays(5) = 1 Then
                                                                SQLQuery = SQLQuery & "datFdSat = " & 1 & ","
                                                            End If
                                                            If ilFdDays(6) = 1 Then
                                                                SQLQuery = SQLQuery & "datFdSun = " & 1 & ","
                                                            End If
                                                            If ilPdDays(0) = 1 Then
                                                                SQLQuery = SQLQuery & "datPdMon = " & 1 & ","
                                                            End If
                                                            If ilPdDays(1) = 1 Then
                                                                SQLQuery = SQLQuery & "datPdTue = " & 1 & ","
                                                            End If
                                                            If ilPdDays(2) = 1 Then
                                                                SQLQuery = SQLQuery & "datPdWed = " & 1 & ","
                                                            End If
                                                            If ilPdDays(3) = 1 Then
                                                                SQLQuery = SQLQuery & "datPdThu = " & 1 & ","
                                                            End If
                                                            If ilPdDays(4) = 1 Then
                                                                SQLQuery = SQLQuery & "datPdFri = " & 1 & ","
                                                            End If
                                                            If ilPdDays(5) = 1 Then
                                                                SQLQuery = SQLQuery & "datPdSat = " & 1 & ","
                                                            End If
                                                            If ilPdDays(6) = 1 Then
                                                                SQLQuery = SQLQuery & "datPdSun = " & 1 & ","
                                                            End If
                                                            'Remove Comma
                                                            SQLQuery = Left$(SQLQuery, Len(SQLQuery) - 1)
                                                            SQLQuery = SQLQuery & " WHERE (datCode = " & tmPledgeInfo(ilPd).lCode & ")"
                                                            cnn.BeginTrans
                                                            'cnn.Execute SQLQuery, rdExecDirect
                                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                                '6/11/16: Replaced GoSub
                                                                'GoSub ErrHand:
                                                                Screen.MousePointer = vbDefault
                                                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFilePledge_Sv"
                                                                cnn.RollbackTrans
                                                                ilRet = 1
                                                            End If
                                                            If ilRet = 0 Then
                                                                cnn.CommitTrans
                                                            End If
                                                        Else
                                                            ''Don't add record
                                                            'If Not ilDuplMsg Then
                                                            '    lbcError.AddItem "Pledge Previously Defined: See Output Text File"
                                                            '    ilDuplMsg = True
                                                            'End If
                                                            'Print #hmMsg, "Pledge Previously Defined : " & Trim$(slLine)
                                                        End If
                                                    End If
                                                End If
                                            Next ilPd
                                        Else
                                            ilMerged = True
                                            ilFound = False
                                            For ilLoop = 0 To UBound(smMissingTime) - 1 Step 1
                                                'If StrComp(smFields(3), smMissingTime(ilLoop), 1) = 0 Then
                                                If StrComp(smFields(2), smMissingTime(ilLoop), 1) = 0 Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilLoop
                                            If Not ilFound Then
                                                'lbcError.AddItem "Avail Times Missing from: " & Trim$(smFields(1)) & ", " & Trim$(smFields(3))
                                                lbcError.AddItem "Avail Times Missing from: " & Trim$(smFields(0)) & ", " & Trim$(smFields(2))
                                                'smMissingTime(UBound(smMissingTime)) = smFields(3)
                                                smMissingTime(UBound(smMissingTime)) = smFields(2)
                                                ReDim Preserve smMissingTime(0 To UBound(smMissingTime) + 1) As String
                                            End If
                                            Print #hmMsg, "Avail Times Missing : " & Trim$(slLine)
                                        End If
                                        If Not ilMerged Then
                                            'Add to Agreement
                                            SQLQuery = "INSERT INTO dat (datAtfCode, datShfCode, datVefCode, "
                                            'SQLQuery = SQLQuery & "datDACode, datFdMon, datFdTue, "
                                            SQLQuery = SQLQuery & "datFdMon, datFdTue, "
                                            SQLQuery = SQLQuery & "datFdWed, datFdThu, datFdFri, "
                                            SQLQuery = SQLQuery & "datFdSat, datFdSun, datFdStTime, "
                                            SQLQuery = SQLQuery & "datFdEdTime, datFdStatus, datPdMon, "
                                            SQLQuery = SQLQuery & "datPdTue, datPdWed, datPdThu, "
                                            SQLQuery = SQLQuery & "datPdFri, datPdSat, datPdSun, "
                                            SQLQuery = SQLQuery & "datPdDayFed, datPdStTime, datPdEdTime)"
                                            SQLQuery = SQLQuery & " VALUES (" & llAttCode & "," & ilShttCode & "," & ilVefCode & ","
                                            'SQLQuery = SQLQuery & "1" & "," & Val(smFields(4)) & "," & Val(smFields(5)) & ","
                                            'SQLQuery = SQLQuery & Val(smFields(4)) & "," & Val(smFields(5)) & ","
                                            SQLQuery = SQLQuery & Val(smFields(3)) & "," & Val(smFields(4)) & ","
                                            'SQLQuery = SQLQuery & Val(smFields(6)) & "," & Val(smFields(7)) & "," & Val(smFields(8)) & ","
                                            SQLQuery = SQLQuery & Val(smFields(5)) & "," & Val(smFields(6)) & "," & Val(smFields(7)) & ","
                                            'SQLQuery = SQLQuery & Val(smFields(9)) & "," & Val(smFields(10)) & ",'" & Format$(slFdSTime, sgSQLTimeForm) & "',"
                                            SQLQuery = SQLQuery & Val(smFields(8)) & "," & Val(smFields(9)) & ",'" & Format$(slFdSTime, sgSQLTimeForm) & "',"
                                            'SQLQuery = SQLQuery & "'" & Format$(slFdETime, sgSQLTimeForm) & "'," & ilStatus & "," & Val(smFields(14)) & ","
                                            SQLQuery = SQLQuery & "'" & Format$(slFdETime, sgSQLTimeForm) & "'," & ilStatus & "," & Val(smFields(13)) & ","
                                            'SQLQuery = SQLQuery & Val(smFields(15)) & "," & Val(smFields(16)) & "," & Val(smFields(17)) & ","
                                            SQLQuery = SQLQuery & Val(smFields(14)) & "," & Val(smFields(15)) & "," & Val(smFields(16)) & ","
                                            'SQLQuery = SQLQuery & Val(smFields(18)) & "," & Val(smFields(19)) & "," & Val(smFields(20)) & ","
                                            SQLQuery = SQLQuery & Val(smFields(17)) & "," & Val(smFields(18)) & "," & Val(smFields(19)) & ","
                                            SQLQuery = SQLQuery & "'A', " & "'" & Format$(slPdSTime, sgSQLTimeForm) & "','" & Format$(slPdETime, sgSQLTimeForm) & "')"
                                            cnn.BeginTrans
                                            'cnn.Execute SQLQuery, rdExecDirect
                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                '6/11/16: Replaced GoSub
                                                'GoSub ErrHand:
                                                Screen.MousePointer = vbDefault
                                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFilePledge_Sv"
                                                cnn.RollbackTrans
                                                ilRet = 1
                                            End If
                                            If ilRet = 0 Then
                                                cnn.CommitTrans
                                            End If
                                        End If
                                    End If
                                ElseIf ilFoundType = 3 Then
                                    If Not ilOverlapMsg Then
                                        lbcError.AddItem "Agreement(s) Overlap: See Output Text File"
                                        ilOverlapMsg = True
                                    End If
                                    'Print #hmMsg, "Agreement Overlap: " & Trim$(smFields(1)) & " " & Trim$(smFields(3)) & " " & Format$(llEffSDate, "m/d/yyyy") & "-" & Format$(llEffEDate, "m/d/yyyy")
                                    Print #hmMsg, "Agreement Overlap: " & Trim$(smFields(0)) & " " & Trim$(smFields(2)) & " " & Format$(llEffSDate, "m/d/yyyy") & "-" & Format$(llEffEDate, "m/d/yyyy")
                                ElseIf ilFoundType = 0 Then
                                    ilFound = False
                                    For ilLoop = 0 To UBound(tmPledgeCount) - 1 Step 1
                                        If (ilShttCode = tmPledgeCount(ilLoop).iShttCode) And (ilVefCode = tmPledgeCount(ilLoop).iVefCode) Then
                                            ilFound = True
                                            tmPledgeCount(ilLoop).lCount = tmPledgeCount(ilLoop).lCount + 1
                                            Exit For
                                        End If
                                    Next ilLoop
                                    If Not ilFound Then
                                        'lbcError.AddItem "Agreement Pledged Ignored: " & Trim$(smFields(1)) & " " & Trim$(smFields(3)) & " Count in Output Text File"
                                        lbcError.AddItem "Agreement Pledged Ignored: " & Trim$(smFields(0)) & " " & Trim$(smFields(2)) & " Count in Output Text File"
                                        tmPledgeCount(UBound(tmPledgeCount)).iShttCode = ilShttCode
                                        tmPledgeCount(UBound(tmPledgeCount)).iVefCode = ilVefCode
                                        tmPledgeCount(UBound(tmPledgeCount)).lCount = 1
                                        ReDim Preserve tmPledgeCount(0 To UBound(tmPledgeCount) + 1) As MISSINGATT
                                    End If
                                End If
                            End If
                        Next llAtt
                        If Not llAttFound Then
                            'If Not llAttMissingMsg Then
                            '    lbcError.AddItem "Agreement(s) Missing: See Output Text File"
                            '    llAttMissingMsg = True
                            'End If
                            'Print #hmMsg, "Agreement Missing: " & Trim$(smFields(1)) & " " & Trim$(smFields(3))
                            'Print #hmMsg, "Agreement Missing: " & Trim$(smFields(0)) & " " & Trim$(smFields(2))
                            ilFound = False
                            For ilLoop = 0 To UBound(tmMissingAtt) - 1 Step 1
                                If (ilShttCode = tmMissingAtt(ilLoop).iShttCode) And (ilVefCode = tmMissingAtt(ilLoop).iVefCode) Then
                                    ilFound = True
                                    tmMissingAtt(ilLoop).lCount = tmMissingAtt(ilLoop).lCount + 1
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                'lbcError.AddItem "Agreement Missing: " & Trim$(smFields(1)) & " " & Trim$(smFields(3))
                                lbcError.AddItem "Agreement Missing: " & Trim$(smFields(0)) & " " & Trim$(smFields(2))
                                tmMissingAtt(UBound(tmMissingAtt)).iShttCode = ilShttCode
                                tmMissingAtt(UBound(tmMissingAtt)).iVefCode = ilVefCode
                                tmMissingAtt(UBound(tmMissingAtt)).lCount = 1
                                ReDim Preserve tmMissingAtt(0 To UBound(tmMissingAtt) + 1) As MISSINGATT
                            End If
                        End If
                    Else
                        If ilVefCode < 0 Then
                            ilFound = False
                            For ilLoop = 0 To UBound(smMissingVef) - 1 Step 1
                                'If StrComp(smFields(3), smMissingVef(ilLoop), 1) = 0 Then
                                If StrComp(smFields(2), smMissingVef(ilLoop), 1) = 0 Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                'lbcError.AddItem "Vehicle Missing: " & Trim$(smFields(3))
                                lbcError.AddItem "Vehicle Missing: " & Trim$(smFields(2))
                                'Print #hmMsg, "Vehicle Missing: " & Trim$(smFields(3))
                                Print #hmMsg, "Vehicle Missing: " & Trim$(smFields(2))
                                'smMissingVef(UBound(smMissingVef)) = smFields(3)
                                smMissingVef(UBound(smMissingVef)) = smFields(2)
                                ReDim Preserve smMissingVef(0 To UBound(smMissingVef) + 1) As String
                            End If
                        End If
                    End If
                Else
                    ilFound = False
                    For ilLoop = 0 To UBound(smMissingShtt) - 1 Step 1
                        'If StrComp(smFields(1), smMissingShtt(ilLoop), 1) = 0 Then
                        If StrComp(smFields(0), smMissingShtt(ilLoop), 1) = 0 Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        'lbcError.AddItem "Station Missing: " & Trim$(smFields(1))
                        lbcError.AddItem "Station Missing: " & Trim$(smFields(0))
                        'Print #hmMsg, "Station Missing: " & Trim$(smFields(1))
                        Print #hmMsg, "Station Missing: " & Trim$(smFields(0))
                        'smMissingShtt(UBound(smMissingShtt)) = smFields(1)
                        smMissingShtt(UBound(smMissingShtt)) = smFields(0)
                        ReDim Preserve smMissingShtt(0 To UBound(smMissingShtt) + 1) As String
                    End If
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
            llPercent = (lmProcessedNoBytes / lmTotalNoBytes) * 100
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmFloodPercent <> llPercent Then
                lmFloodPercent = llPercent
                lbcPercent.Caption = Str$(llPercent) & "%"
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    For ilPledge = 0 To UBound(tmPledgeCount) - 1 Step 1
        slCallLetters = ""
        For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(ilLoop).iCode = tmPledgeCount(ilPledge).iShttCode Then
                slCallLetters = Trim$(tgStationInfo(ilLoop).sCallLetters)
                Exit For
            End If
        Next ilLoop
        slVehicle = ""
        For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            If tgVehicleInfo(ilLoop).iCode = tmPledgeCount(ilPledge).iVefCode Then
                slVehicle = Trim$(tgVehicleInfo(ilLoop).sVehicle)
                Exit For
            End If
        Next ilLoop
        Print #hmMsg, "Agreement Pledged Ignored: " & slCallLetters & " " & slVehicle & " " & tmPledgeCount(ilPledge).lCount
    Next ilPledge
    For ilPledge = 0 To UBound(tmMissingAtt) - 1 Step 1
        slCallLetters = ""
        For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(ilLoop).iCode = tmMissingAtt(ilPledge).iShttCode Then
                slCallLetters = Trim$(tgStationInfo(ilLoop).sCallLetters)
                Exit For
            End If
        Next ilLoop
        slVehicle = ""
        For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            If tgVehicleInfo(ilLoop).iCode = tmMissingAtt(ilPledge).iVefCode Then
                slVehicle = Trim$(tgVehicleInfo(ilLoop).sVehicle)
                Exit For
            End If
        Next ilLoop
        Print #hmMsg, "Agreement Missing: " & slCallLetters & " " & slVehicle & " " & tmMissingAtt(ilPledge).lCount
    Next ilPledge
    If ilRet <> 0 Then
        mReadFilePledge_Sv = False
    Else
        mReadFilePledge_Sv = True
        lmFloodPercent = 100
        lbcPercent.Caption = "100%"
    End If
    Print #hmMsg, "** Import Agreement Pledges Finished: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    Exit Function
mReadFilePledge_SvErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportCSV-mReadFilePledge_Sv"
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFilePledge                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Private Function mReadFilePledge(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilLoop As Integer
    Dim slCurDate As String
    Dim slCurTime As String
    Dim llPercent As Long
    Dim ilOverlapMsg As Integer
    Dim llExcludeDate As Long
    Dim llTFNDate As Long
    Dim ilSave As Integer
    Dim slChar As String
    Dim llPrevAttCode As Long
    Dim slAttCode As String
    Dim llLineCount As Long
    Dim llStartLineCount As Long

        
    slCurDate = Format(gNow(), sgShowDateForm)
    slCurTime = Format(gNow(), sgShowTimeWSecForm)
    ilOverlapMsg = False
    llExcludeDate = DateValue("1/1/2000")
    llTFNDate = DateValue("12/31/1999")
    ReDim smMissingVef(0 To 0) As String
    ReDim smMissingShtt(0 To 0) As String
    ReDim smMissingTime(0 To 0) As String
    ReDim smStationNotMatching(0 To 0) As String
    ReDim smVehicleNotMatching(0 To 0) As String
    ReDim smStationNameError(0 To 0) As String
    ReDim smVehicleNameError(0 To 0) As String
    
    ReDim tmAgreeID(0 To 0) As AGREEID
    ReDim tmPledgeCount(0 To 0) As MISSINGATT
    ReDim tmMissingAtt(0 To 0) As MISSINGATT
    ReDim tmPledgeInfo(0 To 0) As PledgeInfo
    ReDim slAttPledgeLines(0 To 0) As String
    ReDim llLineNo(0 To 0) As Long
    llLineCount = 0
    ilRet = 0
    On Error GoTo ErrHand
    'ilRet = 0
    'On Error GoTo mReadFilePledgeErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        lbcError.AddItem "Open " & slFromFile & " error#" & Str$(ilRet)
        Print #hmMsg, "Open " & slFromFile & " error#" & Str$(ilRet)
        Close hmFrom
        mReadFilePledge = False
        Exit Function
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    lmProcessedNoBytes = 0
    llPrevAttCode = -1
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mReadFilePledgeErr:
        'Line Input #hmFrom, slLine
        slLine = ""
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                Exit Do
            ElseIf slChar <> sgCR Then
                slLine = slLine & slChar
            End If
        Loop
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Print #hmMsg, "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Close hmFrom
            mReadFilePledge = False
            Exit Function
        End If
        llLineCount = llLineCount + 1
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                If (UCase(Left(slLine, 4)) <> "CALL") Or (lmProcessedNoBytes <> 0) Then
                    'ilRet = gParseItem(slLine, 5, ",", slAttCode)
                    gParseCDFields slLine, False, smFields()
                    'slAttCode = smFields(5)
                    slAttCode = smFields(4)
                    If Val(slAttCode) <> llPrevAttCode Then
                        If llPrevAttCode <> -1 Then
                            mProcessPledge slAttPledgeLines, llLineNo
                            ReDim slAttPledgeLines(0 To 0) As String
                            ReDim llLineNo(0 To 0) As Long
                        End If
                        llPrevAttCode = Val(slAttCode)
                        llStartLineCount = llLineCount
                    End If
                    ilSave = True
                    For ilLoop = 0 To UBound(slAttPledgeLines) - 1 Step 1
                        If StrComp(Trim$(slAttPledgeLines(ilLoop)), Trim$(slLine), vbBinaryCompare) = 0 Then
                            ilSave = False
                            Print #hmMsg, "Line, " & llLineCount & "," & "matches line," & llLineNo(ilLoop) & ",Line bypassed"
                            Exit For
                        End If
                    Next ilLoop
                    If ilSave Then
                        slAttPledgeLines(UBound(slAttPledgeLines)) = slLine
                        llLineNo(UBound(llLineNo)) = llLineCount
                        ReDim Preserve slAttPledgeLines(0 To UBound(slAttPledgeLines) + 1) As String
                        ReDim Preserve llLineNo(0 To UBound(llLineNo) + 1) As Long
                    End If
                End If
            End If
        End If
        lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
        llPercent = (lmProcessedNoBytes / lmTotalNoBytes) * 100
        If llPercent >= 100 Then
            If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                llPercent = 99
            Else
                llPercent = 100
            End If
        End If
        If lmFloodPercent <> llPercent Then
            lmFloodPercent = llPercent
            lbcPercent.Caption = Str$(llPercent) & "%"
        End If
        ilRet = 0
    Loop
    If UBound(slAttPledgeLines) > 0 Then
        mProcessPledge slAttPledgeLines, llLineNo
    End If
    Close hmFrom
    If ilRet <> 0 Then
        mReadFilePledge = False
    Else
        mReadFilePledge = True
        lmFloodPercent = 100
        lbcPercent.Caption = "100%"
    End If
    Print #hmMsg, "** Import Agreement Pledges Finished: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    Exit Function
mReadFilePledgeErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    gHandleError "ImptPldg.Txt", "Import-mReadFilePledge"
    ilRet = 1
    Resume Next
ErrHand1:
    gHandleError "ImptPldg.Txt", "Import-mReadFilePledge"
    Return
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFileAffiliateSpots         *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Private Function mReadFileAffiliateSpots(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim ilMaxCol As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilMatch As Integer
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilCode As Integer
    Dim slCallLetters As String
    Dim ilShttCode As Integer
    Dim slVehicle As String
    Dim ilVefCode As Integer
    Dim llAttCode As Long
    Dim llLstCode As Long
    Dim ilStatus As Integer
    Dim ilSelected As Integer
    Dim slCurDate As String
    Dim slCurTime As String
    Dim ilPos As Integer
    Dim slMsg As String
    Dim llPercent As Long
    Dim ilUpper As Integer
    Dim llTstDate As Long
    Dim slName As String
    Dim ilAdfCode As Integer
    Dim slCntrNo As String
    Dim llCntrNo As Long
    Dim slProd As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim slFdDate As String
    Dim slFdTime As String
    Dim slPdDate As String
    Dim slPdSTime As String
    Dim slPdETime As String
    Dim llAttMissingMsg As Long
    Dim ilLstMissingMsg As Integer
    Dim ilAirDateMsg As Integer
    Dim slStr As String
    Dim llStationID As Long
    Dim ilPledgeStatus As Integer
    Dim slZone As String
    Dim ilZone As Integer
    Dim ilLocalAdj As Integer
    Dim slLstDate As String
    Dim slLstTime As String
    Dim llLstDate As Long
    Dim llLstTime As Long
    Dim ilLstVefCode As Integer
    Dim ilLstIndex As Integer
    Dim ilFirstSpotTest As Integer
    Dim llLogSpotDate As Long
    Dim slChar As String
    Dim llDATCode As Long
    Dim llCpfCode As Long
    Dim llRsfCode As Long
    Dim slStationCompliant As String
    Dim slAgencyCompliant As String
    Dim slAffidavitSource As String
    Dim rstDat As ADODB.Recordset
    Dim ilLen As Integer
    Dim ilVef As Integer
        
    slCurDate = Format(gNow(), sgShowDateForm)
    slCurTime = Format(gNow(), sgShowTimeWSecForm)
    llAttMissingMsg = False
    ilLstMissingMsg = False
    ilAirDateMsg = False
    ilFirstSpotTest = False
    llLogSpotDate = -1
    llDATCode = 0
    llCpfCode = 0
    llRsfCode = 0
    slStationCompliant = ""
    slAgencyCompliant = ""
    slAffidavitSource = ""

    ReDim smMissingVef(0 To 0) As String
    ReDim smMissingShtt(0 To 0) As String
    ilUpper = 0
    ReDim tmAgreeID(0 To 0) As AGREEID
    ReDim tmMissingAtt(0 To 0) As MISSINGATT
    ilRet = 0
    On Error GoTo ErrHand
    SQLQuery = "SELECT attShfCode, attVefCode, attOnAir, attOffAir, attDropDate, attCode"
    SQLQuery = SQLQuery + " FROM att"
    Set rst = gSQLSelectCall(SQLQuery)
    If ilRet <> 0 Then
        mReadFileAffiliateSpots = False
        Exit Function
    End If
    While Not rst.EOF
        ilUpper = UBound(tmAgreeID)
        tmAgreeID(ilUpper).lCode = rst!attCode
        tmAgreeID(ilUpper).iShttCode = rst!attshfCode
        tmAgreeID(ilUpper).iVefCode = rst!attvefCode
        tmAgreeID(ilUpper).lOnAir = DateValue(gAdjYear(rst!attOnAir))
        tmAgreeID(ilUpper).lOffAir = DateValue(gAdjYear(rst!attOffAir))
        tmAgreeID(ilUpper).lDropDate = DateValue(gAdjYear(rst!attDropDate))
        If tmAgreeID(ilUpper).lDropDate < tmAgreeID(ilUpper).lOffAir Then
            tmAgreeID(ilUpper).lEndDate = tmAgreeID(ilUpper).lDropDate
        Else
            tmAgreeID(ilUpper).lEndDate = tmAgreeID(ilUpper).lOffAir
        End If
        ilUpper = ilUpper + 1
        ReDim Preserve tmAgreeID(0 To ilUpper) As AGREEID
        rst.MoveNext
    Wend

    'ilRet = 0
    'On Error GoTo mReadFileAffiliateSpotsErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        lbcError.AddItem "Open " & slFromFile & " error#" & Str$(ilRet)
        Print #hmMsg, "Open " & slFromFile & " error#" & Str$(ilRet)
        Close hmFrom
        mReadFileAffiliateSpots = False
        Exit Function
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    lmProcessedNoBytes = 0
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mReadFileAffiliateSpotsErr:
        Line Input #hmFrom, slLine
        'slLine = ""
        'Do While Not EOF(hmFrom)
        '    slChar = Input(1, #hmFrom)
        '    If slChar = sgLF Then
        '        Exit Do
        '    ElseIf slChar <> sgCR Then
        '        slLine = slLine & slChar
        '    End If
        'Loop
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Print #hmMsg, "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Close hmFrom
            mReadFileAffiliateSpots = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, True, smFields()
                For ilLoop = LBound(smFields) To UBound(smFields) Step 1
                    smFields(ilLoop) = Trim$(smFields(ilLoop))
                Next ilLoop
                'smFields(1) = UCase$(Trim$(smFields(1)))
                smFields(0) = UCase$(Trim$(smFields(0)))
                'slCallLetters = smFields(1)
                slCallLetters = smFields(0)
                ilShttCode = -1
                If Len(slCallLetters) > 40 Then
                    lbcError.AddItem slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                    Print #hmMsg, slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                    slCallLetters = Left$(slCallLetters, 40)
                End If
                'llStationID = Val(smFields(2))
                llStationID = Val(smFields(1))
                For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                    'If StrComp(Trim$(tgStationInfo(ilLoop).sCallLetters), slCallLetters, vbTextCompare) = 0 Then
                    If tgStationInfo(ilLoop).lID = llStationID Then
                        ilShttCode = tgStationInfo(ilLoop).iCode
                        slZone = Trim$(tgStationInfo(ilLoop).sZone)
                        Exit For
                    End If
                Next ilLoop
                
                If ilShttCode > 0 Then
                    'slVehicle = Trim$(smFields(3))
                    slVehicle = Trim$(smFields(2))
                    ilVefCode = -1
                    ilLocalAdj = 0
                    ilSelected = False
                    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                        If sgShowByVehType = "Y" Then
                            smVehName = Mid$(tgVehicleInfo(ilLoop).sVehicle, 3)
                        Else
                            smVehName = tgVehicleInfo(ilLoop).sVehicle
                        End If
                        'If StrComp(Trim$(smVehName), smFields(3), vbTextCompare) = 0 Then
                        If StrComp(Trim$(smVehName), smFields(2), vbTextCompare) = 0 Then
                            ilVefCode = tgVehicleInfo(ilLoop).iCode
                            For ilZone = LBound(tgVehicleInfo(ilLoop).sZone) To UBound(tgVehicleInfo(ilLoop).sZone) Step 1
                                If tgVehicleInfo(ilLoop).sZone(ilZone) = slZone Then
                                    ilLocalAdj = -tgVehicleInfo(ilLoop).iLocalAdj(ilZone)
                                    Exit For
                                End If
                            Next ilZone
                            'If lbcNames.Selected(ilLoop) Then
                            '    ilSelected = True
                            'End If
                            'Exit For
                            ilFound = False
                            For ilVef = 0 To lbcNames.ListCount - 1 Step 1
                                If lbcNames.ItemData(ilVef) = ilVefCode Then
                                    If lbcNames.Selected(ilLoop) Then
                                        ilSelected = True
                                    End If
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilVef
                            If ilFound Then
                                Exit For
                            End If
                        End If
                    Next ilLoop
                    If (ilVefCode > 0) And (ilSelected) Then
                        llAttCode = -1
                        For ilLoop = LBound(tmAgreeID) To UBound(tmAgreeID) - 1 Step 1
                            If (tmAgreeID(ilLoop).iShttCode = ilShttCode) And (tmAgreeID(ilLoop).iVefCode = ilVefCode) Then
                                'If smFields(10) <> "" Then
                                If smFields(9) <> "" Then
                                    'llTstDate = DateValue(gAdjYear(smFields(10)))
                                    llTstDate = DateValue(gAdjYear(smFields(9)))
                                Else
                                    'llTstDate = DateValue(gAdjYear(smFields(7)))
                                    llTstDate = DateValue(gAdjYear(smFields(6)))
                                End If
                                If (llTstDate >= tmAgreeID(ilLoop).lOnAir) And (llTstDate <= tmAgreeID(ilLoop).lEndDate) Then
                                    llAttCode = tmAgreeID(ilLoop).lCode
                                    Exit For
                                End If
                            End If
                        Next ilLoop
                        'slStr = smFields(4)
                        slStr = smFields(3)
                        ilPos = InStr(slStr, "-")
                        If ilPos > 0 Then
                            slCntrNo = Left$(slStr, ilPos - 1) & Mid$(slStr, ilPos + 1)
                        Else
                            slCntrNo = slStr
                        End If
                        llCntrNo = Val(slCntrNo)
                        If llAttCode >= 0 Then
                            'slName = smFields(5)
                            slName = smFields(4)
                            ilAdfCode = -1
                            For ilLoop = 0 To UBound(tgAdvtInfo) - 1 Step 1
                                'If StrComp(Trim$(tgAdvtInfo(ilLoop).sAdvtName), smFields(3), vbTextCompare) = 0 Then
                                If StrComp(Trim$(tgAdvtInfo(ilLoop).sAdvtName), smFields(2), vbTextCompare) = 0 Then
                                    ilAdfCode = tgAdvtInfo(ilLoop).iCode
                                    Exit For
                                End If
                            Next ilLoop
                            'slProd = gFixQuote(smFields(6))
                            slProd = gFixQuote(smFields(5))
                            'If (Len(Trim$(smFields(7))) > 0) Or (Len(Trim$(smFields(10))) > 0) Then
                            If (Len(Trim$(smFields(6))) > 0) Or (Len(Trim$(smFields(9))) > 0) Then
                                'If (Len(Trim$(smFields(7))) > 0) Then
                                If (Len(Trim$(smFields(6))) > 0) Then
                                    'slAirDate = Format$(smFields(7), sgShowDateForm)
                                    slAirDate = Format$(smFields(6), sgShowDateForm)
                                Else
                                    'slAirDate = Format$(smFields(10), sgShowDateForm)
                                    slAirDate = Format$(smFields(9), sgShowDateForm)
                                End If
                                'If (Len(Trim$(smFields(8))) > 0) Then
                                If (Len(Trim$(smFields(9))) > 0) Then
                                    'slAirTime = Format$(smFields(8), sgShowTimeWSecForm)
                                    slAirTime = Format$(smFields(7), sgShowTimeWSecForm)
                                Else
                                    'slAirTime = Format$(smFields(11), sgShowTimeWSecForm)
                                    slAirTime = Format$(smFields(10), sgShowTimeWSecForm)
                                End If
                                If (DateValue(gAdjYear(slAirDate)) >= lmStartDate) And (DateValue(gAdjYear(slAirDate)) <= lmLastDate) Then
                                    'ilStatus = Val(smFields(9))
                                    ilStatus = Val(smFields(8))
                                    If ilStatus > 0 Then
                                        ilStatus = ilStatus - 1
                                    End If
                                    If ilStatus = 6 Then 'Change status 7 to 2
                                        ilStatus = 1
                                    End If
                                    'If ilStatus > 0 Then
                                        'slFdDate = Format$(smFields(10), sgShowDateForm)
                                        slFdDate = Format$(smFields(9), sgShowDateForm)
                                        'Truncate seconds
                                        'slFdTime = Format$(Format$(smFields(11), "h:mmam/pm"), "hh:mm:ss")
                                        slFdTime = Format$(Format$(smFields(10), "h:mmam/pm"), "hh:mm:ss")
                                        'If smFields(12) <> "" Then
                                        If smFields(11) <> "" Then
                                            'slPdDate = Format$(smFields(12), sgShowDateForm)
                                            slPdDate = Format$(smFields(11), sgShowDateForm)
                                        Else
                                            slPdDate = slFdDate
                                        End If
                                        'If smFields(13) <> "" Then
                                        If smFields(12) <> "" Then
                                            'If gTimeToLong(Format$(smFields(11), "h:mm:ssam/pm"), False) = gTimeToLong(Format$(smFields(13), "h:mm:ssam/pm"), False) Then
                                            If gTimeToLong(Format$(smFields(10), "h:mm:ssam/pm"), False) = gTimeToLong(Format$(smFields(12), "h:mm:ssam/pm"), False) Then
                                                slPdSTime = slFdTime
                                            Else
                                                'slPdSTime = Format$(smFields(13), "hh:mm:ss")
                                                slPdSTime = Format$(smFields(12), "hh:mm:ss")
                                            End If
                                        Else
                                            slPdSTime = slFdTime
                                        End If
                                        slPdETime = Format$(gTimeToLong(slPdSTime, False) + 60, "hh:mm:ss")
                                        'If gTimeToLong(Format$(slFdTime, sgShowTimeWSecForm), False) = gTimeToLong(Format$(slPdSTime, sgShowTimeWSecForm), False) Then
                                            ilPledgeStatus = 0
                                        'Else
                                        '    ilPledgeStatus = 2
                                        'End If
                                        'Find Lst
                                        llLstDate = DateValue(gAdjYear(slFdDate))
                                        llLstTime = gTimeToLong(slFdTime, False)
                                        
                                        'SQLQuery = "SELECT datDACode FROM dat"
                                        'SQLQuery = SQLQuery + " WHERE (datShfCode = " & ilShttCode & ")"
                                        'Set rstDat = gSQLSelectCall(SQLQuery)
                                        
                                        'If rstDat!datDACode <> 2 Then
                                            llLstTime = llLstTime + 3600 * ilLocalAdj
                                            If llLstTime < 0 Then
                                                llLstTime = llLstTime + 86400
                                                llLstDate = llLstDate - 1
                                            ElseIf llLstTime > 86400 Then
                                                llLstTime = llLstTime - 86400
                                                llLstDate = llLstDate + 1
                                            End If
                                        'End If
                                         slLstTime = Format$(gLongToTime(llLstTime), sgShowTimeWSecForm)
                                        slLstDate = Format$(llLstDate, sgShowDateForm)
                                        
                                        On Error GoTo ErrHand
                                        'SQLQuery = "SELECT lst.lstCode FROM lst "
                                        'SQLQuery = SQLQuery + " WHERE (lst.lstLogVefCode = " & ilVefCode
                                        'SQLQuery = SQLQuery + " AND lst.lstLogDate = '" & slLstDate & "'"
                                        'SQLQuery = SQLQuery & " AND lstCntrNo = " & llCntrNo
                                        'SQLQuery = SQLQuery & " AND lstLogTime = " & slLstTime & ")"
                                        'Set rst = gSQLSelectCall(SQLQuery)
                                        'If Not rst.EOF Then
                                        '    llLstCode = rst!lstCode
                                        'End If
                                        If llLogSpotDate <> llLstDate Then
                                            llLogSpotDate = llLstDate
                                            ReDim tmLogSpotInfo(0 To 0) As LOGSPOTINFO
                                            SQLQuery = "SELECT lstLogVefCode, lstLogTime, lstCntrNo, lstCode FROM lst "
                                            SQLQuery = SQLQuery + " WHERE (lstLogDate = '" & Format$(slLstDate, sgSQLDateForm) & "')"
                                            SQLQuery = SQLQuery & " ORDER BY lstLogVefCode"
                                            Set rst = gSQLSelectCall(SQLQuery)
                                            While Not rst.EOF
                                                ilUpper = UBound(tmLogSpotInfo)
                                                tmLogSpotInfo(ilUpper).lCode = rst!lstCode
                                                tmLogSpotInfo(ilUpper).iVefCode = rst!lstLogVefCode
                                                tmLogSpotInfo(ilUpper).lLogTime = gTimeToLong(rst!lstLogTime, False)
                                                tmLogSpotInfo(ilUpper).lCntrNo = rst!lstCntrNo
                                                ReDim Preserve tmLogSpotInfo(0 To ilUpper + 1) As LOGSPOTINFO
                                                rst.MoveNext
                                            Wend
                                            ilLstVefCode = -1
                                            ilLstIndex = 0
                                        End If
                                        llLstCode = -1
                                        If ilVefCode <> ilLstVefCode Then
                                            ilLstIndex = 0
                                        End If
                                        For ilLoop = ilLstIndex To UBound(tmLogSpotInfo) - 1 Step 1
                                            If (ilVefCode <> ilLstVefCode) And (ilVefCode = tmLogSpotInfo(ilLoop).iVefCode) Then
                                                ilLstIndex = ilLoop
                                                ilLstVefCode = ilVefCode
                                            End If
                                            If (ilVefCode = tmLogSpotInfo(ilLoop).iVefCode) And (llCntrNo = tmLogSpotInfo(ilLoop).lCntrNo) And (llLstTime = tmLogSpotInfo(ilLoop).lLogTime) Then
                                                llLstCode = tmLogSpotInfo(ilLoop).lCode
                                                Exit For
                                            End If
                                        Next ilLoop
                                        'Insert into database
                                        ilRet = 0
                                        If llLstCode > 0 Then
                                            If Not ilFirstSpotTest Then
                                                ilFirstSpotTest = True
                                                '********************************************
                                                'Remove test- Jim
                                                '
                                                'SQLQuery = "SELECT ast.astCode FROM ast WHERE (ast.astatfCode = " & llAttCode & " AND ast.astFeedDate = '" & slFdDate & "  & " ')"
                                                'Set rst = gSQLSelectCall(SQLQuery)
                                                'If Not rst.EOF Then
                                                '    gMsgBox "Affiliate Spots Previously Created, Terminating Import", vbCritical + vbOKOnly
                                                '    mReadFileAffiliateSpots = False
                                                '    Exit Function
                                                'End If
                                                '**********************************************
                                            End If
                                            ilLen = 0
                                            SQLQuery = "INSERT INTO ast (astAtfCode, astShfCode, astVefCode, "
                                            SQLQuery = SQLQuery & "astSdfCode, astLsfCode, astAirDate, "
                                            SQLQuery = SQLQuery & "astAirTime, astStatus, astCPStatus, "
                                            '12/13/13: New AST structure
                                            'SQLQuery = SQLQuery & "astFeedDate, astFeedTime, astPledgeDate, "
                                            'SQLQuery = SQLQuery & "astPledgeStartTime, astPledgeEndTime, astPledgeStatus )"
                                            SQLQuery = SQLQuery & "astFeedDate, astFeedTime, "
                                            SQLQuery = SQLQuery + "astAdfCode, astDatCode, astCpfCode, astRsfCode, astStationCompliant, astAgencyCompliant, astAffidavitSource, astCntrNo, astLen, astLkAstCode, astMissedMnfCode, astUstCode)"
                                            SQLQuery = SQLQuery & " VALUES (" & llAttCode & "," & ilShttCode & "," & ilVefCode & ","
                                            SQLQuery = SQLQuery & "0" & "," & llLstCode & ",'" & Format$(slAirDate, sgSQLDateForm) & "',"
                                            SQLQuery = SQLQuery & "'" & Format$(slAirTime, sgSQLTimeForm) & "'," & ilStatus & "," & "0" & ","
                                            'SQLQuery = SQLQuery & "'" & Format$(slFdDate, sgSQLDateForm) & "','" & Format$(slFdTime, sgSQLTimeForm) & "','" & Format$(slPdDate, sgSQLDateForm) & "',"
                                            'SQLQuery = SQLQuery & "'" & Format$(slPdSTime, sgSQLTimeForm) & "','" & Format$(slPdETime, sgSQLTimeForm) & "'," & ilPledgeStatus & ")"
                                            SQLQuery = SQLQuery & "'" & Format$(slFdDate, sgSQLDateForm) & "','" & Format$(slFdTime, sgSQLTimeForm) & "',"
                                            SQLQuery = SQLQuery & ilAdfCode & ", " & llDATCode & ", " & llCpfCode & ", " & llRsfCode & ", "
                                            SQLQuery = SQLQuery & "'" & slStationCompliant & "', '" & slAgencyCompliant & "', '" & slAffidavitSource & "', " & llCntrNo & ", " & ilLen & ", " & 0 & ", " & 0 & ", " & igUstCode & ")"
                                            cnn.BeginTrans
                                            'cnn.Execute SQLQuery, rdExecDirect
                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                '6/11/16: Replaced GoSub
                                                'GoSub ErrHand:
                                                Screen.MousePointer = vbDefault
                                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileAffiliateSpots"
                                                cnn.RollbackTrans
                                                ilRet = 1
                                            End If
                                            If ilRet = 0 Then
                                                cnn.CommitTrans
                                            End If
                                        Else
                                            If Not ilLstMissingMsg Then
                                                lbcError.AddItem "Log Spot(s) Missing: See Output Text File"
                                                ilLstMissingMsg = True
                                            End If
                                            'Print #hmMsg, "Log Spot Missing: " & Trim$(smFields(1)) & " " & Trim$(smFields(3)) & Str$(llCntrNo) & " " & smFields(10) & " " & smFields(11)
                                            Print #hmMsg, "Log Spot Missing: " & Trim$(smFields(0)) & " " & Trim$(smFields(2)) & Str$(llCntrNo) & " " & smFields(9) & " " & smFields(10)
                                        End If
                                    'End If
                                End If
                            Else
                                If Not ilAirDateMsg Then
                                    lbcError.AddItem "Air Date Missing: See Output Text File"
                                    ilAirDateMsg = True
                                End If
                                'Print #hmMsg, "Air Date Missing: " & Trim$(smFields(1)) & " " & Trim$(smFields(3)) & Str$(llCntrNo) & " " & smFields(10) & " " & smFields(11)
                                Print #hmMsg, "Air Date Missing: " & Trim$(smFields(0)) & " " & Trim$(smFields(2)) & Str$(llCntrNo) & " " & smFields(9) & " " & smFields(10)
                            End If
                        Else
                            'If Not llAttMissingMsg Then
                            '    lbcError.AddItem "Agreement(s) Missing: See Output Text File"
                            '    llAttMissingMsg = True
                            'End If
                            'Print #hmMsg, "Agreement Missing: " & Trim$(smFields(1)) & " " & Trim$(smFields(3)) & Str$(llCntrNo) & " " & smFields(10) & " " & smFields(11)
                            'Print #hmMsg, "Agreement Missing: " & Trim$(smFields(0)) & " " & Trim$(smFields(2)) & Str$(llCntrNo) & " " & smFields(9) & " " & smFields(10)
                            ilFound = False
                            For ilLoop = 0 To UBound(tmMissingAtt) - 1 Step 1
                                If (ilShttCode = tmMissingAtt(ilLoop).iShttCode) And (ilVefCode = tmMissingAtt(ilLoop).iVefCode) Then
                                    ilFound = True
                                    tmMissingAtt(ilLoop).lCount = tmMissingAtt(ilLoop).lCount + 1
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                'lbcError.AddItem "Agreement Missing: " & Trim$(smFields(1)) & " " & Trim$(smFields(3))
                                lbcError.AddItem "Agreement Missing: " & Trim$(smFields(0)) & " " & Trim$(smFields(2))
                                tmMissingAtt(UBound(tmMissingAtt)).iShttCode = ilShttCode
                                tmMissingAtt(UBound(tmMissingAtt)).iVefCode = ilVefCode
                                tmMissingAtt(UBound(tmMissingAtt)).lCount = 1
                                ReDim Preserve tmMissingAtt(0 To UBound(tmMissingAtt) + 1) As MISSINGATT
                            End If
                        End If
                    Else
                        If ilVefCode < 0 Then
                            ilFound = False
                            For ilLoop = 0 To UBound(smMissingVef) - 1 Step 1
                                'If StrComp(smFields(3), smMissingVef(ilLoop), 1) = 0 Then
                                If StrComp(smFields(2), smMissingVef(ilLoop), 1) = 0 Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                'lbcError.AddItem "Vehicle Missing: " & Trim$(smFields(3))
                                lbcError.AddItem "Vehicle Missing: " & Trim$(smFields(2))
                                'Print #hmMsg, "Vehicle Missing: " & Trim$(smFields(3))
                                Print #hmMsg, "Vehicle Missing: " & Trim$(smFields(2))
                                'smMissingVef(UBound(smMissingVef)) = smFields(3)
                                smMissingVef(UBound(smMissingVef)) = smFields(2)
                                ReDim Preserve smMissingVef(0 To UBound(smMissingVef) + 1) As String
                            End If
                        End If
                    End If
                Else
                    ilFound = False
                    For ilLoop = 0 To UBound(smMissingShtt) - 1 Step 1
                        'If StrComp(smFields(1), smMissingShtt(ilLoop), 1) = 0 Then
                        If StrComp(smFields(0), smMissingShtt(ilLoop), 1) = 0 Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        'lbcError.AddItem "Station Missing: " & Trim$(smFields(1))
                        lbcError.AddItem "Station Missing: " & Trim$(smFields(0))
                        'Print #hmMsg, "Station Missing: " & Trim$(smFields(1))
                        Print #hmMsg, "Station Missing: " & Trim$(smFields(0))
                        'smMissingShtt(UBound(smMissingShtt)) = smFields(1)
                        smMissingShtt(UBound(smMissingShtt)) = smFields(0)
                        ReDim Preserve smMissingShtt(0 To UBound(smMissingShtt) + 1) As String
                    End If
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
            llPercent = (lmProcessedNoBytes / lmTotalNoBytes) * 100
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmFloodPercent <> llPercent Then
                lmFloodPercent = llPercent
                lbcPercent.Caption = Str$(llPercent) & "%"
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    If ilRet <> 0 Then
        mReadFileAffiliateSpots = False
    Else
        mReadFileAffiliateSpots = True
        lmFloodPercent = 100
        lbcPercent.Caption = "100%"
    End If
    Print #hmMsg, "** Import Affiliate Spots Finished: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    Exit Function
mReadFileAffiliateSpotsErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportCSV-mReadFileAffiliateSpots"
    Exit Function
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFileAirDates               *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Private Function mReadFileAirDates(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim ilMaxCol As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilMatch As Integer
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilFoundID As Integer
    Dim ilCode As Integer
    Dim slCallLetters As String
    Dim llStationID As Long
    Dim ilShttCode As Integer
    Dim slVehicle As String
    Dim ilVefCode As Integer
    Dim ilSelected As Integer
    Dim slCurDate As String
    Dim slCurTime As String
    Dim slAgreeStart As String
    Dim slAgreeEnd As String
    Dim slOnAir As String
    Dim slOffAir As String
    Dim slDropDate As String
    Dim slEndDate As String
    Dim llAgreementID As Long
    Dim ilPostType As Integer
    Dim slWklyClear As String
    Dim slHrlyClear As String
    Dim slHrUsed As String
    Dim ilPos As Integer
    Dim slMsg As String
    Dim llPercent As Long
    Dim ilUpper As Integer
    Dim ilID As Integer
    Dim ilAgreementType As Integer
    Dim llAttCode As Long
    Dim slFdStTime As String
    Dim slFdEdTime As String
    Dim slPdStTime As String
    Dim slPdEdTime As String
    Dim ilUpdateID As Integer
    Dim ilStartMissingMsg As Integer
    Dim llTemp As Long
    Dim ilVef As Integer
    Dim slPledgeType As String
    Dim VehCombo_rst As ADODB.Recordset
        
    slCurDate = Format(gNow(), sgShowDateForm)
    slCurTime = Format(gNow(), sgShowTimeWSecForm)
    lmStartDate = DateValue("12/28/1998")
    ReDim smMissingVef(0 To 0) As String
    ReDim smMissingShtt(0 To 0) As String
    ilStartMissingMsg = False
    ilUpper = 0
    ReDim tmAgreeID(0 To 0) As AGREEID
    ilRet = 0
    On Error GoTo ErrHand
    SQLQuery = "SELECT attAgreementID, attShfCode, attVefCode, attOnAir, attOffAir, attDropDate, attCode"
    SQLQuery = SQLQuery + " FROM att"
    Set rst = gSQLSelectCall(SQLQuery)
    If ilRet <> 0 Then
        mReadFileAirDates = False
        Exit Function
    End If
    While Not rst.EOF
        ilUpper = UBound(tmAgreeID)
        tmAgreeID(ilUpper).lCode = rst!attCode
        tmAgreeID(ilUpper).iShttCode = rst!attshfCode
        tmAgreeID(ilUpper).iVefCode = rst!attvefCode
        tmAgreeID(ilUpper).lOnAir = DateValue(gAdjYear(rst!attOnAir))
        tmAgreeID(ilUpper).lOffAir = DateValue(gAdjYear(rst!attOffAir))
        tmAgreeID(ilUpper).lDropDate = DateValue(gAdjYear(rst!attDropDate))
        If tmAgreeID(ilUpper).lDropDate < tmAgreeID(ilUpper).lOffAir Then
            tmAgreeID(ilUpper).lEndDate = tmAgreeID(ilUpper).lDropDate
        Else
            tmAgreeID(ilUpper).lEndDate = tmAgreeID(ilUpper).lOffAir
        End If
        tmAgreeID(ilUpper).lAgreementID = rst!attAgreementID
        ilUpper = ilUpper + 1
        ReDim Preserve tmAgreeID(0 To ilUpper) As AGREEID
        rst.MoveNext
    Wend

    'ilRet = 0
    'On Error GoTo mReadFileAirDatesErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        lbcError.AddItem "Open " & slFromFile & " error#" & Str$(ilRet)
        Print #hmMsg, "Open " & slFromFile & " error#" & Str$(ilRet)
        Close hmFrom
        mReadFileAirDates = False
        Exit Function
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    lmProcessedNoBytes = 0
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mReadFileAirDatesErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Print #hmMsg, "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Close hmFrom
            mReadFileAirDates = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, True, smFields()
                For ilLoop = LBound(smFields) To UBound(smFields) Step 1
                    smFields(ilLoop) = Trim$(smFields(ilLoop))
                Next ilLoop
                'smFields(1) = UCase$(Trim$(smFields(1)))
                smFields(0) = UCase$(Trim$(smFields(0)))
                'slCallLetters = smFields(1)
                slCallLetters = smFields(0)
                'llStationID = Val(smFields(11))
                llStationID = Val(smFields(10))
                ilShttCode = -1
                If Len(slCallLetters) > 40 Then
                    lbcError.AddItem slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                    Print #hmMsg, slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                    slCallLetters = Left$(slCallLetters, 40)
                End If
                For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                    'If StrComp(Trim$(tgStationInfo(ilLoop).sCallLetters), slCallLetters, vbTextCompare) = 0 Then
                    If tgStationInfo(ilLoop).lID = llStationID Then
                        ilShttCode = tgStationInfo(ilLoop).iCode
                        Exit For
                    End If
                Next ilLoop
                
                If ilShttCode > 0 Then
                    'slVehicle = Trim$(smFields(2))
                    slVehicle = Trim$(smFields(1))
                    ilVefCode = -1
                    ilSelected = False
                    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                        If sgShowByVehType = "Y" Then
                            smVehName = Mid$(tgVehicleInfo(ilLoop).sVehicle, 3)
                        Else
                            smVehName = tgVehicleInfo(ilLoop).sVehicle
                        End If
                        'If StrComp(Trim$(smVehName), smFields(2), vbTextCompare) = 0 Then
                        If StrComp(Trim$(smVehName), smFields(1), vbTextCompare) = 0 Then
                            'ilVefCode = tgVehicleInfo(ilLoop).icode
                            'If lbcNames.Selected(ilLoop) Then
                            '    ilSelected = True
                            'End If
                            'Exit For
                            ilFound = False
                            ilVefCode = tgVehicleInfo(ilLoop).iCode
                            For ilVef = 0 To lbcNames.ListCount - 1 Step 1
                                If lbcNames.ItemData(ilVef) = ilVefCode Then
                                    If lbcNames.Selected(ilLoop) Then
                                        ilSelected = True
                                    End If
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilVef
                            If ilFound Then
                                Exit For
                            End If
                        End If
                    Next ilLoop
                    If (ilVefCode > 0) And (ilSelected) Then
                        'llAgreementID = Val(smFields(3))
                        llAgreementID = Val(smFields(2))
                        'If Len(smFields(4)) <> 0 Then
                        If Len(smFields(3)) <> 0 Then
                            'slOnAir = Format$(smFields(4), sgShowDateForm)
                            slOnAir = Format$(smFields(3), sgShowDateForm)
                        'Else
                        '    slOnAir = "1/1/1970"
                        'End If
                            'If Len(smFields(5)) <> 0 Then
                            If Len(smFields(4)) <> 0 Then
                                'slOffAir = Format$(smFields(5), sgShowDateForm)
                                slOffAir = Format$(smFields(3), sgShowDateForm)
                            Else
                                slOffAir = "12/31/2069"
                            End If
                            'If Len(smFields(10)) <> 0 Then
                            If Len(smFields(9)) <> 0 Then
                                'slDropDate = Format$(smFields(10), sgShowDateForm)
                                slDropDate = Format$(smFields(9), sgShowDateForm)
                            Else
                                slDropDate = "12/31/2069"
                            End If
                            If DateValue(gAdjYear(slDropDate)) < DateValue(gAdjYear(slOffAir)) Then
                                slEndDate = slDropDate
                            Else
                                slEndDate = slOffAir
                            End If
                            If DateValue(gAdjYear(slEndDate)) >= lmStartDate Then
                                ilFoundID = -1
                                ilAgreementType = -1
                                For ilID = 0 To UBound(tmAgreeID) - 1 Step 1
                                    'If tmAgreeID(ilID).lAgreementID = llAgreementID Then
                                    '    'If (DateValue(slOnAir) = tmAgreeID(ilID).lOnAir) And (DateValue(slOffAir) = tmAgreeID(ilID).lOffAir) Then
                                    '    '    ilFoundID = ilID
                                    '    '    ilAgreementType = 2 'Update
                                    '    '    Exit For
                                    '    'ElseIf (DateValue(slOnAir) <= tmAgreeID(ilID).lOffAir) And (DateValue(slOffAir) >= tmAgreeID(ilID).lOnAir) Then
                                    '    '    'Termimate and Make new Agreement
                                    '    '    ilFoundID = ilID
                                    '    '    ilAgreementType = 3
                                    '    '    Exit For
                                    '    'Else
                                    '    '    ilAgreementType = 1 'New
                                    '    '    ilFoundID = ilID
                                    '    'End If
                                    '    'Problem- dates should be tested but other agreements might be getting
                                    '    'date changes also.
                                    '    If ilAgreementID = -1 Then
                                    '        ilFoundID = ilID
                                    '        ilAgreementType = 2 'Update
                                    '    Else
                                     '
                                     '   End If
                                    'ElseIf (ilShttCode = tmAgreeID(ilID).iShttCode) And (ilVefCode = tmAgreeID(ilID).iVefCode) Then
                                    If (tmAgreeID(ilID).lAgreementID = llAgreementID) Or (ilShttCode = tmAgreeID(ilID).iShttCode) And (ilVefCode = tmAgreeID(ilID).iVefCode) Then
                                        If ((DateValue(gAdjYear(slEndDate)) < tmAgreeID(ilID).lOnAir) Or (DateValue(gAdjYear(slOnAir)) > tmAgreeID(ilID).lEndDate)) Then
                                            If ilAgreementType = -1 Then
                                                ilAgreementType = 1 'New
                                                ilFoundID = ilID
                                            End If
                                        Else
                                            If ilAgreementType = -1 Then
                                                ilFoundID = ilID
                                                ilAgreementType = 2 'Update
                                            Else
                                                If ilAgreementType = 1 Then
                                                    ilAgreementType = 2
                                                    ilFoundID = ilID
                                                Else
                                                    'Agreements overlap
                                                    ilAgreementType = 4 'Overlap message
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                Next ilID
                                'slWklyClear = smFields(6)
                                slWklyClear = smFields(5)
                                If Len(slWklyClear) > 12 Then
                                    slWklyClear = Left$(slWklyClear, 12)
                                End If
                                'slHrlyClear = smFields(7)
                                slHrlyClear = smFields(6)
                                If Len(slHrlyClear) > 12 Then
                                    slHrlyClear = Left$(slHrlyClear, 12)
                                End If
                                'slHrUsed = smFields(8)
                                slHrUsed = smFields(7)
                                If Len(slHrUsed) > 12 Then
                                    slHrUsed = Left$(slHrUsed, 12)
                                End If
                                'ilPostType = Val(smFields(9))
                                ilPostType = Val(smFields(8))
                                If ilAgreementType = 2 Then   'Adjust dates
                                    ilUpdateID = False
                                    If tmAgreeID(ilFoundID).lEndDate < DateValue(gAdjYear(slEndDate)) Then
                                        ilUpdateID = True
                                    End If
                                    slMsg = ""
                                    If tmAgreeID(ilFoundID).lOnAir < DateValue(gAdjYear(slOnAir)) Then
                                        slOnAir = Format$(tmAgreeID(ilFoundID).lOnAir, sgShowDateForm)
                                        slMsg = slMsg & " Retained On Air"
                                    ElseIf tmAgreeID(ilFoundID).lOnAir <> DateValue(gAdjYear(slOnAir)) Then
                                        slMsg = slMsg & " Changed On Air from " & Format$(tmAgreeID(ilFoundID).lOnAir, sgShowDateForm) & " to " & slOnAir
                                    End If
                                    If tmAgreeID(ilFoundID).lOffAir > DateValue(gAdjYear(slOffAir)) Then
                                        slOffAir = Format$(tmAgreeID(ilFoundID).lOffAir, sgShowDateForm)
                                        slMsg = slMsg & " Retained Off Air"
                                    ElseIf tmAgreeID(ilFoundID).lOffAir <> DateValue(gAdjYear(slOffAir)) Then
                                        slMsg = slMsg & " Changed Off Air from " & Format$(tmAgreeID(ilFoundID).lOffAir, sgShowDateForm) & " to " & slOffAir
                                    End If
                                    If tmAgreeID(ilFoundID).lDropDate > DateValue(gAdjYear(slDropDate)) Then
                                        slOffAir = Format$(tmAgreeID(ilFoundID).lDropDate, sgShowDateForm)
                                        slMsg = slMsg & " Retained Drop Date"
                                    ElseIf tmAgreeID(ilFoundID).lDropDate <> DateValue(gAdjYear(slDropDate)) Then
                                        slMsg = slMsg & " Changed Drop Date from " & Format$(tmAgreeID(ilFoundID).lDropDate, sgShowDateForm) & " to " & slDropDate
                                    End If
                                    If slMsg <> "" Then
                                        'Print #hmMsg, "Agreements Changed: " & Trim$(smFields(1)) & " " & Trim$(smFields(2)) & slMsg
                                        Print #hmMsg, "Agreements Changed: " & Trim$(smFields(0)) & " " & Trim$(smFields(1)) & slMsg
                                    End If
                                End If
                                'Changed to allow lst/cp's to be generated passed the Agreement End Date unless Drop Date defined.
                                slAgreeStart = slOnAir
                                slAgreeEnd = slOffAir
                                slOffAir = slDropDate
                                'Insert or Update database
                                ilRet = 0
                                On Error GoTo ErrHand
                                If (ilAgreementType = -1) Or (ilAgreementType = 1) Or (ilAgreementType = 3) Then
                                    'D.S. 8/2/05
                                    llTemp = gFindAttHole()
                                    If llTemp = -1 Then
                                        Screen.MousePointer = vbDefault
                                        mReadFileAirDates = False
                                        Exit Function
                                    End If
                                    slPledgeType = "A"
                                    SQLQuery = "INSERT INTO att (attCode, attShfCode, attVefCode, attAgreeStart, "
                                    SQLQuery = SQLQuery & "attAgreeEnd, attOnAir, attOffAir, "
                                    SQLQuery = SQLQuery & "attSigned, attSignDate, attLoad, "
                                    SQLQuery = SQLQuery & "attTimeType, attComp, attBarCode, "
                                    SQLQuery = SQLQuery & "attDropDate, attUsfCode, "
                                    SQLQuery = SQLQuery & "attEnterDate, attEnterTime, attNotice, "
                                    SQLQuery = SQLQuery & "attCarryCmml, attNoCDs, attSendTape, "
                                    SQLQuery = SQLQuery & "attACName, attACPhone, attGenLog, "
                                    SQLQuery = SQLQuery & "attGenCP, attPostingType, attPrintCP, "
                                    SQLQuery = SQLQuery & "attComments, attGenOther, attAgreementID, "
                                    SQLQuery = SQLQuery & "attWklyClear, attHrlyClear, attHrUsed, attPledgeType)"
                                    SQLQuery = SQLQuery & " VALUES (" & llTemp & ", " & ilShttCode & "," & ilVefCode & ",'" & Format$(slAgreeStart, sgSQLDateForm) & "',"
                                    SQLQuery = SQLQuery & "'" & Format$(slAgreeEnd, sgSQLDateForm) & "','" & Format$(slOnAir, sgSQLDateForm) & "','" & Format$(slOffAir, sgSQLDateForm) & "',"
                                    SQLQuery = SQLQuery & "0,'" & Format$("12/31/2069", sgSQLDateForm) & "'," & 1 & ","
                                    SQLQuery = SQLQuery & "1,0,1,"
                                    SQLQuery = SQLQuery & "'" & Format$(slDropDate, sgSQLDateForm) & "'," & 1 & ","
                                    SQLQuery = SQLQuery & "'" & Format$(slCurDate, sgSQLDateForm) & "','" & Format$(slCurTime, sgSQLTimeForm) & "','" & "" & "',"
                                    SQLQuery = SQLQuery & "0,1,0,"
                                    SQLQuery = SQLQuery & "'" & "" & "','" & "" & "','" & "" & "',"
                                    SQLQuery = SQLQuery & "'" & "" & "'," & ilPostType & ",1,"
                                    SQLQuery = SQLQuery & "'" & "" & "','" & "" & "'," & llAgreementID & ","
                                    SQLQuery = SQLQuery & "'" & slWklyClear & "','" & slHrlyClear & "','" & slHrUsed & "','" & slPledgeType & "'" & ")"
                                    cnn.BeginTrans
                                    'cnn.Execute SQLQuery, rdExecDirect
                                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                        '6/11/16: Replaced GoSub
                                        'GoSub ErrHand:
                                        Screen.MousePointer = vbDefault
                                        gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileAirDates"
                                        cnn.RollbackTrans
                                        ilRet = 1
                                    End If
                                    If ilRet = 0 Then
                                        cnn.CommitTrans
                                        SQLQuery = "Select MAX(attCode) from att"
                                        Set rst = gSQLSelectCall(SQLQuery)
                                        ilUpper = UBound(tmAgreeID)
                                        If llTemp = 0 Then
                                            llAttCode = rst(0).Value
                                            tmAgreeID(ilUpper).lCode = rst(0).Value
                                        Else
                                            llAttCode = llTemp
                                            tmAgreeID(ilUpper).lCode = llTemp
                                        End If
                                        tmAgreeID(ilUpper).lAgreementID = llAgreementID
                                        tmAgreeID(ilUpper).lAgreementID = llAgreementID
                                        tmAgreeID(ilUpper).iShttCode = ilShttCode
                                        tmAgreeID(ilUpper).iVefCode = ilVefCode
                                        tmAgreeID(ilUpper).lOnAir = DateValue(gAdjYear(slOnAir))
                                        tmAgreeID(ilUpper).lOffAir = DateValue(gAdjYear(slOffAir))
                                        tmAgreeID(ilUpper).lDropDate = DateValue(gAdjYear(slDropDate))
                                        If tmAgreeID(ilUpper).lDropDate < tmAgreeID(ilUpper).lOffAir Then
                                            tmAgreeID(ilUpper).lEndDate = tmAgreeID(ilUpper).lDropDate
                                        Else
                                            tmAgreeID(ilUpper).lEndDate = tmAgreeID(ilUpper).lOffAir
                                        End If
                                        ilUpper = ilUpper + 1
                                        ReDim Preserve tmAgreeID(0 To ilUpper) As AGREEID
                                        ReDim tgDat(0 To 0) As DAT
                                        
                                        If (ilAgreementType = -1) Then
                                            SQLQuery = "Select vefCombineVefCode from VEF_Vehicles Where vefCode = " & imVefCode
                                            Set VehCombo_rst = gSQLSelectCall(SQLQuery)
                                            If Not VehCombo_rst.EOF Then
                                                imVefCombo = VehCombo_rst!vefCombineVefCode
                                            End If
                                            gGetAvails llAttCode, ilShttCode, ilVefCode, imVefCombo, slOnAir, True
                                        Else
                                            'Get pledge time from other agreement
                                            SQLQuery = "SELECT * FROM DAT WHERE datAtfCode = " & tmAgreeID(ilFoundID).lCode
                                            Set rst = gSQLSelectCall(SQLQuery)
                                            While Not rst.EOF
                                                ilUpper = UBound(tgDat)
                                                tgDat(ilUpper).iStatus = 1
                                                tgDat(ilUpper).lCode = 0
                                                tgDat(ilUpper).lAtfCode = llAttCode  '(1).Value
                                                tgDat(ilUpper).iShfCode = ilShttCode  '(2).Value
                                                tgDat(ilUpper).iVefCode = ilVefCode  '(3).Value
                                                'tgDat(ilUpper).iDACode = rst!datDACode    '(4).Value
                                                tgDat(ilUpper).iFdDay(0) = rst!datFdMon   '(5).Value
                                                tgDat(ilUpper).iFdDay(1) = rst!datFdTue   '(6).Value
                                                tgDat(ilUpper).iFdDay(2) = rst!datFdWed   '(7).Value
                                                tgDat(ilUpper).iFdDay(3) = rst!datFdThu   '(8).Value
                                                tgDat(ilUpper).iFdDay(4) = rst!datFdFri   '(9).Value
                                                tgDat(ilUpper).iFdDay(5) = rst!datFdSat   '(10).Value
                                                tgDat(ilUpper).iFdDay(6) = rst!datFdSun   '(11).Value
                                                tgDat(ilUpper).sFdSTime = Format$(CStr(rst!datFdStTime), "hh:mm:ss")
                                                tgDat(ilUpper).sFdETime = Format$(CStr(rst!datFdEdTime), "hh:mm:ss")
                                                tgDat(ilUpper).iFdStatus = rst!datFdStatus    '(14).Value
                                                tgDat(ilUpper).iPdDay(0) = rst!datPdMon   '(15).Value
                                                tgDat(ilUpper).iPdDay(1) = rst!datPdTue   '(16).Value
                                                tgDat(ilUpper).iPdDay(2) = rst!datPdWed   '(17).Value
                                                tgDat(ilUpper).iPdDay(3) = rst!datPdThu   '(18).Value
                                                tgDat(ilUpper).iPdDay(4) = rst!datPdFri   '(19).Value
                                                tgDat(ilUpper).iPdDay(5) = rst!datPdSat   '(20).Value
                                                tgDat(ilUpper).iPdDay(6) = rst!datPdSun   '(21).Value
                                                tgDat(ilUpper).sPdDayFed = rst!datPdDayFed
                                                If tgDat(ilUpper).iStatus <= 1 Then
                                                    tgDat(ilUpper).sPdSTime = Format$(CStr(rst!datPdStTime), "hh:mm:ss")
                                                    If (tgDat(ilUpper).iFdStatus = 1) Or (tgDat(ilUpper).iFdStatus = 9) Or (tgDat(ilUpper).iFdStatus = 10) Then
                                                        tgDat(ilUpper).sPdETime = Format$(CStr(rst!datPdEdTime), "hh:mm:ss")
                                                    Else
                                                        tgDat(ilUpper).sPdETime = ""
                                                    End If
                                                Else
                                                    tgDat(ilUpper).sPdSTime = ""
                                                    tgDat(ilUpper).sPdETime = ""
                                                End If
                                                tgDat(ilUpper).iAirPlayNo = 1
                                                '7/15/14
                                                tgDat(ilUpper).sEstimatedTime = "N"
                                                tgDat(ilUpper).sEmbeddedOrROS = "R"
                                                ilUpper = ilUpper + 1
                                                ReDim Preserve tgDat(0 To ilUpper) As DAT
                                                rst.MoveNext
                                            Wend
                                        End If
                                        For ilLoop = 0 To UBound(tgDat) - 1 Step 1
                                            ilRet = 0
                                            slFdStTime = Format$(tgDat(ilLoop).sFdSTime, "hh:mm:ss")
                                            slFdEdTime = Format$(tgDat(ilLoop).sFdETime, "hh:mm:ss")
                                            slPdStTime = Format$(tgDat(ilLoop).sPdSTime, "hh:mm:ss")
                                            slPdEdTime = Format$(tgDat(ilLoop).sPdETime, "hh:mm:ss")
                                            'To avoid sql error for null time, set if null
                                            If Len(Trim$(slPdStTime)) = 0 Then
                                                slPdStTime = slFdStTime
                                            End If
                                            If Len(Trim$(slPdEdTime)) = 0 Then
                                                slPdEdTime = slPdStTime
                                            End If
                                            'SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, datDACode, "
                                            SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, "
                                            SQLQuery = SQLQuery & "datFdMon, datFdTue, datFdWed, datFdThu, "
                                            SQLQuery = SQLQuery & "datFdFri, datFdSat, datFdSun, datFdStTime, datFdEdTime, datFdStatus, "
                                            SQLQuery = SQLQuery & "datPdMon, datPdTue, datPdWed, datPdThu, datPdFri, "
                                            SQLQuery = SQLQuery & "datPdSat, datPdSun, datPdDayFed, datPdStTime, datPdEdTime"
                                            SQLQuery = SQLQuery & "datAirPlayNo, datEstimatedTime, datEmbeddedOrROS)"
                                            SQLQuery = SQLQuery & " VALUES (" & 0 & ", " & llAttCode & ", " & tgDat(ilLoop).iShfCode & ", " & tgDat(ilLoop).iVefCode & ", "
                                            'SQLQuery = SQLQuery & tgDat(ilLoop).iStatus & ", "
                                            SQLQuery = SQLQuery & tgDat(ilLoop).iFdDay(0) & ", " & tgDat(ilLoop).iFdDay(1) & ", "
                                            SQLQuery = SQLQuery & tgDat(ilLoop).iFdDay(2) & ", " & tgDat(ilLoop).iFdDay(3) & ", "
                                            SQLQuery = SQLQuery & tgDat(ilLoop).iFdDay(4) & ", " & tgDat(ilLoop).iFdDay(5) & ", "
                                            SQLQuery = SQLQuery & tgDat(ilLoop).iFdDay(6) & ", "
                                            SQLQuery = SQLQuery & "'" & Format$(slFdStTime, sgSQLTimeForm) & "', '" & Format$(slFdEdTime, sgSQLTimeForm) & "', "
                                            SQLQuery = SQLQuery & tgDat(ilLoop).iFdStatus & ", "
                                            SQLQuery = SQLQuery & tgDat(ilLoop).iPdDay(0) & ", " & tgDat(ilLoop).iPdDay(1) & ", "
                                            SQLQuery = SQLQuery & tgDat(ilLoop).iPdDay(2) & ", " & tgDat(ilLoop).iPdDay(3) & ", "
                                            SQLQuery = SQLQuery & tgDat(ilLoop).iPdDay(4) & ", " & tgDat(ilLoop).iPdDay(5) & ", "
                                            SQLQuery = SQLQuery & tgDat(ilLoop).iPdDay(6) & ", "
                                            SQLQuery = SQLQuery & "'" & tgDat(ilLoop).sPdDayFed & "', "
                                            SQLQuery = SQLQuery & "'" & Format$(slPdStTime, sgSQLTimeForm) & "', '" & Format$(slPdEdTime, sgSQLTimeForm) & "', "
                                            SQLQuery = SQLQuery & tgDat(ilLoop).iAirPlayNo & ", '" & tgDat(ilLoop).sEstimatedTime & "', '" & tgDat(ilLoop).sEmbeddedOrROS & "')"
                                            cnn.BeginTrans
                                            'cnn.Execute SQLQuery, rdExecDirect
                                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                                '6/11/16: Replaced GoSub
                                                'GoSub ErrHand:
                                                Screen.MousePointer = vbDefault
                                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileAirDates"
                                                cnn.RollbackTrans
                                                ilRet = 1
                                            End If
                                            If ilRet = 0 Then
                                                cnn.CommitTrans
                                                SQLQuery = "SELECT MAX(datCode) from dat"
                                                Set rst = gSQLSelectCall(SQLQuery)
                                                If Not rst.EOF Then
                                                    tgDat(ilLoop).lCode = rst(0).Value
                                                End If
                                            End If
                                        Next ilLoop
                                    End If
                                End If
                                If (ilAgreementType = 2) Or (ilAgreementType = 3) Then
                                    If ilAgreementType = 3 Then
                                        If (DateValue(gAdjYear(slOnAir)) > tmAgreeID(ilFoundID).lOnAir) Then
                                            slOffAir = Format$(DateValue(gAdjYear(slOnAir)) - 1, sgShowDateForm)
                                        End If
                                        If (DateValue(gAdjYear(slOffAir)) < tmAgreeID(ilFoundID).lOffAir) Then
                                            slOnAir = Format$(DateValue(gAdjYear(slOffAir)) + 1, sgShowDateForm)
                                        End If
                                    End If
                                    SQLQuery = "UPDATE att"
                                    SQLQuery = SQLQuery & " SET attShfCode = " & ilShttCode & ","
                                    SQLQuery = SQLQuery & "attVefCode = " & ilVefCode & ","
                                    SQLQuery = SQLQuery & "attAgreeStart = '" & Format$(slAgreeStart, sgSQLDateForm) & "',"
                                    SQLQuery = SQLQuery & "attAgreeEnd = '" & Format$(slAgreeEnd, sgSQLDateForm) & "',"
                                    SQLQuery = SQLQuery & "attOnAir = '" & Format$(slOnAir, sgSQLDateForm) & "',"
                                    SQLQuery = SQLQuery & "attOffAir = '" & Format$(slOffAir, sgSQLDateForm) & "',"
                                    If ilUpdateID Then
                                        SQLQuery = SQLQuery & "attDropDate = '" & Format$(slDropDate, sgSQLDateForm) & "',"
                                        SQLQuery = SQLQuery & "attAgreementID = " & llAgreementID & ","
                                        SQLQuery = SQLQuery & "attWklyClear = '" & slWklyClear & "',"
                                        SQLQuery = SQLQuery & "attHrlyClear = '" & slHrlyClear & "',"
                                        SQLQuery = SQLQuery & "attHrUsed = '" & slHrUsed & "'"
                                    Else
                                        SQLQuery = SQLQuery & "attDropDate = '" & Format$(slDropDate, sgSQLDateForm) & "'"
                                    End If
                                    SQLQuery = SQLQuery & " WHERE (attCode = " & tmAgreeID(ilFoundID).lCode & ")"
                                    cnn.BeginTrans
                                    'cnn.Execute SQLQuery, rdExecDirect
                                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                        '6/11/16: Replaced GoSub
                                        'GoSub ErrHand:
                                        Screen.MousePointer = vbDefault
                                        gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileAirDates"
                                        cnn.RollbackTrans
                                        ilRet = 1
                                    End If
                                    If ilRet = 0 Then
                                        cnn.CommitTrans
                                        tmAgreeID(ilFoundID).lAgreementID = llAgreementID
                                        tmAgreeID(ilFoundID).lOnAir = DateValue(gAdjYear(slOnAir))
                                        tmAgreeID(ilFoundID).lOffAir = DateValue(gAdjYear(slOffAir))
                                        tmAgreeID(ilFoundID).lDropDate = DateValue(gAdjYear(slDropDate))
                                        If tmAgreeID(ilFoundID).lDropDate < tmAgreeID(ilFoundID).lOffAir Then
                                            tmAgreeID(ilFoundID).lEndDate = tmAgreeID(ilFoundID).lDropDate
                                        Else
                                            tmAgreeID(ilFoundID).lEndDate = tmAgreeID(ilFoundID).lOffAir
                                        End If
                                    End If
                                End If
                                If ilAgreementType = 4 Then
                                    'lbcError.AddItem "Agreements Overlap: " & Trim$(smFields(1)) & " " & Trim$(smFields(2))
                                    lbcError.AddItem "Agreements Overlap: " & Trim$(smFields(0)) & " " & Trim$(smFields(1))
                                    Print #hmMsg, "Agreements Overlap: " & Trim$(slLine)
                                End If
                            End If
                        Else
                            ''start date missing
                            'If Not ilStartMissingMsg Then
                            '    lbcError.AddItem "Start Date Missing: See Output Text File"
                            '    ilStartMissingMsg = True
                            'End If
                            'Print #hmMsg, "Start Date Missing: " & Trim$(slLine)
                        End If
                    Else
                        If ilVefCode < 0 Then
                            ilFound = False
                            For ilLoop = 0 To UBound(smMissingVef) - 1 Step 1
                                'If StrComp(smFields(2), smMissingVef(ilLoop), 1) = 0 Then
                                If StrComp(smFields(1), smMissingVef(ilLoop), 1) = 0 Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                'lbcError.AddItem "Vehicle Missing: " & Trim$(smFields(2))
                                lbcError.AddItem "Vehicle Missing: " & Trim$(smFields(1))
                                'Print #hmMsg, "Vehicle Missing: " & Trim$(smFields(2))
                                Print #hmMsg, "Vehicle Missing: " & Trim$(smFields(1))
                                'smMissingVef(UBound(smMissingVef)) = smFields(2)
                                smMissingVef(UBound(smMissingVef)) = smFields(1)
                                ReDim Preserve smMissingVef(0 To UBound(smMissingVef) + 1) As String
                            End If
                        End If
                    End If
                Else
                    ilFound = False
                    For ilLoop = 0 To UBound(smMissingShtt) - 1 Step 1
                        'If StrComp(smFields(1), smMissingShtt(ilLoop), 1) = 0 Then
                        If StrComp(smFields(0), smMissingShtt(ilLoop), 1) = 0 Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        'lbcError.AddItem "Station Missing: " & Trim$(smFields(1))
                        lbcError.AddItem "Station Missing: " & Trim$(smFields(0))
                        'Print #hmMsg, "Station Missing: " & Trim$(smFields(1))
                        Print #hmMsg, "Station Missing: " & Trim$(smFields(0))
                        'smMissingShtt(UBound(smMissingShtt)) = smFields(1)
                        smMissingShtt(UBound(smMissingShtt)) = smFields(0)
                        ReDim Preserve smMissingShtt(0 To UBound(smMissingShtt) + 1) As String
                    End If
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
            llPercent = (lmProcessedNoBytes / lmTotalNoBytes) * 100
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmFloodPercent <> llPercent Then
                lmFloodPercent = llPercent
                lbcPercent.Caption = Str$(llPercent) & "%"
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    'Set OffAir to TFN for the latest agreement
    For ilLoop = 0 To UBound(tmAgreeID) - 1 Step 1
        ilIndex = ilLoop
        llAgreementID = tmAgreeID(ilIndex).lAgreementID
        ilShttCode = tmAgreeID(ilIndex).iShttCode
        ilVefCode = tmAgreeID(ilIndex).iVefCode
        tmAgreeID(ilIndex).lAgreementID = -1
        tmAgreeID(ilIndex).iShttCode = -1
        If (llAgreementID <> -1) And (ilShttCode <> -1) Then
            For ilID = ilLoop + 1 To UBound(tmAgreeID) - 1 Step 1
                If (tmAgreeID(ilID).lAgreementID = llAgreementID) Or ((ilShttCode = tmAgreeID(ilID).iShttCode) And (ilVefCode = tmAgreeID(ilID).iVefCode)) Then
                    If tmAgreeID(ilID).lEndDate > tmAgreeID(ilIndex).lEndDate Then
                        ilIndex = ilID
                    End If
                    tmAgreeID(ilID).lAgreementID = -1
                    tmAgreeID(ilID).iShttCode = -1
                End If
            Next ilID
            SQLQuery = "UPDATE att"
            SQLQuery = SQLQuery & " SET attOffAir = '" & Format$("12/31/2069", sgSQLDateForm) & "'"
            SQLQuery = SQLQuery & " WHERE (attCode = " & tmAgreeID(ilIndex).lCode & ")"
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileAirDates"
                cnn.RollbackTrans
                ilRet = 1
            End If
        End If
    Next ilLoop
    If ilRet <> 0 Then
        mReadFileAirDates = False
    Else
        mReadFileAirDates = True
        lmFloodPercent = 100
        lbcPercent.Caption = "100%"
    End If
    Print #hmMsg, "** Import Agreements Finished: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    Exit Function
mReadFileAirDatesErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportCSV-mReadFileAirDates"
End Function


Private Sub mFillVehicle()
    Dim iLoop As Integer
    lbcNames.Clear
    chkAll.Value = 0
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcNames.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcNames.ItemData(lbcNames.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFileLogSpots               *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Private Function mReadFileLogSpots(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim ilMaxCol As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilMatch As Integer
    Dim slCurDate As String
    Dim ilFound As Integer
    Dim ilVefCode As Integer
    Dim ilAdfCode As Integer
    Dim ilPos As Integer
    Dim slCntrNo As String
    Dim llCntrNo As Long
    Dim slName As String
    Dim slAdvtName As String
    Dim slAbbr As String
    Dim slProd As String
    Dim slLogDate As String
    Dim slLogTime As String
    Dim slZone As String
    Dim slCart As String
    Dim ilStatus As Integer
    Dim ilLen As Integer
    Dim ilSpotType As Integer
    Dim iUpper As Integer
    Dim slStr As String
    Dim slMsg As String
    Dim slChar As String
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilSelected As Integer
    Dim llPercent As Long
    Dim ilAdfAdded As Integer
    Dim ilFirstSpotTest As Integer
    Dim ilVef As Integer
        
    slCurDate = Format(gNow(), sgShowDateForm)
    gGetSyncDateTime slSyncDate, slSyncTime
    ilAdfAdded = False
    ilFirstSpotTest = False
    ReDim smMissingVef(0 To 0) As String
    ReDim smZoneError(0 To 0) As String
    ilRet = 0
    'On Error GoTo mReadFileLogSpotsErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        lbcError.AddItem "Open " & slFromFile & " error#" & Str$(ilRet)
        Print #hmMsg, "Open " & slFromFile & " error#" & Str$(ilRet)
        Close hmFrom
        mReadFileLogSpots = False
        Exit Function
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    lmProcessedNoBytes = 0
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mReadFileLogSpotsErr:
        'Line Input #hmFrom, slLine
        slLine = ""
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            If slChar = sgLF Then
                Exit Do
            ElseIf slChar <> sgCR Then
                slLine = slLine & slChar
            End If
        Loop
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Print #hmMsg, "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Close hmFrom
            mReadFileLogSpots = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, True, smFields()
                For ilLoop = LBound(smFields) To UBound(smFields) Step 1
                    smFields(ilLoop) = Trim$(smFields(ilLoop))
                Next ilLoop
                ilVefCode = -1
                ilSelected = False
                For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
                    If sgShowByVehType = "Y" Then
                        smVehName = Mid$(tgVehicleInfo(ilLoop).sVehicle, 3)
                    Else
                        smVehName = tgVehicleInfo(ilLoop).sVehicle
                    End If
                    'If StrComp(Trim$(smVehName), smFields(1), vbTextCompare) = 0 Then
                    If StrComp(Trim$(smVehName), smFields(0), vbTextCompare) = 0 Then
                        'ilVefCode = tgVehicleInfo(ilLoop).icode
                        'If lbcNames.Selected(ilLoop) Then
                        '    ilSelected = True
                        'End If
                        ilFound = False
                        ilVefCode = tgVehicleInfo(ilLoop).iCode
                        For ilVef = 0 To lbcNames.ListCount - 1 Step 1
                            If lbcNames.ItemData(ilVef) = ilVefCode Then
                                If lbcNames.Selected(ilLoop) Then
                                    ilSelected = True
                                End If
                                ilFound = True
                                Exit For
                            End If
                        Next ilVef
                        If ilFound Then
                            Exit For
                        End If
                    End If
                Next ilLoop
                If (ilVefCode > 0) And (ilSelected) Then
                    'slStr = smFields(2)
                    slStr = smFields(1)
                    ilPos = InStr(slStr, "-")
                    If ilPos > 0 Then
                        slCntrNo = Left$(slStr, ilPos - 1) & Mid$(slStr, ilPos + 1)
                    Else
                        slCntrNo = slStr
                    End If
                    llCntrNo = Val(slCntrNo)
                    'slName = smFields(3)
                    slName = smFields(2)
                    If StrComp(Left$(slName, 4), "MYL-", vbTextCompare) = 0 Then
                        slName = "MYL"
                    End If
                    slAdvtName = slName
                    ilAdfCode = -1
                    For ilLoop = 0 To UBound(tgAdvtInfo) - 1 Step 1
                        If StrComp(Trim$(tgAdvtInfo(ilLoop).sAdvtName), slAdvtName, vbTextCompare) = 0 Then
                            ilAdfCode = tgAdvtInfo(ilLoop).iCode
                            Exit For
                        End If
                    Next ilLoop
                    slName = gFixQuote(slName)
                    slAbbr = Left$(slAdvtName, 7)
                    ilPos = InStr(slAbbr, "'")
                    If ilPos > 0 Then
                        If ilPos <> 7 Then
                            slAbbr = Left$(slAbbr, ilPos) & "'" & right$(slAbbr, Len(slAbbr) - ilPos)
                        Else
                            slAbbr = Left$(slAbbr, 6)
                        End If
                    End If
                    'slProd = gFixQuote(smFields(4))
                    slProd = gFixQuote(smFields(3))
                    'ilSpotType = Val(smFields(5))
                    ilSpotType = Val(smFields(4))
                    'slLogDate = Format$(smFields(6), sgShowDateForm)
                    slLogDate = Format$(smFields(5), sgShowDateForm)
                    'Truncate seconds
                    'slLogTime = Format$(smFields(7), "hh:mm") & ":00" '"hh:mm:ss")
                    slLogTime = Format$(smFields(6), "hh:mm") & ":00" '"hh:mm:ss")
                    If (DateValue(gAdjYear(slLogDate)) >= lmStartDate) And (DateValue(gAdjYear(slLogDate)) <= lmLastDate) Then
                        'If Len(Trim$(smFields(8))) <> 0 Then
                        If Len(Trim$(smFields(7))) <> 0 Then
                            'slZone = UCase$(Left$(smFields(8), 1)) & "ST"
                            slZone = UCase$(Left$(smFields(7), 1)) & "ST"
                            Select Case slZone
                                Case "EST", "CST", "MST", "PST"
                                Case Else
                                    ilFound = False
                                    For ilLoop = 0 To UBound(smZoneError) - 1 Step 1
                                        If StrComp(slZone, smZoneError(ilLoop), 1) = 0 Then
                                            ilFound = True
                                            Exit For
                                        End If
                                    Next ilLoop
                                    If Not ilFound Then
                                        lbcError.AddItem "Zone Ignored: " & slZone
                                        Print #hmMsg, "Zone Ignored: " & slZone
                                        smZoneError(UBound(smZoneError)) = slZone
                                        ReDim Preserve smZoneError(0 To UBound(smZoneError) + 1) As String
                                    End If
                                    slZone = ""
                            End Select
                        Else
                            slZone = ""
                        End If
                        'slCart = smFields(9)
                        slCart = smFields(8)
                        'ilStatus = Val(smFields(10))
                        ilStatus = Val(smFields(9))
                        If ilStatus > 0 Then
                            ilStatus = ilStatus - 1
                        End If
                        If ilStatus = 6 Then 'Change status 7 to 2
                            ilStatus = 1
                        End If
                        'ilLen = 60 * Val(Left$(smFields(11), 2)) + Val(right$(smFields(11), 2))
                        ilLen = 60 * Val(Left$(smFields(10), 2)) + Val(right$(smFields(10), 2))
                        On Error GoTo ErrHand
                        If ilAdfCode = -1 Then
                            'Add advertiser
                            'SQLQuery = "INSERT INTO ADF_Advertisers Adf (adfName, adfAbbr, adfProd, "
                            SQLQuery = "INSERT INTO " & "ADF_Advertisers"
                            SQLQuery = SQLQuery & " (adfName, adfAbbr, adfProd, "
                            SQLQuery = SQLQuery & "adfSlfCode, adfAgfCode, "
                            SQLQuery = SQLQuery & "adfCodeRep, adfCodeAgy, adfCodeStn, "
                            SQLQuery = SQLQuery & "adfMnfComp1, adfMnfComp2, adfMnfExcl1, "
                            SQLQuery = SQLQuery & "adfMnfExcl2, adfCppCpm, adfMnfDemo1, "
                            SQLQuery = SQLQuery & "adfMnfDemo2, adfMnfDemo3, adfMnfDemo4, "
                            SQLQuery = SQLQuery & "adfTarget1, adfTarget2, adfTarget3, "
                            SQLQuery = SQLQuery & "adfTarget4, adfCreditRestr, adfCreditLimit, "
                            SQLQuery = SQLQuery & "adfPaymRating, adfISCI, adfMnfSort, "
                            SQLQuery = SQLQuery & "adfBilAgyDir, adfCntrAddr1, adfCntrAddr2, "
                            SQLQuery = SQLQuery & "adfCntrAddr3, adfBillAddr1, adfBillAddr2, "
                            SQLQuery = SQLQuery & "adfBillAddr3, adfArfLkCode, adfArfContrCode, "
                            SQLQuery = SQLQuery & "adfArfInvCode, adfCntrPrtSz, adfSlsTax1, "
                            SQLQuery = SQLQuery & "adfSlsTax2, adfCrdApp, adfCrdRtg, "
                            SQLQuery = SQLQuery & "adfPnfBuyer, adfPnfPay, adfPct90, "
                            SQLQuery = SQLQuery & "adfCurrAR, adfUnbilled, adfHiCredit, "
                            SQLQuery = SQLQuery & "adfTotalGross, adfDateEntrd, adfNSFChks, "
                            SQLQuery = SQLQuery & "adfDateLstInv, adfDateLstPaym, adfAvgToPay, "
                            SQLQuery = SQLQuery & "adfLstToPay, adfNoInvPd, adfNewBus, "
                            SQLQuery = SQLQuery & "adfEndDate, adfMerge, adfUrfCode, "
                            SQLQuery = SQLQuery & "adfState, adfCrdAppDate, adfCrdAppTime, "
                            SQLQuery = SQLQuery & "adfPkInvShow, adfBkoutPoolStatus) "
                            SQLQuery = SQLQuery & " VALUES ('" & slName & "','" & slAbbr & "','" & "" & "',"
                            SQLQuery = SQLQuery & 0 & "," & 0 & ","
                            SQLQuery = SQLQuery & "'" & "" & "','" & "" & "','" & "" & "',"
                            SQLQuery = SQLQuery & 0 & "," & 0 & "," & 0 & ","
                            SQLQuery = SQLQuery & 0 & ",'" & "N" & "'," & 0 & ","
                            SQLQuery = SQLQuery & 0 & "," & 0 & "," & 0 & ","
                            SQLQuery = SQLQuery & 0 & "," & 0 & "," & 0 & ","
                            SQLQuery = SQLQuery & 0 & ",'" & "N" & "'," & 0 & ","
                            SQLQuery = SQLQuery & "'" & "1" & "','" & "" & "'," & 0 & ","
                            SQLQuery = SQLQuery & "'" & "A" & "','" & "" & "','" & "" & "',"
                            SQLQuery = SQLQuery & "'" & "" & "','" & "" & "','" & "" & "',"
                            SQLQuery = SQLQuery & "'" & "" & "'," & 0 & "," & 0 & ","
                            SQLQuery = SQLQuery & 0 & ",'" & "" & "','" & "" & "',"
                            SQLQuery = SQLQuery & "'" & "" & "','" & "A" & "','" & "" & "',"
                            SQLQuery = SQLQuery & 0 & "," & 0 & "," & 0 & ","
                            SQLQuery = SQLQuery & "0" & "," & "0" & "," & "0" & ","
                            SQLQuery = SQLQuery & "0" & ",'" & Format$(slCurDate, sgSQLDateForm) & "'," & 0 & ","
                            SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "','" & Format$("1/1/1970", sgSQLDateForm) & "'," & 0 & ","
                            SQLQuery = SQLQuery & 0 & "," & 0 & ",'" & "Y" & "',"
                            SQLQuery = SQLQuery & "'" & Format$("12/31/2069", sgSQLDateForm) & "'," & 0 & "," & 2 & ","
                            SQLQuery = SQLQuery & "'" & "A" & "','" & Format$("1/1/1970", sgSQLDateForm) & "','" & Format$("00:00:00", sgSQLTimeForm) & "',"
                            SQLQuery = SQLQuery & "'" & "" & "', 'N'" & ")"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/11/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileLogSpots"
                                cnn.RollbackTrans
                                ilRet = 1
                            End If
                            If ilRet = 0 Then
                                cnn.CommitTrans
                                ilAdfAdded = True
                                'SQLQuery = "Select MAX(adfCode) from ADF_Advertisers"
                                SQLQuery = "Select MAX(adfCode) FROM ADF_Advertisers"
                                Set rst = gSQLSelectCall(SQLQuery)
                                iUpper = UBound(tgAdvtInfo)
                                ilAdfCode = rst(0).Value
                                tgAdvtInfo(iUpper).iCode = ilAdfCode
                                tgAdvtInfo(iUpper).sAdvtName = slAdvtName
                                tgAdvtInfo(iUpper).sAdvtAbbr = Left$(slAdvtName, 7)
                                iUpper = iUpper + 1
                                ReDim Preserve tgAdvtInfo(0 To iUpper) As ADVTINFO
                                ''SQLQuery = "UPDATE ADF_Advertisers adf"
                                'SQLQuery = "UPDATE " & "ADF_Advertisers"
                                'SQLQuery = SQLQuery & " SET adfRemoteID = " & 0 & ","
                                'SQLQuery = SQLQuery & "adfAutoCode = " & tgAdvtInfo(iUpper - 1).iCode & ","
                                'SQLQuery = SQLQuery & "adfSyncDate = '" & Format$(slSyncDate, sgSQLDateForm) & "',"
                                'SQLQuery = SQLQuery & "adfSyncTime = '" & Format$(slSyncTime, sgSQLTimeForm) & "',"
                                'SQLQuery = SQLQuery & "adfSourceID = " & 0
                                'SQLQuery = SQLQuery & " WHERE (adfCode = " & ilAdfCode & ")"
                                'cnn.BeginTrans
                                ''cnn.Execute SQLQuery, rdExecDirect
                                'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '    GoSub ErrHand:
                                'End If
                                'If ilRet = 0 Then
                                '    cnn.CommitTrans
                                'End If
                                lbcError.AddItem "Added: " & Trim$(slAdvtName)
                                Print #hmMsg, "Added: " & Trim$(slAdvtName)
                            End If
                        End If
                        If ilAdfCode > 0 Then
                            If Not ilFirstSpotTest Then
                                ilFirstSpotTest = True
                                '*************************
                                'Remove so Jim could read in same day because not
                                'all records were processed do to backup running
                                'SQLQuery = "SELECT lst.lstCode FROM lst WHERE (lst.lstLogVefCode = " & ilVefCode & " AND lst.lstLogDate = '" & slLogDate & "')"
                                'Set rst = gSQLSelectCall(SQLQuery)
                                'If Not rst.EOF Then
                                '    gMsgBox "Log Spots Previously Created, Terminating Import", vbCritical + vbOKOnly
                                '    mReadFileLogSpots = False
                                '    Exit Function
                                'End If
                                '*****************************
                            End If
                            SQLQuery = "INSERT INTO lst (lstType, lstSdfCode, lstCntrNo, "
                            SQLQuery = SQLQuery & "lstAdfCode, lstAgfCode, lstProd, "
                            SQLQuery = SQLQuery & "lstLineNo, lstLnVefCode, lstStartDate, "
                            SQLQuery = SQLQuery & "lstEndDate, lstMon, lstTue, "
                            SQLQuery = SQLQuery & "lstWed, lstThu, lstFri, "
                            SQLQuery = SQLQuery & "lstSat, lstSun, "
                            SQLQuery = SQLQuery & "lstSpotsWk, lstPriceType, lstPrice, "
                            SQLQuery = SQLQuery & "lstSpotType, lstLogVefCode, lstLogDate, "
                            SQLQuery = SQLQuery & "lstLogTime, lstDemo, lstAud, "
                            SQLQuery = SQLQuery & "lstISCI, lstWkNo, lstBreakNo, "
                            SQLQuery = SQLQuery & "lstSeqNo, lstZone, lstCart, "
                            SQLQuery = SQLQuery & "lstCpfCode, lstCrfCsfCode, lstStatus, "
                            SQLQuery = SQLQuery & "lstLen, lstUnits, lstCifCode, "
                            '12/28/06
                            'SQLQuery = SQLQuery & "lstAnfCode)"
                            SQLQuery = SQLQuery & "lstAnfCode, lstEvtIDCefCode, lstSplitNetwork, "
                            SQLQuery = SQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, "
                            SQLQuery = SQLQuery & "lstLnStartTime, lstLnEndTIme, lstUnused)"
                            SQLQuery = SQLQuery & " VALUES (" & 0 & "," & 0 & "," & llCntrNo & ","
                            SQLQuery = SQLQuery & ilAdfCode & "," & 0 & ",'" & slProd & "',"
                            SQLQuery = SQLQuery & "0,0,'" & Format$("1/1/1970", sgSQLDateForm) & "',"
                            SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "'," & 0 & ","
                            SQLQuery = SQLQuery & "0,0,0,"
                            SQLQuery = SQLQuery & "0,0,0,"
                            SQLQuery = SQLQuery & "0,0,0,"
                            SQLQuery = SQLQuery & ilSpotType & "," & ilVefCode & ",'" & Format$(slLogDate, sgSQLDateForm) & "',"
                            SQLQuery = SQLQuery & "'" & Format$(slLogTime, sgSQLTimeForm) & "','',0,"
                            SQLQuery = SQLQuery & "'',0,0,"
                            SQLQuery = SQLQuery & "0,'" & slZone & "','" & slCart & "',"
                            SQLQuery = SQLQuery & "0,0," & ilStatus & ","
                            SQLQuery = SQLQuery & ilLen & ",0,0,"
                            '12/28/06
                            'SQLQuery = SQLQuery & "0)"
                            SQLQuery = SQLQuery & 0 & ", " & 0 & ", '" & "N" & "', "
                            SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", '" & "N" & "', " & 0 & ", "
                            SQLQuery = SQLQuery & "'" & Format("12am", sgSQLTimeForm) & "', '" & Format("12am", sgSQLTimeForm) & "', '" & "" & "'" & ")"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/11/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileLogSpots"
                                cnn.RollbackTrans
                                ilRet = 1
                            End If
                            If ilRet = 0 Then
                                cnn.CommitTrans
                            End If
                        End If
                    End If
                Else
                    If ilVefCode < 0 Then
                        ilFound = False
                        For ilLoop = 0 To UBound(smMissingVef) - 1 Step 1
                            'If StrComp(smFields(1), smMissingVef(ilLoop), 1) = 0 Then
                            If StrComp(smFields(0), smMissingVef(ilLoop), 1) = 0 Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilLoop
                        If Not ilFound Then
                            'lbcError.AddItem "Vehicle Missing: " & Trim$(smFields(1))
                            lbcError.AddItem "Vehicle Missing: " & Trim$(smFields(0))
                            'Print #hmMsg, "Vehicle Missing: " & Trim$(smFields(1))
                            Print #hmMsg, "Vehicle Missing: " & Trim$(smFields(0))
                            'smMissingVef(UBound(smMissingVef)) = smFields(1)
                            smMissingVef(UBound(smMissingVef)) = smFields(0)
                            ReDim Preserve smMissingVef(0 To UBound(smMissingVef) + 1) As String
                        End If
                    End If
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
            llPercent = (lmProcessedNoBytes / lmTotalNoBytes) * 100
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmFloodPercent <> llPercent Then
                lmFloodPercent = llPercent
                lbcPercent.Caption = Str$(llPercent) & "%"
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    If ilAdfAdded Then
        On Error GoTo ErrHand
        iUpper = 0
        ReDim tgAdvtInfo(0 To 0) As ADVTINFO
        SQLQuery = "SELECT adfName, adfAbbr, adfCode"
        SQLQuery = SQLQuery + " FROM ADF_Advertisers"
        SQLQuery = SQLQuery + " ORDER BY adfName"
        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF
            iUpper = UBound(tgAdvtInfo)
            tgAdvtInfo(iUpper).iCode = rst!adfCode
            tgAdvtInfo(iUpper).sAdvtName = rst!adfName
            tgAdvtInfo(iUpper).sAdvtAbbr = rst!adfAbbr
            iUpper = iUpper + 1
            ReDim Preserve tgAdvtInfo(0 To iUpper) As ADVTINFO
            rst.MoveNext
        Wend
    End If
    If ilRet <> 0 Then
        mReadFileLogSpots = False
    Else
        mReadFileLogSpots = True
        lmFloodPercent = 100
        lbcPercent.Caption = "100%"
    End If
    Print #hmMsg, "** Import Log Spots Info Finished: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    Exit Function
mReadFileLogSpotsErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportCSV-mReadFileLogSpots"
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFileMAI                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Private Function mReadFileMAI(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim ilMaxCol As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilMatch As Integer
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilCode As Integer
    Dim slCallLetters As String
    Dim llStationID As Long
    Dim slCurDate As String
    Dim slCurTime As String
    Dim slLastActiveDate As String
    Dim slFirstName As String
    Dim slLastName As String
    Dim slFax As String
    Dim ilPos As Integer
    Dim slAddr1 As String
    Dim slAddr2 As String
    Dim slCity As String
    Dim slState As String
    Dim slCountry As String
    Dim slZip As String
    Dim slPhone As String
    Dim slPDName As String
    Dim slTDName As String
    Dim slACName As String
    Dim slMsg As String
    'Dim slCityMarket As String
    Dim slMarket As String
    Dim slPDPhone As String
    Dim slTDPhone As String
    Dim slACPhone As String
    Dim slEMail As String
    Dim slZone As String
    Dim ilRank As Integer
    Dim slONAddr1 As String
    Dim slONAddr2 As String
    Dim slONCity As String
    Dim slONState As String
    Dim slONZip As String
    Dim slLicCity As String
    Dim slLicState As String
    Dim slSerialNo1 As String
    Dim slSerialNo2 As String
    Dim ilDaylight As Integer
    Dim ilChecked As Integer
    Dim llPercent As Long
    Dim ilLp As Integer
    Dim slFrequency As String
    Dim llPermanentStationID As Long

        
    slCurDate = Format(gNow(), sgShowDateForm)
    slCurTime = Format(gNow(), sgShowTimeWSecForm)
    slLastActiveDate = Format(DateValue(gAdjYear(slCurDate)) - 1, sgShowDateForm)
    ReDim smZoneError(0 To 0) As String
    slSerialNo1 = ""
    slSerialNo2 = ""
    slFrequency = ""
    llPermanentStationID = 0
    'ilRet = 0
    'On Error GoTo mReadFileMAIErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        lbcError.AddItem "Open " & slFromFile & " error#" & Str$(ilRet)
        Print #hmMsg, "Open " & slFromFile & " error#" & Str$(ilRet)
        Close hmFrom
        mReadFileMAI = False
        Exit Function
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    lmProcessedNoBytes = 0
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mReadFileMAIErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Print #hmMsg, "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Close hmFrom
            mReadFileMAI = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, True, smFields()
                For ilLoop = LBound(smFields) To UBound(smFields) Step 1
                    smFields(ilLoop) = Trim$(smFields(ilLoop))
                Next ilLoop
                ilIndex = -1
                'smFields(1) = UCase$(Trim$(smFields(1)))
                smFields(0) = UCase$(Trim$(smFields(0)))
                'slCallLetters = smFields(1)
                slCallLetters = smFields(0)
                'llStationID = Val(smFields(26))
                llStationID = Val(smFields(25))
                If (Asc(slCallLetters) >= Asc("A")) And (Asc(slCallLetters) <= Asc("Z")) Then
                    If Len(slCallLetters) > 40 Then
                        lbcError.AddItem slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                        Print #hmMsg, slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                        slCallLetters = Left$(slCallLetters, 40)
                    End If
                    For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                        'If StrComp(Trim$(tgStationInfo(ilLoop).sCallLetters), slCallLetters, vbTextCompare) = 0 Then
                        If tgStationInfo(ilLoop).lID = llStationID Then
                            ilIndex = ilLoop
                            ilCode = tgStationInfo(ilLoop).iCode
                            Exit For
                        End If
                    Next ilLoop
                    
                    'slAddr1 = Trim$(smFields(2))
                    slAddr1 = Trim$(smFields(1))
                    'If InStr(1, slAddr1, "PO ", vbTextCompare) = 1 Then
                    '    Mid$(slAddr1, 1, 2) = "PO"
                    'ElseIf InStr(1, slAddr1, "P.O. ", vbTextCompare) = 1 Then
                    '    Mid$(slAddr1, 1, 4) = "P.O."
                    'ElseIf InStr(1, slAddr1, "P.O ", vbTextCompare) = 1 Then
                    '    Mid$(slAddr1, 1, 3) = "P.O"
                    'End If
                    If Len(slAddr1) > 40 Then
                        lbcError.AddItem slCallLetters & ": Address " & slAddr1 & " truncated to " & Left$(slAddr1, 40)
                        Print #hmMsg, slCallLetters & ": Address " & slAddr1 & " truncated to " & Left$(slAddr1, 40)
                        slAddr1 = Left$(slAddr1, 40)
                    End If
                    slAddr1 = gFixQuote(slAddr1)
                    'slAddr2 = Trim$(smFields(3))
                    slAddr2 = Trim$(smFields(3))
                    If Len(slAddr2) > 40 Then
                        lbcError.AddItem slCallLetters & ": Address " & slAddr2 & " truncated to " & Left$(slAddr2, 40)
                        Print #hmMsg, slCallLetters & ": Address " & slAddr2 & " truncated to " & Left$(slAddr2, 40)
                        slAddr2 = Left$(slAddr2, 40)
                    End If
                    slAddr2 = gFixQuote(slAddr2)
                    'slCity = Trim$(smFields(4))
                    slCity = Trim$(smFields(3))
                    If Len(slCity) > 40 Then
                        lbcError.AddItem slCallLetters & ": City " & slCity & " truncated to " & Left$(slCity, 40)
                        Print #hmMsg, slCallLetters & ": City " & slCity & " truncated to " & Left$(slCity, 40)
                        slCity = Left$(slCity, 40)
                    End If
                    slCity = gFixQuote(slCity)
                    'slState = UCase$(Trim$(smFields(5)))
                    slState = UCase$(Trim$(smFields(4)))
                    If Len(slState) > 40 Then
                        lbcError.AddItem slCallLetters & ": State " & slState & " truncated to " & Left$(slState, 40)
                        Print #hmMsg, slCallLetters & ": State " & slState & " truncated to " & Left$(slState, 40)
                        slState = Left$(slState, 40)
                    End If
                    slState = gFixQuote(slState)
                    'slCountry = UCase$(Trim$(smFields(6)))
                    slCountry = UCase$(Trim$(smFields(5)))
                    If Len(slCountry) > 40 Then
                        lbcError.AddItem slCallLetters & ": Country " & slCountry & " truncated to " & Left$(slCountry, 40)
                        Print #hmMsg, slCallLetters & ": Country " & slCountry & " truncated to " & Left$(slCountry, 40)
                        slState = Left$(slState, 40)
                    End If
                    slCountry = gFixQuote(slCountry)
                    'slZip = Trim$(smFields(7))
                    slZip = Trim$(smFields(6))
                    If Len(slZip) > 20 Then
                        lbcError.AddItem slCallLetters & ": Zip " & slZip & " truncated to " & Left$(slZip, 20)
                        Print #hmMsg, slCallLetters & ": Zip " & slZip & " truncated to " & Left$(slZip, 20)
                        slZip = Left$(slZip, 20)
                    End If

                    'slEMail = UCase$(Trim$(smFields(8)))
                    slEMail = UCase$(Trim$(smFields(7)))
                    If Len(slEMail) > 70 Then
                        lbcError.AddItem slCallLetters & ": E-Mail " & slEMail & " truncated to " & Left$(slEMail, 40)
                        Print #hmMsg, slCallLetters & ": E-Mail " & slEMail & " truncated to " & Left$(slEMail, 40)
                        slEMail = Left$(slEMail, 70)
                    End If
                    'slFax = Trim$(smFields(9))
                    slFax = Trim$(smFields(8))
                    If InStr(1, slFax, "1-", vbTextCompare) = 1 Then
                        slFax = right$(slFax, Len(slFax) - 2)
                    End If
                    If Len(slFax) > 20 Then
                        lbcError.AddItem slCallLetters & ": Fax " & slFax & " truncated to " & Left$(slFax, 20)
                        Print #hmMsg, slCallLetters & ": Fax " & slFax & " truncated to " & Left$(slFax, 20)
                        slFax = Left$(slFax, 20)
                    End If
                    'slPhone = Trim$(smFields(10))
                    slPhone = Trim$(smFields(9))
                    If Len(slPhone) > 20 Then
                        lbcError.AddItem slCallLetters & ": Phone " & slPhone & " truncated to " & Left$(slPhone, 20)
                        Print #hmMsg, slCallLetters & ": Phone " & slPhone & " truncated to " & Left$(slPhone, 20)
                        slPhone = Left$(slPhone, 20)
                    End If
                    'slZone = UCase$(Trim$(smFields(11)))
                    slZone = UCase$(Trim$(smFields(10)))
                    Select Case slZone
                        Case "E", "C", "M", "P"
                            'slZone = UCase$(Trim$(smFields(11))) & "ST"
                            slZone = UCase$(Trim$(smFields(10))) & "ST"
                        Case Else
                            ilFound = False
                            For ilLoop = 0 To UBound(smZoneError) - 1 Step 1
                                If StrComp(slZone, smZoneError(ilLoop), 1) = 0 Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                lbcError.AddItem "Zone Ignored: " & slZone
                                Print #hmMsg, "Zone Ignored: " & slZone
                                smZoneError(UBound(smZoneError)) = slZone
                                ReDim Preserve smZoneError(0 To UBound(smZoneError) + 1) As String
                            End If
                            slZone = ""
                    End Select
                    'If StrComp(smFields(12), "N", 1) = 0 Then
                    If StrComp(smFields(11), "N", 1) = 0 Then
                        ilDaylight = 1
                    Else
                        ilDaylight = 0
                    End If
                    
                    'slPDName = Trim$(smFields(13))
                    slPDName = Trim$(smFields(12))
                    If Len(slPDName) > 80 Then
                        lbcError.AddItem slCallLetters & ": Program Directory Name " & slPDName & " truncated to " & Left$(slPDName, 80)
                        Print #hmMsg, slCallLetters & ": Program Directory Name " & slPDName & " truncated to " & Left$(slPDName, 80)
                        slPDName = Left$(slPDName, 80)
                    End If
                    slPDName = gFixQuote(slPDName)
                    If slPDName <> "" Then
                        'slPDPhone = Trim$(smFields(14))
                        slPDPhone = Trim$(smFields(13))
                    Else
                        slPDPhone = ""
                    End If
                    'slTDName = Trim$(smFields(15))
                    slTDName = Trim$(smFields(14))
                    If Len(slTDName) > 80 Then
                        lbcError.AddItem slCallLetters & ": Traffic Directory Name " & slTDName & " truncated to " & Left$(slTDName, 80)
                        Print #hmMsg, slCallLetters & ": Traffic Directory Name " & slTDName & " truncated to " & Left$(slTDName, 80)
                        slTDName = Left$(slTDName, 80)
                    End If
                    slTDName = gFixQuote(slTDName)
                    If slTDName <> "" Then
                        'slTDPhone = Trim$(smFields(16))
                        slTDPhone = Trim$(smFields(15))
                    Else
                        slTDPhone = ""
                    End If
                    ilChecked = -1
                    If slPDName <> "" Then
                        slACName = slPDName
                        slACPhone = slPDPhone
                        ilChecked = 0
                    ElseIf slTDName <> "" Then
                        slACName = slTDName
                        slACPhone = slTDPhone
                        ilChecked = 2
                    Else
                        'slACName = Trim$(smFields(17))
                        slACName = Trim$(smFields(16))
                        If Len(slACName) > 80 Then
                            lbcError.AddItem slCallLetters & ": Affidavit Contact Name " & slACName & " truncated to " & Left$(slACName, 80)
                            Print #hmMsg, slCallLetters & ": Affidavit Contact Name " & slACName & " truncated to " & Left$(slACName, 80)
                            slACName = Left$(slACName, 80)
                        End If
                        If slACName <> "" Then
                            slACName = gFixQuote(slACName)
                            'slACPhone = Trim$(smFields(18))
                            slACPhone = Trim$(smFields(17))
                        Else
                            slACPhone = ""
                        End If
                    End If
                    
                    'slMarket = Trim$(smFields(19))
                    slMarket = Trim$(smFields(18))
                    If Len(slMarket) > 60 Then
                        slMarket = Left$(slMarket, 60)
                    End If
                    slMarket = gFixQuote(slMarket)
                    'ilRank = Val(smFields(20))
                    ilRank = Val(smFields(19))
                    
                    'slONAddr1 = Trim$(smFields(21))
                    slONAddr1 = Trim$(smFields(20))
                    If Len(slONAddr1) > 40 Then
                        lbcError.AddItem slCallLetters & ": O.N. Address " & slONAddr1 & " truncated to " & Left$(slONAddr1, 40)
                        Print #hmMsg, slCallLetters & ": O.N. Address " & slONAddr1 & " truncated to " & Left$(slONAddr1, 40)
                        slONAddr1 = Left$(slONAddr1, 40)
                    End If
                    slONAddr1 = gFixQuote(slONAddr1)
                    'slONAddr2 = Trim$(smFields(22))
                    slONAddr2 = Trim$(smFields(21))
                    If Len(slONAddr2) > 40 Then
                        lbcError.AddItem slCallLetters & ": O.N. Address " & slONAddr2 & " truncated to " & Left$(slONAddr2, 40)
                        Print #hmMsg, slCallLetters & ": O.N. Address " & slONAddr2 & " truncated to " & Left$(slONAddr2, 40)
                        slONAddr2 = Left$(slONAddr2, 40)
                    End If
                    slONAddr2 = gFixQuote(slONAddr2)
                    'slONCity = Trim$(smFields(23))
                    slONCity = Trim$(smFields(22))
                    If Len(slONCity) > 40 Then
                        lbcError.AddItem slCallLetters & ": O.N. City " & slONCity & " truncated to " & Left$(slONCity, 40)
                        Print #hmMsg, slCallLetters & ": O.N. City " & slONCity & " truncated to " & Left$(slONCity, 40)
                        slONCity = Left$(slONCity, 40)
                    End If
                    slONCity = gFixQuote(slONCity)
                    'slONState = UCase$(Trim$(smFields(24)))
                    slONState = UCase$(Trim$(smFields(23)))
                    If Len(slONState) > 40 Then
                        lbcError.AddItem slCallLetters & ": O.N. State " & slONState & " truncated to " & Left$(slONState, 40)
                        Print #hmMsg, slCallLetters & ": O.N. State " & slONState & " truncated to " & Left$(slONState, 40)
                        slONState = Left$(slONState, 40)
                    End If
                    slONState = gFixQuote(slONState)
                    'slONZip = Trim$(smFields(25))
                    slONZip = Trim$(smFields(24))
                    If Len(slONZip) > 20 Then
                        lbcError.AddItem slCallLetters & ": O.N. Zip " & slONZip & " truncated to " & Left$(slONZip, 20)
                        Print #hmMsg, slCallLetters & ": O.N. Zip " & slONZip & " truncated to " & Left$(slONZip, 20)
                        slONZip = Left$(slONZip, 20)
                    End If

                    
                    'slLicCity = Trim$(smFields(27))
                    slLicCity = Trim$(smFields(26))
                    If Len(slLicCity) > 40 Then
                        lbcError.AddItem slCallLetters & ": City License " & slLicCity & " truncated to " & Left$(slLicCity, 40)
                        Print #hmMsg, slCallLetters & ": City License " & slLicCity & " truncated to " & Left$(slLicCity, 40)
                        slLicCity = Left$(slLicCity, 40)
                    End If
                    slLicCity = gFixQuote(slLicCity)
                    'slLicState = UCase$(Trim$(smFields(28)))
                    slLicState = UCase$(Trim$(smFields(27)))
                    If Len(slLicState) > 2 Then
                        lbcError.AddItem slCallLetters & ": State License " & slLicState & " truncated to " & Left$(slLicState, 2)
                        Print #hmMsg, slCallLetters & ": State License " & slLicState & " truncated to " & Left$(slLicState, 2)
                        slLicState = Left$(slLicState, 2)
                    End If
                    
                    If (Len(Trim$(slAddr1)) = 0) And (Len(Trim$(slAddr2)) = 0) And (Len(Trim$(slCity)) = 0) And (Len(Trim$(slState)) = 0) And (Len(Trim$(slZip)) = 0) Then
                        slAddr1 = slONAddr1
                        slAddr2 = slONAddr2
                        slCity = slONCity
                        slState = slONState
                        slZip = slONZip
                    End If
                    'Insert or Update database
                    'Test if call letters currently used
                    For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                        If (StrComp(tgStationInfo(ilLoop).sCallLetters, slCallLetters, vbTextCompare) = 0) And (ilIndex <> ilLoop) Then
                            'Test if any agreements exist- if not then remove station
                            'if so, then don't import station
                            On Error GoTo ErrHand
                            ilRet = 0
                            SQLQuery = "SELECT attCode FROM att"
                            SQLQuery = SQLQuery + " WHERE (attShfCode = " & tgStationInfo(ilLoop).iCode & ")"
                            Set rst = gSQLSelectCall(SQLQuery)
                            If (rst.EOF = False) And (ilRet = 0) Then
                                lbcError.AddItem slCallLetters & ": previously defined and Used, not added"
                                Print #hmMsg, slCallLetters & ": previously defined and Used, not added"
                                ilIndex = -2
                            ElseIf (rst.EOF = True) And (ilRet = 0) Then
                                'Delete station
                                ilRet = 0
                                cnn.BeginTrans
                                SQLQuery = "DELETE FROM clt WHERE (cltShfCode = " & tgStationInfo(ilLoop).iCode & ")"
                                'cnn.Execute SQLQuery, rdExecDirect
                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    '6/11/16: Replaced GoSub
                                    'GoSub ErrHand:
                                    Screen.MousePointer = vbDefault
                                    gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileMAI"
                                    cnn.RollbackTrans
                                    ilRet = 1
                                End If
                                If ilRet = 0 Then
                                    SQLQuery = "DELETE FROM shtt WHERE (shttCode = " & tgStationInfo(ilLoop).iCode & ")"
                                    'cnn.Execute SQLQuery, rdExecDirect
                                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                        '6/11/16: Replaced GoSub
                                        'GoSub ErrHand:
                                        Screen.MousePointer = vbDefault
                                        gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileMAI"
                                        cnn.RollbackTrans
                                        ilRet = 1
                                    End If
                                    If ilRet = 0 Then
                                        cnn.CommitTrans
                                        For ilLp = ilLoop To UBound(tgStationInfo) - 1 Step 1
                                            tgStationInfo(ilLp) = tgStationInfo(ilLp + 1)
                                        Next ilLp
                                        ReDim Preserve tgStationInfo(0 To UBound(tgStationInfo) - 1) As STATIONINFO
                                    Else
                                        ilIndex = -2
                                    End If
                                Else
                                    ilIndex = -2
                                End If
                            Else
                                ilIndex = -2
                            End If
                            Exit For
                        End If
                    Next ilLoop
                    
                    'Determine if call letters changed, if so, add old to history
                    If ilIndex >= 0 Then
                        If (StrComp(tgStationInfo(ilIndex).sCallLetters, slCallLetters, vbTextCompare) <> 0) Then
                            SQLQuery = "INSERT INTO clt (cltShfCode, cltCallLetters, cltEndDate) "
                            SQLQuery = SQLQuery & " VALUES ( " & tgStationInfo(ilIndex).iCode & ", '" & tgStationInfo(ilIndex).sCallLetters & "', '" & Format$(slLastActiveDate, sgSQLDateForm) & "')"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/11/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileMAI"
                                cnn.RollbackTrans
                                ilRet = 1
                            End If
                            If ilRet = 0 Then
                                cnn.CommitTrans
                            Else
                                ilIndex = -2
                            End If
                        End If
                    End If
                    
                    If ilIndex <> -2 Then
                        ilRet = 0
                        On Error GoTo ErrHand
                        If ilIndex < 0 Then
                            SQLQuery = "INSERT INTO shtt (shttCallLetters, shttAddress1, shttAddress2, "
                            SQLQuery = SQLQuery & "shttCity, shttState, shttCountry, shttZip, "
                            SQLQuery = SQLQuery & "shttSelected, shttEMail, shttFax, shttPhone, "
                            SQLQuery = SQLQuery & "shttTimeZone, shttHomePage, shttPDName, shttPDPhone,"
                            'SQLQuery = SQLQuery & "shttTDName, shttTDPhone, shttMDName, shttMDPhone, "
                            SQLQuery = SQLQuery & "shttTDName, shttTDPhone, shttMDName, "
                            'SQLQuery = SQLQuery & "shttPC, shttHdDrive, shttACName, shttACPhone,"
                            SQLQuery = SQLQuery & "shttACName, shttACPhone, "
                            SQLQuery = SQLQuery & "shttMntCode, shttChecked, shttMarket, shttRank, "
                            SQLQuery = SQLQuery & "shttUsfCode, shttEnterDate, shttEnterTime, "
                            SQLQuery = SQLQuery & "shttType, shttONAddress1, shttONAddress2, shttONCity, "
                            SQLQuery = SQLQuery & "shttONState, shttONZip, shttStationID, shttCityLic, shttStateLic, shttAckDaylight, "
                            SQLQuery = SQLQuery & "shttSerialNo1, shttSerialNo2, shttFrequency, shttPermStationID, shttSpotsPerWebPage)"
                            SQLQuery = SQLQuery & " VALUES ('" & slCallLetters & "','" & slAddr1 & "','" & slAddr2 & "',"
                            SQLQuery = SQLQuery & "'" & slCity & "','" & slState & "','" & slCountry & "', '" & slZip & "',"
                            SQLQuery = SQLQuery & "-1,'" & slEMail & "','" & slFax & "','" & slPhone & "','" & slZone & "','http://www.',"
                            SQLQuery = SQLQuery & "'" & slPDName & "', '" & slPDPhone & "',"
                            'SQLQuery = SQLQuery & "'" & slTDName & "','" & slTDPhone & "','','',"
                            SQLQuery = SQLQuery & "'" & slTDName & "','" & slTDPhone & "','',"
                            SQLQuery = SQLQuery & "'" & slACName & "','" & slACPhone & "',"
                            SQLQuery = SQLQuery & "0," & ilChecked & ",'" & slMarket & "'," & ilRank & ",0,"
                            SQLQuery = SQLQuery & "'" & Format$(slCurDate, sgSQLDateForm) & "','" & Format$(slCurTime, sgSQLTimeForm) & "',0,'" & slONAddr1 & "','" & slONAddr2 & "','" & slONCity & "',"
                            SQLQuery = SQLQuery & "'" & slONState & "','" & slONZip & "'," & llStationID & ",'" & slLicCity & "','" & slLicState & "'," & ilDaylight & ", "
                            SQLQuery = SQLQuery & "'" & slSerialNo1 & "', '" & slSerialNo2 & "','" & slFrequency & "'," & llPermanentStationID & "," & 0 & ")"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/11/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileMAI"
                                cnn.RollbackTrans
                                ilRet = 1
                            End If
                            If ilRet = 0 Then
                                cnn.CommitTrans
                                SQLQuery = "Select MAX(shttCode) from shtt"
                                Set rst = gSQLSelectCall(SQLQuery)
                                tgStationInfo(UBound(tgStationInfo)).iCode = rst(0).Value
                                tgStationInfo(UBound(tgStationInfo)).sCallLetters = slCallLetters
                                tgStationInfo(UBound(tgStationInfo)).sMarket = slMarket
                                tgStationInfo(UBound(tgStationInfo)).iType = 0
                                tgStationInfo(UBound(tgStationInfo)).lID = llStationID
                                ReDim Preserve tgStationInfo(0 To UBound(tgStationInfo) + 1) As STATIONINFO
                            End If
                        Else
                            'UPDATE existing rep
                            SQLQuery = "UPDATE shtt"
                            SQLQuery = SQLQuery & " SET shttCallLetters = '" & slCallLetters & "',"
                            If slAddr1 <> "" Then
                                SQLQuery = SQLQuery & "shttAddress1 = '" & slAddr1 & "',"
                            End If
                            If slAddr2 <> "" Then
                                SQLQuery = SQLQuery & "shttAddress2 = '" & slAddr2 & "',"
                            End If
                            If slCity <> "" Then
                                SQLQuery = SQLQuery & "shttCity = '" & slCity & "',"
                            End If
                            If slState <> "" Then
                                SQLQuery = SQLQuery & "shttState = '" & slState & "',"
                            End If
                            If slCountry <> "" Then
                                SQLQuery = SQLQuery & "shttCountry = '" & slCountry & "',"
                            End If
                            If slZip <> "" Then
                                SQLQuery = SQLQuery & "shttZip = '" & slZip & "',"
                            End If
                            'SQLQuery = SQLQuery & "shttSelected = -1" & ","
                            If slEMail <> "" Then
                                SQLQuery = SQLQuery & "shttEMail = '" & slEMail & "',"
                            End If
                            If slFax <> "" Then
                                SQLQuery = SQLQuery & "shttFax = '" & slFax & "',"
                            End If
                            If slPhone <> "" Then
                                SQLQuery = SQLQuery & "shttPhone = '" & slPhone & "',"
                            End If
                            If slZone <> "" Then
                                SQLQuery = SQLQuery & "shttTimeZone = '" & slZone & "',"
                            End If
                            If slPDName <> "" Then
                                SQLQuery = SQLQuery & "shttPDName = '" & slPDName & "',"
                            End If
                            If slPDPhone <> "" Then
                                SQLQuery = SQLQuery & "shttPDPhone = '" & slPDPhone & "',"
                            End If
                            If slTDName <> "" Then
                                SQLQuery = SQLQuery & "shttTDName = '" & slTDName & "',"
                            End If
                            If slTDPhone <> "" Then
                                SQLQuery = SQLQuery & "shttTDPhone = '" & slTDPhone & "',"
                            End If
                            If slACName <> "" Then
                                SQLQuery = SQLQuery & "shttACName = '" & slACName & "',"
                            End If
                            If slACPhone <> "" Then
                                SQLQuery = SQLQuery & "shttACPhone = '" & slACPhone & "',"
                            End If
                            'SQLQuery = SQLQuery & "shttPDPhone = '" & "',"
                            SQLQuery = SQLQuery & "shttChecked =" & ilChecked & ","
                            If slMarket <> "" Then
                                SQLQuery = SQLQuery & "shttMarket = '" & slMarket & "',"
                                SQLQuery = SQLQuery & "shttRank = " & ilRank & ","
                            End If
                            SQLQuery = SQLQuery & "shttEnterDate = '" & Format$(slCurDate, sgSQLDateForm) & "',"
                            SQLQuery = SQLQuery & "shttEnterTime = '" & Format$(slCurTime, sgSQLTimeForm) & "',"
                            If slONAddr1 <> "" Then
                                SQLQuery = SQLQuery & "shttONAddress1 = '" & slONAddr1 & "',"
                            End If
                            If slONAddr2 <> "" Then
                                SQLQuery = SQLQuery & "shttONAddress2 = '" & slONAddr2 & "',"
                            End If
                            If slONCity <> "" Then
                                SQLQuery = SQLQuery & "shttONCity = '" & slONCity & "',"
                            End If
                            If slONState <> "" Then
                                SQLQuery = SQLQuery & "shttONState = '" & slONState & "',"
                            End If
                            If slONZip <> "" Then
                                SQLQuery = SQLQuery & "shttONZip = '" & slONZip & "',"
                            End If
                            SQLQuery = SQLQuery & "shttStationID =" & llStationID & ","
                            If slLicCity <> "" Then
                                SQLQuery = SQLQuery & "shttCityLic = '" & slLicCity & "',"
                            End If
                            If slLicState <> "" Then
                                SQLQuery = SQLQuery & "shttStateLic = '" & slLicState & "',"
                            End If
                            SQLQuery = SQLQuery & "shttAckDaylight =" & ilDaylight
                            SQLQuery = SQLQuery & " WHERE (shttCode = " & ilCode & ")"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/11/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileMAI"
                                cnn.RollbackTrans
                                ilRet = 1
                            End If
                            If ilRet = 0 Then
                                cnn.CommitTrans
                            End If
                            tgStationInfo(ilIndex).sCallLetters = slCallLetters
                            tgStationInfo(ilIndex).sMarket = slMarket
                        End If
                    End If
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
            llPercent = (lmProcessedNoBytes / lmTotalNoBytes) * 100
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmFloodPercent <> llPercent Then
                lmFloodPercent = llPercent
                lbcPercent.Caption = Str$(llPercent) & "%"
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    If ilRet <> 0 Then
        mReadFileMAI = False
    Else
        mReadFileMAI = True
        lmFloodPercent = 100
        lbcPercent.Caption = "100%"
    End If
    Print #hmMsg, "** Import Station Info Finished: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    '11/26/17: Set Changed date/time
    gFileChgdUpdate "shtt.mkd", True
    Exit Function
mReadFileMAIErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportCSV-mReadFileMAI"
End Function


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
Private Function mOpenMsgFile(ilSplit As Integer) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer

    'On Error GoTo mOpenMsgFileErr:
    slToFile = smMsgFile    '"ImptCSV.Txt"
    slNowDate = Format$(gNow(), sgShowDateForm)
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, sgShowDateForm)
        If DateValue(gAdjYear(slFileDate)) = DateValue(gAdjYear(slNowDate)) Then  'Append
            On Error GoTo 0
            'ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Close hmMsg
                hmMsg = -1
                gMsgBox "Open File " & slToFile & " error #" & Str$(Err.Number), vbOKOnly
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            'ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Close hmMsg
                hmMsg = -1
                gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        'ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    If igImportSelection = 0 Then
        Print #hmMsg, "** Import Station Info- CSV: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    ElseIf igImportSelection = 1 Then
        Print #hmMsg, "** Import Station Info- Vantive: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    ElseIf igImportSelection = 2 Then
        Print #hmMsg, "** Import Log Spots Info: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    ElseIf igImportSelection = 3 Then
        If ilSplit Then
            Print #hmMsg, "** Creating Affiliate Spots Split Files: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
        Else
            Print #hmMsg, "** Import Affiliate Spots Info: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
        End If
    ElseIf igImportSelection = 4 Then
        Print #hmMsg, "** Import Agreement Dates Info: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    ElseIf igImportSelection = 5 Then
        Print #hmMsg, "** Import Agreement Pledge Info: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    ElseIf igImportSelection = 6 Then
        Print #hmMsg, "** Import CP Info: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    ElseIf igImportSelection = 7 Then
        Print #hmMsg, "** Import MYL Spots Info: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    ElseIf igImportSelection = 8 Then           'another form of MAI layout
        Print #hmMsg, "** Import CSV Info: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
      
    End If
    Print #hmMsg, ""
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFileGlobal                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Private Function mReadFileGlobal(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim ilMaxCol As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilMatch As Integer
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilCode As Integer
    Dim slCallLetters As String
    Dim slCurDate As String
    Dim slCurTime As String
    Dim slFirstName As String
    Dim slLastName As String
    Dim slFax As String
    Dim ilPos As Integer
    Dim slAddr1 As String
    Dim slAddr2 As String
    Dim slCity As String
    Dim slState As String
    Dim slZip As String
    Dim slPhone As String
    Dim slPDName As String
    Dim slACName As String
    Dim slMsg As String
    Dim slCityMarket As String
    Dim slMarket As String
    Dim slPDPhone As String
    Dim slACPhone As String
    Dim llPercent As Long
    Dim slSerialNo1 As String
    Dim slSerialNo2 As String
    Dim slFrequency As String
    Dim llPermanentStationID As Long
        
    slCurDate = Format(gNow(), sgShowDateForm)
    slCurTime = Format(gNow(), sgShowTimeWSecForm)
    slSerialNo1 = ""
    slSerialNo2 = ""
    slFrequency = ""
    llPermanentStationID = 0
    'ilRet = 0
    'On Error GoTo mReadFileGlobalErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        lbcError.AddItem "Open " & slFromFile & " error#" & Str$(ilRet)
        Print #hmMsg, "Open " & slFromFile & " error#" & Str$(ilRet)
        Close hmFrom
        mReadFileGlobal = False
        Exit Function
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    lmProcessedNoBytes = 0
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mReadFileGlobalErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Print #hmMsg, "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Close hmFrom
            mReadFileGlobal = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, True, smFields()
                For ilLoop = LBound(smFields) To UBound(smFields) Step 1
                    smFields(ilLoop) = Trim$(smFields(ilLoop))
                Next ilLoop
                ilFound = -1
                'smFields(1) = UCase(smFields(1))
                smFields(1) = UCase(smFields(0))
                'smFields(2) = UCase(smFields(2))
                smFields(2) = UCase(smFields(1))
                'If smFields(2) <> "" Then
                If smFields(1) <> "" Then
                    'slCallLetters = smFields(1) & "-" & smFields(2)
                    slCallLetters = smFields(0) & "-" & smFields(1)
                Else
                    'slCallLetters = smFields(1)
                    slCallLetters = smFields(0)
                End If
                If (Asc(slCallLetters) >= Asc("A")) And (Asc(slCallLetters) <= Asc("Z")) Then
                    If Len(slCallLetters) > 40 Then
                        lbcError.AddItem slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                        Print #hmMsg, slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                        slCallLetters = Left$(slCallLetters, 40)
                    End If
                    For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                        If StrComp(Trim$(tgStationInfo(ilLoop).sCallLetters), slCallLetters, vbTextCompare) = 0 Then
                            ilFound = ilLoop
                            ilCode = tgStationInfo(ilLoop).iCode
                            Exit For
                        End If
                    Next ilLoop
                    'slFirstName = smFields(3)
                    slFirstName = smFields(2)
                    'slLastName = smFields(4)
                    slLastName = smFields(3)
                    slPDName = slFirstName & " " & slLastName
                    If Len(slPDName) > 40 Then
                        lbcError.AddItem slCallLetters & ": Program Directory Name " & slPDName & " truncated to " & Left$(slPDName, 40)
                        Print #hmMsg, slCallLetters & ": Program Directory Name " & slPDName & " truncated to " & Left$(slPDName, 40)
                        slPDName = Left$(slPDName, 40)
                    End If
                    slPDName = gFixQuote(slPDName)
                    'slFirstName = smFields(15)
                    slFirstName = smFields(14)
                    'slLastName = smFields(16)
                    slLastName = smFields(15)
                    slACName = slFirstName & " " & slLastName
                    If Len(slACName) > 40 Then
                        lbcError.AddItem slCallLetters & ": Affiliate Contact Name " & slACName & " truncated to " & Left$(slACName, 40)
                        Print #hmMsg, slCallLetters & ": Affiliate Contact Name " & slACName & " truncated to " & Left$(slACName, 40)
                        slACName = Left$(slACName, 40)
                    End If
                    slACName = gFixQuote(slACName)
                    'slAddr1 = smFields(5)
                    slAddr1 = smFields(4)
                    If InStr(1, slAddr1, "PO ", vbTextCompare) = 1 Then
                        Mid$(slAddr1, 1, 2) = "PO"
                    ElseIf InStr(1, slAddr1, "P.O. ", vbTextCompare) = 1 Then
                        Mid$(slAddr1, 1, 4) = "P.O."
                    ElseIf InStr(1, slAddr1, "P.O ", vbTextCompare) = 1 Then
                        Mid$(slAddr1, 1, 3) = "P.O"
                    End If
                    If Len(slAddr1) > 40 Then
                        lbcError.AddItem slCallLetters & ": Address " & slAddr1 & " truncated to " & Left$(slAddr1, 40)
                        Print #hmMsg, slCallLetters & ": Address " & slAddr1 & " truncated to " & Left$(slAddr1, 40)
                        slAddr1 = Left$(slAddr1, 40)
                    End If
                    slAddr1 = gFixQuote(slAddr1)
                    slAddr2 = ""
                    'slCity = smFields(6)
                    slCity = smFields(5)
                    If Len(slCity) > 20 Then
                        lbcError.AddItem slCallLetters & ": City " & slCity & " truncated to " & Left$(slCity, 20)
                        Print #hmMsg, slCallLetters & ": City " & slCity & " truncated to " & Left$(slCity, 20)
                        slCity = Left$(slCity, 20)
                    End If
                    slCity = gFixQuote(slCity)
                    slCityMarket = slCity
                    'slMarket = smFields(6)
                    slMarket = smFields(5)
                    If Len(slMarket) > 20 Then
                        slMarket = Left$(slMarket, 20)
                    End If
                    'slState = UCase$(smFields(7))
                    slState = UCase$(smFields(6))
                    If Len(slState) > 2 Then
                        lbcError.AddItem slCallLetters & ": State " & slState & " truncated to " & Left$(slState, 2)
                        Print #hmMsg, slCallLetters & ": State " & slState & " truncated to " & Left$(slState, 2)
                        slState = Left$(slState, 2)
                    End If
                    'slZip = smFields(8)
                    slZip = smFields(7)
                    If Len(slZip) > 20 Then
                        lbcError.AddItem slCallLetters & ": Zip " & slZip & " truncated to " & Left$(slZip, 20)
                        Print #hmMsg, slCallLetters & ": Zip " & slZip & " truncated to " & Left$(slZip, 20)
                        slZip = Left$(slZip, 20)
                    End If
                    'slPhone = smFields(9)
                    slPhone = smFields(8)
                    If Len(slPhone) > 20 Then
                        lbcError.AddItem slCallLetters & ": Phone " & slPhone & " truncated to " & Left$(slPhone, 20)
                        Print #hmMsg, slCallLetters & ": Phone " & slPhone & " truncated to " & Left$(slPhone, 20)
                        slPhone = Left$(slPhone, 20)
                    End If
                    If slPDName <> "" Then
                        slPDPhone = slPhone
                    Else
                        slPDPhone = ""
                    End If
                    'slFax = smFields(10)
                    slFax = smFields(9)
                    If InStr(1, slFax, "1-", vbTextCompare) = 1 Then
                        slFax = right$(slFax, Len(slFax) - 2)
                    End If
                    If Len(slFax) > 20 Then
                        lbcError.AddItem slCallLetters & ": Fax " & slFax & " truncated to " & Left$(slFax, 20)
                        Print #hmMsg, slCallLetters & ": Fax " & slFax & " truncated to " & Left$(slFax, 20)
                        slFax = Left$(slFax, 20)
                    End If
                    'Insert or Update database
                    ilRet = 0
                    On Error GoTo ErrHand
                    If ilFound < 0 Then
                        SQLQuery = "INSERT INTO shtt (shttCallLetters, shttAddress1, shttAddress2, "
                        SQLQuery = SQLQuery & "shttCity, shttState, shttCountry, shttZip, "
                        SQLQuery = SQLQuery & "shttSelected, shttEMail, shttFax, shttPhone, "
                        SQLQuery = SQLQuery & "shttTimeZone, shttHomePage, shttPDName, shttPDPhone,"
                        'SQLQuery = SQLQuery & "shttTDName, shttTDPhone, shttMDName, shttMDPhone, "
                        SQLQuery = SQLQuery & "shttTDName, shttTDPhone, shttMDName, "
                        'SQLQuery = SQLQuery & "shttPC, shttHdDrive, shttACName, shttACPhone,"
                        SQLQuery = SQLQuery & "shttACName, shttACPhone, "
                        SQLQuery = SQLQuery & "shttMntCode, shttChecked, shttMarket, shttRank, "
                        SQLQuery = SQLQuery & "shttUsfCode, shttEnterDate, shttEnterTime, "
                        SQLQuery = SQLQuery & "shttType, shttONAddress1, shttONAddress2, shttONCity, "
                        SQLQuery = SQLQuery & "shttONState, shttONZip, shttSerialNo1, shttSerialNo2, shttFrequency, shttPermStationID, shttSpotsPerWebPage)"
                        SQLQuery = SQLQuery & " VALUES ('" & slCallLetters & "','" & slAddr1 & "','" & slAddr2 & "',"
                        SQLQuery = SQLQuery & "'" & slCity & "','" & slState & "','" & "" & "','" & slZip & "',"
                        SQLQuery = SQLQuery & "-1,'','" & slFax & "','" & slPhone & "','','http://www.',"
                        SQLQuery = SQLQuery & "'" & slPDName & "', '" & slPDPhone & "',"
                        'SQLQuery = SQLQuery & "'','','','',"
                        SQLQuery = SQLQuery & "'','','',"
                        'If StrComp(Trim$(smFields(15)), "Affidavit", 1) = 0 Then
                        If StrComp(Trim$(smFields(14)), "Affidavit", 1) = 0 Then
                            SQLQuery = SQLQuery & "'" & slPDName & "','" & slPDPhone & "',"
                        Else
                            If slACName <> "" Then
                                slACPhone = slPhone
                            Else
                                slACPhone = ""
                            End If
                            SQLQuery = SQLQuery & "'" & slACName & "','" & slACPhone & "',"
                        End If
                        SQLQuery = SQLQuery & "0, -1,'" & slCityMarket & "',0,0,"
                        SQLQuery = SQLQuery & "'" & Format$(slCurDate, sgSQLDateForm) & "','" & Format$(slCurTime, sgSQLTimeForm) & "',0,'','','','','',"
                        SQLQuery = SQLQuery & "'" & slSerialNo1 & "', '" & slSerialNo2 & "','" & slFrequency & "'," & llPermanentStationID & "," & 0 & ")"
                        cnn.BeginTrans
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/11/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileGlobal"
                            cnn.RollbackTrans
                            ilRet = 1
                        End If
                        If ilRet = 0 Then
                            cnn.CommitTrans
                        End If
                        SQLQuery = "Select MAX(shttCode) from shtt"
                        Set rst = gSQLSelectCall(SQLQuery)
                        tgStationInfo(UBound(tgStationInfo)).iCode = rst(0).Value
                        tgStationInfo(UBound(tgStationInfo)).sCallLetters = slCallLetters
                        tgStationInfo(UBound(tgStationInfo)).sMarket = slMarket
                        tgStationInfo(UBound(tgStationInfo)).iType = 0
                        ReDim Preserve tgStationInfo(0 To UBound(tgStationInfo) + 1) As STATIONINFO
                    Else
                        'UPDATE existing rep
                        SQLQuery = "UPDATE shtt"
                        SQLQuery = SQLQuery & " SET shttCallLetters = '" & slCallLetters & "',"
                        If slAddr1 <> "" Then
                            SQLQuery = SQLQuery & "shttAddress1 = '" & slAddr1 & "',"
                        End If
                        If slAddr2 <> "" Then
                            SQLQuery = SQLQuery & "shttAddress2 = '" & slAddr2 & "',"
                        End If
                        If slCity <> "" Then
                            SQLQuery = SQLQuery & "shttCity = '" & slCity & "',"
                        End If
                        If slState <> "" Then
                            SQLQuery = SQLQuery & "shttState = '" & slState & "',"
                        End If
                        If slZip <> "" Then
                            SQLQuery = SQLQuery & "shttZip = '" & slZip & "',"
                        End If
                        'SQLQuery = SQLQuery & "shttSelected = -1" & ","
                        If slFax <> "" Then
                            SQLQuery = SQLQuery & "shttFax = '" & slFax & "',"
                        End If
                        If slPhone <> "" Then
                            SQLQuery = SQLQuery & "shttPhone = '" & slPhone & "',"
                        End If
                        If slPDName <> "" Then
                            SQLQuery = SQLQuery & "shttPDName = '" & slPDName & "',"
                        End If
                        If slPDPhone <> "" Then
                            SQLQuery = SQLQuery & "shttPDPhone = '" & slPDPhone & "',"
                        End If
                        'SQLQuery = SQLQuery & "shttPDPhone = '" & "',"
                        'If StrComp(Trim$(smFields(15)), "Affidavit", 1) = 0 Then
                        If StrComp(Trim$(smFields(14)), "Affidavit", 1) = 0 Then
                            If slPDName <> "" Then
                                SQLQuery = SQLQuery & "shttACName = '" & slPDName & "',"
                                'SQLQuery = SQLQuery & "shttACPhone = '" & "',"
                                If slPDPhone <> "" Then
                                    SQLQuery = SQLQuery & "shttACPhone = '" & slPDPhone & "',"
                                End If
                            End If
                        Else
                            If slACName <> "" Then
                                SQLQuery = SQLQuery & "shttACName = '" & slACName & "',"
                                slACPhone = slPhone
                                If slACPhone <> "" Then
                                    SQLQuery = SQLQuery & "shttACPhone = '" & slACPhone & "',"
                                End If
                                'SQLQuery = SQLQuery & "shttACPhone = '" & "',"
                            End If
                        End If
                        If slCityMarket <> "" Then
                            SQLQuery = SQLQuery & "shttMarket = '" & slCityMarket & "',"
                        End If
                        SQLQuery = SQLQuery & "shttEnterDate = '" & Format$(slCurDate, sgSQLDateForm) & "',"
                        SQLQuery = SQLQuery & "shttEnterTime = '" & Format$(slCurTime, sgSQLTimeForm) & "'"
                        SQLQuery = SQLQuery & " WHERE (shttCode = " & ilCode & ")"
                        cnn.BeginTrans
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/11/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileGlobal"
                            cnn.RollbackTrans
                            ilRet = 1
                        End If
                        If ilRet = 0 Then
                            cnn.CommitTrans
                        End If
                        tgStationInfo(ilFound).sCallLetters = slCallLetters
                        tgStationInfo(ilFound).sMarket = slMarket
                    End If
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
            llPercent = (lmProcessedNoBytes / lmTotalNoBytes) * 100
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmFloodPercent <> llPercent Then
                lmFloodPercent = llPercent
                lbcPercent.Caption = Str$(llPercent) & "%"
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    If ilRet <> 0 Then
        mReadFileGlobal = False
    Else
        mReadFileGlobal = True
        lmFloodPercent = 100
        lbcPercent.Caption = "100%"
    End If
    Print #hmMsg, "** Import Station Info Finished: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    '11/26/17: Set Changed date/time
    gFileChgdUpdate "shtt.mkd", True
    Exit Function
mReadFileGlobalErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportCSV-mReadFileGlobal"
End Function

Private Sub chkAll_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcNames.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcNames.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcNames.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllClick = False
    End If

End Sub

Private Sub cmcBrowse_Click()

    Dim slCurDir As String
    
    slCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
    "(*.txt)|*.txt|CSV Files (*.csv)|*.csv"
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

Private Sub cmcCancel_Click()
    If imImporting Then
        imTerminate = True
        Exit Sub
    End If
    Unload frmImportCSV
End Sub

Private Sub cmcImport_Click()
    Dim ilRet As Integer
    Dim slFromFile As String
    Dim slDate As String
    Dim ilVefCode As Integer
    Dim ilSelected As Integer
    Dim ilLoop As Integer
    
    If txtFile.Text = "" Then
        gMsgBox "Import File must be specified.", vbOKOnly
        txtFile.SetFocus
        Exit Sub
    End If
    lmProcessedNoBytes = 0
    lmStartDate = DateValue("12/28/1998")
    lbcError.Clear
    lbcMsg.Caption = ""
    lbcPercent.Caption = ""
    imImporting = True
    slFromFile = txtFile.Text
    ilRet = mOpenMsgFile(False)
    Screen.MousePointer = vbHourglass
    Print #hmMsg, "Import File: " & slFromFile
    If igImportSelection = 0 Then
        ilRet = mReadFileGlobal(slFromFile)
    ElseIf igImportSelection = 1 Then
        ilRet = mReadFileMAI(slFromFile)
    ElseIf igImportSelection = 8 Then
        ilRet = mReadFileUSRN(slFromFile)
    ElseIf igImportSelection = 2 Then
        On Error GoTo ErrHand
        SQLQuery = "Select MAX(lstCode) from lst"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst(0).Value <> 0 Then
            If gMsgBox("Log Spots Previously Created, Continue", vbYesNo) = vbNo Then
                imImporting = False
                Screen.MousePointer = vbDefault
                Close hmMsg
                Exit Sub
            End If
        End If
        slDate = Trim$(txtDate.Text)
        If Len(slDate) <= 0 Then
            imImporting = False
            Screen.MousePointer = vbDefault
            gMsgBox "Date must be Specified", vbCritical
            txtDate.SetFocus
            Close hmMsg
            Exit Sub
        End If
        If Not gIsDate(slDate) Then
            imImporting = False
            Screen.MousePointer = vbDefault
            gMsgBox "Date must be specified Correctly", vbCritical
            txtDate.SetFocus
            Close hmMsg
            Exit Sub
        End If
        lmLastDate = DateValue(gAdjYear(slDate))
        ilRet = mReadFileLogSpots(slFromFile)
    ElseIf igImportSelection = 3 Then   'Import Affiliate Spots
        On Error GoTo ErrHand
        SQLQuery = "Select MAX(astCode) from ast"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst(0).Value <> 0 Then
            If gMsgBox("Affiliate Spots Previously Created, Continue", vbYesNo) = vbNo Then
                imImporting = False
                Screen.MousePointer = vbDefault
                Close hmMsg
                Exit Sub
            End If
        End If
        slDate = Trim$(txtDate.Text)
        If Len(slDate) <= 0 Then
            imImporting = False
            Screen.MousePointer = vbDefault
            gMsgBox "Date must be Specified", vbCritical
            txtDate.SetFocus
            Close hmMsg
            Exit Sub
        End If
        If Not gIsDate(slDate) Then
            imImporting = False
            Screen.MousePointer = vbDefault
            gMsgBox "Date must be specified Correctly", vbCritical
            txtDate.SetFocus
            Close hmMsg
            Exit Sub
        End If
        lmLastDate = DateValue(gAdjYear(slDate))
        ilRet = mReadFileAffiliateSpots(slFromFile)
    ElseIf igImportSelection = 4 Then   'Import Agreements
        ilRet = mReadFileAirDates(slFromFile)
    ElseIf igImportSelection = 5 Then   'Import Agreements
        ilRet = mReadFilePledge(slFromFile)
    ElseIf igImportSelection = 6 Then   'Import CP's
        On Error GoTo ErrHand
        SQLQuery = "Select MAX(cpttCode) from cptt"
        Set rst = gSQLSelectCall(SQLQuery)
        If rst(0).Value <> 0 Then
            If gMsgBox("CP's Previously Created, Continue", vbYesNo) = vbNo Then
                imImporting = False
                Screen.MousePointer = vbDefault
                Close hmMsg
                Exit Sub
            End If
        End If
        slDate = Trim$(txtDate.Text)
        If Len(slDate) <= 0 Then
            imImporting = False
            Screen.MousePointer = vbDefault
            gMsgBox "Date must be Specified", vbCritical
            txtDate.SetFocus
            Close hmMsg
            Exit Sub
        End If
        If Not gIsDate(slDate) Then
            imImporting = False
            Screen.MousePointer = vbDefault
            gMsgBox "Date must be specified Correctly", vbCritical
            txtDate.SetFocus
            Close hmMsg
            Exit Sub
        End If
        lmLastDate = DateValue(gAdjYear(slDate))
        ilRet = mReadFileCP(slFromFile)
    ElseIf igImportSelection = 7 Then
        ilSelected = False
        For ilLoop = 0 To lbcNames.ListCount - 1 Step 1
            If lbcNames.Selected(ilLoop) Then
                If ilSelected Then
                    imImporting = False
                    Screen.MousePointer = vbDefault
                    gMsgBox "Select Only One Vehicle.", vbOKOnly
                    lbcNames.SetFocus
                    Close hmMsg
                    Exit Sub
                Else
                    ilSelected = True
                    'ilVefCode = tgVehicleInfo(ilLoop).icode
                    ilVefCode = lbcNames.ItemData(ilLoop)
                End If
            End If
        Next ilLoop
        If Not ilSelected Then
            imImporting = False
            Screen.MousePointer = vbDefault
            gMsgBox "One Vehicle Must be Selected.", vbOKOnly
            lbcNames.SetFocus
            Close hmMsg
            Exit Sub
        End If
        ilRet = mReadFileMYLSpots(ilVefCode, slFromFile)
    End If
    Close hmMsg
    Screen.MousePointer = vbDefault
    lbcMsg.Caption = "See " & smMsgFile & " for Messages"
    cmcImport.Enabled = False
    imImporting = False
    If Not imTerminate Then
        cmcCancel.Caption = "&Done"
        cmcCancel.SetFocus
    Else
        Unload frmImportCSV
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErorLog.txt", "frmImportCSV-cmcImport"
End Sub

Private Sub cmcSplit_Click()
    Dim ilRet As Integer
    Dim slFromFile As String
    Dim slDate As String
    
    If txtFile.Text = "" Then
        gMsgBox "Import File must be specified.", vbOKOnly
        txtFile.SetFocus
        Exit Sub
    End If
    smMsgFile = "ImptAf_C.Txt"
    lmProcessedNoBytes = 0
    lmStartDate = DateValue("12/28/1998")
    lbcError.Clear
    lbcMsg.Caption = ""
    lbcPercent.Caption = ""
    imImporting = True
    slFromFile = txtFile.Text
    ilRet = mOpenMsgFile(True)
    Screen.MousePointer = vbHourglass
    Print #hmMsg, "Import File: " & slFromFile
    If igImportSelection = 3 Then   'Import Affiliate Spots
        slDate = Trim$(txtDate.Text)
        If Len(slDate) <= 0 Then
            imImporting = False
            Screen.MousePointer = vbDefault
            gMsgBox "Date must be Specified", vbCritical
            txtDate.SetFocus
            Exit Sub
        End If
        If Not gIsDate(slDate) Then
            imImporting = False
            Screen.MousePointer = vbDefault
            gMsgBox "Date must be specified Correctly", vbCritical
            txtDate.SetFocus
            Exit Sub
        End If
        lmLastDate = DateValue(gAdjYear(slDate))
        ilRet = mReadAndSplit(slFromFile)
    End If
    Close hmMsg
    Screen.MousePointer = vbDefault
    lbcMsg.Caption = "See " & smMsgFile & " for Messages"
    cmcImport.Enabled = False
    imImporting = False
    If Not imTerminate Then
        cmcCancel.Caption = "&Done"
        cmcCancel.SetFocus
    Else
        Unload frmImportCSV
    End If
    Exit Sub
End Sub

Private Sub Form_Load()
    Dim iUpper As Integer
    Dim ilRet As Integer
    
    frmImportCSV.Caption = "Import Station Information - " & sgClientName
    imAllClick = False
    smCurDir = CurDir
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    imImporting = False
    imTerminate = False
    If igImportSelection = 0 Then
        lbcNames.Visible = False
        chkAll.Visible = False
        smMsgFile = "ImptStat.Txt"
        frmImportCSV.Caption = "Import Station Information- CSV"
    ElseIf igImportSelection = 1 Then
        lbcNames.Visible = False
        chkAll.Visible = False
        smMsgFile = "ImptStat.Txt"
        frmImportCSV.Caption = "Import Station Information- Vantive"
    ElseIf igImportSelection = 2 Then
        mFillVehicle
        lbcError.Move lbcNames.Left + lbcNames.Width + 120, lbcError.Top, lbcError.Width - lbcNames.Width - 120
        lbcNames.Visible = True
        chkAll.Visible = True
        lacDate.Visible = True
        txtDate.Visible = True
        smMsgFile = "ImptLog.Txt"
        frmImportCSV.Caption = "Import Spot Information- Logs"
    ElseIf igImportSelection = 3 Then
        mFillVehicle
        lbcError.Move lbcNames.Left + lbcNames.Width + 120, lbcError.Top, lbcError.Width - lbcNames.Width - 120
        lbcNames.Visible = True
        chkAll.Visible = True
        lacDate.Visible = True
        txtDate.Visible = True
        cmcSplit.Visible = True
        smMsgFile = "ImptAff.Txt"
        frmImportCSV.Caption = "Import Spot Information- Affiliate"
    ElseIf igImportSelection = 4 Then
        mFillVehicle
        lbcError.Move lbcNames.Left + lbcNames.Width + 120, lbcError.Top, lbcError.Width - lbcNames.Width - 120
        lbcNames.Visible = True
        chkAll.Visible = True
        smMsgFile = "ImptAgre.Txt"
        frmImportCSV.Caption = "Import Agreement Information- Air Dates"
    ElseIf igImportSelection = 5 Then
        mFillVehicle
        chkAll.Value = vbChecked
        'lbcError.Move lbcNames.Left + lbcNames.Width + 120, lbcError.Top, lbcError.Width - lbcNames.Width - 120
        lbcNames.Visible = False    'True
        chkAll.Visible = False  'True
        smMsgFile = "ImptPldg.Csv"
        smMsgFile = sgMsgDirectory & smMsgFile
        frmImportCSV.Caption = "Import Agreement Information- Pledged"
    ElseIf igImportSelection = 6 Then
        mFillVehicle
        lbcError.Move lbcNames.Left + lbcNames.Width + 120, lbcError.Top, lbcError.Width - lbcNames.Width - 120
        lbcNames.Visible = True
        chkAll.Visible = True
        lacDate.Visible = True
        txtDate.Visible = True
        smMsgFile = "ImptCPs.Txt"
        frmImportCSV.Caption = "Import Agreements Information- CP's"
    ElseIf igImportSelection = 7 Then
        mFillVehicle
        lbcError.Move lbcNames.Left + lbcNames.Width + 120, lbcError.Top, lbcError.Width - lbcNames.Width - 120
        lbcNames.Visible = True
        chkAll.Visible = False
        smMsgFile = "ImptMYL.Txt"
        frmImportCSV.Caption = "Import Spot Information- MYL"
    ElseIf igImportSelection = 8 Then      'USRN or CSV2 import
        smMsgFile = "ImptCSV2.Txt"
        lbcNames.Visible = False
        chkAll.Visible = False
        frmImportCSV.Caption = "Import Station Information- CSV2"
    End If
    'SQLQuery = "SELECT shttCallLetters, shttMarket, shttCode FROM shtt ORDER BY shttCallLetters, shttMarket"
    'Set rst = gSQLSelectCall(SQLQuery)
    'iUpper = 0
    'ReDim tmStationInfo(0 To 0) As STATIONINFO
    'While Not rst.EOF
    '    tmStationInfo(iUpper).iCode = rst(2).Value
    '    tmStationInfo(iUpper).sCallLetters = rst(0).Value
    '    tmStationInfo(iUpper).sMarket = rst(1).Value
    '    iUpper = iUpper + 1
    '    ReDim Preserve tmStationInfo(0 To iUpper) As STATIONINFO
    '    rst.MoveNext
    'Wend
    ilRet = gPopAvailNames()
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportCSV-FormLoad"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If InStr(1, smCurDir, ":") > 0 Then
        ChDrive Left$(smCurDir, 1)
        ChDir smCurDir
    End If
    Erase smMissingVef
    Erase smMissingShtt
    Erase smStationNotMatching
    Erase smVehicleNotMatching
    Erase smStationNameError
    Erase smVehicleNameError
    Erase smMissingTime
    Erase tmPledgeInfo
    Erase tmAgreeID
    Erase smZoneError
    Erase tmMissingAtt
    Erase tmPledgeCount
    Erase tmLstMYLInfo
    Erase tmLogSpotInfo
    Set frmImportCSV = Nothing
End Sub

Private Sub lbcNames_Click()
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = 1 Then
        imAllClick = True
        chkAll.Value = 0
        imAllClick = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFileUSRN                   *
'*                                                     *
'*             Created:10/13/05       By:D. Hosaka     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File Comma Delimited      *
'*      This file layout is the same as MAI with the   *
'*      following exceptions:  Col 6 is A or F (AM/FM) *
'*      The time zone doesnt exists and is hard-coded
'       based on the table smZoneTable
'   Columns1:   Call letters (+ AM or FM)
'          2:   Address 1
'          3:   Address 2
'          4:   City
'          5:   State
'          6:   Country
'          7:   Zip
'          8:   Email address
'          9:   Station Fax
'         10:   Station Phone
'         11:   Zone  (determined by hardcoded array)
'         12:   Daylight Savings
'         13:   PDName
'         14:   PD Phone
'         15:   Traf director name
'         16:   Traf director phone
'         17:   Affidavit contact name
'         18:   Affidavit contact phone
'         19:   Market
'         20:   Rank
'         21:   Overnite Address 1
'         22:   Overnite address 2
'         23:   Overnite City
'         24:   Overnite State
'         25:   Overnite zip
'         26:   Station Code        'previously used to test duplicate stations, use call letters instead
'         27:   License city
'         28:   License State
'
'*******************************************************
Private Function mReadFileUSRN(slFromFile As String) As Integer
Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim ilMaxCol As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilMatch As Integer
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilCode As Integer
    Dim slCallLetters As String
    Dim llStationID As Long
    Dim slCurDate As String
    Dim slCurTime As String
    Dim slLastActiveDate As String
    Dim slFirstName As String
    Dim slLastName As String
    Dim slFax As String
    Dim ilPos As Integer
    Dim slAddr1 As String
    Dim slAddr2 As String
    Dim slCity As String
    Dim slState As String
    Dim slBand As String
    Dim slCountry As String
    Dim slZip As String
    Dim slPhone As String
    Dim slPDName As String
    Dim slTDName As String
    Dim slACName As String
    Dim slMsg As String
    'Dim slCityMarket As String
    Dim slMarket As String
    Dim slPDPhone As String
    Dim slTDPhone As String
    Dim slACPhone As String
    Dim slEMail As String
    Dim slZone As String
    Dim ilRank As Integer
    Dim slONAddr1 As String
    Dim slONAddr2 As String
    Dim slONCity As String
    Dim slONState As String
    Dim slONZip As String
    Dim slLicCity As String
    Dim slLicState As String
    Dim ilDaylight As Integer
    Dim ilChecked As Integer
    Dim llPercent As Long
    Dim ilLp As Integer
    Dim llZoneIndex As Long
    Dim slZoneCode As String * 1
    Dim slSerialNo1 As String
    Dim slSerialNo2 As String
    Dim slFrequency As String
    Dim llPermanentStationID As Long

        
    slCurDate = Format(gNow(), sgShowDateForm)
    slCurTime = Format(gNow(), sgShowTimeWSecForm)
    slLastActiveDate = Format(DateValue(gAdjYear(slCurDate)) - 1, sgShowDateForm)
    ReDim smZoneError(0 To 0) As String
    slSerialNo1 = ""
    slSerialNo2 = ""
    slFrequency = ""
    llPermanentStationID = 0
    'ilRet = 0
    'On Error GoTo mReadFileUSRNErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        lbcError.AddItem "Open " & slFromFile & " error#" & Str$(ilRet)
        Print #hmMsg, "Open " & slFromFile & " error#" & Str$(ilRet)
        Close hmFrom
        mReadFileUSRN = False
        Exit Function
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    lmProcessedNoBytes = 0
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mReadFileUSRNErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Print #hmMsg, "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Close hmFrom
            mReadFileUSRN = False
            Exit Function
        End If
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, True, smFields()
                For ilLoop = LBound(smFields) To UBound(smFields) Step 1
                    smFields(ilLoop) = Trim$(smFields(ilLoop))
                Next ilLoop
                ilIndex = -1
                'smFields(1) = UCase$(Trim$(smFields(1)))
                smFields(1) = UCase$(Trim$(smFields(0)))
                'slCallLetters = smFields(1)
                slCallLetters = smFields(0)
                'llStationID = Val(smFields(26))
                llStationID = Val(smFields(25))
                If (Asc(slCallLetters) >= Asc("A")) And (Asc(slCallLetters) <= Asc("Z")) Then
                    If Len(slCallLetters) > 40 Then
                        lbcError.AddItem slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                        Print #hmMsg, slCallLetters & " truncated to " & Left$(slCallLetters, 40)
                        slCallLetters = Left$(slCallLetters, 40)
                    End If
                    For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                        If StrComp(Trim$(tgStationInfo(ilLoop).sCallLetters), slCallLetters, vbTextCompare) = 0 Then
                        'If (tgStationInfo(ilLoop).lID = llStationID) And (tgStationInfo(ilLoop).lID <> 0) Then
                            ilIndex = ilLoop
                            ilCode = tgStationInfo(ilLoop).iCode
                            Exit For
                        End If
                    Next ilLoop
                    
                    'If Trim$(smFields(6)) <> "" Then
                    If Trim$(smFields(5)) <> "" Then
                        'If Trim$(smFields(6)) = "A" Or Trim$(smFields(6)) = "F" Then
                        If Trim$(smFields(5)) = "A" Or Trim$(smFields(5)) = "F" Then
                            'slCallLetters = slCallLetters & "-" & Trim$(smFields(6)) & "M"
                            slCallLetters = slCallLetters & "-" & Trim$(smFields(5)) & "M"
                        Else
                            'slCallLetters = slCallLetters & "-" & Trim$(smFields(6))
                            slCallLetters = slCallLetters & "-" & Trim$(smFields(5))
                        End If
                    End If
                    
                    'slAddr1 = Trim$(smFields(2))
                    slAddr1 = Trim$(smFields(1))
                    'If InStr(1, slAddr1, "PO ", vbTextCompare) = 1 Then
                    '    Mid$(slAddr1, 1, 2) = "PO"
                    'ElseIf InStr(1, slAddr1, "P.O. ", vbTextCompare) = 1 Then
                    '    Mid$(slAddr1, 1, 4) = "P.O."
                    'ElseIf InStr(1, slAddr1, "P.O ", vbTextCompare) = 1 Then
                    '    Mid$(slAddr1, 1, 3) = "P.O"
                    'End If
                    If Len(slAddr1) > 40 Then
                        lbcError.AddItem slCallLetters & ": Address " & slAddr1 & " truncated to " & Left$(slAddr1, 40)
                        Print #hmMsg, slCallLetters & ": Address " & slAddr1 & " truncated to " & Left$(slAddr1, 40)
                        slAddr1 = Left$(slAddr1, 40)
                    End If
                    slAddr1 = gFixQuote(slAddr1)
                    'slAddr2 = Trim$(smFields(3))
                    slAddr2 = Trim$(smFields(2))
                    If Len(slAddr2) > 40 Then
                        lbcError.AddItem slCallLetters & ": Address " & slAddr2 & " truncated to " & Left$(slAddr2, 40)
                        Print #hmMsg, slCallLetters & ": Address " & slAddr2 & " truncated to " & Left$(slAddr2, 40)
                        slAddr2 = Left$(slAddr2, 40)
                    End If
                    slAddr2 = gFixQuote(slAddr2)
                    'slCity = Trim$(smFields(4))
                    slCity = Trim$(smFields(3))
                    If Len(slCity) > 40 Then
                        lbcError.AddItem slCallLetters & ": City " & slCity & " truncated to " & Left$(slCity, 40)
                        Print #hmMsg, slCallLetters & ": City " & slCity & " truncated to " & Left$(slCity, 40)
                        slCity = Left$(slCity, 40)
                    End If
                    slCity = gFixQuote(slCity)
                    'slState = UCase$(Trim$(smFields(5)))
                    slState = UCase$(Trim$(smFields(4)))
                    If Len(slState) > 40 Then
                        lbcError.AddItem slCallLetters & ": State " & slState & " truncated to " & Left$(slState, 40)
                        Print #hmMsg, slCallLetters & ": State " & slState & " truncated to " & Left$(slState, 40)
                        slState = Left$(slState, 40)
                    End If
                    slState = gFixQuote(slState)
                    slCountry = ""                      'this field not imported
                    'slCountry omitted, field 6 changed to band width
                    ''slCountry = UCase$(Trim$(smFields(6)))
                    'slCountry = UCase$(Trim$(smFields(5)))
                    'If Len(slCountry) > 40 Then
                    '    lbcError.AddItem slCallLetters & ": Country " & slCountry & " truncated to " & Left$(slCountry, 40)
                    '    Print #hmMsg, slCallLetters & ": Country " & slCountry & " truncated to " & Left$(slCountry, 40)
                    '    slState = Left$(slState, 40)
                    'End If
                    'slCountry = gFixQuote(slCountry)
                    'slZip = Trim$(smFields(7))
                    slZip = Trim$(smFields(6))
                    If Len(slZip) > 20 Then
                        lbcError.AddItem slCallLetters & ": Zip " & slZip & " truncated to " & Left$(slZip, 20)
                        Print #hmMsg, slCallLetters & ": Zip " & slZip & " truncated to " & Left$(slZip, 20)
                        slZip = Left$(slZip, 20)
                    End If

                    'slEMail = UCase$(Trim$(smFields(8)))
                    slEMail = UCase$(Trim$(smFields(7)))
                    If Len(slEMail) > 70 Then
                        lbcError.AddItem slCallLetters & ": E-Mail " & slEMail & " truncated to " & Left$(slEMail, 40)
                        Print #hmMsg, slCallLetters & ": E-Mail " & slEMail & " truncated to " & Left$(slEMail, 40)
                        slEMail = Left$(slEMail, 70)
                    End If
                    'slFax = Trim$(smFields(9))
                    slFax = Trim$(smFields(8))
                    If InStr(1, slFax, "1-", vbTextCompare) = 1 Then
                        slFax = right$(slFax, Len(slFax) - 2)
                    End If
                    If Len(slFax) > 20 Then
                        lbcError.AddItem slCallLetters & ": Fax " & slFax & " truncated to " & Left$(slFax, 20)
                        Print #hmMsg, slCallLetters & ": Fax " & slFax & " truncated to " & Left$(slFax, 20)
                        slFax = Left$(slFax, 20)
                    End If
                    'slPhone = Trim$(smFields(10))
                    slPhone = Trim$(smFields(10))
                    If Len(slPhone) > 20 Then
                        lbcError.AddItem slCallLetters & ": Phone " & slPhone & " truncated to " & Left$(slPhone, 20)
                        Print #hmMsg, slCallLetters & ": Phone " & slPhone & " truncated to " & Left$(slPhone, 20)
                        slPhone = Left$(slPhone, 20)
                    End If
                    ''slZONE = UCase$(Trim$(smFields(11)))
                    'slZONE = UCase$(Trim$(smFields(10)))
                    'get zone based on state field in constant array. if null state, force the pos of found match to 0, it returns if state is null
                    If Trim$(slState) <> "" Then
                        llZoneIndex = InStr(1, smZone, Trim$(slState))
                    Else
                        llZoneIndex = 0     'if state if null, return is 1
                    End If
                    slZone = ""
                    If llZoneIndex > 0 Then     'is there a index to location of match?
                        slZoneCode = Mid$(smZone, llZoneIndex + 2, 1)
                        If slZoneCode = "0" Then
                            slZone = "E"
                        ElseIf slZoneCode = "1" Then
                            slZone = "C"
                        ElseIf slZoneCode = "2" Then
                            slZone = "M"
                        ElseIf slZoneCode = "3" Then
                            slZone = "P"
                        End If
                    End If
                    
                    Select Case slZone
                        Case "E", "C", "M", "P"
                            slZone = Trim$(slZone) & "ST"
                        Case Else
                            ilFound = False
                            For ilLoop = 0 To UBound(smZoneError) - 1 Step 1
                                If StrComp(Trim$(slCallLetters), smZoneError(ilLoop), 1) = 0 Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                lbcError.AddItem "State Missing or Invalid, no Zone determined: " & Trim$(slZone) & " for " & slCallLetters
                                Print #hmMsg, "State Missing or Invalid, no Zone determined: " & Trim$(slZone); " for " & slCallLetters
                                smZoneError(UBound(smZoneError)) = Trim$(slCallLetters)
                                ReDim Preserve smZoneError(0 To UBound(smZoneError) + 1) As String
                            End If
                            slZone = ""
                    End Select
                    'If StrComp(smFields(12), "N", 1) = 0 Then
                    If StrComp(smFields(11), "N", 1) = 0 Then
                        ilDaylight = 1
                    Else
                        ilDaylight = 0
                    End If
                    
                    'slPDName = Trim$(smFields(13))
                    slPDName = Trim$(smFields(12))
                    If Len(slPDName) > 80 Then
                        lbcError.AddItem slCallLetters & ": Program Directory Name " & slPDName & " truncated to " & Left$(slPDName, 80)
                        Print #hmMsg, slCallLetters & ": Program Directory Name " & slPDName & " truncated to " & Left$(slPDName, 80)
                        slPDName = Left$(slPDName, 80)
                    End If
                    slPDName = gFixQuote(slPDName)
                    If slPDName <> "" Then
                        'slPDPhone = Trim$(smFields(14))
                        slPDPhone = Trim$(smFields(13))
                    Else
                        slPDPhone = ""
                    End If
                    'slTDName = Trim$(smFields(15))
                    slTDName = Trim$(smFields(14))
                    If Len(slTDName) > 80 Then
                        lbcError.AddItem slCallLetters & ": Traffic Directory Name " & slTDName & " truncated to " & Left$(slTDName, 80)
                        Print #hmMsg, slCallLetters & ": Traffic Directory Name " & slTDName & " truncated to " & Left$(slTDName, 80)
                        slTDName = Left$(slTDName, 80)
                    End If
                    slTDName = gFixQuote(slTDName)
                    If slTDName <> "" Then
                        'slTDPhone = Trim$(smFields(16))
                        slTDPhone = Trim$(smFields(15))
                    Else
                        slTDPhone = ""
                    End If
                    ilChecked = -1
                    If slPDName <> "" Then
                        slACName = slPDName
                        slACPhone = slPDPhone
                        ilChecked = 0
                    ElseIf slTDName <> "" Then
                        slACName = slTDName
                        slACPhone = slTDPhone
                        ilChecked = 2
                    Else
                        'slACName = Trim$(smFields(17))
                        slACName = Trim$(smFields(16))
                        If Len(slACName) > 80 Then
                            lbcError.AddItem slCallLetters & ": Affidavit Contact Name " & slACName & " truncated to " & Left$(slACName, 80)
                            Print #hmMsg, slCallLetters & ": Affidavit Contact Name " & slACName & " truncated to " & Left$(slACName, 80)
                            slACName = Left$(slACName, 80)
                        End If
                        If slACName <> "" Then
                            slACName = gFixQuote(slACName)
                            'slACPhone = Trim$(smFields(18))
                            slACPhone = Trim$(smFields(17))
                        Else
                            slACPhone = ""
                        End If
                    End If
                    
                    'slMarket = Trim$(smFields(19))
                    slMarket = Trim$(smFields(18))
                    If Len(slMarket) > 60 Then
                        slMarket = Left$(slMarket, 60)
                    End If
                    slMarket = gFixQuote(slMarket)
                    'ilRank = Val(smFields(20))
                    ilRank = Val(smFields(19))
                    
                    'slONAddr1 = Trim$(smFields(21))
                    slONAddr1 = Trim$(smFields(20))
                    If Len(slONAddr1) > 40 Then
                        lbcError.AddItem slCallLetters & ": O.N. Address " & slONAddr1 & " truncated to " & Left$(slONAddr1, 40)
                        Print #hmMsg, slCallLetters & ": O.N. Address " & slONAddr1 & " truncated to " & Left$(slONAddr1, 40)
                        slONAddr1 = Left$(slONAddr1, 40)
                    End If
                    slONAddr1 = gFixQuote(slONAddr1)
                    'slONAddr2 = Trim$(smFields(22))
                    slONAddr2 = Trim$(smFields(21))
                    If Len(slONAddr2) > 40 Then
                        lbcError.AddItem slCallLetters & ": O.N. Address " & slONAddr2 & " truncated to " & Left$(slONAddr2, 40)
                        Print #hmMsg, slCallLetters & ": O.N. Address " & slONAddr2 & " truncated to " & Left$(slONAddr2, 40)
                        slONAddr2 = Left$(slONAddr2, 40)
                    End If
                    slONAddr2 = gFixQuote(slONAddr2)
                    'slONCity = Trim$(smFields(23))
                    slONCity = Trim$(smFields(22))
                    If Len(slONCity) > 40 Then
                        lbcError.AddItem slCallLetters & ": O.N. City " & slONCity & " truncated to " & Left$(slONCity, 40)
                        Print #hmMsg, slCallLetters & ": O.N. City " & slONCity & " truncated to " & Left$(slONCity, 40)
                        slONCity = Left$(slONCity, 40)
                    End If
                    slONCity = gFixQuote(slONCity)
                    'slONState = UCase$(Trim$(smFields(24)))
                    slONState = UCase$(Trim$(smFields(23)))
                    If Len(slONState) > 40 Then
                        lbcError.AddItem slCallLetters & ": O.N. State " & slONState & " truncated to " & Left$(slONState, 40)
                        Print #hmMsg, slCallLetters & ": O.N. State " & slONState & " truncated to " & Left$(slONState, 40)
                        slONState = Left$(slONState, 40)
                    End If
                    slONState = gFixQuote(slONState)
                    'slONZip = Trim$(smFields(25))
                    slONZip = Trim$(smFields(24))
                    If Len(slONZip) > 20 Then
                        lbcError.AddItem slCallLetters & ": O.N. Zip " & slONZip & " truncated to " & Left$(slONZip, 20)
                        Print #hmMsg, slCallLetters & ": O.N. Zip " & slONZip & " truncated to " & Left$(slONZip, 20)
                        slONZip = Left$(slONZip, 20)
                    End If

                    
                    'slLicCity = Trim$(smFields(27))
                    slLicCity = Trim$(smFields(26))
                    If Len(slLicCity) > 40 Then
                        lbcError.AddItem slCallLetters & ": City License " & slLicCity & " truncated to " & Left$(slLicCity, 40)
                        Print #hmMsg, slCallLetters & ": City License " & slLicCity & " truncated to " & Left$(slLicCity, 40)
                        slLicCity = Left$(slLicCity, 40)
                    End If
                    slLicCity = gFixQuote(slLicCity)
                    'slLicState = UCase$(Trim$(smFields(28)))
                    slLicState = UCase$(Trim$(smFields(27)))
                    If Len(slLicState) > 2 Then
                        lbcError.AddItem slCallLetters & ": State License " & slLicState & " truncated to " & Left$(slLicState, 2)
                        Print #hmMsg, slCallLetters & ": State License " & slLicState & " truncated to " & Left$(slLicState, 2)
                        slLicState = Left$(slLicState, 2)
                    End If
                    
                    If (Len(Trim$(slAddr1)) = 0) And (Len(Trim$(slAddr2)) = 0) And (Len(Trim$(slCity)) = 0) And (Len(Trim$(slState)) = 0) And (Len(Trim$(slZip)) = 0) Then
                        slAddr1 = slONAddr1
                        slAddr2 = slONAddr2
                        slCity = slONCity
                        slState = slONState
                        slZip = slONZip
                    End If
                    'Insert or Update database
                    'Test if call letters currently used
                    For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                        If (StrComp(tgStationInfo(ilLoop).sCallLetters, slCallLetters, vbTextCompare) = 0) And (ilIndex <> ilLoop) Then
                            'Test if any agreements exist- if not then remove station
                            'if so, then don't import station
                            On Error GoTo ErrHand
                            ilRet = 0
                            SQLQuery = "SELECT attCode FROM att"
                            SQLQuery = SQLQuery + " WHERE (attShfCode = " & tgStationInfo(ilLoop).iCode & ")"
                            Set rst = gSQLSelectCall(SQLQuery)
                            If (rst.EOF = False) And (ilRet = 0) Then
                                lbcError.AddItem slCallLetters & ": previously defined and Used, not added"
                                Print #hmMsg, slCallLetters & ": previously defined and Used, not added"
                                ilIndex = -2
                            ElseIf (rst.EOF = True) And (ilRet = 0) Then
                                'Delete station
                                ilRet = 0
                                cnn.BeginTrans
                                SQLQuery = "DELETE FROM clt WHERE (cltShfCode = " & tgStationInfo(ilLoop).iCode & ")"
                                'cnn.Execute SQLQuery, rdExecDirect
                                 If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    '6/11/16: Replaced GoSub
                                    'GoSub ErrHand:
                                    Screen.MousePointer = vbDefault
                                    gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileUSRN"
                                    cnn.RollbackTrans
                                    ilRet = 1
                                End If
                                If ilRet = 0 Then
                                    SQLQuery = "DELETE FROM shtt WHERE (shttCode = " & tgStationInfo(ilLoop).iCode & ")"
                                    'cnn.Execute SQLQuery, rdExecDirect
                                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                        '6/11/16: Replaced GoSub
                                        'GoSub ErrHand:
                                        Screen.MousePointer = vbDefault
                                        gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileUSRN"
                                        cnn.RollbackTrans
                                        ilRet = 1
                                    End If
                                    If ilRet = 0 Then
                                        cnn.CommitTrans
                                        For ilLp = ilLoop To UBound(tgStationInfo) - 1 Step 1
                                            tgStationInfo(ilLp) = tgStationInfo(ilLp + 1)
                                        Next ilLp
                                        ReDim Preserve tgStationInfo(0 To UBound(tgStationInfo) - 1) As STATIONINFO
                                    Else
                                        ilIndex = -2
                                    End If
                                Else
                                    ilIndex = -2
                                End If
                            Else
                                ilIndex = -2
                            End If
                            Exit For
                        End If
                    Next ilLoop
                    
                    'Determine if call letters changed, if so, add old to history
                    If ilIndex >= 0 Then
                        If (StrComp(tgStationInfo(ilIndex).sCallLetters, slCallLetters, vbTextCompare) <> 0) Then
                            SQLQuery = "INSERT INTO clt (cltShfCode, cltCallLetters, cltEndDate) "
                            SQLQuery = SQLQuery & " VALUES ( " & tgStationInfo(ilIndex).iCode & ", '" & tgStationInfo(ilIndex).sCallLetters & "', '" & Format$(slLastActiveDate, sgSQLDateForm) & "')"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/11/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileUSRN"
                                cnn.RollbackTrans
                                ilRet = 1
                            End If
                            If ilRet = 0 Then
                                cnn.CommitTrans
                            Else
                                ilIndex = -2
                            End If
                        End If
                    End If
                    
                    If ilIndex <> -2 Then
                        ilRet = 0
                        On Error GoTo ErrHand
                        If ilIndex < 0 Then
                            SQLQuery = "INSERT INTO shtt (shttCallLetters, shttAddress1, shttAddress2, "
                            SQLQuery = SQLQuery & "shttCity, shttState, shttCountry, shttZip, "
                            SQLQuery = SQLQuery & "shttSelected, shttEMail, shttFax, shttPhone, "
                            SQLQuery = SQLQuery & "shttTimeZone, shttHomePage, shttPDName, shttPDPhone,"
                            'SQLQuery = SQLQuery & "shttTDName, shttTDPhone, shttMDName, shttMDPhone, "
                            SQLQuery = SQLQuery & "shttTDName, shttTDPhone, shttMDName, "
                            'SQLQuery = SQLQuery & "shttPC, shttHdDrive, shttACName, shttACPhone,"
                            SQLQuery = SQLQuery & "shttACName, shttACPhone, "
                            SQLQuery = SQLQuery & "shttMntCode, shttChecked, shttMarket, shttRank, "
                            SQLQuery = SQLQuery & "shttUsfCode, shttEnterDate, shttEnterTime, "
                            SQLQuery = SQLQuery & "shttType, shttONAddress1, shttONAddress2, shttONCity, "
                            SQLQuery = SQLQuery & "shttONState, shttONZip, shttStationID, shttCityLic, shttStateLic, shttAckDaylight, "
                            SQLQuery = SQLQuery & "shttSerialNo1, shttSerialNo2, shttFrequency, shttPermStationID, shttSpotsPerWebPage)"
                            SQLQuery = SQLQuery & " VALUES ('" & slCallLetters & "','" & slAddr1 & "','" & slAddr2 & "',"
                            SQLQuery = SQLQuery & "'" & slCity & "','" & slState & "','" & slCountry & "', '" & slZip & "',"
                            SQLQuery = SQLQuery & "-1,'" & slEMail & "','" & slFax & "','" & slPhone & "','" & slZone & "','http://www.',"
                            SQLQuery = SQLQuery & "'" & slPDName & "', '" & slPDPhone & "',"
                            'SQLQuery = SQLQuery & "'" & slTDName & "','" & slTDPhone & "','','',"
                            SQLQuery = SQLQuery & "'" & slTDName & "','" & slTDPhone & "','',"
                            SQLQuery = SQLQuery & "'" & slACName & "','" & slACPhone & "',"
                            SQLQuery = SQLQuery & "0," & ilChecked & ",'" & slMarket & "'," & ilRank & ",0,"
                            SQLQuery = SQLQuery & "'" & Format$(slCurDate, sgSQLDateForm) & "','" & Format$(slCurTime, sgSQLTimeForm) & "',0,'" & slONAddr1 & "','" & slONAddr2 & "','" & slONCity & "',"
                            SQLQuery = SQLQuery & "'" & slONState & "','" & slONZip & "'," & llStationID & ",'" & slLicCity & "','" & slLicState & "'," & ilDaylight & ","
                            SQLQuery = SQLQuery & "'" & slSerialNo1 & "', '" & slSerialNo2 & "','" & slFrequency & "'," & llPermanentStationID & "," & 0 & ")"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/11/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileUSRN"
                                cnn.RollbackTrans
                                ilRet = 1
                            End If
                            If ilRet = 0 Then
                                cnn.CommitTrans
                                SQLQuery = "Select MAX(shttCode) from shtt"
                                Set rst = gSQLSelectCall(SQLQuery)
                                tgStationInfo(UBound(tgStationInfo)).iCode = rst(0).Value
                                tgStationInfo(UBound(tgStationInfo)).sCallLetters = slCallLetters
                                tgStationInfo(UBound(tgStationInfo)).sMarket = slMarket
                                tgStationInfo(UBound(tgStationInfo)).iType = 0
                                tgStationInfo(UBound(tgStationInfo)).lID = llStationID
                                ReDim Preserve tgStationInfo(0 To UBound(tgStationInfo) + 1) As STATIONINFO
                            End If
                        Else
                            'UPDATE existing rep
                            SQLQuery = "UPDATE shtt"
                            SQLQuery = SQLQuery & " SET shttCallLetters = '" & slCallLetters & "',"
                            If slAddr1 <> "" Then
                                SQLQuery = SQLQuery & "shttAddress1 = '" & slAddr1 & "',"
                            End If
                            If slAddr2 <> "" Then
                                SQLQuery = SQLQuery & "shttAddress2 = '" & slAddr2 & "',"
                            End If
                            If slCity <> "" Then
                                SQLQuery = SQLQuery & "shttCity = '" & slCity & "',"
                            End If
                            If slState <> "" Then
                                SQLQuery = SQLQuery & "shttState = '" & slState & "',"
                            End If
                            If slCountry <> "" Then
                                SQLQuery = SQLQuery & "shttCountry = '" & slCountry & "',"
                            End If
                            If slZip <> "" Then
                                SQLQuery = SQLQuery & "shttZip = '" & slZip & "',"
                            End If
                            'SQLQuery = SQLQuery & "shttSelected = -1" & ","
                            If slEMail <> "" Then
                                SQLQuery = SQLQuery & "shttEMail = '" & slEMail & "',"
                            End If
                            If slFax <> "" Then
                                SQLQuery = SQLQuery & "shttFax = '" & slFax & "',"
                            End If
                            If slPhone <> "" Then
                                SQLQuery = SQLQuery & "shttPhone = '" & slPhone & "',"
                            End If
                            If slZone <> "" Then
                                SQLQuery = SQLQuery & "shttTimeZone = '" & slZone & "',"
                            End If
                            If slPDName <> "" Then
                                SQLQuery = SQLQuery & "shttPDName = '" & slPDName & "',"
                            End If
                            If slPDPhone <> "" Then
                                SQLQuery = SQLQuery & "shttPDPhone = '" & slPDPhone & "',"
                            End If
                            If slTDName <> "" Then
                                SQLQuery = SQLQuery & "shttTDName = '" & slTDName & "',"
                            End If
                            If slTDPhone <> "" Then
                                SQLQuery = SQLQuery & "shttTDPhone = '" & slTDPhone & "',"
                            End If
                            If slACName <> "" Then
                                SQLQuery = SQLQuery & "shttACName = '" & slACName & "',"
                            End If
                            If slACPhone <> "" Then
                                SQLQuery = SQLQuery & "shttACPhone = '" & slACPhone & "',"
                            End If
                            'SQLQuery = SQLQuery & "shttPDPhone = '" & "',"
                            SQLQuery = SQLQuery & "shttChecked =" & ilChecked & ","
                            If slMarket <> "" Then
                                SQLQuery = SQLQuery & "shttMarket = '" & slMarket & "',"
                                SQLQuery = SQLQuery & "shttRank = " & ilRank & ","
                            End If
                            SQLQuery = SQLQuery & "shttEnterDate = '" & Format$(slCurDate, sgSQLDateForm) & "',"
                            SQLQuery = SQLQuery & "shttEnterTime = '" & Format$(slCurTime, sgSQLTimeForm) & "',"
                            If slONAddr1 <> "" Then
                                SQLQuery = SQLQuery & "shttONAddress1 = '" & slONAddr1 & "',"
                            End If
                            If slONAddr2 <> "" Then
                                SQLQuery = SQLQuery & "shttONAddress2 = '" & slONAddr2 & "',"
                            End If
                            If slONCity <> "" Then
                                SQLQuery = SQLQuery & "shttONCity = '" & slONCity & "',"
                            End If
                            If slONState <> "" Then
                                SQLQuery = SQLQuery & "shttONState = '" & slONState & "',"
                            End If
                            If slONZip <> "" Then
                                SQLQuery = SQLQuery & "shttONZip = '" & slONZip & "',"
                            End If
                            SQLQuery = SQLQuery & "shttStationID =" & llStationID & ","
                            If slLicCity <> "" Then
                                SQLQuery = SQLQuery & "shttCityLic = '" & slLicCity & "',"
                            End If
                            If slLicState <> "" Then
                                SQLQuery = SQLQuery & "shttStateLic = '" & slLicState & "',"
                            End If
                            SQLQuery = SQLQuery & "shttAckDaylight =" & ilDaylight
                            SQLQuery = SQLQuery & " WHERE (shttCode = " & ilCode & ")"
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/11/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "ImportCSV-mReadFileUSRN"
                                cnn.RollbackTrans
                                ilRet = 1
                            End If
                            If ilRet = 0 Then
                                cnn.CommitTrans
                            End If
                            tgStationInfo(ilIndex).sCallLetters = slCallLetters
                            tgStationInfo(ilIndex).sMarket = slMarket
                        End If
                    End If
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
            llPercent = (lmProcessedNoBytes / lmTotalNoBytes) * 100
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmFloodPercent <> llPercent Then
                lmFloodPercent = llPercent
                lbcPercent.Caption = Str$(llPercent) & "%"
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    If ilRet <> 0 Then
        mReadFileUSRN = False
    Else
        mReadFileUSRN = True
        lmFloodPercent = 100
        lbcPercent.Caption = "100%"
    End If
    Print #hmMsg, "** Import Station Info Finished: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Print #hmMsg, ""
    '11/26/17: Set Changed date/time
    gFileChgdUpdate "shtt.mkd", True
    Exit Function
mReadFileUSRNErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportCSV-"
End Function

Private Sub mProcessPledge(slAttPledgeLines() As String, llLineNo() As Long)
    Dim ilPledge As Integer
    Dim ilRet As Integer
    Dim llAttCode As Long
    Dim ilAttVefCode As Integer
    Dim ilAttShttCode As Integer
    Dim ilVefCode As Integer
    Dim ilShttCode As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilDat As Integer
    Dim llFdSTime As Long
    Dim llFdETime As Long
    Dim slCallLetters As String
    Dim slVehicleName As String
    Dim llStationID As Long
    Dim ilSelected As Integer
    Dim slOnAir As String
    Dim ilStatus As Integer
    Dim slFdDate As String
    Dim slFdTime As String
    Dim slPdDate As String
    Dim slPdTime As String
    Dim slFdSTime As String
    Dim slFdETime As String
    Dim slPdSTime As String
    Dim slPdETime As String
    Dim slEstimedTime As String
    Dim llDATCode As Long
    Dim ilSeqNo As Integer
    Dim llCode As Long
    Dim ilDay As Integer
    Dim ilDayMatch As Integer
    Dim ilTimeMatch As Integer
    Dim ilAvailCount As Integer
    Dim ilAirPlayNumber As Integer
    Dim ilField As Integer
    Dim ilUpper As Integer
    Dim ilMerge As Integer
    ReDim slFields(0 To 60) As String
    Dim VehCombo_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    If UBound(slAttPledgeLines) <= LBound(slAttPledgeLines) Then
        Exit Sub
    End If

    gParseCDFields slAttPledgeLines(0), False, smFields()
    For ilLoop = LBound(smFields) To UBound(smFields) Step 1
        smFields(ilLoop) = Trim$(smFields(ilLoop))
    Next ilLoop
    For ilLoop = UBound(smFields) - 1 To LBound(smFields) Step -1
        smFields(ilLoop + 1) = Trim$(smFields(ilLoop))
    Next ilLoop
    smFields(0) = ""
    ilRet = 0
    llStationID = Val(smFields(2))
    llAttCode = Val(smFields(5))
    ilVefCode = Val(smFields(4))
    
    ilSelected = False
    For ilLoop = 0 To lbcNames.ListCount - 1 Step 1
        If ilVefCode = lbcNames.ItemData(ilLoop) Then
            If lbcNames.Selected(ilLoop) Then
                slVehicleName = lbcNames.List(ilLoop)
                ilSelected = True
            End If
            Exit For
        End If
    Next ilLoop
    If Not ilSelected Then
        Exit Sub
    End If
    SQLQuery = "SELECT *"
    SQLQuery = SQLQuery + " FROM att"
    SQLQuery = SQLQuery & " WHERE (attCode = " & llAttCode & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF Then
        lbcError.AddItem "Agreement Missing: " & smFields(1) & " " & smFields(3) & " Agreement " & smFields(5)
        Print #hmMsg, "Agreement Missing," & smFields(1) & "," & smFields(3) & ",Agreement," & smFields(5) & "," & "Line," & llLineNo(0) & "-" & llLineNo(UBound(llLineNo) - 1)
        Exit Sub
    End If
    ilAttVefCode = rst!attvefCode
    ilAttShttCode = rst!attshfCode
    slOnAir = Format(rst!attOnAir, sgShowDateForm)
    ilAirPlayNumber = 0

    SQLQuery = "SELECT Count(datCode)"
    SQLQuery = SQLQuery + " FROM dat"
    SQLQuery = SQLQuery & " WHERE (datAtfCode = " & llAttCode & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If rst(0).Value > 0 Then
        lbcError.AddItem "Pledge Information previously defined: " & smFields(1) & " " & smFields(3) & " Agreement " & smFields(5)
        Print #hmMsg, "Pledge Information previously defined," & smFields(1) & "," & smFields(3) & ",Agreement," & smFields(5) & "," & "Line," & llLineNo(0) & "-" & llLineNo(UBound(llLineNo) - 1)
        Exit Sub
    End If
    
    ilShttCode = -1
    For ilLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        'If StrComp(Trim$(tgStationInfo(ilLoop).sCallLetters), slCallLetters, vbTextCompare) = 0 Then
        If tgStationInfo(ilLoop).lPermStationID = llStationID Then
            ilShttCode = tgStationInfo(ilLoop).iCode
            slCallLetters = Trim$(tgStationInfo(ilLoop).sCallLetters)
            Exit For
        End If
    Next ilLoop
    If ilShttCode < 0 Then
        ilFound = False
        For ilLoop = 0 To UBound(smMissingShtt) - 1 Step 1
            If StrComp(Trim$(smFields(1)) & " " & Trim$(smFields(2)), smMissingShtt(ilLoop), 1) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            lbcError.AddItem "Station Missing: " & Trim$(smFields(1)) & " " & Trim$(smFields(2))
            Print #hmMsg, "Station Missing," & Trim$(smFields(1)) & "," & Trim$(smFields(2)) & "," & "Line," & llLineNo(0) & "-" & llLineNo(UBound(llLineNo) - 1)
            smMissingShtt(UBound(smMissingShtt)) = smFields(1) & " " & Trim$(smFields(2))
            ReDim Preserve smMissingShtt(0 To UBound(smMissingShtt) + 1) As String
        End If
        Exit Sub
    End If
    If ilAttVefCode <> ilVefCode Then
        ilFound = False
        For ilLoop = 0 To UBound(smMissingVef) - 1 Step 1
            If StrComp(Trim$(smFields(3)) & " " & Trim$(smFields(4)), smMissingVef(ilLoop), 1) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            lbcError.AddItem "Vehicle Missing: " & Trim$(smFields(3)) & " " & Trim$(smFields(4))
            Print #hmMsg, "Vehicle Missing," & Trim$(smFields(3)) & "," & Trim$(smFields(4)) & "," & "Line," & llLineNo(0) & "-" & llLineNo(UBound(llLineNo) - 1)
            smMissingVef(UBound(smMissingVef)) = Trim$(smFields(3)) & " " & Trim$(smFields(4))
            ReDim Preserve smMissingVef(0 To UBound(smMissingVef) + 1) As String
        End If
        Exit Sub
    End If

    If ilAttShttCode <> ilShttCode Then
        ilFound = False
        For ilLoop = 0 To UBound(smStationNotMatching) - 1 Step 1
            If StrComp(Trim$(smFields(1)) & " " & Trim$(smFields(2)), smStationNotMatching(ilLoop), 1) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            lbcError.AddItem "Agreement and Station not matching: " & Trim$(smFields(1)) & " " & Trim$(smFields(2))
            Print #hmMsg, "Agreement and Station not matching," & Trim$(smFields(1)) & "," & Trim$(smFields(2)) & "," & "Line," & llLineNo(0) & "-" & llLineNo(UBound(llLineNo) - 1)
            smStationNotMatching(UBound(smStationNotMatching)) = Trim$(smFields(1)) & " " & Trim$(smFields(2))
            ReDim Preserve smStationNotMatching(0 To UBound(smStationNotMatching) + 1) As String
        End If
        Exit Sub
    End If
    
    'Warning
    If UCase(smFields(1)) <> UCase(slCallLetters) Then
        'Output Warning message
        ilFound = False
        For ilLoop = 0 To UBound(smStationNameError) - 1 Step 1
            If StrComp(Trim$(smFields(1)), smStationNameError(ilLoop), 1) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            lbcError.AddItem "Call Letter Name not matching: " & Trim$(smFields(1))
            Print #hmMsg, "Call Letter Name not matching," & Trim$(smFields(1)) & "," & "Line," & llLineNo(0) & "-" & llLineNo(UBound(llLineNo) - 1)
            smStationNameError(UBound(smStationNameError)) = Trim$(smFields(1))
            ReDim Preserve smStationNameError(0 To UBound(smStationNameError) + 1) As String
        End If
    End If

     If UCase(smFields(3)) <> UCase(slVehicleName) Then
        'Output warning message
        ilFound = False
        For ilLoop = 0 To UBound(smVehicleNameError) - 1 Step 1
            If StrComp(Trim$(smFields(3)), smVehicleNameError(ilLoop), 1) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            lbcError.AddItem "Vehicle Name not matching: " & Trim$(smFields(3))
            Print #hmMsg, "Vehicle Name not matching," & Trim$(smFields(3)) & "," & "Line," & llLineNo(0) & "-" & llLineNo(UBound(llLineNo) - 1)
            smVehicleNameError(UBound(smVehicleNameError)) = Trim$(smFields(3))
            ReDim Preserve smVehicleNameError(0 To UBound(smVehicleNameError) + 1) As String
        End If
    End If
   
    ReDim tgDat(0 To 0) As DAT
    SQLQuery = "Select vefCombineVefCode from VEF_Vehicles Where vefCode = " & ilAttVefCode
    Set VehCombo_rst = gSQLSelectCall(SQLQuery)
    If Not VehCombo_rst.EOF Then
        imVefCombo = VehCombo_rst!vefCombineVefCode
    End If
    gGetAvails llAttCode, ilAttShttCode, ilAttVefCode, imVefCombo, slOnAir, True
    
    '6/16/14: Merge records matching days and airplay so that they can be compared to Avails
    For ilDat = LBound(tgDat) To UBound(tgDat) - 1 Step 1
        For ilDay = 0 To 6 Step 1
            If tgDat(ilDat).iFdDay(ilDay) <> 1 Then
                tgDat(ilDat).iFdDay(ilDay) = 0
            End If
        Next ilDay
    Next ilDat
    ilUpper = 0
    ReDim slMergeAttPledgeLines(0 To UBound(slAttPledgeLines)) As String
    ReDim llMergeLineNo(0 To UBound(slAttPledgeLines)) As Long
    For ilPledge = 0 To UBound(slAttPledgeLines) - 1 Step 1
        ilMerge = False
        gParseCDFields slAttPledgeLines(ilPledge), False, smFields()
        For ilLoop = LBound(smFields) To UBound(smFields) Step 1
            smFields(ilLoop) = Trim$(smFields(ilLoop))
        Next ilLoop
        For ilLoop = UBound(smFields) - 1 To LBound(smFields) Step -1
            smFields(ilLoop + 1) = Trim$(smFields(ilLoop))
        Next ilLoop
        smFields(0) = ""
        If Len(smFields(14)) <> 0 Then  'Feed time defined?
            llFdSTime = gTimeToLong(Format$(smFields(14), "h:mm:ssam/pm"), False)
            llFdETime = gTimeToLong(Format$(smFields(15), "h:mm:ssam/pm"), False)
            For ilLoop = 0 To ilUpper - 1 Step 1
                gParseCDFields slMergeAttPledgeLines(ilLoop), False, slFields()
                For ilField = LBound(slFields) To UBound(slFields) Step 1
                    slFields(ilField) = Trim$(slFields(ilField))
                Next ilField
                For ilField = UBound(slFields) - 1 To LBound(slFields) Step -1
                    slFields(ilField + 1) = Trim$(slFields(ilField))
                Next ilField
                slFields(0) = ""
                If Len(slFields(14)) <> 0 Then  'Feed time defined?
                    'If (llFdSTime = gTimeToLong(Format$(slFields(14), "h:mm:ssam/pm"), False)) And (Val(smFields(6)) = Val(slFields(6))) Then
                    If (llFdSTime = gTimeToLong(Format$(slFields(14), "h:mm:ssam/pm"), False)) And (llFdETime = gTimeToLong(Format$(slFields(15), "h:mm:ssam/pm"), False)) And (Val(smFields(6)) = Val(slFields(6))) Then
                        ilMerge = True
                        'Merge days
                        For ilDay = 7 To 13 Step 1
                            If (smFields(ilDay) = "1") And (slFields(ilDay) = "1") Then
                                lbcError.AddItem "Import Pledge days in conflict: " & smFields(1) & " " & smFields(3) & " Agreement " & smFields(5)
                                Print #hmMsg, "Import Pledge days in conflict," & smFields(1) & "," & smFields(3) & ",Agreement," & smFields(5) & "," & "Line," & llLineNo(ilPledge) & "," & llMergeLineNo(ilLoop)
                                Exit Sub
                            End If
                            If smFields(ilDay) = "1" Then
                                slFields(ilDay) = "1"
                            End If
                        Next ilDay
                        slMergeAttPledgeLines(ilLoop) = """" & slFields(1) & """"
                        For ilField = 2 To UBound(slFields) Step 1
                            If ilField = 3 Then
                                slMergeAttPledgeLines(ilLoop) = slMergeAttPledgeLines(ilLoop) & "," & """" & slFields(ilField) & """"
                            Else
                                slMergeAttPledgeLines(ilLoop) = slMergeAttPledgeLines(ilLoop) & "," & slFields(ilField)
                            End If
                        Next ilField
                    End If
                End If
            Next ilLoop
        End If
        If Not ilMerge Then
            slMergeAttPledgeLines(ilUpper) = """" & smFields(1) & """"
            For ilField = 2 To UBound(smFields) Step 1
                If ilField = 3 Then
                    slMergeAttPledgeLines(ilUpper) = slMergeAttPledgeLines(ilUpper) & "," & """" & smFields(ilField) & """"
                Else
                    slMergeAttPledgeLines(ilUpper) = slMergeAttPledgeLines(ilUpper) & "," & smFields(ilField)
                End If
            Next ilField
            llMergeLineNo(ilUpper) = llLineNo(ilPledge)
            ilUpper = ilUpper + 1
        End If
    Next ilPledge
    ReDim Preserve slMergeAttPledgeLines(0 To ilUpper) As String
    'Check that pledge information matches import pledge information
    ilAvailCount = 0
    For ilPledge = 0 To UBound(slMergeAttPledgeLines) - 1 Step 1
        gParseCDFields slMergeAttPledgeLines(ilPledge), False, smFields()
        For ilLoop = UBound(smFields) - 1 To LBound(smFields) Step -1
            smFields(ilLoop + 1) = Trim$(smFields(ilLoop))
        Next ilLoop
        If Val(Trim$(smFields(6))) = 1 Then
            ilAvailCount = ilAvailCount + 1
        End If
    Next ilPledge
    If ilAvailCount <> UBound(tgDat) Then
        lbcError.AddItem "Pledge Break Count Information not matching: " & smFields(1) & " " & smFields(3) & " Agreement " & smFields(5)
        Print #hmMsg, "Pledge Break Count Information not matching," & smFields(1) & "," & smFields(3) & ",Agreement," & smFields(5) & "," & "Line," & llLineNo(0) & "-" & llLineNo(UBound(llLineNo) - 1) & "," & "Import Count " & ilAvailCount & "," & "Agreement Count " & UBound(tgDat)
        Exit Sub
    End If
    For ilPledge = 0 To UBound(slMergeAttPledgeLines) - 1 Step 1
        gParseCDFields slMergeAttPledgeLines(ilPledge), False, smFields()
        For ilLoop = LBound(smFields) To UBound(smFields) Step 1
            smFields(ilLoop) = Trim$(smFields(ilLoop))
        Next ilLoop
        For ilLoop = UBound(smFields) - 1 To LBound(smFields) Step -1
            smFields(ilLoop + 1) = Trim$(smFields(ilLoop))
        Next ilLoop
        If Len(smFields(14)) <> 0 Then
            'Truncate seconds
            ilTimeMatch = False
            ilFound = False
            llFdSTime = gTimeToLong(Format$(smFields(14), "h:mm:ssam/pm"), False)
            For ilDat = LBound(tgDat) To UBound(tgDat) - 1 Step 1
                If (llFdSTime >= gTimeToLong(tgDat(ilDat).sFdSTime, False)) And (llFdSTime <= gTimeToLong(tgDat(ilDat).sFdETime, False)) Then
                    ilTimeMatch = True
                    ilDayMatch = True
                    For ilDay = 0 To 6 Step 1
                        If Val(smFields(7 + ilDay)) <> tgDat(ilDat).iFdDay(ilDay) Then
                            ilDayMatch = False
                            Exit For
                        End If
                    Next ilDay
                    If ilDayMatch Then
                        ilFound = True
                        Exit For
                    End If
                End If
            Next ilDat
            If Not ilFound Then
                If ilTimeMatch Then
                    lbcError.AddItem "Pledge Feed Day Information not matching: " & slMergeAttPledgeLines(ilPledge)
                    Print #hmMsg, "Pledge Feed Day Information not matching," & slMergeAttPledgeLines(ilPledge) & "," & "Line," & llMergeLineNo(ilPledge)
                Else
                    lbcError.AddItem "Pledge Feed Time Information not matching: " & slMergeAttPledgeLines(ilPledge)
                    Print #hmMsg, "Pledge Feed Time Information not matching," & slMergeAttPledgeLines(ilPledge) & "," & "Line," & llMergeLineNo(ilPledge)
                End If
                Exit Sub
            End If
        Else
            'Output error
            lbcError.AddItem "Pledge Feed Time Information missing: " & slMergeAttPledgeLines(ilPledge)
            Print #hmMsg, "Pledge Feed Time Information missing," & slMergeAttPledgeLines(ilPledge) & "," & "Line," & llMergeLineNo(ilPledge)
            Exit Sub
        End If
    Next ilPledge
    '6/16/14: end of change
    
    'Process records as they passed all test
    For ilPledge = 0 To UBound(slAttPledgeLines) - 1 Step 1
        'Process Input
        gParseCDFields slAttPledgeLines(ilPledge), False, smFields()
        For ilLoop = LBound(smFields) To UBound(smFields) Step 1
            smFields(ilLoop) = Trim$(smFields(ilLoop))
        Next ilLoop
        For ilLoop = UBound(smFields) - 1 To LBound(smFields) Step -1
            smFields(ilLoop + 1) = Trim$(smFields(ilLoop))
        Next ilLoop
        smFields(0) = ""
        ilRet = 0
                
        If Len(smFields(14)) <> 0 Then
            'Truncate seconds
            slFdSTime = Format$(smFields(14), "h:mm:ssam/pm")
        Else
            slFdSTime = ""
        End If
        If Len(smFields(15)) <> 0 Then
            slFdETime = Format$(smFields(15), "h:mm:ssam/pm")
        Else
            slFdETime = Format$(slFdSTime, "h:mm:ssam/pm")
        End If
        ilStatus = Val(smFields(16))
        If ilStatus > 0 Then
            ilStatus = ilStatus - 1
        End If
        If ilStatus = 6 Then 'Change status 7 to 2
            ilStatus = 1
        End If
        If Len(smFields(24)) <> 0 Then
            slPdSTime = Format$(smFields(24), "h:mm:ssam/pm")
        Else
            slPdSTime = slFdSTime
        End If
        If Len(smFields(25)) <> 0 Then
            slPdETime = Format$(smFields(25), "h:mm:ssam/pm")
        Else
            slPdETime = slPdSTime
        End If

        slEstimedTime = "N"
        For ilLoop = 26 To UBound(smFields) Step 3
            If (smFields(ilLoop) <> "") Then
                slEstimedTime = "Y"
                Exit For
            End If
        Next ilLoop
        If Val(smFields(6)) > ilAirPlayNumber Then
            ilAirPlayNumber = Val(smFields(6))
        End If
        'Add to Agreement
        SQLQuery = "INSERT INTO dat (datCode, datAtfCode, datShfCode, datVefCode, "
        'SQLQuery = SQLQuery & "datDACode, datFdMon, datFdTue, "
        SQLQuery = SQLQuery & "datFdMon, datFdTue, "
        SQLQuery = SQLQuery & "datFdWed, datFdThu, datFdFri, "
        SQLQuery = SQLQuery & "datFdSat, datFdSun, datFdStTime, "
        SQLQuery = SQLQuery & "datFdEdTime, datFdStatus, datPdMon, "
        SQLQuery = SQLQuery & "datPdTue, datPdWed, datPdThu, "
        SQLQuery = SQLQuery & "datPdFri, datPdSat, datPdSun, "
        SQLQuery = SQLQuery & "datPdStTime, datPdEdTime, "
        SQLQuery = SQLQuery & "datPdDayFed, datAirPlayNo, datEstimatedTime" & ")"
        SQLQuery = SQLQuery & " VALUES (" & "Replace" & "," & llAttCode & "," & ilShttCode & "," & ilVefCode & ","
        'SQLQuery = SQLQuery & "1" & "," & Val(smFields(4)) & "," & Val(smFields(5)) & ","
        SQLQuery = SQLQuery & Val(smFields(7)) & "," & Val(smFields(8)) & ","
        SQLQuery = SQLQuery & Val(smFields(9)) & "," & Val(smFields(10)) & "," & Val(smFields(11)) & ","
        SQLQuery = SQLQuery & Val(smFields(12)) & "," & Val(smFields(13)) & ",'" & Format$(slFdSTime, sgSQLTimeForm) & "',"
        SQLQuery = SQLQuery & "'" & Format$(slFdETime, sgSQLTimeForm) & "'," & ilStatus & "," & Val(smFields(17)) & ","
        SQLQuery = SQLQuery & Val(smFields(18)) & "," & Val(smFields(19)) & "," & Val(smFields(20)) & ","
        SQLQuery = SQLQuery & Val(smFields(21)) & "," & Val(smFields(22)) & "," & Val(smFields(23)) & ","
        SQLQuery = SQLQuery & "'" & Format$(slPdSTime, sgSQLTimeForm) & "','" & Format$(slPdETime, sgSQLTimeForm) & "',"
        SQLQuery = SQLQuery & "'" & "A" & "', " & Val(smFields(6)) & ",'" & slEstimedTime & "')"
        llDATCode = gInsertAndReturnCode(SQLQuery, "dat", "datCode", "Replace")
        If llDATCode <= 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHand1:
            gHandleError "ImptPldg.Txt", "ImportCSV-mProcessPledge"
        End If
        On Error GoTo ErrHand
        If slEstimedTime = "Y" Then
            ilSeqNo = 1
            For ilLoop = 26 To UBound(smFields) Step 3
                If (smFields(ilLoop) <> "") Then
                    SQLQuery = "Insert Into ept ( "
                    SQLQuery = SQLQuery & "eptCode, "
                    SQLQuery = SQLQuery & "eptDatCode, "
                    SQLQuery = SQLQuery & "eptSeqNo, "
                    SQLQuery = SQLQuery & "eptAttCode, "
                    SQLQuery = SQLQuery & "eptShttCode, "
                    SQLQuery = SQLQuery & "eptVefCode, "
                    SQLQuery = SQLQuery & "eptFdAvailDay, "
                    SQLQuery = SQLQuery & "eptFdAvailTime, "
                    SQLQuery = SQLQuery & "eptEstimatedDay, "
                    SQLQuery = SQLQuery & "eptEstimatedTime, "
                    SQLQuery = SQLQuery & "eptUnused "
                    SQLQuery = SQLQuery & ") "
                    SQLQuery = SQLQuery & "Values ( "
                    SQLQuery = SQLQuery & "Replace" & ", "
                    SQLQuery = SQLQuery & llDATCode & ", "
                    SQLQuery = SQLQuery & ilSeqNo & ", "
                    SQLQuery = SQLQuery & llAttCode & ", "
                    SQLQuery = SQLQuery & ilShttCode & ", "
                    SQLQuery = SQLQuery & ilVefCode & ", "
                    SQLQuery = SQLQuery & "'" & gFixQuote(smFields(ilLoop)) & "', "
                    SQLQuery = SQLQuery & "'" & Format$(slFdSTime, sgSQLTimeForm) & "', "
                    If Trim$(smFields(ilLoop + 2)) <> "" Then
                        SQLQuery = SQLQuery & "'" & gFixQuote(smFields(ilLoop + 1)) & "', "
                        SQLQuery = SQLQuery & "'" & Format$(smFields(ilLoop + 2), sgSQLTimeForm) & "', "
                    Else
                        SQLQuery = SQLQuery & "'" & "" & "', "
                        SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "
                    End If
                    SQLQuery = SQLQuery & "'" & "" & "' "
                    SQLQuery = SQLQuery & ") "
                    llCode = gInsertAndReturnCode(SQLQuery, "ept", "eptCode", "Replace")
                End If
            Next ilLoop
        End If
    Next ilPledge
    SQLQuery = "UPDATE att"
    SQLQuery = SQLQuery & " SET attNoAirPlays = " & ilAirPlayNumber
    SQLQuery = SQLQuery & " WHERE (attCode = " & llAttCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "ImptPldg.Txt", "ImportCSV-mProcessPledge"
    End If
    Exit Sub
ErrHand:
    gHandleError "ImptPldg.Txt", "Import-mProcessPledge"
    ilRet = 1
    Resume Next
'ErrHand1:
'    gHandleError "ImptPldg.Txt", "Import-mProcessPledge"
'    Return
End Sub
