VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ResetMGSpec 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Reset MG Specifications"
   ClientHeight    =   1965
   ClientLeft      =   4260
   ClientTop       =   3645
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "ResetMGSpec.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1965
   ScaleWidth      =   4680
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2475
      TabIndex        =   1
      Top             =   1260
      Width           =   1050
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "&Reset"
      Height          =   285
      Left            =   1110
      TabIndex        =   0
      Top             =   1245
      Width           =   1050
   End
   Begin VB.Timer tmcTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3885
      Top             =   1500
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   285
      Left            =   885
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lacMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   195
      TabIndex        =   3
      Top             =   315
      Width           =   4170
   End
End
Attribute VB_Name = "ResetMGSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ResetMGSpec.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Set Advertiser/Agency Credit code
Option Explicit
Option Compare Text
Dim lmTotalNoRecs As Long
Dim lmProcessedNoRecs As Long
'Spot file
Dim hmSdf As Integer
Dim tmSdf As SDF            'SDF record image
Dim tmSdfSrchKey0 As SDFKEY0 'SDF key record image (agency)
Dim tmSdfSrchKey1 As SDFKEY1 'SDF key record image (advt)
Dim tmSdfSrchKey2 As SDFKEY2 'SDF key record image (advt)
Dim tmSdfSrchKey3 As LONGKEY0 'SDF key record image (code)
Dim imSdfRecLen As Integer     'SDF record length
Dim hmSmf As Integer    'file handle
Dim imSmfRecLen As Integer  'Record length
Dim tmSmf As SMF
Dim tmSmfSrchKey1 As LONGKEY0
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim smNowDate As String
Dim lmNowDate As Long
Dim smCommandDate As String
Dim imOkToUnload As Integer
Dim lmSdfNoRecs As Long
Dim lmSmfNoRecs As Long

Private Sub cmcCancel_Click()
    Dim ilRes As Integer
    If Not imOkToUnload Then
        ilRes = MsgBox("Select OK to Terminate Resetting MG Specifications, Select Cancel to Continue Setting MG Specifications", vbOKCancel + vbQuestion + vbDefaultButton2, "Set Credit")
        If ilRes = vbCancel Then
            Exit Sub
        Else
            imTerminate = True
            Exit Sub
        End If
    End If
    mTerminate
End Sub

Private Sub cmcGen_Click()
    Screen.MousePointer = vbHourglass
    plcGauge.Visible = True
    imOkToUnload = False
    mResetMGSpec
    imOkToUnload = True
    Screen.MousePointer = vbDefault
    If imTerminate Then
        lacMsg.Caption = "Reset MG Spec: Canceled"
        cmcCancel.SetFocus
    Else
        lacMsg.Caption = "Reset MG Spec: Completed"
        cmcCancel.Caption = "&Done"
        cmcCancel.SetFocus
        cmcGen.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
'    gShowBranner
End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    If App.PrevInstance Then
        End
    End If
    gSetBkgdMode
    igBkgdProg = 13
    mInit
'    If imTerminate Then
'        mTerminate
'    Else
        tmcTimer.Enabled = True
'    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ilRes As Integer
    Dim ilRet As Integer
    Cancel = 0
    If (Not imOkToUnload) And (UnloadMode > 1) Then
        ilRes = MsgBox("Select OK to Terminate Resetting MG Specifications, Select Cancel to Continue Setting MG Specifications", vbOKCancel + vbQuestion + vbDefaultButton2, "Set Credit")
        If ilRes = vbCancel Then
            Cancel = 1  'Any non-zero value stop the unload
        Else
        End If
    Else
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    btrExtClear hmSdf   'Clear any previous extend operation
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    btrExtClear hmSmf   'Clear any previous extend operation
    ilRet = btrClose(hmSmf)
    btrDestroy hmSmf
    btrStopAppl
    Set ResetMGSpec = Nothing   'Remove data segment
    End
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim llAdfNoRecs As Integer
    Dim llAgfNoRecs As Integer
    imOkToUnload = True
    imTerminate = False
    
    'Screen.MousePointer = vbHourGlass
    mParseCmmdLine
    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    
    If Not mCheckRecLength() Then
        imTerminate = True
        Exit Sub
    End If
    
    lacMsg.BackColor = Me.BackColor
    imTerminate = False
    imFirstFocus = True
    smNowDate = Format$(Now, "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    hmSdf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", ResetMGSpec
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    lmSdfNoRecs = btrRecords(hmSdf) - 1
    If lmSdfNoRecs < 0 Then
        imTerminate = True
        Exit Sub
    End If
    hmSmf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Smf.Btr)", ResetMGSpec
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)
    lmSmfNoRecs = btrRecords(hmSmf) - 1
    If lmSmfNoRecs < 0 Then
        imTerminate = True
        Exit Sub
    End If
    lmTotalNoRecs = 2 * lmSmfNoRecs
    gCenterStdAlone ResetMGSpec

    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    slCommand = Command$
    sgCommandStr = "debug"
    'slCommand = sgCommandStr  'Command$
    'If (StrComp(slCommand, "Debug", 1) = 0) Or (StrComp(slCommand, "CS", 1) = 0) Then
        igStdAloneMode = True 'Switch from/to stand alone mode
        sgCallAppName = ""
        slStr = "Guide"
        ilTestSystem = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
    '    ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
    '    If Trim$(slStr) = "" Then
    '        MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
    '        End
    '    End If
    '    ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
    '    ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
    '    If StrComp(slTestSystem, "Test", 1) = 0 Then
    '        ilTestSystem = True
    '    Else
    '        ilTestSystem = False
    '    End If
    '    ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    ilPos = InStr(1, slCommand, "/IniLoc:", vbTextCompare)
    If ilPos > 0 Then
        igDirectCall = -1
        ilSpace = InStr(ilPos, slCommand, " ")
        If ilSpace = 0 Then
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8))
        Else
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8, ilSpace - ilPos - 3))
        End If
        If right$(sgIniPath, 1) <> "\" Then
            sgIniPath = sgIniPath & "\"
        End If
    End If
    gInitStdAlone ResetMGSpec, slStr, ilTestSystem
    smCommandDate = ""
    ilPos = InStr(1, slCommand, "/D:", 1)
    If ilPos > 0 Then
        ilSpace = InStr(ilPos, slCommand, " ")
        If ilSpace = 0 Then
            smCommandDate = Trim$(Mid$(slCommand, ilPos + 3))
        Else
            smCommandDate = Trim$(Mid$(slCommand, ilPos + 3, ilSpace - ilPos - 3))
        End If
        If gValidDate(smCommandDate) Then
            smCommandDate = gAdjYear(smCommandDate)
        Else
            smCommandDate = ""
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'
 
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Traffic
    Unload ResetMGSpec
    igManUnload = NO
End Sub
Private Sub tmcTimer_Timer()
    Dim ilLoop As Integer
    tmcTimer.Enabled = False
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
End Sub



Private Function mCheckRecLength() As Integer
    If Not gRecLengthOk("Sdf.Btr", Len(tmSdf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Smf.Btr", Len(tmSmf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    mCheckRecLength = True
End Function


Private Sub mResetMGSpec()
    Dim ilPass As Integer
    Dim llMissedDate As Long
    Dim ll12_26_11 As Long
    Dim ilRet As Integer
    Dim llPercent As Long
    Dim llPrevPercent As Long
    Dim llPrevSmfCode As Long
    Dim ilRecRemoved As Long
    Dim blFound As Boolean
    Dim llSdf As Long
    ReDim llSdfCode(0 To 0) As Long
    
    ll12_26_11 = gDateValue("12/26/11")
    lmProcessedNoRecs = 0
    llPrevPercent = -1
    llPrevSmfCode = 0
    For ilPass = 0 To 1 Step 1
        ilRet = btrGetFirst(hmSmf, tmSmf, imSmfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record
        Do While ilRet = BTRV_ERR_NONE
            DoEvents
            If imTerminate Then
                Exit Sub
            End If
            ilRecRemoved = False
            gUnpackDateLong tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), llMissedDate
            If llMissedDate >= ll12_26_11 Then
                Do
                    
                    blFound = False
                    If ilPass = 0 Then
                        For llSdf = 0 To UBound(llSdfCode) - 1 Step 1
                            If llSdfCode(llSdf) = tmSmf.lSdfCode Then
                                blFound = True
                                Exit For
                            End If
                        Next llSdf
                    End If
                    If Not blFound Then
                        tmSdfSrchKey3.lCode = tmSmf.lSdfCode
                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilPass = 0 Then
                            If tmSdf.lSmfCode <> tmSmf.lCode Then
                                tmSdf.lSmfCode = tmSmf.lCode
                                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                                llSdfCode(UBound(llSdfCode)) = tmSmf.lSdfCode
                                ReDim Preserve llSdfCode(0 To UBound(llSdfCode) + 1) As Long
                            End If
                        Else
                            If tmSdf.lSmfCode <> tmSmf.lCode Then
                                ilRet = btrDelete(hmSmf)
                                ilRecRemoved = True
                            End If
                        End If
                    Else
                        ilRet = BTRV_ERR_NONE
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    gLogBtrError ilRet, "mResetMGSpec: Pass #: " & ilPass
                End If
            End If
            If ilPass = 1 Then
                If ilRecRemoved Then
                    If llPrevSmfCode = 0 Then
                        ilRet = btrGetFirst(hmSmf, tmSmf, imSmfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record
                    Else
                        tmSmfSrchKey1.lCode = llPrevSmfCode
                        ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    End If
                Else
                    llPrevSmfCode = tmSmf.lCode
                End If
            End If
            lmProcessedNoRecs = lmProcessedNoRecs + 1
            lacMsg.Caption = lmProcessedNoRecs & " of " & lmTotalNoRecs
            llPercent = (lmProcessedNoRecs * 100) \ lmTotalNoRecs
            If llPercent >= 100 Then
                llPercent = 100
            End If
            If llPercent <> llPrevPercent Then
                plcGauge.Value = llPercent
            End If
            llPrevPercent = llPercent
            ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get next record
        Loop
    Next ilPass
End Sub
