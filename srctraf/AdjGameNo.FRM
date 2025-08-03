VERSION 5.00
Begin VB.Form AdjGameNo 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Adjust Game #"
   ClientHeight    =   285
   ClientLeft      =   4260
   ClientTop       =   3645
   ClientWidth     =   1800
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
   Icon            =   "AdjGameNo.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   285
   ScaleWidth      =   1800
   WindowState     =   1  'Minimized
   Begin VB.Timer tmcTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3270
      Top             =   435
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2505
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   195
      Visible         =   0   'False
      Width           =   525
   End
End
Attribute VB_Name = "AdjGameNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: AdjGameNo.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Set Advertiser/Agency Credit code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim tmCgfSrchKey As LONGKEY0    'AGf key record image
Dim hmCgf As Integer    'Agency file handle
Dim imCgfRecLen As Integer        'Cif record length
Dim tmCgf As CGF
Dim tmClf As CLF        'Clf record image
Dim tmClfSrchKey2 As LONGKEY0    'Clf key record image
Dim hmClf As Integer    'Contract Line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim smNowDate As String
Dim lmNowDate As Long
Dim smCommandDate As String
Dim lmCgfNoRecs As Long
Dim imOkToUnload As Integer



Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
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
    igBkgdProg = 2
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
        ilRes = MsgBox("Select OK to Terminate Setting of Game Numbers, Select Cancel to Continue Setting Game Numbers and Stop the Exit from Windows", vbOKCancel + vbQuestion + vbDefaultButton2, "Set Credit")
        If ilRes = vbCancel Then
            Cancel = 1  'Any non-zero value stop the unload
        Else
            'Do
            '    ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            '    tmSpf.sLkCredit = ""
            '    ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
            'Loop While ilRet = BTRV_ERR_CONFLICT
        End If
        'ilRet = btrUnlock(hmUrf, BTRV_UNLOCK_MULTIPLE)
    Else
        'Do
        '    ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        '    tmSpf.sLkCredit = ""
        '    ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
        'Loop While ilRet = BTRV_ERR_CONFLICT
        ''ilRet = btrUnlock(hmUrf, BTRV_UNLOCK_MULTIPLE)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    'Reset used instead of Close to cause # Clients on network to be decrement
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    btrStopAppl
    Set AdjGameNo = Nothing   'Remove data segment
    End
End Sub
Private Sub mSetGameNo()
    Dim ilRet As Integer
    
    'Update Cgf
    ilRet = btrGetFirst(hmCgf, tmCgf, imCgfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        If tmCgf.iGameNo >= 132 Then
            tmClfSrchKey2.lCode = tmCgf.lClfCode
            ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                If tmClf.iVefCode = 16 Then
                    tmCgf.iGameNo = tmCgf.iGameNo + 1
                    ilRet = btrUpdate(hmCgf, tmCgf, imCgfRecLen)
                End If
            End If
        End If
        ilRet = btrGetNext(hmCgf, tmCgf, imCgfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
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
    'AdjGameNo.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    'gCenterStdAlone AdjGameNo
    'AdjGameNo.Show
    'Screen.MousePointer = vbHourGlass
'    mInitDDE
    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    
    If Not mCheckRecLength() Then
        imTerminate = True
        Exit Sub
    End If
    
    imTerminate = False
    imFirstFocus = True
    smNowDate = Format$(Now, "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    
    hmCgf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmCgf, "", sgDBPath & "Cgf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cgf.Btr)", AdjGameNo
    On Error GoTo 0
    imCgfRecLen = Len(tmCgf)
    lmCgfNoRecs = btrRecords(hmCgf) - 1
    If lmCgfNoRecs < 0 Then
        imTerminate = True
        Exit Sub
    End If
    hmClf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", AdjGameNo
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    If imTerminate Then
        Exit Sub
    End If
'    gCenterModalForm AdjGameNo
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitDDE                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize DDE                 *
'*                                                     *
'*******************************************************
Private Sub mInitDDE()
'    Dim ilRet As Integer
'    Dim ilLoop As Integer
'    If Not igStdAloneMode Then
'        'Set up DDE for Branner and help message area within traffic
'        edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
'        edcLinkDestHelpMsg.LinkTopic = "Traffic|HelpMsg"
'        edcLinkDestHelpMsg.LinkItem = "edcLinkSrceHelpMsg"
'        edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
'    End If
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
    gInitStdAlone AdjGameNo, slStr, ilTestSystem
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
    Dim ilRet As Integer
    btrExtClear hmClf   'Clear any previous extend operation
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    btrExtClear hmCgf   'Clear any previous extend operation
    ilRet = btrClose(hmCgf)
    btrDestroy hmCgf
    igParentRestarted = False
    If Not igStdAloneMode Then
        'If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
        '    edcLinkDestHelpMsg.LinkExecute "@" & "Done"
        'Else
        '    edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
        '    edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
        '    edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
        '    edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
        '    edcLinkDestHelpMsg.LinkExecute "Done"
        'End If
        'Do While Not igParentRestarted
        '    DoEvents
        'Loop
    End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Traffic
    Unload AdjGameNo
    igManUnload = NO
End Sub
Private Sub tmcTimer_Timer()
    Dim ilLoop As Integer
    tmcTimer.Enabled = False
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
    imOkToUnload = False
    mSetGameNo
    imOkToUnload = True
    mTerminate
End Sub


Private Function mCheckRecLength() As Integer
    If Not gRecLengthOk("Cgf.Btr", Len(tmCgf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Clf.Btr", Len(tmClf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    mCheckRecLength = True
End Function

