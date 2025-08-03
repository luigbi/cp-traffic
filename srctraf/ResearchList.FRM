VERSION 5.00
Begin VB.Form ResearchList 
   Appearance      =   0  'Flat
   Caption         =   "Import Research"
   ClientHeight    =   5085
   ClientLeft      =   585
   ClientTop       =   2880
   ClientWidth     =   4410
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
   Icon            =   "ResearchList.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5085
   ScaleWidth      =   4410
   Begin VB.Timer tmcAutoRun 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   45
      Top             =   4530
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3630
      Top             =   4545
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2415
      TabIndex        =   4
      Top             =   4575
      Width           =   1050
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   1770
      Width           =   75
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      ScaleHeight     =   240
      ScaleWidth      =   5010
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -30
      Width           =   5010
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "C&ontinue"
      Default         =   -1  'True
      Height          =   285
      Left            =   705
      TabIndex        =   3
      Top             =   4575
      Width           =   1035
   End
   Begin VB.PictureBox plcExport 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4200
      Left            =   105
      ScaleHeight     =   4140
      ScaleWidth      =   4035
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   225
      Width           =   4095
      Begin VB.ListBox lbcResearch 
         Appearance      =   0  'Flat
         Height          =   3810
         ItemData        =   "ResearchList.frx":08CA
         Left            =   105
         List            =   "ResearchList.frx":08CC
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   3825
      End
   End
End
Attribute VB_Name = "ResearchList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ResearchList.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ResearchList.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim smPassCommands As String
Dim smScreenCaption As String

Dim hmSaf As Integer
Dim tmSaf As SAF            'Schedule Attributes record image
Dim imSafRecLen As Integer

Dim imIgnoreChg As Integer
Dim smCommandKeys As String
Dim bmInModalModule As Boolean
Dim smImportName As String
Dim bmAutoMode As Boolean





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
Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slStartIn As String
    Dim ilPos As Integer
    Dim ilLen As Integer

    sgCommandStr = Command$
    slStartIn = CurDir$
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
    igShowVersionNo = 0
    If (InStr(1, slStartIn, "Prod", vbTextCompare) = 0) And (InStr(1, slStartIn, "Test", vbTextCompare) = 0) Then
        igShowVersionNo = 1
        If InStr(1, sgCommandStr, "Debug", vbTextCompare) > 0 Then
            igShowVersionNo = 2
        End If
    End If
    slCommand = sgCommandStr    'Command$
    smPassCommands = slCommand
    lgCurrHRes = GetDeviceCaps(Traffic!pbcList.hdc, HORZRES)
    lgCurrVRes = GetDeviceCaps(Traffic!pbcList.hdc, VERTRES)
    lgCurrBPP = GetDeviceCaps(Traffic!pbcList.hdc, BITSPIXEL)
    mTestPervasive
    
    DoEvents                    'try to prevent open 3012 error on auf
    'hgAuf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    hgAuf = CBtrvTable(ONEHANDLE) 'CBtrvObj()

    ilRet = btrOpen(hgAuf, "", sgDBPath & "AUF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        MsgBox "Unable to Open Alert File, Error = " & Str$(ilRet), vbOKOnly + vbInformation, "Warning"
    End If
    hgUlf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hgUlf, "", sgDBPath & "ULF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        MsgBox "Unable to Open User Log File, Error = " & Str$(ilRet), vbOKOnly + vbInformation, "Warning"
    End If

    '4/2/11: Add setting of value
    lgUlfCode = 0
    
    'examples of auto export command line
    'Auto-Matrix Section-Matrix Std TNet
    'Auto-Efficio Revenue Section-Efficio Cal TNet
    If (Trim$(sgCommandStr) = "") Or (Trim$(sgCommandStr) = "/UserInput") Then
        bmAutoMode = False
        Signon.Show vbModal
        If igExitTraffic Then
            imTerminate = True
            Exit Sub
        End If
        slStr = sgUserName
        'dan M 12/01/10 redundant
        'gCallNetReporter CsiReportCall.StartReports
        'Sleep 1000
        sgCallAppName = "Traffic"
        igRptCallType = -2
        If Not igTestSystem Then
            smPassCommands = "Traffic^Prod\" & slStr & "\-2\0"
        Else
            smPassCommands = "Traffic^Test\" & slStr & "\-2\0"
        End If
    Else
        If InStr(1, sgCommandStr, "Act 1", vbTextCompare) <> 0 Then
            igBkgdProg = 17                     'for msg logging, gmsgbox
            bmAutoMode = True
            smImportName = "Act 1"
        ElseIf InStr(1, sgCommandStr, "CSV", vbTextCompare) <> 0 Then
            igBkgdProg = 17                     'for msg logging,gmsgbox
            bmAutoMode = True
            smImportName = "CSV"
        ElseIf InStr(1, sgCommandStr, "RADAR", vbTextCompare) <> 0 Then
            igBkgdProg = 17                     'for msg logging,gmsgbox
            bmAutoMode = True
            smImportName = "RADAR"
        ElseIf InStr(1, sgCommandStr, "Vehicle", vbTextCompare) <> 0 Then
            igBkgdProg = 17                     'for msg logging,gmsgbox
            bmAutoMode = True
            smImportName = "Vehicle"
        ElseIf InStr(1, sgCommandStr, "Impressions", vbTextCompare) <> 0 Then
            igBkgdProg = 17                     'for msg logging,gmsgbox
            bmAutoMode = True
            smImportName = "Impressions"
        Else
            igBkgdProg = 0
            bmAutoMode = False
            smImportName = ""
        End If
        igSportsSystem = 0
        sgCallAppName = "Traffic"
        igRptCallType = 0
        sgUrfStamp = "~" 'Clear time stamp incase same name
        sgUserName = "Guide"
        gUrfRead Signon, sgUserName, True, tgUrf(), False  'Obtain user records
        mGetUlfCode
        DoEvents
        gInitStdAlone
        mCheckForDate
        '4/2/11: Add setting and call.  Note: The call in _Load will be ignored
        ilRet = gObtainSAF()
        igLogActivityStatus = 32123
        gUserActivityLog "L", "ResearchList.Frm"
        Exit Sub
    End If
    'End If
    DoEvents
'    gInitStdAlone ResearchList, slStr, igTestSystem
    gInitStdAlone
    mCheckForDate
    mGetImportName
    '4/2/11: Add setting and call.  Note: The call in _Load will be ignored
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "ResearchList.Frm"
    'ilRet = gParseItem(slCommand, 3, "\", slStr)
    'igRptCallType = Val(slStr)
    
    'debug
    'MsgBox "igExporttype = " & str$(igExportType) & " smExportName = " & smExportName & " sgExportIniSectionName=" & sgExportIniSectionName, vbOKOnly
End Sub

Private Sub cmcCancel_Click()
    mTerminate 'True
End Sub
Private Sub cmcCancel_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
End Sub
'
'
'               3-10-00 Create duplicate of rptselct , named rptselcb to separate the "bridge" reports
'                       Send appropirate command line string
'
Private Sub cmcDone_Click()
    Dim ilShell As Integer
    Dim ilRet As Integer
    Dim slStr As String
    
    On Error Resume Next
    If lbcResearch.ListIndex < 0 Then
        Exit Sub
    End If
    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(CALLNONE))
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(CALLNONE))
    End If
    sgCommandStr = slStr
    Select Case lbcResearch.ItemData(lbcResearch.ListIndex)
        Case 0  'Act1
            ImptMark.Show vbModal
        Case 1  'CSV
            ImptIMS.Show vbModal
        Case 2  'RADAR
            ImptRad.Show vbModal
        Case 3  'Vehicle
            ImptVeh.Show vbModal
        Case 4  'Impression
            ImptImpression.Show vbModal
    End Select
    ''Dan M   need to find if counterpoint date has been changed in traffic.
    'mGetCsiDate
    bmInModalModule = False
    If bmAutoMode Then        'if auto mode, kill module
        'cmcCancel_Click
        tmcDelay.Enabled = True
    End If
    'gChDrDir        '3-25-03
    'Screen.MousePointer = vbDefault    'Default
End Sub
Private Function mGetCsiDate() As Boolean
'Dan M 9/21/10 get date from text file: when date can't be retrieved from command parameter because app already running
    Dim oMyFileObj As FileSystemObject
    Dim MyFile As TextStream
    Dim slFullPath As String
    Dim slCommand As String
    Dim ilPos As Integer
    Dim slNewDate As String
    
    mGetCsiDate = True
        '5676 remove hard coded c:
    'slFullPath = "C:\csi\ReportPasser.txt"
    slFullPath = sgRootDrive & "csi\ReportPasser.txt"
    Set oMyFileObj = New FileSystemObject
    If oMyFileObj.FILEEXISTS(slFullPath) Then
        On Error GoTo ErrCatch
       Set MyFile = oMyFileObj.OpenTextFile(slFullPath, ForReading, False)
       slCommand = MyFile.ReadLine
       ilPos = InStr(1, UCase(slCommand), "/D:")
       If ilPos > 0 Then
           slNewDate = Mid(slCommand, ilPos + 3, Len(slCommand) - 3)
           gCsiSetName slNewDate
       Else
           mGetCsiDate = False
       End If
       MyFile.Close
       Set MyFile = Nothing
On Error GoTo FIXFILE
       oMyFileObj.DeleteFile slFullPath, True
On Error GoTo 0
    End If
Cleanup:
    Set oMyFileObj = Nothing
    Exit Function
ErrCatch:
    gMsgBox "traffic exports couldn't read values in " & slFullPath & ".  Form_GotFocus", vbOKOnly, "Problem reading values from file."
    mGetCsiDate = False
    GoTo Cleanup
    Exit Function
FIXFILE:
    'couldn't delete file...probably open, so erase the value
    If oMyFileObj.FILEEXISTS(slFullPath) Then
        Set MyFile = oMyFileObj.OpenTextFile(slFullPath, ForWriting, False)
        MyFile.WriteLine ("")
        MyFile.Close
        Set MyFile = Nothing
        GoTo Cleanup
    End If
End Function
Private Sub cmcDone_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
'    gShowBranner
    Me.KeyPreview = True
    'ResearchList.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    'Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    Me.WindowState = vbNormal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ilRet As Integer
    ilRet = KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim ilRet As Integer
    ilRet = KeyAscii
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    'If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
    '    gFunctionKeyBranch KeyCode
    'End If
' Dan M 9/21/10 not using: replaced with textfile to pass date ( see cmcDone)
    Dim slKey As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilLoop As Integer

    slKey = Chr(KeyCode)
    If slKey = "~" Then
        smCommandKeys = ""
    Else
        If slKey <> "#" Then
            smCommandKeys = smCommandKeys + slKey
        Else
            Me.WindowState = vbNormal
        End If
        If Not bmInModalModule Then
            ilRet = gParseItem(smCommandKeys, 3, "\", slStr)
        End If
    End If
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        End
    Else
        mInit
        If imTerminate Then
            'mTerminate 'True
            tmcDelay.Enabled = True
        Else
            mCheckGG
            If igGGFlag = 0 Then
                imTerminate = True
                tmcDelay.Enabled = True
            Else
                If smImportName <> "" Then
                    tmcAutoRun.Enabled = True
                End If
            End If
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    ilRet = btrClose(hgAuf)
    btrDestroy hgAuf
    ilRet = btrClose(hgUlf)
    btrDestroy hgUlf
    
    '4/2/11: Add setting and call.
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    Set RESEARCHLIST = Nothing   'Remove data segment
    'Reset used instead of Close to cause # Clients on network to be decrement
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    'btrStopAppl
    gEraseGlobalVar True
    'Dan M 9/21/10 delete text file passing the date from traffic--just in case.
    mEraseOrChangeDateFile
    End


End Sub
Private Sub mEraseOrChangeDateFile()
    Dim oMyFileObj As FileSystemObject
    Dim MyFile As TextStream
    Dim slFullPath As String
    
        '5676 remove hard coded c:
    'slFullPath = "C:\csi\ReportPasser.txt"
    slFullPath = sgRootDrive & "csi\ReportPasser.txt"
    Set oMyFileObj = New FileSystemObject
    If oMyFileObj.FILEEXISTS(slFullPath) Then
On Error GoTo FIXFILE
        oMyFileObj.DeleteFile slFullPath, True
    End If
Cleanup:
    Set oMyFileObj = Nothing
    Exit Sub
FIXFILE:
    'couldn't delete file...probably open, so erase the value
    If oMyFileObj.FILEEXISTS(slFullPath) Then
        Set MyFile = oMyFileObj.OpenTextFile(slFullPath, ForWriting, False)
        MyFile.WriteLine ("")
        MyFile.Close
        Set MyFile = Nothing
        GoTo Cleanup
    End If

End Sub


Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcResearch_DblClick()
    cmcDone_Click
End Sub

Private Sub lbcResearch_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
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
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    bmInModalModule = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    If (tgSpf.sGUsePropSys <> "Y") Then
        imTerminate = True
        Exit Sub
    End If
    ReDim tgNtrSortInfo(0 To 0) As NTRSORTINFO              'array of vehicles with start/end indices pointing to TGNTRInfo array of SBF records
    RESEARCHLIST.Height = cmcDone.Top + 3 * cmcDone.Height '/ 3
    If smImportName <> "" Then
        RESEARCHLIST.Left = -RESEARCHLIST.Width - 100
    Else
        gCenterStdAlone RESEARCHLIST
    End If
    'ResearchList.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    smScreenCaption = "Import Research Selection"
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imFirstFocus = True
    imIgnoreChg = False
    mPopulate
    ilRet = gObtainRcfRifRdf()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainAdvt()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainAgency() 'Build into tgCommAgf
    If ilRet = False Then
        imTerminate = True
    End If
    ilRet = gObtainVef() 'Build into tgMVef
    If ilRet = False Then
        imTerminate = True
    End If
    ilRet = gVffRead()
    If ilRet = False Then
        imTerminate = True
    End If

    ilRet = gObtainSalesperson() 'Build into tgMSlf
    If ilRet = False Then
        imTerminate = True
    End If
    ilRet = gObtainComp()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainAvail()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainSAF()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainMnfForType("R", sgRevSetStamp, tgRevSet())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainMnfForType("O", sgShareBudgetStamp, tgShareBudget())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainMnfForType("X", sgExclMnfStamp, tgExclMnf())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainMnfForType("B", sgBusCatMnfStamp, tgBusCatMnf())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainMnfForType("P", sgPotMnfStamp, tgPotMnf())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainMnfForType("D", sgDemoMnfStamp, tgDemoMnf())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    ilRet = gObtainMnfForType("F", sgSocEcoMnfStamp, tgSocEcoMnf())
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    lbcResearch.Visible = True
    If imTerminate Then
        Exit Sub
    End If
    
    plcScreen_Paint
'    gCenterModalForm ResearchList
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate tree structure list   *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
    Dim ilAllowed As Integer
    Dim ilRet As Integer
    Dim ilVff As Integer
    
    
    lbcResearch.AddItem "Act 1"
    lbcResearch.ItemData(lbcResearch.NewIndex) = 0
    lbcResearch.AddItem "CSV"
    lbcResearch.ItemData(lbcResearch.NewIndex) = 1
    lbcResearch.AddItem "RADAR"
    lbcResearch.ItemData(lbcResearch.NewIndex) = 2
    lbcResearch.AddItem "Vehicle"
    lbcResearch.ItemData(lbcResearch.NewIndex) = 3
    If tgSaf(0).sHideDemoOnBR = "Y" Then
        lbcResearch.AddItem "Impressions"
        lbcResearch.ItemData(lbcResearch.NewIndex) = 4
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
Private Sub mTerminate() 'ilFromCancel As Integer)
'
'   mTerminate
'   Where:
'
'    Dim ilRet As Integer
'    Erase tmSelSrf
'    Erase tmResearchList
'    Erase tgRtfList
'    Erase tgRnfList
'    'gObtainRNF hmRnf
'    btrExtClear hmRtf   'Clear any previous extend operation
'    ilRet = btrClose(hmRtf)
'    btrDestroy hmRtf
'    btrExtClear hmSrf   'Clear any previous extend operation
'    ilRet = btrClose(hmSrf)
'    btrDestroy hmSrf
'    btrExtClear hmSnf   'Clear any previous extend operation
'    ilRet = btrClose(hmSnf)
'    btrDestroy hmSnf
'    btrExtClear hmRnf   'Clear any previous extend operation
'    ilRet = btrClose(hmRnf)
'    btrDestroy hmRnf
    'If ilFromCancel Then
    '    igParentRestarted = False
    '    If Not igStdAloneMode Then
    '        If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
    '            edcLinkDestHelpMsg.LinkExecute "@" & "Done"
    '        Else
    '            edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
    '            edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
    '            edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
    '            edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
    '            edcLinkDestHelpMsg.LinkExecute "Done"
    '        End If
    '        Do While Not igParentRestarted
    '            DoEvents
    '        Loop
    '    End If
    'End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload RESEARCHLIST
'    Set ResearchList = Nothing   'Remove data segment
    igManUnload = NO
End Sub

Private Sub pbcClickFocus_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub




Private Sub tmcAutoRun_Timer()
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim slStr As String
    
    tmcAutoRun.Enabled = False
    Me.Left = -Me.Width - 100
    For ilLoop = 0 To lbcResearch.ListCount - 1 Step 1
        slStr = lbcResearch.List(ilLoop)
        ilPos = InStr(1, smImportName, slStr, vbTextCompare)
        If ilPos = 1 Then
            lbcResearch.ListIndex = ilLoop
            cmcDone_Click
            mTerminate
            Exit Sub
        End If
    Next ilLoop
    mTerminate
End Sub

Private Sub tmcDelay_Timer()
    tmcDelay.Enabled = False
    mTerminate 'False
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub


Private Sub mTestPervasive()
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim hlSpf As Integer
    Dim tlSpf As SPF

    gInitGlobalVar
    hlSpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlSpf, "", sgDBPath & "Spf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    ilRecLen = Len(tlSpf)
    ilRet = btrGetFirst(hlSpf, tlSpf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    btrDestroy hlSpf
End Sub

Private Sub mCheckForDate()
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slDate As String
    Dim slSetDate As String
    Dim ilRet As Integer
    
    ilPos = InStr(1, sgCommandStr, "/D:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommandStr, " ")
        If ilSpace = 0 Then
            slDate = Trim$(Mid$(sgCommandStr, ilPos + 3))
        Else
            slDate = Trim$(Mid$(sgCommandStr, ilPos + 3, ilSpace - ilPos - 3))
        End If
        If gValidDate(slDate) Then
            slDate = gAdjYear(slDate)
            slSetDate = slDate
        End If
    End If
    If Trim$(slSetDate) = "" Then
        If (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) > 0) Or (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) > 0) Then
            slSetDate = "12/15/1999"
            slDate = slSetDate
        End If
    End If
    If Trim$(slSetDate) <> "" Then
        'Dan M 9/20/10 problems with gGetCSIName("SYSDate") in v57 reports.exe... change to global variable
     '   ilRet = csiSetName("SYSDate", slDate)
        ilRet = gCsiSetName(slDate)
    End If
    
End Sub
'4/2/11: Add routine
Private Sub mGetUlfCode()
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    ilPos = InStr(1, sgCommandStr, "/ULF:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommandStr, " ")
        If ilSpace = 0 Then
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5)))
        Else
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5, ilSpace - ilPos - 3)))
        End If
    End If
End Sub

Private Sub mGetImportName()
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    smImportName = ""
    ilPos = InStr(1, sgCommandStr, "/Import:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommandStr, "~")
        If ilSpace = 0 Then
            smImportName = Trim$(Mid$(sgCommandStr, ilPos + 8))
        Else
            smImportName = Trim$(Mid$(sgCommandStr, ilPos + 8, ilSpace - ilPos - 8))
        End If
        
    End If
End Sub

Private Sub mCheckGG()
    Dim ilRet As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim llNow As Long
    Dim llDate As Long
    Dim slStr As String
    Dim ilLoop As Integer
    
    On Error Resume Next
    
    'If imLastHourGGChecked = Hour(Now) Then
    '    Exit Sub
    'End If
    'imLastHourGGChecked = Hour(Now)
    igGGFlag = 1
    hmSaf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    imSafRecLen = Len(tmSaf)
    ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    ilField1 = Asc(tmSaf.sName)
    slStr = Mid$(tmSaf.sName, 2, 5)
    llDate = Val(slStr)
    ilField2 = Asc(Mid$(tmSaf.sName, 11, 1))
    llNow = gDateValue(Format$(Now, "m/d/yy"))
    If (ilField1 = 0) And (ilField2 = 1) Then
        If llDate <= llNow Then
            ilField2 = 0
        End If
    End If
    If (ilField1 = 0) And (ilField2 = 0) Then
        igGGFlag = 0
    End If
    'gSetRptGGFlag tmSaf
    btrDestroy hmSaf
End Sub


