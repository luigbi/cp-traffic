VERSION 5.00
Begin VB.Form ReportList 
   Appearance      =   0  'Flat
   Caption         =   "CSI Reports"
   ClientHeight    =   6105
   ClientLeft      =   585
   ClientTop       =   2880
   ClientWidth     =   9210
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
   Icon            =   "Reportlist.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6105
   ScaleWidth      =   9210
   Begin VB.Frame frcGen 
      Caption         =   "Insert last-used report settings"
      Height          =   555
      Left            =   1605
      TabIndex        =   6
      Top             =   5505
      Width           =   4575
      Begin VB.CommandButton cmcDone 
         Appearance      =   0  'Flat
         Caption         =   "No"
         Height          =   285
         Index           =   2
         Left            =   3330
         TabIndex        =   9
         Top             =   210
         Width           =   1035
      End
      Begin VB.CommandButton cmcDone 
         Appearance      =   0  'Flat
         Caption         =   "Yes-Except Dates"
         Height          =   285
         Index           =   1
         Left            =   1350
         TabIndex        =   8
         Top             =   210
         Width           =   1860
      End
      Begin VB.CommandButton cmcDone 
         Appearance      =   0  'Flat
         Caption         =   "Yes"
         Default         =   -1  'True
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   210
         Width           =   1035
      End
   End
   Begin VB.Timer tmcSetCntrls 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8970
      Top             =   4770
   End
   Begin VB.Timer tmcClock 
      Interval        =   60000
      Left            =   9060
      Top             =   5325
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8955
      Top             =   5835
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   6420
      TabIndex        =   10
      Top             =   5730
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
      TabIndex        =   13
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
   Begin VB.PictureBox plcInvNo 
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
      Height          =   5250
      Left            =   120
      ScaleHeight     =   5190
      ScaleWidth      =   8940
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   9000
      Begin V81TrafficReports.CSI_ComboBoxMS CSI_ComboBoxMS1 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   1005
         BorderStyle     =   1
      End
      Begin VB.OptionButton rbcShowBy 
         Caption         =   "Quick Find"
         Height          =   210
         Index           =   2
         Left            =   3360
         TabIndex        =   18
         Top             =   15
         Width           =   1305
      End
      Begin VB.OptionButton rbcShowBy 
         Caption         =   "Report Name"
         Height          =   210
         Index           =   1
         Left            =   1875
         TabIndex        =   17
         Top             =   15
         Width           =   1425
      End
      Begin VB.OptionButton rbcShowBy 
         Caption         =   "Group"
         Height          =   210
         Index           =   0
         Left            =   930
         TabIndex        =   16
         Top             =   15
         Width           =   870
      End
      Begin VB.HScrollBar hbcRptSample 
         Height          =   240
         LargeChange     =   8490
         Left            =   105
         SmallChange     =   8490
         TabIndex        =   14
         Top             =   4935
         Width           =   8490
      End
      Begin VB.ListBox lbcRpt 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   4485
      End
      Begin VB.VScrollBar vbcRptSample 
         Height          =   2595
         LargeChange     =   1455
         Left            =   8595
         SmallChange     =   1455
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2550
         Width           =   255
      End
      Begin VB.PictureBox pbcRptSample 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   2595
         Index           =   0
         Left            =   105
         ScaleHeight     =   2565
         ScaleWidth      =   8460
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2550
         Width           =   8490
         Begin VB.PictureBox pbcRptSample 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
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
            Index           =   1
            Left            =   15
            ScaleHeight     =   165
            ScaleWidth      =   6060
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   -15
            Width           =   6060
         End
      End
      Begin VB.TextBox edcRptDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2340
         Left            =   4680
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   150
         Visible         =   0   'False
         Width           =   4155
      End
      Begin VB.Label lacShowBy 
         Caption         =   "Show by"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   15
         Width           =   750
      End
      Begin VB.Label lacAltDownMsg 
         Appearance      =   0  'Flat
         Caption         =   "Alt-Down-Arrow to tab to next Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   2355
         Width           =   4410
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   195
      Top             =   5625
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "ReportList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ReportList.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text

'***
'***  When adding new Report modules, you need to add setting of fgReportForm prior to the Show Modal call
'***


'Program library dates Field Areas
Dim tmRnf As RNF        'Rnf record image
Dim tmRnfSrchKey As INTKEY0    'Rnf key record image
Dim hmRnf As Integer    'Report Names file handle
Dim imRnfRecLen As Integer        'RvF record length
Dim tmSnf As SNF        'Snf record image
Dim tmSnfSrchKey As INTKEY0    'Snf key record image
Dim hmSnf As Integer    'Report Set file handle
Dim imSnfRecLen As Integer        'ADF record length
Dim tmSrf As SRF        'Snf record image
Dim tmSrfSrchKey1 As INTKEY0    'Snf key record image
Dim hmSrf As Integer    'Report Set file handle
Dim imSrfRecLen As Integer        'ADF record length
Dim tmRtf As RTF        'Rtf record image
Dim hmRtf As Integer    'Report Tree file handle
Dim imRtfRecLen As Integer        'RtF record length
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim tmSelSrf() As SRF
Dim tmReportList() As RPTLST
Dim smPassCommands As String
Dim smScreenCaption As String
Dim tmRptNoSelNameMap(0 To 50) As RPTNAMEMAP
Dim tmRptSelNameMap(0 To 115) As RPTNAMEMAP     '4-9-12 expand the array for rptsel module from 105 to 110, 1-24-18 expand from 110 to 115
Dim tmRptSelCtNameMap(0 To 50) As RPTNAMEMAP    '4-5-06 increase from 40 to 45
Dim tmRptSelPjNameMap(0 To 4) As RPTNAMEMAP
Dim tmRptSelRINameMap(0 To 3) As RPTNAMEMAP '8-29-02
Dim tmRptSelNTNameMap(0 To 2) As RPTNAMEMAP
Dim tmRptSelCCNameMap(0 To 4) As RPTNAMEMAP '1-15-04
Dim tmRptSelFDNameMap(0 To 4) As RPTNAMEMAP '8-18-04 Station Pledge report
Dim tmRptSelRSnameMap(0 To 5) As RPTNAMEMAP '9-10-15 Research reports (expand from 2 to 5)
Dim tmRptSelSRNameMap(0 To 2) As RPTNAMEMAP     '9-19-06  Split region list
Dim tmRptSelCANameMap(0 To 2) As RPTNAMEMAP     '12-18-07 Combo Avails versions
Dim tmRptSelSNNameMap(0 To 2) As RPTNAMEMAP     '04-10-08 Split Network Avails
Dim tmRptSelADNameMAP(0 To 2) As RPTNAMEMAP      '5-06-08 Post buy analysis, combine with Aud Delivery
Dim tmRptSelSpotBBNameMAP(0 To 2) As RPTNAMEMAP     '11-4-14 Rev on books
Dim tmRptSelAvgCompareNameMAP(0 To 2) As RPTNAMEMAP     '11-1-18 Average Rate/Spot Price Comparison
'Dim tmRptSelPodcastBillingMap(0 To 2) As RPTNAMEMAP     '12-28-20 Podcast Spot billing reports

Dim imIgnoreChg As Integer
Dim smCommandKeys As String
Dim bmInModalModule As Boolean

Dim tmSaf As SAF
Dim hmSaf As Integer
Dim imSafRecLen As Integer
Dim imLastHourGGChecked As Integer

Private smExe As String
Private imRnfCode As Integer





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
    '4/2/11: Add setting of value
    lgUlfCode = 0
    bgDevEnv = IsDevEnv()
    bgInternalGuide = False
'    If (Trim$(sgCommandStr) = "") Or (Trim$(sgCommandStr) = "/UserInput") Then
    If (Trim$(sgCommandStr) = "") Or (InStr(1, UCase(sgCommandStr), "/USERINPUT", vbTextCompare) > 0) Then
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
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide" '"rn48616" '"Guide"
    '    ilTestSystem = False
    '    imShowHelpMsg = False
    '    If ilTestSystem Then
    '        slCommand = "ReportList^TEST^NoHelp\ADVERTISERSLIST\1"
    '    Else
    '        slCommand = "ReportList^Prod^NoHelp\ADVERTISERSLIST\1"
    '    End If
    '    smPassCommands = slCommand
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
        igSportsSystem = 0
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        'If Trim$(slStr) = "" Then
        '    MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
        '    'End
        '    imTerminate = True
        '    Exit Sub
        'End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        'If StrComp(slTestSystem, "Test", 1) = 0 Then
        '    ilTestSystem = True
        'Else
        '    ilTestSystem = False
        'End If
    '    imShowHelpMsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpMsg = False
    '    End If
        ilRet = gParseItem(slCommand, 3, "\", slStr)
        igRptCallType = Val(slStr)
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
        'btrStopAppl
        'Sleep 1000
        'gInitGlobalVar   'Initialize global variables
        'Sleep 1000
        sgUrfStamp = "~" 'Clear time stamp incase same name
        sgUserName = Trim$(slStr)
        '6/20/09:  Jim requested that the Guide sign in be changed to CSI for internal Guide only
        If StrComp(sgUserName, "CSI", vbTextCompare) = 0 Then
            sgUserName = "Guide"
        End If
        gUrfRead Signon, sgUserName, True, tgUrf(), False  'Obtain user records
        If StrComp(slStr, "CSI", vbTextCompare) = 0 Then
            gExpandGuideAsUser tgUrf(0)
        End If
        'dan M 12/01/10 redundant
        'gCallNetReporter CsiReportCall.StartReports
       ' Sleep 1000
        '4/2/11: Add call to routine
        mGetUlfCode
    End If
    'End If
    DoEvents
'    gInitStdAlone ReportList, slStr, igTestSystem
    gInitStdAlone
    mCheckForDate
    '4/2/11: Add setting and call.  Note: The call in _Load will be ignored
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "ReportList.Frm"
    'ilRet = gParseItem(slCommand, 3, "\", slStr)
    'igRptCallType = Val(slStr)
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
Private Sub cmcDone_Click(Index As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        slTestName                                              *
'******************************************************************************************

    Dim slStr As String
    Dim ilRet As Integer
    Dim slName As String
    Dim slExe As String
    Dim slRptCallType As String
    Dim ilFound As Integer
    Dim slStr1 As String
    Dim slStr2 As String
    Dim slStr3 As String
    Dim ilPos As Integer

    If (Not bgInternalGuide) And (igGGFlag = 0) And (igRptGGFlag = 0) Then
        tmcDelay.Enabled = True
        Exit Sub
    End If
    If rbcShowBy(2).Value = True Then
        If CSI_ComboBoxMS1.Text = "" Then Exit Sub
        If CSI_ComboBoxMS1.ListIndex < 0 Then Exit Sub
    Else
        If lbcRpt.ListIndex < 0 Then
            lbcRpt.SetFocus
            Exit Sub
        End If
        
        If tmReportList(lbcRpt.ListIndex).tRnf.sType = "C" Then
            lbcRpt.SetFocus
            Exit Sub
        End If
    End If
    'Dan M   need to find if counterpoint date has been changed in traffic.
    mGetCsiDate
    slExe = Trim$(tmReportList(lbcRpt.ListIndex).tRnf.sRptExe)
    ilPos = InStr(1, slExe, ".", vbTextCompare)
    If ilPos > 0 Then
        slExe = Left$(slExe, ilPos - 1)
    End If
    slName = Trim$(tmReportList(lbcRpt.ListIndex).tRnf.sName)
    '4/2/11: Add Seeting of value
    sgReportListName = Trim$(slName)
    slRptCallType = Trim$(str$(igRptCallType))
    If StrComp(slExe, "RptSel", 1) = 0 Then
        'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
        '    Exit Sub
        'End If

        ilFound = mFindNameInMap(tmRptSelNameMap, slName, slRptCallType)

        'ilFound = False
        'For ilLoop = 0 To UBound(tmRptSelNameMap) Step 1
        '    slTestName = Trim$(tmRptSelNameMap(ilLoop).sName)
        '    If Len(slTestName) = 0 Then
        '        Exit For
        '    End If
        '    If StrComp(slname, slTestName, 1) = 0 Then
        '        ilFound = True
        '        slRptCallType = Trim$(Str$(tmRptSelNameMap(ilLoop).iRptCallType))
        '        Exit For
        '    End If
        'Next ilLoop
        If Not ilFound Then
            'ilRet = MsgBox("Report Name " & slname & " not found in Mapping Table", vbOkOnly + vbInformation, "Report List")
            'cmcCancel.SetFocus
            Exit Sub
        End If
        If Val(slRptCallType) = PROGRAMMINGJOB Then
            ilRet = gParseItem(smPassCommands, 1, "\", slStr1)
            ilRet = gParseItem(smPassCommands, 2, "\", slStr2)
            ilRet = gParseItem(smPassCommands, 3, "\", slStr3)
            If StrComp(slName, "Program Libraries", 1) = 0 Then
                'Replace report type with 3 (4th parse item)
                smPassCommands = slStr1 & "\" & slStr2 & "\" & slStr3 & "\" & "3"
            Else
                'Replace report type with 0 (4 parse item)
                smPassCommands = slStr1 & "\" & slStr2 & "\" & slStr3 & "\" & "0"
            End If
        End If
        If Val(slRptCallType) = LOGSJOB Then
            ilRet = gParseItem(smPassCommands, 1, "\", slStr1)
            ilRet = gParseItem(smPassCommands, 2, "\", slStr2)
            ilRet = gParseItem(smPassCommands, 3, "\", slStr3)
            If Val(slStr3) = LOGSJOB Then   'Leave the command line
                'smPassCommands = slStr1 & "\" & slStr2 & "\" & slStr3 & "\" & "3"   '3=Reprint
            Else
                smPassCommands = slStr1 & "\" & slStr2 & "\" & slStr3 & "\" & "3" & "\" & Trim$(str$(tgUrf(0).iCode)) & "\\\\\0\"   '3=Reprint
            End If
        End If
    ElseIf StrComp(slExe, "RptSelCt", 1) = 0 Or StrComp(slExe, "RptSelCb", 1) = 0 Then
        'If Not gWinRoom(igNoExeWinRes(RPTSELCTEXE)) Then
        '    Exit Sub
        'End If

        ilFound = mFindNameInMap(tmRptSelCtNameMap, slName, slRptCallType)
        'ilFound = False
        'For ilLoop = 0 To UBound(tmRptSelCtNameMap) Step 1
        '    slTestName = Trim$(tmRptSelCtNameMap(ilLoop).sName)
        '    If Len(slTestName) = 0 Then
        '        Exit For
        '    End If
        '    If StrComp(slName, slTestName, 1) = 0 Then
        '        ilFound = True
        '        slRptCallType = Trim$(Str$(tmRptSelCtNameMap(ilLoop).iRptCallType))
        '        Exit For
        '    End If
        'Next ilLoop
        If Not ilFound Then
            'ilRet = MsgBox("Report Name " & slName & " not found in Mapping Table", vbOkOnly + vbInformation, "Report List")
            'cmcCancel.SetFocus
            Exit Sub
        End If
        If Val(slRptCallType) = CONTRACTSJOB Then
            ilRet = gParseItem(smPassCommands, 1, "\", slStr1)
            ilRet = gParseItem(smPassCommands, 2, "\", slStr2)
            ilRet = gParseItem(smPassCommands, 3, "\", slStr3)
            'Replace report type with 1 (4 parse item)
            smPassCommands = slStr1 & "\" & slStr2 & "\" & slStr3 & "\" & "1"
        End If
    ElseIf StrComp(slExe, "RptNoSel", 1) = 0 Then
        'If Not gWinRoom(igNoExeWinRes(RPTNOSELEXE)) Then
        '    Exit Sub
        'End If

        ilFound = mFindNameInMap(tmRptNoSelNameMap, slName, slRptCallType)
        'ilFound = False
        'For ilLoop = 0 To UBound(tmRptNoSelNameMap) Step 1
        '    slTestName = Trim$(tmRptNoSelNameMap(ilLoop).sName)
        '    If Len(slTestName) = 0 Then
        '        Exit For
        '    End If
        '    If StrComp(slName, slTestName, 1) = 0 Then
        '        ilFound = True
        '        slRptCallType = Trim$(Str$(tmRptNoSelNameMap(ilLoop).iRptCallType))
        '        Exit For
        '    End If
        'Next ilLoop
        If Not ilFound Then
            'ilRet = MsgBox("Report Name " & slName & " not found in Mapping Table", vbOkOnly + vbInformation, "Report List")
            'cmcCancel.SetFocus
            Exit Sub
        End If
    ElseIf StrComp(slExe, "RptSelAD", 1) = 0 Then
        ilFound = mFindNameInMap(tmRptSelADNameMAP, slName, slRptCallType)
        If Not ilFound Then
            Exit Sub
        End If

    ElseIf StrComp(slExe, "RptSelpj", 1) = 0 Then
        'If Not gWinRoom(igNoExeWinRes(RPTNOSELEXE) + 20) Then
        '    Exit Sub
        'End If

        ilFound = mFindNameInMap(tmRptSelPjNameMap, slName, slRptCallType)
        'ilFound = False
        'For ilLoop = 0 To UBound(tmRptSelPjNameMap) Step 1
        '    slTestName = Trim$(tmRptSelPjNameMap(ilLoop).sName)
        '    If Len(slTestName) = 0 Then
        '        Exit For
        '    End If
        '    If StrComp(slName, slTestName, 1) = 0 Then
        '        ilFound = True
        '        slRptCallType = Trim$(Str$(tmRptSelPjNameMap(ilLoop).iRptCallType))
        '        Exit For
        '    End If
        'Next ilLoop
        If Not ilFound Then
            'ilRet = MsgBox("Report Name " & slName & " not found in Mapping Table", vbOkOnly + vbInformation, "Report List")
            'cmcCancel.SetFocus
            Exit Sub
        End If
    ElseIf StrComp(slExe, "RptSelRI", 1) = 0 Then      '8-29-02

        ilFound = mFindNameInMap(tmRptSelRINameMap, slName, slRptCallType)

        'ilFound = False
        'For ilLoop = 0 To UBound(tmRptSelRINameMap) Step 1
        '    slTestName = Trim$(tmRptSelRINameMap(ilLoop).sName)
        '    If Len(slTestName) = 0 Then
        '        Exit For
        '    End If
        '    If StrComp(slName, slTestName, 1) = 0 Then
        '        ilFound = True
        '        slRptCallType = Trim$(Str$(tmRptSelRINameMap(ilLoop).iRptCallType))
        '        Exit For
        '    End If
        'Next ilLoop
        If Not ilFound Then
            'ilRet = MsgBox("Report Name " & slName & " not found in Mapping Table", vbOkOnly + vbInformation, "Report List")
            'cmcCancel.SetFocus
            Exit Sub
        End If

    ElseIf StrComp(slExe, "RptSelNT", 1) = 0 Then      '4-2-03

        ilFound = mFindNameInMap(tmRptSelNTNameMap, slName, slRptCallType)

        'ilFound = False
        'For ilLoop = 0 To UBound(tmRptSelNTNameMap) Step 1
         '   slTestName = Trim$(tmRptSelNTNameMap(ilLoop).sName)
         '   If Len(slTestName) = 0 Then
         '       Exit For
         '   End If
         '   If StrComp(slName, slTestName, 1) = 0 Then
         '       ilFound = True
         '       slRptCallType = Trim$(Str$(tmRptSelNTNameMap(ilLoop).iRptCallType))
         '       Exit For
         '   End If
        'Next ilLoop
        If Not ilFound Then
            'ilRet = MsgBox("Report Name " & slName & " not found in Mapping Table", vbOkOnly + vbInformation, "Report List")
            'cmcCancel.SetFocus
            Exit Sub
        End If
    ElseIf StrComp(slExe, "RptSelCC", 1) = 0 Then      '1-15-04

        ilFound = mFindNameInMap(tmRptSelCCNameMap, slName, slRptCallType)

        'ilFound = False
        'For ilLoop = 0 To UBound(tmRptSelCCNameMap) Step 1
        '    slTestName = Trim$(tmRptSelCCNameMap(ilLoop).sName)
        '    If Len(slTestName) = 0 Then
        '        Exit For
        '    End If
        '    If StrComp(slName, slTestName, 1) = 0 Then
        '        ilFound = True
        '        slRptCallType = Trim$(Str$(tmRptSelCCNameMap(ilLoop).iRptCallType))
        '        Exit For
        '    End If
        'Next ilLoop
        If Not ilFound Then
            'ilRet = MsgBox("Report Name " & slName & " not found in Mapping Table", vbOkOnly + vbInformation, "Report List")
            'cmcCancel.SetFocus
            Exit Sub
        End If
    ElseIf StrComp(slExe, "RptSelFD", 1) = 0 Then      '8-18-04 Feed Reports
        ilFound = mFindNameInMap(tmRptSelFDNameMap, slName, slRptCallType)

        'ilFound = False
        'For ilLoop = 0 To UBound(tmRptSelFDNameMap) Step 1
        '    slTestName = Trim$(tmRptSelFDNameMap(ilLoop).sName)
        '    If Len(slTestName) = 0 Then
        '        Exit For
        '    End If
        '    If StrComp(slName, slTestName, 1) = 0 Then
        '        ilFound = True
        '        slRptCallType = Trim$(Str$(tmRptSelFDNameMap(ilLoop).iRptCallType))
        '        Exit For
        '    End If
        'Next ilLoop
        If Not ilFound Then
            'ilRet = MsgBox("Report Name " & slName & " not found in Mapping Table", vbOkOnly + vbInformation, "Report List")
            'cmcCancel.SetFocus
            Exit Sub
        End If

    ElseIf StrComp(slExe, "RptSelRS", 1) = 0 Then      '2-8-06 Research reports
        ilFound = mFindNameInMap(tmRptSelRSnameMap, slName, slRptCallType)

        'ilFound = False
        'For ilLoop = 0 To UBound(tmRptSelRSnameMap) Step 1
        '    slTestName = Trim$(tmRptSelRSnameMap(ilLoop).sName)
        '    If Len(slTestName) = 0 Then
        '        Exit For
        '    End If
        '    If StrComp(slName, slTestName, 1) = 0 Then
        '        ilFound = True
        '        slRptCallType = Trim$(Str$(tmRptSelRSnameMap(ilLoop).iRptCallType))
        '        Exit For
        '    End If
        'Next ilLoop
        If Not ilFound Then
            'ilRet = MsgBox("Report Name " & slName & " not found in Mapping Table", vbOkOnly + vbInformation, "Report List")
            'cmcCancel.SetFocus
            Exit Sub
        End If
    ElseIf StrComp(slExe, "RptSelSR", 1) = 0 Then      '9-19-06 Split Regions
        ilFound = mFindNameInMap(tmRptSelSRNameMap, slName, slRptCallType)
        If Not ilFound Then
            Exit Sub
        End If
    ElseIf StrComp(slExe, "RptSelSN", 1) = 0 Then
        ilFound = mFindNameInMap(tmRptSelSNNameMap, slName, slRptCallType)
        If Not ilFound Then
            Exit Sub
        End If
    ElseIf StrComp(slExe, "RptSelSpotBB", 1) = 0 Then
        ilFound = mFindNameInMap(tmRptSelSpotBBNameMAP, slName, slRptCallType)
        If Not ilFound Then
            Exit Sub
        End If
    ElseIf StrComp(slExe, "RptSelAvgCmp", 1) = 0 Then
        ilFound = mFindNameInMap(tmRptSelAvgCompareNameMAP, slName, slRptCallType)
        If Not ilFound Then
            Exit Sub
        End If
'    ElseIf StrComp(slExe, "RptSelPodBil", 1) = 0 Then
'        ilFound = mFindNameInMap(tmRptSelPodcastBillingMap, slName, slRptCallType)
'        If Not ilFound Then
'            Exit Sub
'        End If
    Else
        'If Not gWinRoom(igNoExeWinRes(RPTNOSELEXE) + 20) Then
        '    Exit Sub
        'End If
    End If
    'Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    slStr = smPassCommands & "\||" & slRptCallType & "\" & slName
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
    '    If igTestSystem Then
    '        slStr = slStr & "ReportList^Test\" & slCallRptType & "\" & slName
    '    Else
    '        slStr = slStr & "ReportList^Prod\" & slCallRptType & "\" & slName
    '    End If
    'Else
    '    If igTestSystem Then
    '        slStr = slStr & "ReportList^Test^NOHELP\" & slCallRptType & "\" & slName
    '    Else
    '        slStr = slStr & "ReportList^Prod^NOHELP\" & slCallRptType & "\" & slName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & slExe & " " & slStr, 1)
    bmInModalModule = True
    sgCommandStr = slStr
    DoEvents                '2-27-03
    On Error Resume Next
    
    '5/24/18: Reset the Form controls
    'gSetReportCtrlsSetting slExe, tmReportList(lbcRpt.ListIndex).tRnf.iCode
    sgReportFormExe = slExe
    igReportRnfCode = tmReportList(lbcRpt.ListIndex).tRnf.iCode
    'Bypass ReRate as it handle refresh internally within the code
    If StrComp(UCase(slExe), UCase("ExptReRate"), 1) <> 0 And StrComp(UCase(slExe), UCase("ReRate"), 1) <> 0 Then
        igReportButtonIndex = Index
        tmcSetCntrls.Enabled = True
    End If
    '***
    '***  When adding new Report modules, you need to add setting of fgReportForm prior to the Show Modal call
    '***
    If StrComp(slExe, "RptNoSel", 1) = 0 Then
        Set fgReportForm = RptNoSel
        RptNoSel.Show vbModal
    ElseIf StrComp(slExe, "RptSel", 1) = 0 Then
        Set fgReportForm = RptSel
        RptSel.Show vbModal
    ElseIf StrComp(slExe, "RptSel30", 1) = 0 Then           '6-12-13 cpp/cpm 30"unit
        Set fgReportForm = RptSel30
        RptSel30.Show vbModal
    ElseIf StrComp(slExe, "RptSelaa", 1) = 0 Then
        Set fgReportForm = RptSelAA
        RptSelAA.Show vbModal
    ElseIf StrComp(slExe, "RptSelac", 1) = 0 Then
        Set fgReportForm = RptSelAc
        RptSelAc.Show vbModal
    ElseIf StrComp(slExe, "RptSelAcqPay", 1) = 0 Then       '8-5-15
        Set fgReportForm = RptSelAcqPay
        RptSelAcqPay.Show vbModal
    ElseIf StrComp(slExe, "RptSelad", 1) = 0 Then
        Set fgReportForm = RptSelAD
        RptSelAD.Show vbModal
    ElseIf StrComp(slExe, "RptSelal", 1) = 0 Then       '4-14-04
        Set fgReportForm = RptSelAL
        RptSelAL.Show vbModal
    ElseIf StrComp(slExe, "RptSelalloc", 1) = 0 Then       '12-13-18   Revenue Allocation
        Set fgReportForm = RptSelALLOC
        RptSelALLOC.Show vbModal
    ElseIf StrComp(slExe, "RptSelap", 1) = 0 Then
        Set fgReportForm = RptSelAp
        RptSelAp.Show vbModal
    ElseIf StrComp(slExe, "RptSelas", 1) = 0 Then
        Set fgReportForm = RptSelAS
        RptSelAS.Show vbModal
    ElseIf StrComp(slExe, "RptSelav", 1) = 0 Then
        Set fgReportForm = RptSelAv
        RptSelAv.Show vbModal
    ElseIf StrComp(slExe, "RptSelBO", 1) = 0 Then       '7-22-11 Sales Breakout
        Set fgReportForm = RptSelBO
        RptSelBO.Show vbModal
    ElseIf StrComp(slExe, "RptSelcb", 1) = 0 Then
        Set fgReportForm = RptSelCb
        RptSelCb.Show vbModal
    ElseIf StrComp(slExe, "RptSelcc", 1) = 0 Then       '1-15-04 Producer/Provider reports
        Set fgReportForm = RptSelCC
        RptSelCC.Show vbModal
    ElseIf StrComp(slExe, "RptSelcm", 1) = 0 Then       '10-5-10 Competitive Categories
        Set fgReportForm = RptSelCM
        RptSelCM.Show vbModal
    ElseIf StrComp(slExe, "RptSelcp", 1) = 0 Then
        Set fgReportForm = RptSelCp
        RptSelCp.Show vbModal
    ElseIf StrComp(slExe, "RptSelCt", 1) = 0 Then
        Set fgReportForm = RptSelCt
        RptSelCt.Show vbModal
    ElseIf StrComp(slExe, "RptSeldb", 1) = 0 Then
        Set fgReportForm = RptSelDB
        RptSelDB.Show vbModal
    ElseIf StrComp(slExe, "RptSeldf", 1) = 0 Then
        Set fgReportForm = RptSelDF
        RptSelDF.Show vbModal
    ElseIf StrComp(slExe, "RptSelEL", 1) = 0 Then           '7-1-19 Engineering links
        Set fgReportForm = RptSelEngrLk
        RptSelEngrLk.Show vbModal
    ElseIf StrComp(slExe, "RptSelds", 1) = 0 Then
        Set fgReportForm = RptSelDS
        RptSelDS.Show vbModal
    'ElseIf StrComp(slExe, "RptSelEd", 1) = 0 Then
    '    RptSelED.Show vbModal
    ElseIf StrComp(slExe, "RptSelFd", 1) = 0 Then       '8-4-04 Feed Report
        Set fgReportForm = rptSelFD
        rptSelFD.Show vbModal
    ElseIf StrComp(slExe, "RptSelia", 1) = 0 Then
        Set fgReportForm = RptSelIA
        RptSelIA.Show vbModal
    ElseIf StrComp(slExe, "RptSelid", 1) = 0 Then       '5-21-02
        Set fgReportForm = RptSelID
        RptSelID.Show vbModal
    ElseIf StrComp(slExe, "RptSelin", 1) = 0 Then
        'RptSelIn.Show vbModal
    ElseIf StrComp(slExe, "RptSelir", 1) = 0 Then       '7-13-05
        Set fgReportForm = RptSelIR
        RptSelIR.Show vbModal
    ElseIf StrComp(slExe, "RptSeliv", 1) = 0 Then
        Set fgReportForm = RptSelIv
        RptSelIv.Show vbModal
    ElseIf StrComp(slExe, "RptSellg", 1) = 0 Then
        'RptSellg.Show vbModal
    ElseIf StrComp(slExe, "RptSelNT", 1) = 0 Then       '4-2-03
        Set fgReportForm = RptSelNT
        RptSelNT.Show vbModal
    ElseIf StrComp(slExe, "RptSelOF", 1) = 0 Then       '7-21-06
        Set fgReportForm = RptSelOF
        RptSelOF.Show vbModal
    ElseIf StrComp(slExe, "RptSelos", 1) = 0 Then
        Set fgReportForm = RptSelOS
        RptSelOS.Show vbModal
    ElseIf StrComp(slExe, "RptSelpa", 1) = 0 Then
        Set fgReportForm = RptSelPA
        RptSelPA.Show vbModal
    ElseIf StrComp(slExe, "RptSelMA", 1) = 0 Then       '6-18-13 margin allocation
        Set fgReportForm = RptSelMA
        RptSelMA.Show vbModal
    ElseIf StrComp(slExe, "RptSelParPay", 1) = 0 Then    '8-25-17 Participant Payables
        Set fgReportForm = RptSelParPay
        RptSelParPay.Show vbModal
    ElseIf StrComp(slExe, "RptSelpc", 1) = 0 Then
        Set fgReportForm = RptSelPC
        RptSelPC.Show vbModal
    ElseIf StrComp(slExe, "RptSelpj", 1) = 0 Then
        Set fgReportForm = RptSelPJ
        RptSelPJ.Show vbModal
    ElseIf StrComp(slExe, "RptSelpp", 1) = 0 Then
        Set fgReportForm = RptSelPP
        RptSelPP.Show vbModal
    ElseIf StrComp(slExe, "RptSelpr", 1) = 0 Then       '6-15-04  Proposal Research Recap
        Set fgReportForm = RptSelPr
        RptSelPr.Show vbModal
    ElseIf StrComp(slExe, "RptSelps", 1) = 0 Then
        Set fgReportForm = RptSelPS
        RptSelPS.Show vbModal
    ElseIf StrComp(slExe, "RptSelqb", 1) = 0 Then
        Set fgReportForm = RptSelQB
        RptSelQB.Show vbModal
    ElseIf StrComp(slExe, "RptSelra", 1) = 0 Then
        Set fgReportForm = RptSelRA
        RptSelRA.Show vbModal
    ElseIf StrComp(slExe, "RptSelRD", 1) = 0 Then           '5-13-03
        Set fgReportForm = RptSelRD
        RptSelRD.Show vbModal
    ElseIf StrComp(slExe, "RptSelRevEvt", 1) = 0 Then       '10-15-14 Revenue by Event
        Set fgReportForm = RptSelRevEvt
        RptSelRevEvt.Show vbModal
    ElseIf StrComp(slExe, "RptSelRG", 1) = 0 Then           '12-22-09  Regional copy assignment
        Set fgReportForm = RptSelRg
        RptSelRg.Show vbModal
    ElseIf StrComp(slExe, "RptSelRk", 1) = 0 Then          '7-30-12 Spot Price Ranking
        Set fgReportForm = RptSelRk
        RptSelRk.Show vbModal
    ElseIf StrComp(slExe, "RptSelRI", 1) = 0 Then
        Set fgReportForm = RptSelRI
        RptSelRI.Show vbModal
    ElseIf StrComp(slExe, "RptSelRP", 1) = 0 Then           '11-1-02 Remote Posting
        Set fgReportForm = RptSelRP
        RptSelRP.Show vbModal
    ElseIf StrComp(slExe, "RptSelrs", 1) = 0 Then
        Set fgReportForm = RptSelRS
        RptSelRS.Show vbModal
    ElseIf StrComp(slExe, "RptSelrr", 1) = 0 Then           '6-20-03 Research Revenue
        Set fgReportForm = RptSelRR
        RptSelRR.Show vbModal
    ElseIf StrComp(slExe, "RptSelrv", 1) = 0 Then
        Set fgReportForm = RptSelRV
        RptSelRV.Show vbModal
    ElseIf StrComp(slExe, "RptSelca", 1) = 0 Then      '12/18/07 combo avails
        Set fgReportForm = RptSelCA
        RptSelCA.Show vbModal
    ElseIf StrComp(slExe, "RptSelSN", 1) = 0 Then      '04-10-08 Split Network Avails
        Set fgReportForm = RptSelSN
        RptSelSN.Show vbModal
    ElseIf StrComp(slExe, "RptSelsp", 1) = 0 Then
        Set fgReportForm = RptSelSP
        RptSelSP.Show vbModal
    ElseIf StrComp(slExe, "RptSelspotBB", 1) = 0 Then       'Business on the books for air time, ntr and rep
        Set fgReportForm = RptSelSpotBB
        RptSelSpotBB.Show vbModal
    ElseIf StrComp(slExe, "RptSelSR", 1) = 0 Then           '9-19-06 split regions
        Set fgReportForm = RptSelSR
        RptSelSR.Show vbModal
    ElseIf StrComp(slExe, "RptSelss", 1) = 0 Then
        'RptSelss.Show vbModal
    ElseIf StrComp(slExe, "RptSeltx", 1) = 0 Then
        Set fgReportForm = RptSelTx
        RptSelTx.Show vbModal
    ElseIf StrComp(slExe, "RptSelus", 1) = 0 Then
        Set fgReportForm = RptSelUS
        RptSelUS.Show vbModal
    ElseIf StrComp(slExe, "RptSelAvgCmp", 1) = 0 Then
        Set fgReportForm = RptSelAvgCmp
        RptSelAvgCmp.Show vbModal
    ElseIf StrComp(UCase(slExe), UCase("ExptReRate"), 1) = 0 Or StrComp(UCase(slExe), UCase("ReRate"), 1) = 0 Then
        ExptReRate.Show vbModal
'    ElseIf StrComp(UCase(slExe), UCase("rptselpodbil"), 1) = 0 Then
'        RptSelPodBil.Show vbModal
    Else
        bmInModalModule = False
        MsgBox "Missing test for " & slExe
        Exit Sub
    End If
    bmInModalModule = False
    gChDrDir        '3-25-03
    'ChDrive Left$(sgCurDir, 2)  'Set the default drive
    'ChDir sgCurDir              'set the default path
    'If Not igStdAloneMode Then
    '    'tmcDelay.Enabled = True
    '    mTerminate False
    'Else
    '    ReportList.Enabled = False
    '    Do While Not igChildDone
    '        DoEvents
    '    Loop
    '    slStr = sgDoneMsg
    '    ReportList.Enabled = True
    '    edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    '    For ilLoop = 0 To 10
    '        DoEvents
    '    Next ilLoop
    'End If
    If Not igRptReturn Then
        tmcDelay.Enabled = True
        'mTerminate
    End If
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
    '5676 lose hard coded c: '8903 added urf code
    slFullPath = sgRootDrive & "csi\ReportPasser-" & tgUrf(0).iCode & ".txt"
    'slFullPath = sgRootDrive & "csi\ReportPasser.txt"
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
    gMsgBox "traffic reports couldn't read values in " & slFullPath & ".  Form_GotFocus", vbOKOnly, "Problem reading values from file."
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
Private Sub cmcDone_GotFocus(Index As Integer)
    If imFirstFocus Then
        imFirstFocus = False
    End If
    gCtrlGotFocus ActiveControl
End Sub

Private Sub CSI_ComboBoxMS1_OnChange()
    'select the same item in the old list as the new Combo has selected... (so images & Description's load, etc.)
    Dim illoop As Integer
    If CSI_ComboBoxMS1.ListIndex > -1 Then
        For illoop = 0 To lbcRpt.ListCount - 1
            If lbcRpt.List(illoop) = CSI_ComboBoxMS1.Text Then
                lbcRpt.ListIndex = illoop
            End If
        Next illoop
    End If
End Sub

Private Sub edcRptDescription_GotFocus()
    pbcClickFocus.SetFocus
End Sub
Private Sub edcRptDescription_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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
    'ReportList.Refresh
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
    Dim illoop As Integer

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
            igRptCallType = Val(slStr)
            If rbcShowBy(0).Value Then
                If (igRptCallType >= 20) And (igRptCallType <= 89) Then  'List Function
                    'List Function
                    For illoop = 0 To UBound(tmReportList) - 1 Step 1
                        If (tmReportList(illoop).tRnf.sType = "C") And (tmReportList(illoop).tRnf.iJobListNo = -1) Then
                            lbcRpt.TopIndex = illoop
                            lbcRpt.ListIndex = illoop
                            Exit For
                        End If
                    Next illoop
                ElseIf igRptCallType < 20 Then  'Job function
                    For illoop = 0 To UBound(tmReportList) - 1 Step 1
                        If (tmReportList(illoop).tRnf.sType = "C") And (tmReportList(illoop).tRnf.iJobListNo = igRptCallType) Then
                            lbcRpt.TopIndex = illoop
                            lbcRpt.ListIndex = illoop
                            Exit For
                        End If
                    Next illoop
                End If
            End If
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
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    tmcClock.Enabled = False
    Erase tmSelSrf
    Erase tmReportList
    Erase tgRtfList
    Erase tgRnfList
    'gObtainRNF hmRnf
    btrExtClear hmRtf   'Clear any previous extend operation
    ilRet = btrClose(hmRtf)
    btrDestroy hmRtf
    btrExtClear hmSrf   'Clear any previous extend operation
    ilRet = btrClose(hmSrf)
    btrDestroy hmSrf
    btrExtClear hmSnf   'Clear any previous extend operation
    ilRet = btrClose(hmSnf)
    btrDestroy hmSnf
    btrExtClear hmRnf   'Clear any previous extend operation
    ilRet = btrClose(hmRnf)
    btrDestroy hmRnf
    '4/2/11: Add setting and call.
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    Set ReportList = Nothing   'Remove data segment
    'Reset used instead of Close to cause # Clients on network to be decrement
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    'btrStopAppl
    
    Erase tgAcqComm
    Erase tgAcqCommInx

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

Private Sub hbcRptSample_Change()
    pbcRptSample(1).Left = -hbcRptSample.Value
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcRpt_Click()
    Dim slFromFile As String
    Dim slName As String
    Dim ilRpt As Integer
    Screen.MousePointer = vbHourglass
    vbcRptSample.Value = vbcRptSample.Min
    hbcRptSample.Value = hbcRptSample.Min
    pbcRptSample(1).Move 0, 0
    If lbcRpt.ListIndex >= 0 Then
        slName = Trim$(lbcRpt.List(lbcRpt.ListIndex))
        For ilRpt = 0 To UBound(tmReportList) - 1 Step 1
            If slName = Trim$(tmReportList(ilRpt).tRnf.sName) Then
                'edcRptDescription.Text = Left$(tmReportList(ilRpt).tRnf.sDescription, tmReportList(ilRpt).tRnf.iStrLen)
                edcRptDescription.Text = gStripChr0(tmReportList(ilRpt).tRnf.sDescription)
                slFromFile = tmReportList(ilRpt).tRnf.sRptSample
                On Error GoTo lbcRptErr:
                pbcRptSample(1).Picture = LoadPicture(sgRptPath & slFromFile)
                vbcRptSample.Max = pbcRptSample(1).Height - pbcRptSample(0).Height
                vbcRptSample.Enabled = (pbcRptSample(0).Height < pbcRptSample(1).Height)
                If vbcRptSample.Enabled Then
                    vbcRptSample.SmallChange = pbcRptSample(0).Height
                    vbcRptSample.LargeChange = pbcRptSample(0).Height
                End If
                hbcRptSample.Max = pbcRptSample(1).Width - pbcRptSample(0).Width
                hbcRptSample.Enabled = (pbcRptSample(0).Width < pbcRptSample(1).Width)
                If hbcRptSample.Enabled Then
                    hbcRptSample.SmallChange = pbcRptSample(0).Width
                    hbcRptSample.LargeChange = pbcRptSample(0).Width
                End If
                If InStr(1, UCase$(slName), "RERATE") > 0 Then
                    cmcDone(0).Enabled = False
                    cmcDone(2).Enabled = False
                    cmcDone(1).Caption = "Yes-With Exceptions"
                Else
                    cmcDone(0).Enabled = True
                    cmcDone(2).Enabled = True
                    cmcDone(1).Caption = "Yes-Except Dates"
                End If
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Next ilRpt
    End If
    cmcDone(0).Enabled = True
    cmcDone(2).Enabled = True
    cmcDone(1).Caption = "Yes-Except Dates"
    edcRptDescription.Text = ""
    pbcRptSample(1).Picture = LoadPicture()
    vbcRptSample.Max = vbcRptSample.Min
    vbcRptSample.Enabled = False
    hbcRptSample.Max = hbcRptSample.Min
    hbcRptSample.Enabled = False
    Screen.MousePointer = vbDefault
    Exit Sub
lbcRptErr:
    pbcRptSample(1).Picture = LoadPicture()
    Resume Next
End Sub

Private Sub lbcRpt_DblClick()
    cmcDone_Click 0
End Sub

Private Sub lbcRpt_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
End Sub
Private Sub lbcRpt_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim slName As String
    Dim ilBlanks As Integer
    Dim illoop As Integer
    Dim ilLen As Integer
    If (Shift And vbAltMask) = ALTMASK Then
        If KeyCode = KeyDown Then
            'Find next same level
            If lbcRpt.ListIndex >= 0 Then
                slName = lbcRpt.List(lbcRpt.ListIndex)
                ilBlanks = Len(slName) - Len(LTrim$(slName))
                For illoop = lbcRpt.ListIndex + 1 To lbcRpt.ListCount - 1 Step 1
                    slName = lbcRpt.List(illoop)
                    ilLen = Len(slName) - Len(LTrim$(slName))
                    If ilBlanks = ilLen Then
                        lbcRpt.TopIndex = illoop
                        lbcRpt.ListIndex = illoop
                        Exit Sub
                    ElseIf ilLen < ilBlanks Then
                        lbcRpt.TopIndex = lbcRpt.ListIndex + 1
                        lbcRpt.ListIndex = lbcRpt.ListIndex + 1
                        Exit Sub
                    End If
                Next illoop
                If lbcRpt.ListIndex < lbcRpt.ListCount - 1 Then
                    lbcRpt.TopIndex = lbcRpt.ListIndex + 1
                    lbcRpt.ListIndex = lbcRpt.ListIndex + 1
                    Exit Sub
                End If
            End If
        End If
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
    Dim slDSN As String

    Screen.MousePointer = vbHourglass
    bmInModalModule = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    'ReportList.Height = cmcDone.Top + 3 * cmcDone.Height '/ 3
    ReportList.Height = frcGen.Top + 2.2 * frcGen.Height '/ 3
    gCenterStdAlone ReportList
    'ReportList.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    smScreenCaption = "Report Selection"
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imFirstFocus = True
    hmRnf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmRnf, "", sgDBPath & "Rnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rnf.Btr)", ReportList
    On Error GoTo 0
    imRnfRecLen = Len(tmRnf)
    hmSnf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmSnf, "", sgDBPath & "Snf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Snf.Btr)", ReportList
    On Error GoTo 0
    imSnfRecLen = Len(tmSnf)
    hmSrf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmSrf, "", sgDBPath & "Srf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Srf.Btr)", ReportList
    On Error GoTo 0
    imSrfRecLen = Len(tmSrf)
    hmRtf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmRtf, "", sgDBPath & "Rtf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rtf.Btr)", ReportList
    On Error GoTo 0
    imRtfRecLen = Len(tmRtf)
    mRptNameMap
    imIgnoreChg = True
    If (tgUrf(0).iSlfCode > 0) Or (tgUrf(0).iRemoteUserID > 0) Or (rbcShowBy(1).Value) Then
        rbcShowBy(1).Value = True
        rbcShowBy(0).Enabled = False
    Else
        rbcShowBy(igShowCatOrNames).Value = True
    End If
    imIgnoreChg = False
    mPopulate
    ilRet = gObtainAdvt()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    
    ilRet = gObtainAgency()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    
    ilRet = gVffRead()
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    
    ilRet = gBuildAcqCommInfo(ReportList)
    If ilRet = False Then
        imTerminate = True
        Exit Sub
    End If
    
    lbcRpt.Visible = True
    edcRptDescription.Visible = True
    If imTerminate Then
        Exit Sub
    End If
    plcScreen_Paint
'    gCenterModalForm ReportList
    Screen.MousePointer = vbDefault
    
    ''D.S. 07/20/15 Startup Relational/SQL Engine
    'If Not gLoadOption("Locations", "Name", sgDatabaseName) Then
    '    gMsgBox "Traffic.Ini [Locations] 'Name' key is missing.", vbCritical
    'End If
    'Set cnn = New ADODB.Connection
    'slDSN = sgDatabaseName
    'On Error GoTo ERRNOPERVASIVE
    'ilRet = 0
    'cnn.Open "DSN=" & slDSN
    
    igGGFlag = 1
    imLastHourGGChecked = -1
    tmcClock_Timer
    
    Exit Sub
ERRNOPERVASIVE:
    ilRet = 1
    Resume Next
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainSetSrf                   *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain the selected reports of *
'*                      a set for a user               *
'*                                                     *
'*******************************************************
Private Sub mObtainSrf(ilSnfCode As Integer, hlSrf As Integer, tlSelSrf() As SRF)
'
'   gObtainSrf ilSnfCode, hlSrf, tmSelSrf()
'   Where:
'       gObtainRNF must be called prior to this call to load tgRNFLIST
'
    Dim ilSortCode As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim llRecPos As Long
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    ilSortCode = 0
    ReDim tlSelSrf(0 To 0) As SRF   'VB list box clear (list box used to retain code number so record can be found)
    imSrfRecLen = Len(tlSelSrf(0)) 'btrRecordLength(hlSrf)  'Get and save record length
    ilExtLen = Len(tlSelSrf(0))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSrf) 'Obtain number of records
    btrExtClear hlSrf   'Clear any previous extend operation
    tmSrfSrchKey1.iCode = ilSnfCode
    ilRet = btrGetGreaterOrEqual(hlSrf, tmSrf, imSrfRecLen, tmSrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    tlIntTypeBuff.iType = ilSnfCode
    ilOffSet = 4    'gFieldOffset("Prf", "PrfAdfCode")
    ilRet = btrExtAddLogicConst(hlSrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
    Call btrExtSetBounds(hlSrf, llNoRec, -1, "UC", "SRF", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlSrf, ilOffSet, ilExtLen)  'Extract First Name field
    If ilRet = BTRV_ERR_NONE Then
        'ilRet = btrExtGetNextExt(hlSrf)    'Extract record
        ilRet = btrExtGetNext(hlSrf, tlSelSrf(ilSortCode), ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                ilExtLen = Len(tlSelSrf(ilSortCode))  'Extract operation record size
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSrf, tlSelSrf(ilSortCode), ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    If ilSortCode >= UBound(tlSelSrf) Then
                        ReDim Preserve tlSelSrf(0 To UBound(tlSelSrf) + 100) As SRF
                    End If
                    ilSortCode = ilSortCode + 1
                    ilRet = btrExtGetNext(hlSrf, tlSelSrf(ilSortCode), ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlSrf, tlSelSrf(ilSortCode), ilExtLen, llRecPos)
                    Loop
                Loop
                ReDim Preserve tlSelSrf(0 To ilSortCode) As SRF
            End If
        End If
    End If
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
Private Sub mPopulate(Optional ilWhichList As Integer = 0)
    Dim ilRet As Integer 'btrieve status
    Dim slName As String
    Dim llLen As Long
    Dim illoop As Integer
    Dim ilLevel As Integer
    Dim ilSrf As Integer
    Dim ilRnf As Integer
    Dim ilIndex As Integer
    Dim ilAllowed As Integer
    Dim ilRpt As Integer
'    slGetStamp = gGetCSIStamp("ReportList")
'    If StrComp(slGetStamp, "ReportList", 1) = 0 Then
'        ilRet = csiGetAlloc("ReportList", ilStartIndex, ilEndIndex)
'    Else
'        ilRet = 1
'    End If
    'If (StrComp(slGetStamp, "ReportList", 1) = 0) And (ilRet = 0) Then
    '    ReDim tmReportList(ilStartIndex To ilEndIndex) As RPTLST
    '    For ilLoop = LBound(tmReportList) To UBound(tmReportList) Step 1
    '        ilRet = csiGetRec("ReportList", ilLoop, VarPtr(tmReportList(ilLoop)), LenB(tmReportList(ilLoop)))
    '    Next ilLoop
    '    lbcRpt.Clear
    'Else
        'gObtainRNF hmRnf
        lbcRpt.Visible = False
        CSI_ComboBoxMS1.Visible = False
        CSI_ComboBoxMS1.FontBold = True
        
        lbcRpt.Clear
        CSI_ComboBoxMS1.Clear
        
        gObtainRTF hmRtf, False
        If (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            ilRet = mReadRec()
            If ilRet = False Then
                lbcRpt.Clear
                ReDim tmReportList(0 To 0) As RPTLST
                MsgBox "No Report Selection Allowed", vbOKOnly + vbInformation, "Report List"
                Exit Sub
            End If
        End If
        
        ReDim tmReportList(0 To 0) As RPTLST
        llLen = 0
        If (UBound(tgRtfList) <= LBound(tgRtfList)) Or ((Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName)) Then
            gObtainRNF hmRnf
            'Dan M 4/09/09  Limit guide to one report.
            mLimitGuideReports tgRnfList
        Else
            ReDim tgRnfList(0 To 1) As RNFLIST
        End If
        If UBound(tgRtfList) > LBound(tgRtfList) Then
            For illoop = 0 To UBound(tgRtfList) - 1 Step 1
                If (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                    If tgRtfList(illoop).tRtf.sRnfType <> "C" Then
                        ilAllowed = False
                        For ilSrf = 0 To UBound(tmSelSrf) - 1 Step 1
                            If tgRtfList(illoop).tRtf.iRnfCode = tmSelSrf(ilSrf).iRnfCode Then
                                If tgRtfList(illoop).tRtf.sRnfState <> "D" Then
                                    ilAllowed = True
                                    If (UBound(tgRtfList) <= LBound(tgRtfList)) Or ((Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName)) Then
                                    Else
                                        imRnfRecLen = Len(tmRnf)  'Get and save CmF record length (the read will change the length)
                                        tmRnfSrchKey.iCode = tgRtfList(illoop).tRtf.iRnfCode
                                        ilRet = btrGetEqual(hmRnf, tmRnf, imRnfRecLen, tmRnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet = BTRV_ERR_NONE Then
                                            'slName = tmRnf.sName
                                            'If tmRnf.sType = "C" Then
                                            '    slName = "C|" & slName & "|" & tmRnf.sState & "\" & Trim$(Str$(tmRnf.iCode))
                                            'Else
                                            '    slName = "R|" & slName & "|" & tmRnf.sState & "\" & Trim$(Str$(tmRnf.iCode))
                                            'End If
                                            'tgRnfList(UBound(tgRnfList)).sKey = slName
                                            'tgRnfList(UBound(tgRnfList)).tRnf = tmRnf
                                            'ReDim tgRnfList(0 To UBound(tgRnfList) + 1) As RNFLIST
                                            tgRnfList(0).tRnf = tmRnf
                                        Else
                                            ilAllowed = False
                                        End If
                                    End If
                                End If
                                Exit For
                            End If
                        Next ilSrf
                    Else
                        ilAllowed = True
                        If (UBound(tgRtfList) <= LBound(tgRtfList)) Or ((Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName)) Then
                        Else
                            imRnfRecLen = Len(tmRnf)  'Get and save CmF record length (the read will change the length)
                            tmRnfSrchKey.iCode = tgRtfList(illoop).tRtf.iRnfCode
                            ilRet = btrGetEqual(hmRnf, tmRnf, imRnfRecLen, tmRnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                'slName = tmRnf.sName
                                'If tmRnf.sType = "C" Then
                                '    slName = "C|" & slName & "|" & tmRnf.sState & "\" & Trim$(Str$(tmRnf.iCode))
                                'Else
                                '    slName = "R|" & slName & "|" & tmRnf.sState & "\" & Trim$(Str$(tmRnf.iCode))
                                'End If
                                'tgRnfList(UBound(tgRnfList)).sKey = slName
                                'tgRnfList(UBound(tgRnfList)).tRnf = tmRnf
                                'ReDim tgRnfList(0 To UBound(tgRnfList) + 1) As RNFLIST
                                tgRnfList(0).tRnf = tmRnf
                            Else
                                ilAllowed = False
                            End If
                        End If
                    End If
                Else
                    ilAllowed = True
                End If
                ilIndex = -1
                For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
                    If tgRnfList(ilRnf).tRnf.iCode = tgRtfList(illoop).tRtf.iRnfCode Then
                        ilIndex = ilRnf
                        If tgRnfList(ilRnf).tRnf.sType = "C" Then
                            ilAllowed = True
                        End If
                        Exit For
                    End If
                Next ilRnf
                ilAllowed = mCheckReportName(ilAllowed, ilIndex)
                If (ilAllowed) And (ilIndex >= 0) Then
                    slName = Trim$(tgRnfList(ilIndex).tRnf.sName)
                    For ilLevel = 1 To tgRtfList(illoop).tRtf.iLevel - 1 Step 1
                        slName = "  " & slName
                    Next ilLevel
                    tmReportList(UBound(tmReportList)).sName = slName
                    tmReportList(UBound(tmReportList)).tRnf = tgRnfList(ilIndex).tRnf
                    tmReportList(UBound(tmReportList)).iLevel = tgRtfList(illoop).tRtf.iLevel
                    ReDim Preserve tmReportList(0 To UBound(tmReportList) + 1) As RPTLST
                End If
            Next illoop
        Else
            For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
                If (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                    ilAllowed = False
                    For ilSrf = 0 To UBound(tmSelSrf) - 1 Step 1
                        If tgRnfList(ilRnf).tRnf.iCode = tmSelSrf(ilSrf).iRnfCode Then
                            If tgRtfList(illoop).tRtf.sRnfState <> "D" Then
                                ilAllowed = True
                            End If
                            Exit For
                        End If
                    Next ilSrf
                Else
                    ilAllowed = True
                End If
                ilAllowed = mCheckReportName(ilAllowed, ilRnf)
                If ilAllowed Then
                    slName = Trim$(tgRnfList(ilRnf).tRnf.sName)
                    tmReportList(UBound(tmReportList)).sName = slName
                    tmReportList(UBound(tmReportList)).tRnf = tgRnfList(ilRnf).tRnf
                    tmReportList(UBound(tmReportList)).iLevel = 0
                    ReDim Preserve tmReportList(0 To UBound(tmReportList) + 1) As RPTLST
                End If
            Next ilRnf
        End If
        'Remove Proposal reports if Not using Proposal system
        If (Trim$(tgUrf(0).sName) <> sgCPName) And (tgSpf.sGUsePropSys <> "Y") Then
            illoop = LBound(tmReportList)
            Do
                If tmReportList(illoop).tRnf.sType = "C" Then
                    If (tmReportList(illoop).iLevel <= 1) And (StrComp("Proposals", Trim$(tmReportList(illoop).tRnf.sName), 1) = 0) Then
                        ilIndex = illoop + 1
                        Do
                            If tmReportList(ilIndex).iLevel = tmReportList(illoop).iLevel Then
                                Exit Do
                            End If
                            ilIndex = ilIndex + 1
                        Loop While ilIndex < UBound(tmReportList)
                        For ilRpt = illoop To UBound(tmReportList) - (ilIndex - illoop) Step 1
                            tmReportList(ilRpt) = tmReportList(ilRpt + (ilIndex - illoop))
                        Next ilRpt
                        ReDim Preserve tmReportList(0 To UBound(tmReportList) - (ilIndex - illoop)) As RPTLST
                        Exit Do
                    Else
                        illoop = illoop + 1
                    End If
                Else
                    illoop = illoop + 1
                End If
            Loop While illoop <= UBound(tmReportList) - 1
        End If
        'Remove all level except Proposals and List if Using traffic is No
        If (Trim$(tgUrf(0).sName) <> sgCPName) And (tgSpf.sUsingTraffic = "N") Then
            illoop = LBound(tmReportList)
            Do
                If tmReportList(illoop).tRnf.sType = "C" Then
                    If (tmReportList(illoop).iLevel <= 1) And ((StrComp("Proposals", Trim$(tmReportList(illoop).tRnf.sName), 1) <> 0) And (StrComp("Lists", Trim$(tmReportList(illoop).tRnf.sName), 1) <> 0)) Then
                        ilIndex = illoop + 1
                        Do
                            If tmReportList(ilIndex).iLevel = tmReportList(illoop).iLevel Then
                                Exit Do
                            End If
                            ilIndex = ilIndex + 1
                        Loop While ilIndex < UBound(tmReportList)
                        For ilRpt = illoop To UBound(tmReportList) - (ilIndex - illoop) Step 1
                            tmReportList(ilRpt) = tmReportList(ilRpt + (ilIndex - illoop))
                        Next ilRpt
                        ReDim Preserve tmReportList(0 To UBound(tmReportList) - (ilIndex - illoop)) As RPTLST
                    Else
                        illoop = illoop + 1
                    End If
                Else
                    illoop = illoop + 1
                End If
            Loop While illoop <= UBound(tmReportList) - 1
        End If
        'Remove unused levels b
        illoop = LBound(tmReportList)
        Do
            If tmReportList(illoop).tRnf.sType = "C" Then
                If illoop + 1 >= UBound(tmReportList) Then
                    ReDim Preserve tmReportList(0 To UBound(tmReportList) - 1) As RPTLST
                    illoop = illoop - 1
                Else
                    For ilIndex = illoop + 1 To UBound(tmReportList) - 1 Step 1
                        If tmReportList(ilIndex).tRnf.sType <> "C" Then
                            'ilLoop = ilIndex + 1
                            Exit For
                        Else
                            If tmReportList(ilIndex).iLevel <= tmReportList(illoop).iLevel Then
                                'Remove leveles from ilLoop to ilIndex
                                For ilRpt = illoop To UBound(tmReportList) - (ilIndex - illoop) Step 1
                                    tmReportList(ilRpt) = tmReportList(ilRpt + (ilIndex - illoop))
                                Next ilRpt
                                ReDim Preserve tmReportList(0 To UBound(tmReportList) - (ilIndex - illoop)) As RPTLST
                                illoop = illoop - 1
                                Exit For
                            End If
                        End If
                    Next ilIndex
                    illoop = illoop + 1
                End If
            Else
                illoop = illoop + 1
            End If
        Loop While illoop <= UBound(tmReportList) - 1
        'If Salesperson or Remote User remove categories and duplicates
        If (tgUrf(0).iSlfCode > 0) Or (tgUrf(0).iRemoteUserID > 0) Or (rbcShowBy(1).Value Or rbcShowBy(2).Value) Then
            illoop = LBound(tmReportList)
            Do
                If tmReportList(illoop).tRnf.sType = "C" Then
                    If illoop + 1 >= UBound(tmReportList) Then
                        ReDim Preserve tmReportList(0 To UBound(tmReportList) - 1) As RPTLST
                    Else
                        For ilIndex = illoop + 1 To UBound(tmReportList) - 1 Step 1
                            tmReportList(ilIndex - 1) = tmReportList(ilIndex)
                        Next ilIndex
                        ReDim Preserve tmReportList(0 To UBound(tmReportList) - 1) As RPTLST
                    End If
                Else
                    illoop = illoop + 1
                End If
            Loop While illoop <= UBound(tmReportList) - 1
            illoop = LBound(tmReportList)
            '3/1/10:  Handle case where only one name exist
            Do
                ilIndex = illoop + 1
                '3/1/10:  Handle case where only one name exist
                'Do
                'Do While ilLoop < UBound(tmReportList) - 1
                'Do While ilIndex < UBound(tmReportList) - 1
                '8/13/11: Row x matches UBound()-1 row, it was not removed
                Do While ilIndex <= UBound(tmReportList) - 1
                    If StrComp(Trim$(tmReportList(illoop).sName), Trim$(tmReportList(ilIndex).sName), 1) = 0 Then
                        For ilRpt = ilIndex + 1 To UBound(tmReportList) - 1 Step 1
                            tmReportList(ilRpt - 1) = tmReportList(ilRpt)
                        Next ilRpt
                        ReDim Preserve tmReportList(0 To UBound(tmReportList) - 1) As RPTLST
                    Else
                        ilIndex = ilIndex + 1
                    End If
                'Loop While ilIndex <= UBound(tmReportList) - 1
                Loop
                illoop = illoop + 1
            Loop While illoop < UBound(tmReportList) - 1   '= UBound(tmReportList) - 1
            For illoop = 0 To UBound(tmReportList) - 1 Step 1
                tmReportList(illoop).sName = Trim$(tmReportList(illoop).sName)
            Next illoop
            If UBound(tmReportList) - 1 > 0 Then
                ArraySortTyp fnAV(tmReportList(), 0), UBound(tmReportList), 0, LenB(tmReportList(0)), 0, LenB(tmReportList(0).sName), 0
            End If
        End If
    '    ilRet = csiSetStamp("ReportList", "ReportList")
    '    ilRet = csiSetAlloc("ReportList", LBound(tmReportList), UBound(tmReportList))
    '    For ilLoop = LBound(tmReportList) To UBound(tmReportList) Step 1
    '        ilRet = csiSetRec("ReportList", ilLoop, VarPtr(tmReportList(ilLoop)), LenB(tmReportList(ilLoop)))
    '    Next ilLoop
    'End If
    For illoop = 0 To UBound(tmReportList) - 1 Step 1
        slName = RTrim$(tmReportList(illoop).sName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        If ilWhichList = 0 Then
            lbcRpt.AddItem slName  'Add ID to list box
        Else
            lbcRpt.AddItem slName  'Add ID to list box
            CSI_ComboBoxMS1.AddItem slName
        End If
    Next illoop
    If (igRptCallType >= 20) And (igRptCallType <= 89) Then  'List Function
        'List Function
        For illoop = 0 To UBound(tmReportList) - 1 Step 1
            If (tmReportList(illoop).tRnf.sType = "C") And (tmReportList(illoop).tRnf.iJobListNo = -1) Then
                If ilWhichList = 0 Then
                    lbcRpt.TopIndex = illoop
                    lbcRpt.ListIndex = illoop
                Else
                    CSI_ComboBoxMS1.ListIndex = illoop
                End If
                Exit For
            End If
        Next illoop
    ElseIf igRptCallType < 20 Then  'Job function
        For illoop = 0 To UBound(tmReportList) - 1 Step 1
            If (tmReportList(illoop).tRnf.sType = "C") And (tmReportList(illoop).tRnf.iJobListNo = igRptCallType) Then
                If ilWhichList = 0 Then
                    lbcRpt.TopIndex = illoop
                    lbcRpt.ListIndex = illoop
                Else
                    CSI_ComboBoxMS1.ListIndex = illoop
                End If
                
                Exit For
            End If
        Next illoop
    End If
    If ilWhichList = 0 Then
        lbcRpt.Visible = True
    Else
        CSI_ComboBoxMS1.Visible = True
    End If
    
    If lbcRpt.ListIndex < 0 Then
        If ilWhichList = 0 Then
            If lbcRpt.ListCount > 0 Then
                lbcRpt.ListIndex = 0
            End If
        Else
            If CSI_ComboBoxMS1.ListCount > 0 Then
                'CSI_ComboBoxMS1.SetListIndex = 0
            End If
        End If
    End If
    
End Sub
Private Sub mLimitGuideReports(ByRef tgRnfList() As RNFLIST)
Dim c As Integer
If (Trim$(tgUrf(0).sName) = sgSUName And Not bgInternalGuide) Then
    For c = 0 To UBound(tgRnfList) - 1
        If tgRnfList(c).tRnf.iCode = 22 Then
            tgRnfList(0).sKey = tgRnfList(c).sKey
            tgRnfList(0).tRnf = tgRnfList(c).tRnf
            ReDim Preserve tgRnfList(0 To 1) As RNFLIST
            Exit For
        End If
    Next c
End If

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec() As Integer
'
'   iRet = mReadRec(ilSelectIndex)
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    'slNameCode = tgSnfCode(ilSelectIndex - 1).sKey   'lbcTitleCode.List(ilSelectIndex - 1)
    'ilRet = gParseItem(slNameCode, 2, "\", slCode)
    'On Error GoTo mReadRecErr
    'gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", ReportList
    'On Error GoTo 0
    'slCode = Trim$(slCode)
    tmSnfSrchKey.iCode = tgUrf(0).iSnfCode
    imSnfRecLen = Len(tmSnf)  'Get and save CmF record length (the read will change the length)
    ilRet = btrGetEqual(hmSnf, tmSnf, imSnfRecLen, tmSnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        mReadRec = False
        Exit Function
    End If
    smScreenCaption = "Report Selection- " & Trim$(tmSnf.sName)
    plcScreen_Paint
    mObtainSrf tmSnf.iCode, hmSrf, tmSelSrf()
    mReadRec = True
    Exit Function

    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gRptNameMap                      *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create cross reference Table   *
'*                                                     *
'*******************************************************
Private Sub mRptNameMap()
    'RptNoSel
    tmRptNoSelNameMap(0).sName = "Item Billing Types"
    tmRptNoSelNameMap(0).iRptCallType = ITEMBILLINGTYPESLIST
    tmRptNoSelNameMap(1).sName = "Invoice Sorts"
    tmRptNoSelNameMap(1).iRptCallType = INVOICESORTLIST
    tmRptNoSelNameMap(2).sName = "Program Exclusions"
    tmRptNoSelNameMap(2).iRptCallType = EXCLUSIONSLIST
    tmRptNoSelNameMap(3).sName = "Announcer Names"
    tmRptNoSelNameMap(3).iRptCallType = ANNOUNCERNAMESLIST
    tmRptNoSelNameMap(4).sName = "Genre Names"
    tmRptNoSelNameMap(4).iRptCallType = GENRESLIST
    tmRptNoSelNameMap(5).sName = "Sales Regions"
    tmRptNoSelNameMap(5).iRptCallType = SALESREGIONSLIST
    tmRptNoSelNameMap(6).sName = "Sales Sources"
    tmRptNoSelNameMap(6).iRptCallType = SALESSOURCESLIST
    tmRptNoSelNameMap(7).sName = "Sales Teams"
    tmRptNoSelNameMap(7).iRptCallType = SALESTEAMSLIST
    tmRptNoSelNameMap(8).sName = "Revenue Sets"
    tmRptNoSelNameMap(8).iRptCallType = REVENUESETSLIST
    tmRptNoSelNameMap(9).sName = "Boilerplate"
    tmRptNoSelNameMap(9).iRptCallType = BOILERPLATESLIST
    tmRptNoSelNameMap(10).sName = "Missed Reasons"
    tmRptNoSelNameMap(10).iRptCallType = MISSEDREASONSLIST
    tmRptNoSelNameMap(11).sName = "Product Protection"
    tmRptNoSelNameMap(11).iRptCallType = COMPETITIVESLIST
    tmRptNoSelNameMap(12).sName = "Feed Type"
    tmRptNoSelNameMap(12).iRptCallType = FEEDTYPESLIST
    tmRptNoSelNameMap(13).sName = "Event Types"
    tmRptNoSelNameMap(13).iRptCallType = EVENTTYPESLIST
    tmRptNoSelNameMap(14).sName = "Avail Names"
    tmRptNoSelNameMap(14).iRptCallType = AVAILNAMESLIST
    tmRptNoSelNameMap(15).sName = "Sales Offices"
    tmRptNoSelNameMap(15).iRptCallType = SALESOFFICESLIST
    tmRptNoSelNameMap(16).sName = "Media Definitions"
    tmRptNoSelNameMap(16).iRptCallType = MEDIADEFINITIONSLIST
    tmRptNoSelNameMap(17).sName = "Lock Boxes"
    tmRptNoSelNameMap(17).iRptCallType = LOCKBOXESLIST
    tmRptNoSelNameMap(18).sName = "Agency DP Services"
    tmRptNoSelNameMap(18).iRptCallType = EDISERVICESLIST
    tmRptNoSelNameMap(19).sName = "Transaction Types"
    tmRptNoSelNameMap(19).iRptCallType = TRANSACTIONSLIST
    tmRptNoSelNameMap(20).sName = "Site Options"
    tmRptNoSelNameMap(20).iRptCallType = SITELIST
    'tmRptNoSelNameMap(21).sName = "User Options"
    'tmRptNoSelNameMap(21).iRptCallType = USERLIST
    tmRptNoSelNameMap(22).sName = "Reconcile"
    tmRptNoSelNameMap(22).iRptCallType = COLLECTIONSJOB
    tmRptNoSelNameMap(23).sName = "Vehicle Participants"
    tmRptNoSelNameMap(23).iRptCallType = VEHICLEGROUPSLIST
    tmRptNoSelNameMap(24).sName = "Potential Codes"
    tmRptNoSelNameMap(24).iRptCallType = POTENTIALCODESLIST
    tmRptNoSelNameMap(25).sName = "Business Categories"
    tmRptNoSelNameMap(25).iRptCallType = BUSCATEGORIESLIST
    tmRptNoSelNameMap(26).sName = "Demos List"
    tmRptNoSelNameMap(26).iRptCallType = DEMOSLIST
    tmRptNoSelNameMap(27).sName = "Competitors"
    tmRptNoSelNameMap(27).iRptCallType = COMPETITORSLIST
    'RptSel
    tmRptSelNameMap(0).sName = "Vehicle Summary"
    tmRptSelNameMap(0).iRptCallType = VEHICLESLIST
    tmRptSelNameMap(1).sName = "Vehicle Options"
    tmRptSelNameMap(1).iRptCallType = VEHICLESLIST
    tmRptSelNameMap(2).sName = "Virtual Vehicles"
    tmRptSelNameMap(2).iRptCallType = VEHICLESLIST
    tmRptSelNameMap(3).sName = "Advertiser List"        '9-20-10 chged from Summary to List"
    tmRptSelNameMap(3).iRptCallType = ADVERTISERSLIST
    tmRptSelNameMap(4).sName = "Advertiser Detail"      'N/A
    tmRptSelNameMap(4).iRptCallType = ADVERTISERSLIST
    tmRptSelNameMap(5).sName = "Agency List"            '9-20-10 changed from summary to list
    tmRptSelNameMap(5).iRptCallType = AGENCIESLIST
    tmRptSelNameMap(6).sName = "Agency Detail"          'N/A
    tmRptSelNameMap(6).iRptCallType = AGENCIESLIST
    tmRptSelNameMap(7).sName = "Salespeople Summary"
    tmRptSelNameMap(7).iRptCallType = SALESPEOPLELIST
    'tmRptSelNameMap(8).sName = "Salespeople Budgets"           'unused
    'tmRptSelNameMap(8).iRptCallType = SALESPEOPLELIST
    tmRptSelNameMap(9).sName = "Event Names"
    tmRptSelNameMap(9).iRptCallType = EVENTNAMESLIST
    tmRptSelNameMap(10).sName = "Rate Card"
    tmRptSelNameMap(10).iRptCallType = RATECARDSJOB
    tmRptSelNameMap(11).sName = "Dayparts"
    tmRptSelNameMap(11).iRptCallType = RATECARDSJOB
    tmRptSelNameMap(12).sName = "Budgets"
    tmRptSelNameMap(12).iRptCallType = BUDGETSJOB
    tmRptSelNameMap(13).sName = "Budget Comparisons"
    tmRptSelNameMap(13).iRptCallType = BUDGETSJOB
    tmRptSelNameMap(14).sName = "Program Libraries"
    tmRptSelNameMap(14).iRptCallType = PROGRAMMINGJOB
    tmRptSelNameMap(15).sName = "Selling to Airing Vehicles"
    tmRptSelNameMap(15).iRptCallType = PROGRAMMINGJOB
    tmRptSelNameMap(16).sName = "Airing to Selling Vehicles"
    tmRptSelNameMap(16).iRptCallType = PROGRAMMINGJOB
    tmRptSelNameMap(17).sName = "Vehicle Avail Conflicts"
    tmRptSelNameMap(17).iRptCallType = PROGRAMMINGJOB
    tmRptSelNameMap(18).sName = "Delivery by Vehicle"
    tmRptSelNameMap(18).iRptCallType = PROGRAMMINGJOB
    tmRptSelNameMap(19).sName = "Delivery by Feed"
    tmRptSelNameMap(19).iRptCallType = PROGRAMMINGJOB
    tmRptSelNameMap(20).sName = "Engineering by Vehicle"
    tmRptSelNameMap(20).iRptCallType = PROGRAMMINGJOB
    tmRptSelNameMap(21).sName = "Engineering by Feed"
    tmRptSelNameMap(21).iRptCallType = PROGRAMMINGJOB
    tmRptSelNameMap(22).sName = "Salesperson Projection"
    tmRptSelNameMap(22).iRptCallType = PROPOSALPROJECTION
    tmRptSelNameMap(23).sName = "Vehicle Projection"
    tmRptSelNameMap(23).iRptCallType = PROPOSALPROJECTION
    tmRptSelNameMap(24).sName = "Sales Office Projection"
    tmRptSelNameMap(24).iRptCallType = PROPOSALPROJECTION
    tmRptSelNameMap(25).sName = "Category Projection"
    tmRptSelNameMap(25).iRptCallType = PROPOSALPROJECTION
    tmRptSelNameMap(26).sName = "Office Projection by Potential"
    tmRptSelNameMap(26).iRptCallType = PROPOSALPROJECTION
    tmRptSelNameMap(27).sName = "Cash Receipts or Usage"
    tmRptSelNameMap(27).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(28).sName = "Ageing by Payee"
    tmRptSelNameMap(28).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(29).sName = "Ageing by Salesperson"
    tmRptSelNameMap(29).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(30).sName = "Ageing by Vehicle"
    tmRptSelNameMap(30).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(31).sName = "Delinquent Accounts"
    tmRptSelNameMap(31).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(32).sName = "Statements"
    tmRptSelNameMap(32).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(33).sName = "Cash Payment or Usage History"
    tmRptSelNameMap(33).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(34).sName = "Advertiser and Agency Credit Status"
    tmRptSelNameMap(34).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(35).sName = "Cash Distribution"
    tmRptSelNameMap(35).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(36).sName = "Cash Summary"
    tmRptSelNameMap(36).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(37).sName = "Account History"
    tmRptSelNameMap(37).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(38).sName = "Merchandising/Promotions"
    tmRptSelNameMap(38).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(39).sName = "Merchandising/Promotions Recap"
    tmRptSelNameMap(39).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(40).sName = "Copy Status by Date"
    tmRptSelNameMap(40).iRptCallType = COPYJOB
    tmRptSelNameMap(41).sName = "Copy Status by Advertiser"
    tmRptSelNameMap(41).iRptCallType = COPYJOB
    tmRptSelNameMap(42).sName = "Contracts Missing Copy"
    tmRptSelNameMap(42).iRptCallType = COPYJOB
    tmRptSelNameMap(43).sName = "Copy Rotations by Advertiser"
    tmRptSelNameMap(43).iRptCallType = COPYJOB
    tmRptSelNameMap(44).sName = "Copy Inventory by Number"
    tmRptSelNameMap(44).iRptCallType = COPYJOB
    tmRptSelNameMap(45).sName = "Copy Inventory by ISCI"
    tmRptSelNameMap(45).iRptCallType = COPYJOB
    tmRptSelNameMap(46).sName = "Copy Inventory by Advertiser"
    tmRptSelNameMap(46).iRptCallType = COPYJOB
    tmRptSelNameMap(47).sName = "Copy Inventory by Start Date"
    tmRptSelNameMap(47).iRptCallType = COPYJOB
    tmRptSelNameMap(48).sName = "Copy Inventory by Expiration Date"
    tmRptSelNameMap(48).iRptCallType = COPYJOB
    tmRptSelNameMap(49).sName = "Copy Inventory by Purge Date"
    tmRptSelNameMap(49).iRptCallType = COPYJOB
    tmRptSelNameMap(50).sName = "Copy Inventory by Entry Date"
    tmRptSelNameMap(50).iRptCallType = COPYJOB
    tmRptSelNameMap(51).sName = "Copy Play List by ISCI"
    tmRptSelNameMap(51).iRptCallType = COPYJOB
    tmRptSelNameMap(52).sName = "Unapproved Copy"
    tmRptSelNameMap(52).iRptCallType = COPYJOB
    tmRptSelNameMap(53).sName = "Log Posting Status"
    tmRptSelNameMap(53).iRptCallType = POSTLOGSJOB
    tmRptSelNameMap(54).sName = "Missing ISCI Codes"
    tmRptSelNameMap(54).iRptCallType = POSTLOGSJOB
    tmRptSelNameMap(55).sName = "Log"
    tmRptSelNameMap(55).iRptCallType = LOGSJOB
    tmRptSelNameMap(56).sName = "Commercial Schedule"
    tmRptSelNameMap(56).iRptCallType = LOGSJOB
    tmRptSelNameMap(57).sName = "Commercial Summary"
    tmRptSelNameMap(57).iRptCallType = LOGSJOB
    tmRptSelNameMap(58).sName = "Short Form Certificate of Performance"
    tmRptSelNameMap(58).iRptCallType = LOGSJOB
    tmRptSelNameMap(59).sName = "Long Form Certificate of Performance"
    tmRptSelNameMap(59).iRptCallType = LOGSJOB
    tmRptSelNameMap(60).sName = "Short Form Log"
    tmRptSelNameMap(60).iRptCallType = LOGSJOB
    tmRptSelNameMap(61).sName = "Long Form Log"
    tmRptSelNameMap(61).iRptCallType = LOGSJOB
    tmRptSelNameMap(62).sName = "Log 4"
    tmRptSelNameMap(62).iRptCallType = LOGSJOB
    tmRptSelNameMap(63).sName = "Invoice Register"
    tmRptSelNameMap(63).iRptCallType = INVOICESJOB
    tmRptSelNameMap(64).sName = "View Invoice Export"
    tmRptSelNameMap(64).iRptCallType = INVOICESJOB
    tmRptSelNameMap(65).sName = "Billing Distribution"
    tmRptSelNameMap(65).iRptCallType = INVOICESJOB
    tmRptSelNameMap(66).sName = "Contract Import"
    tmRptSelNameMap(66).iRptCallType = CHFCONVMENU
    tmRptSelNameMap(67).sName = "Dallas Feed"
    tmRptSelNameMap(67).iRptCallType = DALLASFEED
    tmRptSelNameMap(68).sName = "Dallas Studio Log"
    tmRptSelNameMap(68).iRptCallType = DALLASFEED
    tmRptSelNameMap(69).sName = "Dallas Error Log"
    tmRptSelNameMap(69).iRptCallType = DALLASFEED
    'tmRptSelNameMap(70).sName = "New York Feed"
    'tmRptSelNameMap(70).iRptCallType = NYFEED
    'tmRptSelNameMap(71).sName = "New York Error Log"
    'tmRptSelNameMap(71).iRptCallType = NYFEED
    tmRptSelNameMap(70).sName = "Engineering Feed"
    tmRptSelNameMap(70).iRptCallType = NYFEED
    tmRptSelNameMap(71).sName = "Engineering Error Log"
    tmRptSelNameMap(71).iRptCallType = NYFEED
    '7-5-01 remove new York from blackout reports
    tmRptSelNameMap(72).sName = "Blackout Suppression"
    tmRptSelNameMap(72).iRptCallType = NYFEED
    tmRptSelNameMap(73).sName = "Blackout Replacement"
    tmRptSelNameMap(73).iRptCallType = NYFEED
    tmRptSelNameMap(74).sName = "Phoenix Studio Log"
    tmRptSelNameMap(74).iRptCallType = PHOENIXFEED
    tmRptSelNameMap(75).sName = "Phoenix Error Log"
    tmRptSelNameMap(75).iRptCallType = PHOENIXFEED
    tmRptSelNameMap(76).sName = "Commercial Change Export"
    tmRptSelNameMap(76).iRptCallType = CMMLCHG
    tmRptSelNameMap(77).sName = "Affiliate Spots Export"
    tmRptSelNameMap(77).iRptCallType = EXPORTAFFSPOTS
    tmRptSelNameMap(78).sName = "Affiliate Spots Error Log"
    tmRptSelNameMap(78).iRptCallType = EXPORTAFFSPOTS
    tmRptSelNameMap(79).sName = "Bulk Copy Feed"
    tmRptSelNameMap(79).iRptCallType = BULKCOPY
    tmRptSelNameMap(80).sName = "Bulk Copy Cross Reference"
    tmRptSelNameMap(80).iRptCallType = BULKCOPY
    tmRptSelNameMap(81).sName = "Affiliate Bulk Feed by Cart"
    tmRptSelNameMap(81).iRptCallType = BULKCOPY
    tmRptSelNameMap(82).sName = "Affiliate Bulk Feed by Vehicle"
    tmRptSelNameMap(82).iRptCallType = BULKCOPY
    tmRptSelNameMap(83).sName = "Affiliate Bulk Feed by Date"
    tmRptSelNameMap(83).iRptCallType = BULKCOPY
    tmRptSelNameMap(84).sName = "Affiliate Bulk Feed by Advertiser"
    tmRptSelNameMap(84).iRptCallType = BULKCOPY
    tmRptSelNameMap(85).sName = "Copy Play List by Vehicle"
    tmRptSelNameMap(85).iRptCallType = COPYJOB
    tmRptSelNameMap(86).sName = "Copy Play List by Advertiser"
    tmRptSelNameMap(86).iRptCallType = COPYJOB
    tmRptSelNameMap(87).sName = "Ageing by Participant"
    tmRptSelNameMap(87).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(88).sName = "Ageing by Sales Source"
    tmRptSelNameMap(88).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(89).sName = "Ageing by Producer"    '2-10-00
    tmRptSelNameMap(89).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(90).sName = "Unused"        '2-12-09 moved to splitregionlist
    tmRptSelNameMap(90).iRptCallType = COPYJOB
    tmRptSelNameMap(91).sName = "Vehicle Groups"
    tmRptSelNameMap(91).iRptCallType = VEHICLESLIST
    tmRptSelNameMap(92).sName = "Log Closing Schedules"
    tmRptSelNameMap(92).iRptCallType = VEHICLESLIST
    tmRptSelNameMap(93).sName = "On-Account Cash Applied"
    tmRptSelNameMap(93).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(94).sName = "Credit/Debit Memo"
    tmRptSelNameMap(94).iRptCallType = INVOICESJOB
    tmRptSelNameMap(95).sName = "Mailing Labels"
    tmRptSelNameMap(95).iRptCallType = AGENCIESLIST
    tmRptSelNameMap(96).sName = "Invoice Summary"
    tmRptSelNameMap(96).iRptCallType = INVOICESJOB
    tmRptSelNameMap(97).sName = "Copy Book"             '8-30-05
    tmRptSelNameMap(97).iRptCallType = COPYJOB
    tmRptSelNameMap(98).sName = "Live Log Activity"             '12-8-05
    tmRptSelNameMap(98).iRptCallType = POSTLOGSJOB
    tmRptSelNameMap(99).sName = "Tax Register"                  '1-30-07
    tmRptSelNameMap(99).iRptCallType = INVOICESJOB
    tmRptSelNameMap(100).sName = "Vehicle Participant Information"
    tmRptSelNameMap(100).iRptCallType = VEHICLESLIST
    tmRptSelNameMap(101).sName = "Installment Reconciliation"
    tmRptSelNameMap(101).iRptCallType = INVOICESJOB
    tmRptSelNameMap(102).sName = "Split Copy/Blackout Rotation"
    tmRptSelNameMap(102).iRptCallType = COPYJOB
    tmRptSelNameMap(103).sName = "User Options"
    tmRptSelNameMap(103).iRptCallType = USERLIST
    tmRptSelNameMap(104).sName = "User Activity Log"        '5-6-11
    tmRptSelNameMap(104).iRptCallType = USERLIST
    tmRptSelNameMap(105).sName = "Script Affidavits"        '4-9-12
    tmRptSelNameMap(105).iRptCallType = COPYJOB

    tmRptSelNameMap(106).sName = "Copy Inventory Producer Status"        '4-10-13
    tmRptSelNameMap(106).iRptCallType = COPYJOB
    
    tmRptSelNameMap(107).sName = "Airing Vehicle Inventory"        '3-31-15
    tmRptSelNameMap(107).iRptCallType = PROGRAMMINGJOB
    
    tmRptSelNameMap(108).sName = "Unposted Barter Stations"        '8-12-15
    tmRptSelNameMap(108).iRptCallType = INVOICESJOB
    tmRptSelNameMap(109).sName = "Ageing Summary by Month"        '8-21-15
    tmRptSelNameMap(109).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(110).sName = "Standard Package Vehicles"               '9-30-15
    tmRptSelNameMap(110).iRptCallType = VEHICLESLIST
    tmRptSelNameMap(111).sName = "Sales Commissions on Collections"     '1-24-18
    tmRptSelNameMap(111).iRptCallType = COLLECTIONSJOB
    tmRptSelNameMap(112).sName = "Station Posting Activity"             '1-23-19
    tmRptSelNameMap(112).iRptCallType = POSTLOGSJOB
    tmRptSelNameMap(113).sName = "User Summary"                         'Date: 4/15/2020 User Summary report
    tmRptSelNameMap(113).iRptCallType = USERLIST

    'RptSelCt
    tmRptSelCtNameMap(0).sName = "Sales Commissions on Billing"
    tmRptSelCtNameMap(0).iRptCallType = SLSPCOMMSJOB
    tmRptSelCtNameMap(1).sName = "Billed and Booked Commissions"
    tmRptSelCtNameMap(1).iRptCallType = SLSPCOMMSJOB
    tmRptSelCtNameMap(2).sName = "Proposals/Contracts"
    tmRptSelCtNameMap(2).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(3).sName = "Paperwork Summary"
    tmRptSelCtNameMap(3).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(4).sName = "Spots by Advertiser"
    tmRptSelCtNameMap(4).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(5).sName = "Spots by Date & Time"
    tmRptSelCtNameMap(5).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(6).sName = "Business Booked by Contract"
    tmRptSelCtNameMap(6).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(7).sName = "Contract Recap"
    tmRptSelCtNameMap(7).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(8).sName = "Spot Placements"
    tmRptSelCtNameMap(8).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(9).sName = "Spot Discrepancies"
    tmRptSelCtNameMap(9).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(10).sName = "MG's"
    tmRptSelCtNameMap(10).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(11).sName = "Sales Spot Tracking"
    tmRptSelCtNameMap(11).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(12).sName = "Commercial Changes"
    tmRptSelCtNameMap(12).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(13).sName = "Contract History"
    tmRptSelCtNameMap(13).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(14).sName = "Affiliate Spot Tracking"
    tmRptSelCtNameMap(14).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(15).sName = "Spot Sales"
    tmRptSelCtNameMap(15).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(16).sName = "Missed Spots"
    tmRptSelCtNameMap(16).iRptCallType = CONTRACTSJOB
    'tmRptSelCtNameMap(17).sName = "Business Booked by Spot"
    '1-31-00 name change of business booked by spot
    tmRptSelCtNameMap(17).sName = "Spot Business Booked"
    tmRptSelCtNameMap(17).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(18).sName = "Business Booked by Spot Reprint"
    tmRptSelCtNameMap(18).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(19).sName = "Avails"
    tmRptSelCtNameMap(19).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(20).sName = "Average Spot Prices"
    tmRptSelCtNameMap(20).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(21).sName = "Advertiser Units Ordered"
    tmRptSelCtNameMap(21).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(22).sName = "Sales Analysis by CPP & CPM"
    tmRptSelCtNameMap(22).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(23).sName = "Average 30" & """" & " Unit Rate"    ' "Average 30" & """ & " Unit Rate"
    tmRptSelCtNameMap(23).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(24).sName = "Tie-Out"
    tmRptSelCtNameMap(24).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(25).sName = "Billed and Booked"
    tmRptSelCtNameMap(25).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(26).sName = "Weekly Sales Activity by Quarter"
    tmRptSelCtNameMap(26).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(27).sName = "Sales Comparison"
    tmRptSelCtNameMap(27).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(28).sName = "Weekly Sales Activity by Month"
    tmRptSelCtNameMap(28).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(29).sName = "Average Prices to Make Plan"
    tmRptSelCtNameMap(29).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(30).sName = "CPP/CPM by Vehicle"
    tmRptSelCtNameMap(30).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(31).sName = "Sales Analysis Summary"
    tmRptSelCtNameMap(31).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(32).sName = "Insertion Orders"
    tmRptSelCtNameMap(32).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(33).sName = "Makegood Revenue"
    tmRptSelCtNameMap(33).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(34).sName = "Daily Sales Activity by Contract"   '6-5-01
    tmRptSelCtNameMap(34).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(35).sName = "Daily Sales Activity by Month"   '7-25-02
    tmRptSelCtNameMap(35).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(36).sName = "Sales Placement"   '7-5-02
    tmRptSelCtNameMap(36).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(37).sName = "Vehicle Unit Count"   '7-15-03
    tmRptSelCtNameMap(37).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(38).sName = "Billed and Booked Recap"   '4-14-05
    tmRptSelCtNameMap(38).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(39).sName = "Locked Avails"   '4-5-06
    tmRptSelCtNameMap(39).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(40).sName = "Event Summary"   '7-14-06
    tmRptSelCtNameMap(40).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(41).sName = "Accrual/Deferral"   '12-20-06
    tmRptSelCtNameMap(41).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(42).sName = "Paperwork Tax Summary"   '04-09-07
    tmRptSelCtNameMap(42).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(43).sName = "Billed and Booked Comparisons"   '09-13-07
    tmRptSelCtNameMap(43).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(44).sName = "Hi-Lo Spot Rate"         '6-9-10
    tmRptSelCtNameMap(44).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(45).sName = "Contract Verification"         '6-9-10
    tmRptSelCtNameMap(45).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(46).sName = "Insertion Order Activity Log"         '10-3-15
    tmRptSelCtNameMap(46).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(47).sName = "Proposal XML Activity Log"         '4-1-16
    tmRptSelCtNameMap(47).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(48).sName = "Spot Discrepancy Summary by Month"         '6-21-16
    tmRptSelCtNameMap(48).iRptCallType = CONTRACTSJOB
    'TTP 10674 - Spot and Digital Line combo report: new report for revenue by date for spots and digital lines
    tmRptSelCtNameMap(49).sName = "Spot and Digital Line Combo"
    tmRptSelCtNameMap(49).iRptCallType = CONTRACTSJOB
    
    'rptselpj
    tmRptSelPjNameMap(0).sName = "Salesperson Projection"
    tmRptSelPjNameMap(0).iRptCallType = PROPOSALPROJECTION
    tmRptSelPjNameMap(1).sName = "Vehicle Projection"
    tmRptSelPjNameMap(1).iRptCallType = PROPOSALPROJECTION
    tmRptSelPjNameMap(2).sName = "Sales Office Projection"
    tmRptSelPjNameMap(2).iRptCallType = PROPOSALPROJECTION
    tmRptSelPjNameMap(3).sName = "Category Projection"
    tmRptSelPjNameMap(3).iRptCallType = PROPOSALPROJECTION
    tmRptSelPjNameMap(4).sName = "Office Projection by Potential"
    tmRptSelPjNameMap(4).iRptCallType = PROPOSALPROJECTION

    'rptselRI       '8-29-02
    'tmRptSelRINameMap(0).sName = "Remote Invoice Worksheet"
    tmRptSelRINameMap(0).sName = "Rep Invoice Worksheet"        'changed 11-17-02
    tmRptSelRINameMap(0).iRptCallType = INVOICESJOB
    tmRptSelRINameMap(1).sName = "Delinquent Affidavits"
    tmRptSelRINameMap(1).iRptCallType = INVOICESJOB
    tmRptSelRINameMap(2).sName = "Unbillable Invoices"
    tmRptSelRINameMap(2).iRptCallType = INVOICESJOB
    tmRptSelRINameMap(3).sName = "Barter Payments"             '12-28-06
    tmRptSelRINameMap(3).iRptCallType = INVOICESJOB


    'rptselNT       '4-2-03
    tmRptSelNTNameMap(0).sName = "NTR Recap"
    tmRptSelNTNameMap(0).iRptCallType = CONTRACTSJOB
    tmRptSelNTNameMap(1).sName = "NTR Billed and Booked"
    tmRptSelNTNameMap(1).iRptCallType = CONTRACTSJOB
    tmRptSelNTNameMap(2).sName = "Multimedia Billed and Booked"     '1-25-08
    tmRptSelNTNameMap(2).iRptCallType = CONTRACTSJOB

    'rptselCC   1-15-04 Producer/Provider reports
    tmRptSelCCNameMap(0).sName = "Vehicle Producer"
    tmRptSelCCNameMap(0).iRptCallType = PRODUCERLIST
    tmRptSelCCNameMap(1).sName = "Content Provider"
    tmRptSelCCNameMap(1).iRptCallType = PROVIDERLIST

    'rptselFD   8-18-04 Feed Reports
    tmRptSelFDNameMap(0).sName = "Feed Recap"
    tmRptSelFDNameMap(0).iRptCallType = FEEDJOB
    tmRptSelFDNameMap(1).sName = "Feed Pledges"
    tmRptSelFDNameMap(1).iRptCallType = FEEDJOB
    tmRptSelFDNameMap(2).sName = "Pre-Feed"         '5-8-10
    tmRptSelFDNameMap(2).iRptCallType = FEEDJOB
    
    'rptselRS   1-8-06
    tmRptSelRSnameMap(0).sName = "Research"
    tmRptSelRSnameMap(0).iRptCallType = RESEARCHLIST
    tmRptSelRSnameMap(1).sName = "Special Research Summary"
    tmRptSelRSnameMap(1).iRptCallType = RESEARCHLIST
    tmRptSelRSnameMap(2).sName = "Demo Rank"
    tmRptSelRSnameMap(2).iRptCallType = RESEARCHLIST
    
    'rptselSR 9-19-07
    tmRptSelSRNameMap(0).sName = "Copy Regions"     '2-12-09 changed from Split Regions
    tmRptSelSRNameMap(0).iRptCallType = SPLITREGIONLIST


    '12-18-07
    tmRptSelCANameMap(0).sName = "Event & Sports Avails"
    tmRptSelCANameMap(0).iRptCallType = CONTRACTSJOB
    tmRptSelCANameMap(1).sName = "Avails Combo by Day/Week"
    tmRptSelCANameMap(1).iRptCallType = CONTRACTSJOB

    '04-10-08
    tmRptSelSNNameMap(0).sName = "Split Network Avails"
    tmRptSelSNNameMap(0).iRptCallType = CONTRACTSJOB

    '05-06-08
    tmRptSelADNameMAP(0).sName = "Audience Delivery"
    tmRptSelADNameMAP(0).iRptCallType = CONTRACTSJOB
    tmRptSelADNameMAP(1).sName = "Post Buy Analysis"
    tmRptSelADNameMAP(1).iRptCallType = CONTRACTSJOB
    
    tmRptSelSpotBBNameMAP(0).sName = "Revenue on the Books"
    tmRptSelSpotBBNameMAP(0).iRptCallType = CONTRACTSJOB
    tmRptSelSpotBBNameMAP(1).sName = "Spot Revenue Register"
    tmRptSelSpotBBNameMAP(1).iRptCallType = INVOICESJOB
    
    tmRptSelAvgCompareNameMAP(0).sName = "Average Rate Comparison"
    tmRptSelAvgCompareNameMAP(0).iRptCallType = CONTRACTSJOB
    tmRptSelAvgCompareNameMAP(1).sName = "Average Spot Price Comparison"
    tmRptSelAvgCompareNameMAP(1).iRptCallType = CONTRACTSJOB
    
'    '12/28/2020 - Podcast Billing Discrepancy
'    tmRptSelPodcastBillingMap(0).sName = "Podcast Billing Discrepancy"
'    tmRptSelPodcastBillingMap(0).iRptCallType = PODBILLJOB
    
    Exit Sub
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
'    Erase tmReportList
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
    Unload ReportList
'    Set ReportList = Nothing   'Remove data segment
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

Private Sub rbcShowBy_Click(Index As Integer)
    If imIgnoreChg Then
        Exit Sub
    End If
    If rbcShowBy(Index).Value Then
        If Index = 2 Then
            'Quick Find
            igShowCatOrNames = 0
            mPopulate (1)
            lbcRpt.Visible = False
            CSI_ComboBoxMS1.SetDropDownWidth (CSI_ComboBoxMS1.Width)
            CSI_ComboBoxMS1.mShowDropDown
            CSI_ComboBoxMS1.Visible = True
            CSI_ComboBoxMS1.SetFocus
        Else
            CSI_ComboBoxMS1.Visible = False
            lbcRpt.Visible = True
            igShowCatOrNames = Index
            mPopulate
            lbcRpt.SetFocus
        End If
    End If
End Sub

Private Sub tmcClock_Timer()
    mCheckGG
End Sub

Private Sub tmcDelay_Timer()
    tmcDelay.Enabled = False
    mTerminate 'False
End Sub

Private Sub tmcSetCntrls_Timer()
    Dim slName As String
    Dim slStr As String
    
    tmcSetCntrls.Enabled = False
    slName = igReportRnfCode
    Do While Len(slName) < 5
        slName = "0" & slName
    Loop
    slStr = tgUrf(0).iCode
    Do While Len(slStr) < 5
        slStr = "0" & slStr
    Loop
    slName = slName & slStr
    sgReportCtrlSaveName = Left$(sgReportFormExe, 10) & slName
    '5/24/18: Reset the Form controls
    If igReportButtonIndex <> 2 Then
        gSetReportCtrlsSetting
    End If
End Sub

Private Sub vbcRptSample_Change()
    pbcRptSample(1).Top = -vbcRptSample.Value
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub

'
'       find the report name in the rptsel table maps
'
'       <input> report name mapping table
'               report name
'       <output> Report call type
'       return - true if found
'
Private Function mFindNameInMap(tlRptnameMap() As RPTNAMEMAP, slName As String, slRptCallType As String) As Integer
    Dim ilFound As Integer
    Dim illoop As Integer
    Dim slTestName As String
    Dim ilRet As Integer
        ilFound = False
        For illoop = 0 To UBound(tlRptnameMap) Step 1
            slTestName = Trim$(tlRptnameMap(illoop).sName)
            If Len(slTestName) = 0 Then
                Exit For
            End If
            If StrComp(slName, slTestName, 1) = 0 Then
                ilFound = True
                slRptCallType = Trim$(str$(tlRptnameMap(illoop).iRptCallType))
                Exit For
            End If
        Next illoop
        If Not ilFound Then
            ilRet = MsgBox("Report Name " & slName & " not found in Mapping Table", vbOKOnly + vbInformation, "Report List")
            cmcCancel.SetFocus
            mFindNameInMap = ilFound
            Exit Function
        End If
        mFindNameInMap = ilFound
        Exit Function
End Function

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

Private Sub mCheckGG()
    Dim ilRet As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim llNow As Long
    Dim llDate As Long
    Dim slStr As String
    Dim illoop As Integer
    
    On Error Resume Next
    
    If imLastHourGGChecked = Hour(Now) Then
        Exit Sub
    End If
    imLastHourGGChecked = Hour(Now)
    
    If (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) > 0) Or (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) > 0) Then
        Exit Sub
    End If
    If bgInternalGuide Then
        Exit Sub
    End If
        
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
    gSetRptGGFlag tmSaf
    btrDestroy hmSaf
End Sub

Private Function mCheckReportName(ilAllowed As Integer, ilIndex As Integer) As Integer
    mCheckReportName = ilAllowed
    If ilAllowed And (ilIndex >= 0) Then
        mCheckReportName = True
        If ((Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName)) Then
            'If Trim$(tgRnfList(ilIndex).tRnf.sName) = "ReRate" Then
            '    mCheckReportName = False
            'End If
        End If
    End If
End Function
