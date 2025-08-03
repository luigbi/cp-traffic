VERSION 5.00
Begin VB.Form RptList 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6000
   ClientLeft      =   570
   ClientTop       =   2475
   ClientWidth     =   9210
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6000
   ScaleWidth      =   9210
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   7155
      Top             =   5625
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4875
      TabIndex        =   9
      Top             =   5595
      Width           =   1050
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
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
      Left            =   8055
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5685
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
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
      Left            =   8310
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5685
      Visible         =   0   'False
      Width           =   525
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
      Left            =   8610
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5685
      Visible         =   0   'False
      Width           =   525
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
      TabIndex        =   11
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
      Left            =   3315
      TabIndex        =   8
      Top             =   5595
      Width           =   1035
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
         Visible         =   0   'False
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2550
         Width           =   255
      End
      Begin VB.PictureBox pbcRptSample 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
            BackColor       =   &H80000005&
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
            TabIndex        =   6
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
Attribute VB_Name = "RptList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptlist.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptList.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
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
Dim tmRptList() As RPTLST
Dim smPassCommands As String
Dim smScreenCaption As String
Dim tmRptNoSelNameMap(0 To 50) As RPTNAMEMAP
Dim tmRptSelNameMap(0 To 105) As RPTNAMEMAP
Dim tmRptSelCtNameMap(0 To 45) As RPTNAMEMAP    '4-5-06 increase from 40 to 45
Dim tmRptSelPjNameMap(0 To 4) As RPTNAMEMAP
Dim tmRptSelRINameMap(0 To 3) As RPTNAMEMAP '8-29-02
Dim tmRptSelNTNameMap(0 To 2) As RPTNAMEMAP
Dim tmRptSelCCNameMap(0 To 4) As RPTNAMEMAP '1-15-04
Dim tmRptSelFDNameMap(0 To 4) As RPTNAMEMAP '8-18-04 Station Pledge report
Dim tmRptSelRSnameMap(0 To 2) As RPTNAMEMAP '2-8-06 Research reports
Dim tmRptSelSRNameMap(0 To 2) As RPTNAMEMAP     '9-19-06  Split region list
Dim tmRptSelCANameMap(0 To 2) As RPTNAMEMAP     '12-18-07 Combo Avails versions
Dim tmRptSelSNNameMap(0 To 2) As RPTNAMEMAP     '04-10-08 Split Network Avails
Dim tmRptSelADNameMAP(0 To 2) As RPTNAMEMAP      '5-06-08 Post buy analysis, combine with Aud Delivery

Dim imIgnoreChg As Integer


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

    slCommand = sgCommandStr    'Command$
    smPassCommands = slCommand
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide" '"rn48616" '"Guide"
    '    ilTestSystem = False
    '    imShowHelpMsg = False
    '    If ilTestSystem Then
    '        slCommand = "RptList^TEST^NoHelp\ADVERTISERSLIST\1"
    '    Else
    '        slCommand = "RptList^Prod^NoHelp\ADVERTISERSLIST\1"
    '    End If
    '    smPassCommands = slCommand
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            'End
            imTerminate = True
            Exit Sub
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
    '    imShowHelpMsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpMsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone RptList, slStr, ilTestSystem
    'ilRet = gParseItem(slCommand, 3, "\", slStr)
    igRptCallType = Val(slStr)
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

    If lbcRpt.ListIndex < 0 Then
        lbcRpt.SetFocus
        Exit Sub
    End If
    If tmRptList(lbcRpt.ListIndex).tRnf.sType = "C" Then
        lbcRpt.SetFocus
        Exit Sub
    End If
    slExe = Trim$(tmRptList(lbcRpt.ListIndex).tRnf.sRptExe)
    ilPos = InStr(1, slExe, ".", vbTextCompare)
    If ilPos > 0 Then
        slExe = Left$(slExe, ilPos - 1)
    End If
    slName = Trim$(tmRptList(lbcRpt.ListIndex).tRnf.sName)
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
    '        slStr = slStr & "RptList^Test\" & slCallRptType & "\" & slName
    '    Else
    '        slStr = slStr & "RptList^Prod\" & slCallRptType & "\" & slName
    '    End If
    'Else
    '    If igTestSystem Then
    '        slStr = slStr & "RptList^Test^NOHELP\" & slCallRptType & "\" & slName
    '    Else
    '        slStr = slStr & "RptList^Prod^NOHELP\" & slCallRptType & "\" & slName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & slExe & " " & slStr, 1)
    sgCommandStr = slStr
    DoEvents                '2-27-03
    On Error Resume Next
    If StrComp(slExe, "RptNoSel", 1) = 0 Then
        RptNoSel.Show vbModal
    ElseIf StrComp(slExe, "RptSel", 1) = 0 Then
        RptSel.Show vbModal
    ElseIf StrComp(slExe, "RptSelaa", 1) = 0 Then
        RptSelAA.Show vbModal
    ElseIf StrComp(slExe, "RptSelac", 1) = 0 Then
        RptSelAc.Show vbModal
    ElseIf StrComp(slExe, "RptSelad", 1) = 0 Then
        RptSelAD.Show vbModal
    ElseIf StrComp(slExe, "RptSelal", 1) = 0 Then       '4-14-04
        RptSelAL.Show vbModal
    ElseIf StrComp(slExe, "RptSelap", 1) = 0 Then
        RptSelAp.Show vbModal
    ElseIf StrComp(slExe, "RptSelas", 1) = 0 Then
        RptSelAS.Show vbModal
    ElseIf StrComp(slExe, "RptSelav", 1) = 0 Then
        RptSelAv.Show vbModal
    'ElseIf StrComp(slExe, "RptSelbr", 1) = 0 Then
        'RptSelbr.Show vbModal
    ElseIf StrComp(slExe, "RptSelcb", 1) = 0 Then
        RptSelCb.Show vbModal
    ElseIf StrComp(slExe, "RptSelcc", 1) = 0 Then       '1-15-04 Producer/Provider reports
        RptSelCC.Show vbModal
    ElseIf StrComp(slExe, "RptSelcp", 1) = 0 Then
        RptSelCp.Show vbModal
    ElseIf StrComp(slExe, "RptSelCt", 1) = 0 Then
        RptSelCt.Show vbModal
    ElseIf StrComp(slExe, "RptSeldb", 1) = 0 Then
        RptSelDB.Show vbModal
    ElseIf StrComp(slExe, "RptSeldf", 1) = 0 Then
        RptSelDF.Show vbModal
    ElseIf StrComp(slExe, "RptSelds", 1) = 0 Then
        RptSelDS.Show vbModal
    'ElseIf StrComp(slExe, "RptSelEd", 1) = 0 Then
    '    RptSelED.Show vbModal
    ElseIf StrComp(slExe, "RptSelFd", 1) = 0 Then       '8-4-04 Feed Report
        rptSelFD.Show vbModal
    ElseIf StrComp(slExe, "RptSelia", 1) = 0 Then
        RptSelIA.Show vbModal
    ElseIf StrComp(slExe, "RptSelid", 1) = 0 Then       '5-21-02
        RptSelID.Show vbModal
    ElseIf StrComp(slExe, "RptSelin", 1) = 0 Then
        'RptSelIn.Show vbModal
    ElseIf StrComp(slExe, "RptSelir", 1) = 0 Then       '7-13-05
        RptSelIR.Show vbModal
    ElseIf StrComp(slExe, "RptSeliv", 1) = 0 Then
        RptSelIv.Show vbModal
    ElseIf StrComp(slExe, "RptSellg", 1) = 0 Then
        'RptSellg.Show vbModal
    ElseIf StrComp(slExe, "RptSelNT", 1) = 0 Then       '4-2-03
        RptSelNT.Show vbModal
    ElseIf StrComp(slExe, "RptSelOF", 1) = 0 Then       '7-21-06
        RptSelOF.Show vbModal
    ElseIf StrComp(slExe, "RptSelos", 1) = 0 Then
        RptSelOS.Show vbModal
    ElseIf StrComp(slExe, "RptSelpa", 1) = 0 Then
        RptSelPA.Show vbModal
    ElseIf StrComp(slExe, "RptSelpc", 1) = 0 Then
        RptSelPC.Show vbModal
    ElseIf StrComp(slExe, "RptSelpj", 1) = 0 Then
        RptSelPJ.Show vbModal
    ElseIf StrComp(slExe, "RptSelpp", 1) = 0 Then
        RptSelPP.Show vbModal
    ElseIf StrComp(slExe, "RptSelpr", 1) = 0 Then       '6-15-04  Proposal Research Recap
        RptSelPr.Show vbModal
    ElseIf StrComp(slExe, "RptSelps", 1) = 0 Then
        RptSelPS.Show vbModal
    ElseIf StrComp(slExe, "RptSelqb", 1) = 0 Then
        RptSelQB.Show vbModal
    ElseIf StrComp(slExe, "RptSelra", 1) = 0 Then
        RptSelRA.Show vbModal
    ElseIf StrComp(slExe, "RptSelRD", 1) = 0 Then           '5-13-03
        RptSelRD.Show vbModal
    ElseIf StrComp(slExe, "RptSelRG", 1) = 0 Then           '12-22-09  Regional copy assignment
        RptSelRg.Show vbModal
    ElseIf StrComp(slExe, "RptSelRI", 1) = 0 Then
        RptSelRI.Show vbModal
    ElseIf StrComp(slExe, "RptSelRP", 1) = 0 Then           '11-1-02 Remote Posting
        RptSelRP.Show vbModal
    ElseIf StrComp(slExe, "RptSelrs", 1) = 0 Then
        RptSelRS.Show vbModal
    ElseIf StrComp(slExe, "RptSelrr", 1) = 0 Then           '6-20-03 Research Revenue
        RptSelRR.Show vbModal
    ElseIf StrComp(slExe, "RptSelrv", 1) = 0 Then
        RptSelRV.Show vbModal
    ElseIf StrComp(slExe, "RptSelca", 1) = 0 Then      '12/18/07 combo avails
        RptSelCA.Show vbModal
    ElseIf StrComp(slExe, "RptSelSN", 1) = 0 Then      '04-10-08 Split Network Avails
        RptSelSN.Show vbModal
    ElseIf StrComp(slExe, "RptSelsp", 1) = 0 Then
        RptSelSP.Show vbModal
    ElseIf StrComp(slExe, "RptSelSR", 1) = 0 Then           '9-19-06 split regions
        RptSelSR.Show vbModal
    ElseIf StrComp(slExe, "RptSelss", 1) = 0 Then
        'RptSelss.Show vbModal
    ElseIf StrComp(slExe, "RptSeltx", 1) = 0 Then
        RptSelTx.Show vbModal
    ElseIf StrComp(slExe, "RptSelus", 1) = 0 Then
        RptSelUS.Show vbModal
    Else
        MsgBox "Missing test for " & slExe
        Exit Sub
    End If
    gChDrDir        '3-25-03
    'ChDrive Left$(sgCurDir, 2)  'Set the default drive
    'ChDir sgCurDir              'set the default path
    'If Not igStdAloneMode Then
    '    'tmcDelay.Enabled = True
    '    mTerminate False
    'Else
    '    RptList.Enabled = False
    '    Do While Not igChildDone
    '        DoEvents
    '    Loop
    '    slStr = sgDoneMsg
    '    RptList.Enabled = True
    '    edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    '    For ilLoop = 0 To 10
    '        DoEvents
    '    Next ilLoop
    'End If
    If Not igRptReturn Then
        mTerminate
    End If
    'Screen.MousePointer = vbDefault    'Default
End Sub
Private Sub cmcDone_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
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
    RptList.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        mTerminate 'True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Reset used instead of Close to cause # Clients on network to be decrement
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    'btrStopAppl
    'End
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
        For ilRpt = 0 To UBound(tmRptList) - 1 Step 1
            If slName = Trim$(tmRptList(ilRpt).tRnf.sName) Then
                'edcRptDescription.Text = Left$(tmRptList(ilRpt).tRnf.sDescription, tmRptList(ilRpt).tRnf.iStrLen)
                edcRptDescription.Text = gStripChr0(tmRptList(ilRpt).tRnf.sDescription)
                slFromFile = tmRptList(ilRpt).tRnf.sRptSample
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
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        Next ilRpt
    End If
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
    cmcDone_Click
End Sub

Private Sub lbcRpt_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
End Sub
Private Sub lbcRpt_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim slName As String
    Dim ilBlanks As Integer
    Dim ilLoop As Integer
    Dim ilLen As Integer
    If (Shift And vbAltMask) = ALTMASK Then
        If KeyCode = KEYDOWN Then
            'Find next same level
            If lbcRpt.ListIndex >= 0 Then
                slName = lbcRpt.List(lbcRpt.ListIndex)
                ilBlanks = Len(slName) - Len(LTrim$(slName))
                For ilLoop = lbcRpt.ListIndex + 1 To lbcRpt.ListCount - 1 Step 1
                    slName = lbcRpt.List(ilLoop)
                    ilLen = Len(slName) - Len(LTrim$(slName))
                    If ilBlanks = ilLen Then
                        lbcRpt.TopIndex = ilLoop
                        lbcRpt.ListIndex = ilLoop
                        Exit Sub
                    ElseIf ilLen < ilBlanks Then
                        lbcRpt.TopIndex = lbcRpt.ListIndex + 1
                        lbcRpt.ListIndex = lbcRpt.ListIndex + 1
                        Exit Sub
                    End If
                Next ilLoop
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

    Screen.MousePointer = vbHourglass
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    RptList.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone RptList
    'RptList.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    smScreenCaption = "Report Selection"
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imFirstFocus = True
    hmRnf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmRnf, "", sgDBPath & "Rnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rnf.Btr)", RptList
    On Error GoTo 0
    imRnfRecLen = Len(tmRnf)
    hmSnf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmSnf, "", sgDBPath & "Snf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Snf.Btr)", RptList
    On Error GoTo 0
    imSnfRecLen = Len(tmSnf)
    hmSrf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmSrf, "", sgDBPath & "Srf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Srf.Btr)", RptList
    On Error GoTo 0
    imSrfRecLen = Len(tmSrf)
    hmRtf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmRtf, "", sgDBPath & "Rtf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rtf.Btr)", RptList
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
    lbcRpt.Visible = True
    edcRptDescription.Visible = True
    If imTerminate Then
        Exit Sub
    End If
    plcScreen_Paint
'    gCenterModalForm RptList
    Screen.MousePointer = vbDefault
    Exit Sub
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
    Dim ilOffset As Integer
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
    ilOffset = 4    'gFieldOffset("Prf", "PrfAdfCode")
    ilRet = btrExtAddLogicConst(hlSrf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
    Call btrExtSetBounds(hlSrf, llNoRec, -1, "UC", "SRF", "") 'Set extract limits (all records)
    ilOffset = 0
    ilRet = btrExtAddField(hlSrf, ilOffset, ilExtLen)  'Extract First Name field
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
Private Sub mPopulate()
    Dim ilRet As Integer 'btrieve status
    Dim slName As String
    Dim llLen As Long
    Dim ilLoop As Integer
    Dim ilLevel As Integer
    Dim ilSrf As Integer
    Dim ilRnf As Integer
    Dim ilIndex As Integer
    Dim ilAllowed As Integer
    Dim ilRpt As Integer
'    slGetStamp = gGetCSIStamp("RPTLIST")
'    If StrComp(slGetStamp, "RptList", 1) = 0 Then
'        ilRet = csiGetAlloc("RPTLIST", ilStartIndex, ilEndIndex)
'    Else
'        ilRet = 1
'    End If
    'If (StrComp(slGetStamp, "RptList", 1) = 0) And (ilRet = 0) Then
    '    ReDim tmRptList(ilStartIndex To ilEndIndex) As RPTLST
    '    For ilLoop = LBound(tmRptList) To UBound(tmRptList) Step 1
    '        ilRet = csiGetRec("RPTLIST", ilLoop, VarPtr(tmRptList(ilLoop)), LenB(tmRptList(ilLoop)))
    '    Next ilLoop
    '    lbcRpt.Clear
    'Else
        'gObtainRNF hmRnf
        gObtainRTF hmRtf, False
        If (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
            ilRet = mReadRec()
            If ilRet = False Then
                lbcRpt.Clear
                ReDim tmRptList(0 To 0) As RPTLST
                MsgBox "No Report Selection Allowed", vbOkOnly + vbInformation, "Report List"
                Exit Sub
            End If
        End If
        lbcRpt.Clear
        ReDim tmRptList(0 To 0) As RPTLST
        llLen = 0
        If (UBound(tgRtfList) <= LBound(tgRtfList)) Or ((Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName)) Then
            gObtainRNF hmRnf
            'Dan M 4/09/09  Limit guide to one report.
            mLimitGuideReports tgRnfList
        Else
            ReDim tgRnfList(0 To 1) As RNFLIST
        End If
        If UBound(tgRtfList) > LBound(tgRtfList) Then
            For ilLoop = 0 To UBound(tgRtfList) - 1 Step 1
                If (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                    If tgRtfList(ilLoop).tRtf.sRnfType <> "C" Then
                        ilAllowed = False
                        For ilSrf = 0 To UBound(tmSelSrf) - 1 Step 1
                            If tgRtfList(ilLoop).tRtf.iRnfCode = tmSelSrf(ilSrf).iRnfCode Then
                                If tgRtfList(ilLoop).tRtf.sRnfState <> "D" Then
                                    ilAllowed = True
                                    If (UBound(tgRtfList) <= LBound(tgRtfList)) Or ((Trim$(tgUrf(0).sName) = sgCPName) Or (Trim$(tgUrf(0).sName) = sgSUName)) Then
                                    Else
                                        imRnfRecLen = Len(tmRnf)  'Get and save CmF record length (the read will change the length)
                                        tmRnfSrchKey.iCode = tgRtfList(ilLoop).tRtf.iRnfCode
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
                            tmRnfSrchKey.iCode = tgRtfList(ilLoop).tRtf.iRnfCode
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
                    If tgRnfList(ilRnf).tRnf.iCode = tgRtfList(ilLoop).tRtf.iRnfCode Then
                        ilIndex = ilRnf
                        If tgRnfList(ilRnf).tRnf.sType = "C" Then
                            ilAllowed = True
                        End If
                        Exit For
                    End If
                Next ilRnf
                If (ilAllowed) And (ilIndex >= 0) Then
                    slName = Trim$(tgRnfList(ilIndex).tRnf.sName)
                    For ilLevel = 1 To tgRtfList(ilLoop).tRtf.iLevel - 1 Step 1
                        slName = "  " & slName
                    Next ilLevel
                    tmRptList(UBound(tmRptList)).sName = slName
                    tmRptList(UBound(tmRptList)).tRnf = tgRnfList(ilIndex).tRnf
                    tmRptList(UBound(tmRptList)).iLevel = tgRtfList(ilLoop).tRtf.iLevel
                    ReDim Preserve tmRptList(0 To UBound(tmRptList) + 1) As RPTLST
                End If
            Next ilLoop
        Else
            For ilRnf = 0 To UBound(tgRnfList) - 1 Step 1
                If (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
                    ilAllowed = False
                    For ilSrf = 0 To UBound(tmSelSrf) - 1 Step 1
                        If tgRnfList(ilRnf).tRnf.iCode = tmSelSrf(ilSrf).iRnfCode Then
                            If tgRtfList(ilLoop).tRtf.sRnfState <> "D" Then
                                ilAllowed = True
                            End If
                            Exit For
                        End If
                    Next ilSrf
                Else
                    ilAllowed = True
                End If
                If ilAllowed Then
                    slName = Trim$(tgRnfList(ilRnf).tRnf.sName)
                    tmRptList(UBound(tmRptList)).sName = slName
                    tmRptList(UBound(tmRptList)).tRnf = tgRnfList(ilRnf).tRnf
                    tmRptList(UBound(tmRptList)).iLevel = 0
                    ReDim Preserve tmRptList(0 To UBound(tmRptList) + 1) As RPTLST
                End If
            Next ilRnf
        End If
        'Remove Proposal reports if Not using Proposal system
        If (Trim$(tgUrf(0).sName) <> sgCPName) And (tgSpf.sGUsePropSys <> "Y") Then
            ilLoop = LBound(tmRptList)
            Do
                If tmRptList(ilLoop).tRnf.sType = "C" Then
                    If (tmRptList(ilLoop).iLevel <= 1) And (StrComp("Proposals", Trim$(tmRptList(ilLoop).tRnf.sName), 1) = 0) Then
                        ilIndex = ilLoop + 1
                        Do
                            If tmRptList(ilIndex).iLevel = tmRptList(ilLoop).iLevel Then
                                Exit Do
                            End If
                            ilIndex = ilIndex + 1
                        Loop While ilIndex < UBound(tmRptList)
                        For ilRpt = ilLoop To UBound(tmRptList) - (ilIndex - ilLoop) Step 1
                            tmRptList(ilRpt) = tmRptList(ilRpt + (ilIndex - ilLoop))
                        Next ilRpt
                        ReDim Preserve tmRptList(0 To UBound(tmRptList) - (ilIndex - ilLoop)) As RPTLST
                        Exit Do
                    Else
                        ilLoop = ilLoop + 1
                    End If
                Else
                    ilLoop = ilLoop + 1
                End If
            Loop While ilLoop <= UBound(tmRptList) - 1
        End If
        'Remove all level except Proposals and List if Using traffic is No
        If (Trim$(tgUrf(0).sName) <> sgCPName) And (tgSpf.sUsingTraffic = "N") Then
            ilLoop = LBound(tmRptList)
            Do
                If tmRptList(ilLoop).tRnf.sType = "C" Then
                    If (tmRptList(ilLoop).iLevel <= 1) And ((StrComp("Proposals", Trim$(tmRptList(ilLoop).tRnf.sName), 1) <> 0) And (StrComp("Lists", Trim$(tmRptList(ilLoop).tRnf.sName), 1) <> 0)) Then
                        ilIndex = ilLoop + 1
                        Do
                            If tmRptList(ilIndex).iLevel = tmRptList(ilLoop).iLevel Then
                                Exit Do
                            End If
                            ilIndex = ilIndex + 1
                        Loop While ilIndex < UBound(tmRptList)
                        For ilRpt = ilLoop To UBound(tmRptList) - (ilIndex - ilLoop) Step 1
                            tmRptList(ilRpt) = tmRptList(ilRpt + (ilIndex - ilLoop))
                        Next ilRpt
                        ReDim Preserve tmRptList(0 To UBound(tmRptList) - (ilIndex - ilLoop)) As RPTLST
                    Else
                        ilLoop = ilLoop + 1
                    End If
                Else
                    ilLoop = ilLoop + 1
                End If
            Loop While ilLoop <= UBound(tmRptList) - 1
        End If
        'Remove unused levels b
        ilLoop = LBound(tmRptList)
        Do
            If tmRptList(ilLoop).tRnf.sType = "C" Then
                If ilLoop + 1 >= UBound(tmRptList) Then
                    ReDim Preserve tmRptList(0 To UBound(tmRptList) - 1) As RPTLST
                    ilLoop = ilLoop - 1
                Else
                    For ilIndex = ilLoop + 1 To UBound(tmRptList) - 1 Step 1
                        If tmRptList(ilIndex).tRnf.sType <> "C" Then
                            'ilLoop = ilIndex + 1
                            Exit For
                        Else
                            If tmRptList(ilIndex).iLevel <= tmRptList(ilLoop).iLevel Then
                                'Remove leveles from ilLoop to ilIndex
                                For ilRpt = ilLoop To UBound(tmRptList) - (ilIndex - ilLoop) Step 1
                                    tmRptList(ilRpt) = tmRptList(ilRpt + (ilIndex - ilLoop))
                                Next ilRpt
                                ReDim Preserve tmRptList(0 To UBound(tmRptList) - (ilIndex - ilLoop)) As RPTLST
                                ilLoop = ilLoop - 1
                                Exit For
                            End If
                        End If
                    Next ilIndex
                    ilLoop = ilLoop + 1
                End If
            Else
                ilLoop = ilLoop + 1
            End If
        Loop While ilLoop <= UBound(tmRptList) - 1
        'If Salesperson or Remote User remove categories and duplicates
        If (tgUrf(0).iSlfCode > 0) Or (tgUrf(0).iRemoteUserID > 0) Or (rbcShowBy(1).Value) Then
            ilLoop = LBound(tmRptList)
            Do
                If tmRptList(ilLoop).tRnf.sType = "C" Then
                    If ilLoop + 1 >= UBound(tmRptList) Then
                        ReDim Preserve tmRptList(0 To UBound(tmRptList) - 1) As RPTLST
                    Else
                        For ilIndex = ilLoop + 1 To UBound(tmRptList) - 1 Step 1
                            tmRptList(ilIndex - 1) = tmRptList(ilIndex)
                        Next ilIndex
                        ReDim Preserve tmRptList(0 To UBound(tmRptList) - 1) As RPTLST
                    End If
                Else
                    ilLoop = ilLoop + 1
                End If
            Loop While ilLoop <= UBound(tmRptList) - 1
            ilLoop = LBound(tmRptList)
            '3/1/10:  Handle case where only one name exist
            Do
                ilIndex = ilLoop + 1
                '3/1/10:  Handle case where only one name exist
                'Do
                'Do While ilLoop < UBound(tmRptList) - 1
                Do While ilIndex < UBound(tmRptList) - 1
                    If StrComp(Trim$(tmRptList(ilLoop).sName), Trim$(tmRptList(ilIndex).sName), 1) = 0 Then
                        For ilRpt = ilIndex + 1 To UBound(tmRptList) - 1 Step 1
                            tmRptList(ilRpt - 1) = tmRptList(ilRpt)
                        Next ilRpt
                        ReDim Preserve tmRptList(0 To UBound(tmRptList) - 1) As RPTLST
                    Else
                        ilIndex = ilIndex + 1
                    End If
                'Loop While ilIndex <= UBound(tmRptList) - 1
                Loop
                ilLoop = ilLoop + 1
            Loop While ilLoop < UBound(tmRptList) - 1   '= UBound(tmRptList) - 1
            For ilLoop = 0 To UBound(tmRptList) - 1 Step 1
                tmRptList(ilLoop).sName = Trim$(tmRptList(ilLoop).sName)
            Next ilLoop
            If UBound(tmRptList) - 1 > 0 Then
                ArraySortTyp fnAV(tmRptList(), 0), UBound(tmRptList), 0, LenB(tmRptList(0)), 0, LenB(tmRptList(0).sName), 0
            End If
        End If
    '    ilRet = csiSetStamp("RPTLIST", "RptList")
    '    ilRet = csiSetAlloc("RPTLIST", LBound(tmRptList), UBound(tmRptList))
    '    For ilLoop = LBound(tmRptList) To UBound(tmRptList) Step 1
    '        ilRet = csiSetRec("RPTLIST", ilLoop, VarPtr(tmRptList(ilLoop)), LenB(tmRptList(ilLoop)))
    '    Next ilLoop
    'End If
    For ilLoop = 0 To UBound(tmRptList) - 1 Step 1
        slName = RTrim$(tmRptList(ilLoop).sName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcRpt.AddItem slName  'Add ID to list box
    Next ilLoop
    If (igRptCallType >= 20) And (igRptCallType <= 89) Then  'List Function
        'List Function
        For ilLoop = 0 To UBound(tmRptList) - 1 Step 1
            If (tmRptList(ilLoop).tRnf.sType = "C") And (tmRptList(ilLoop).tRnf.iJobListNo = -1) Then
                lbcRpt.TopIndex = ilLoop
                lbcRpt.ListIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    ElseIf igRptCallType < 20 Then  'Job function
        For ilLoop = 0 To UBound(tmRptList) - 1 Step 1
            If (tmRptList(ilLoop).tRnf.sType = "C") And (tmRptList(ilLoop).tRnf.iJobListNo = igRptCallType) Then
                lbcRpt.TopIndex = ilLoop
                lbcRpt.ListIndex = ilLoop
                Exit For
            End If
        Next ilLoop
    End If
    If lbcRpt.ListIndex < 0 Then
        If lbcRpt.ListCount > 0 Then
            lbcRpt.ListIndex = 0
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
    'gCPErrorMsg ilRet, "mReadRec (gParseItem field 2)", RptList
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
    tmRptNoSelNameMap(12).sName = "Feed Types"
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
    tmRptSelNameMap(3).sName = "Advertiser Summary"
    tmRptSelNameMap(3).iRptCallType = ADVERTISERSLIST
    tmRptSelNameMap(4).sName = "Advertiser Detail"
    tmRptSelNameMap(4).iRptCallType = ADVERTISERSLIST
    tmRptSelNameMap(5).sName = "Agency Summary"
    tmRptSelNameMap(5).iRptCallType = AGENCIESLIST
    tmRptSelNameMap(6).sName = "Agency Detail"
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
    tmRptSelNameMap(70).sName = "New York Feed"
    tmRptSelNameMap(70).iRptCallType = NYFEED
    tmRptSelNameMap(71).sName = "New York Error Log"
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

    'RptSelCt
    tmRptSelCtNameMap(0).sName = "Sales Commissions"
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
    tmRptSelCtNameMap(23).sName = "Average Rate"
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
    tmRptSelCtNameMap(40).sName = "Game Summary"   '7-14-06
    tmRptSelCtNameMap(40).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(41).sName = "Accrual/Deferral"   '12-20-06
    tmRptSelCtNameMap(41).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(42).sName = "Paperwork Tax Summary"   '04-09-07
    tmRptSelCtNameMap(42).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(43).sName = "Billed and Booked Comparisons"   '09-13-07
    tmRptSelCtNameMap(43).iRptCallType = CONTRACTSJOB
    tmRptSelCtNameMap(44).sName = "Hi-Lo Spot Rate"   '6-1-10
    tmRptSelCtNameMap(44).iRptCallType = CONTRACTSJOB

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
    tmRptSelFDNameMap(1).sName = "Pre-Feed"         '5-8-10
    tmRptSelFDNameMap(1).iRptCallType = FEEDJOB
    
    'rptselRS   1-8-06
    tmRptSelRSnameMap(0).sName = "Research"
    tmRptSelRSnameMap(0).iRptCallType = RESEARCHLIST
    tmRptSelRSnameMap(1).sName = "Special Research Summary"
    tmRptSelRSnameMap(1).iRptCallType = RESEARCHLIST

    'rptselSR 9-19-07
    tmRptSelSRNameMap(0).sName = "Copy Regions"     '2-12-09 changed from Split Regions
    tmRptSelSRNameMap(0).iRptCallType = SPLITREGIONLIST


    '12-18-07
    tmRptSelCANameMap(0).sName = "Game Avails"
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
    Dim ilRet As Integer
    Erase tmSelSrf
    Erase tmRptList
    Erase tgRtfList
    Erase tgRnfList
    gObtainRNF hmRnf
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
    Unload RptList
    Set RptList = Nothing   'Remove data segment
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
        igShowCatOrNames = Index
        mPopulate
        lbcRpt.SetFocus
    End If
End Sub

Private Sub tmcDelay_Timer()
    tmcDelay.Enabled = False
    mTerminate 'False
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
    Dim ilLoop As Integer
    Dim slTestName As String
    Dim ilRet As Integer
        ilFound = False
        For ilLoop = 0 To UBound(tlRptnameMap) Step 1
            slTestName = Trim$(tlRptnameMap(ilLoop).sName)
            If Len(slTestName) = 0 Then
                Exit For
            End If
            If StrComp(slName, slTestName, 1) = 0 Then
                ilFound = True
                slRptCallType = Trim$(str$(tlRptnameMap(ilLoop).iRptCallType))
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            ilRet = MsgBox("Report Name " & slName & " not found in Mapping Table", vbOkOnly + vbInformation, "Report List")
            cmcCancel.SetFocus
            mFindNameInMap = ilFound
            Exit Function
        End If
        mFindNameInMap = ilFound
        Exit Function
End Function
