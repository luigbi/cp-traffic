VERSION 5.00
Begin VB.Form ExportList 
   Appearance      =   0  'Flat
   Caption         =   "CSI Exports"
   ClientHeight    =   5085
   ClientLeft      =   585
   ClientTop       =   2880
   ClientWidth     =   6195
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
   Icon            =   "Exportlist.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5085
   ScaleWidth      =   6195
   Begin VB.Timer tmcAutoRun 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5325
      Top             =   4710
   End
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4665
      Top             =   4665
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3420
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
      Left            =   1710
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
      ScaleWidth      =   5820
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   225
      Width           =   5880
      Begin VB.ListBox lbcNY 
         Appearance      =   0  'Flat
         Height          =   450
         ItemData        =   "Exportlist.frx":08CA
         Left            =   2970
         List            =   "Exportlist.frx":08D4
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.ListBox lbcCnC 
         Appearance      =   0  'Flat
         Height          =   870
         ItemData        =   "Exportlist.frx":08FB
         Left            =   2970
         List            =   "Exportlist.frx":090B
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   180
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.ListBox lbcExport 
         Appearance      =   0  'Flat
         Height          =   3810
         ItemData        =   "Exportlist.frx":095E
         Left            =   120
         List            =   "Exportlist.frx":0960
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   2715
      End
   End
End
Attribute VB_Name = "ExportList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ExportList.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExportList.Frm
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
Dim tmExportList() As RPTLST
Dim smPassCommands As String
Dim smScreenCaption As String

Dim imIgnoreChg As Integer
Dim smCommandKeys As String
Dim bmInModalModule As Boolean
Dim smExportName As String

Dim tmSaf As SAF
Dim hmSaf As Integer
Dim imSafRecLen As Integer
Dim imLastHourGGChecked As Integer





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
    Dim slDateTime As String
    slDateTime = Format$(gNow(), "m/d/yy") & " " & Format(Now, "hh:mm:ss AMPM")
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
        MsgBox "Unable to Open Alert File, Error = " & str$(ilRet), vbOkOnly + vbInformation, "Warning"
    End If
    hgUlf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hgUlf, "", sgDBPath & "ULF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        MsgBox "Unable to Open User Log File, Error = " & str$(ilRet), vbOkOnly + vbInformation, "Warning"
    End If

    sgExportIniSectionName = ""             'exports.ini section name
    '4/2/11: Add setting of value
    lgUlfCode = 0
    
    bgDevEnv = IsDevEnv()
    
    'examples of auto export command line
    'Auto-Matrix Section-Matrix Std TNet
    'Auto-Efficio Revenue Section-Efficio Cal TNet
    'Auto-Tableau Section-Tableau Std TNet
    'Auto-RAB Section-RAB CalContract
    'Auto-RAB Section-RAB CalSpots
    'Auto-Matrix Section-Matrix Cal Net
    
    '*********
    '      sgCommandSTr - set this field to test auto export and comment out to do auto export in DEBUG
    '       Otherwise, keep it commented out
    ' *********
    'sgCommandStr = "Auto-RAB Section-RAB CalContract"
    If (Trim$(sgCommandStr) = "") Or (Trim$(sgCommandStr) = "/UserInput") Then
        igExportType = 0
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
        If (InStr(1, sgCommandStr, "Auto-", vbTextCompare) = 0) And (InStr(1, sgCommandStr, "Auto -", vbTextCompare) = 0) Then       'in auto export mode?
            igExportType = 1
            igSportsSystem = 0
            ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
            ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
            ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
            ilRet = gParseItem(slCommand, 3, "\", slStr)
            igRptCallType = Val(slStr)
            ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
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
            mGetUlfCode
        Else
            'Running Auto Mode:
            If InStr(1, sgCommandStr, "Matrix", vbTextCompare) <> 0 Then
                igBkgdProg = 15                     'for msg logging, gmsgbox
                igExportType = 4
                smExportName = "Matrix"
                ilPos = InStr(1, sgCommandStr, "Section-", vbTextCompare)
                If ilPos <> 0 Then                              'found match
                    ilLen = Len((sgCommandStr))
                    ilLen = ilLen - ilPos - 7             'get the length of section definition for matrix options, adjust for keyword section
                    sgExportIniSectionName = Mid$(sgCommandStr, ilPos + 8, ilLen)   'pick up description of matrix options for .ini entry
                End If
            ElseIf InStr(1, sgCommandStr, "MillerKapLan", vbTextCompare) <> 0 Then
                igBkgdProg = 16                     'for msg logging,gmsgbox
                igExportType = 3
                smExportName = "MillerKaplan"
                ilPos = InStr(1, sgCommandStr, "Section-", vbTextCompare)           '11-10-14 need to use a section command for parameters
                If ilPos <> 0 Then                              'found match
                    ilLen = Len((sgCommandStr))
                    ilLen = ilLen - ilPos - 7             'get the length of section definition for matrix options, adjust for keyword section
                    sgExportIniSectionName = Mid$(sgCommandStr, ilPos + 8, ilLen)   'pick up description of Miller Kaplan options for .ini entry
                End If
            ElseIf InStr(1, sgCommandStr, "Efficio Revenue", vbTextCompare) <> 0 Then
                igBkgdProg = 16                     'for msg logging,gmsgbox
                igExportType = 3
                smExportName = "Efficio Revenue"
                ilPos = InStr(1, sgCommandStr, "Section-", vbTextCompare)           '11-10-14 need to use a section command for parameters
                If ilPos <> 0 Then                              'found match
                    ilLen = Len((sgCommandStr))
                    ilLen = ilLen - ilPos - 7             'get the length of section definition for matrix options, adjust for keyword section
                    sgExportIniSectionName = Mid$(sgCommandStr, ilPos + 8, ilLen)   'pick up description of Efficio Revenue options for .ini entry
                End If
            ElseIf InStr(1, sgCommandStr, "Efficio Projections", vbTextCompare) <> 0 Then
                igBkgdProg = 16                     'for msg logging,gmsgbox
                igExportType = 2
                smExportName = "Efficio Projections"
                ilPos = InStr(1, sgCommandStr, "Section-", vbTextCompare)           '11-10-14 need to use a section command for parameters
                If ilPos <> 0 Then                              'found match
                    ilLen = Len((sgCommandStr))
                    ilLen = ilLen - ilPos - 7             'get the length of section definition for matrix options, adjust for keyword section
                    sgExportIniSectionName = Mid$(sgCommandStr, ilPos + 8, ilLen)   'pick up description of Efficio Proj options for .ini entry
                End If
            '7-9-15 implement Tableau export - same format as matrix
            ElseIf InStr(1, sgCommandStr, "Tableau", vbTextCompare) <> 0 Then
                igBkgdProg = 18                     'for msg logging,gmsgbox
                igExportType = 5
                smExportName = "Tableau"
                ilPos = InStr(1, sgCommandStr, "Section-", vbTextCompare)           '11-10-14 need to use a section command for parameters
                If ilPos <> 0 Then                              'found match
                    ilLen = Len((sgCommandStr))
                    ilLen = ilLen - ilPos - 7             'get the length of section definition for tableau options, adjust for keyword section
                    sgExportIniSectionName = Mid$(sgCommandStr, ilPos + 8, ilLen)   'pick up description of Efficio Revenue options for .ini entry
                End If
            '1-23-20 implement RAB export - using base Matrix code; add calendar projections from contracts
            ElseIf InStr(1, sgCommandStr, "RAB", vbTextCompare) <> 0 Then
                igBkgdProg = 21                     'for msg logging,gmsgbox
                igExportType = 6
                smExportName = "RAB"
                ilPos = InStr(1, sgCommandStr, "Section-", vbTextCompare)           ' need to use a section command for parameters
                If ilPos <> 0 Then                              'found match
                    ilLen = Len((sgCommandStr))
                    ilLen = ilLen - ilPos - 7             'get the length of section definition for RAB options, adjust for keyword section
                    sgExportIniSectionName = Mid$(sgCommandStr, ilPos + 8, ilLen)   'pick up description of RAB options for .ini entry
                End If
                'TTP 9992
            ElseIf InStr(1, sgCommandStr, "CustomRevenueExport", vbTextCompare) <> 0 Then
                igBkgdProg = 24                     'for msg logging,CustomRevenueExport
                igExportType = 7 'auto-CustomRevenueExport
                smExportName = "CustomRevenueExport"
                ilPos = InStr(1, sgCommandStr, "Section-", vbTextCompare)           ' need to use a section command for parameters
                If ilPos <> 0 Then                              'found match
                    ilLen = Len((sgCommandStr))
                    ilLen = ilLen - ilPos - 7             'get the length of section definition for RAB options, adjust for keyword section
                    sgExportIniSectionName = Mid$(sgCommandStr, ilPos + 8, ilLen)   'pick up description of RAB options for .ini entry
                End If
            ElseIf InStr(1, sgCommandStr, "AdServerBillDisc", vbTextCompare) <> 0 Then
                'igBkgdProg = 24                     'for msg logging,AdServerBillDisc
                igExportType = 8 'auto-AdServerBillDisc
                smExportName = "AdServerBillDisc"
                ilPos = InStr(1, sgCommandStr, "Section-", vbTextCompare)           ' need to use a section command for parameters
                If ilPos <> 0 Then                              'found match
                    ilLen = Len((sgCommandStr))
                    ilLen = ilLen - ilPos - 7             'get the length of section definition for AdServerBillDisc options, adjust for keyword section
                    sgExportIniSectionName = Mid$(sgCommandStr, ilPos + 8, ilLen)   'pick up description of AdServerBillDisc options for .ini entry
                End If
            ElseIf InStr(1, sgCommandStr, "CntrLine", vbTextCompare) <> 0 Then
                'igBkgdProg = 24                     'for msg logging,AdServerBillDisc
                igExportType = 8 'auto-AdServerBillDisc
                smExportName = "CntrLine"
                ilPos = InStr(1, sgCommandStr, "Section-", vbTextCompare)           ' need to use a section command for parameters
                If ilPos <> 0 Then                              'found match
                    ilLen = Len((sgCommandStr))
                    ilLen = ilLen - ilPos - 7             'get the length of section definition for AdServerBillDisc options, adjust for keyword section
                    sgExportIniSectionName = Mid$(sgCommandStr, ilPos + 8, ilLen)   'pick up description of AdServerBillDisc options for .ini entry
                End If
            Else
                igBkgdProg = 0
                igExportType = 1
                smExportName = ""
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
            gUserActivityLog "L", "ExportList.Frm"
            If igExportType > 1 Then 'Startup in Automation mode
                gAutomationAlertAndLogHandler "Starting Exports, Command Line=" & sgCommandStr 'log
            End If
            Exit Sub
        End If
    End If
    'End If
    DoEvents

'    gInitStdAlone ExportList, slStr, igTestSystem
    gInitStdAlone
    mCheckForDate
    mGetExportName
    '4/2/11: Add setting and call.  Note: The call in _Load will be ignored
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "ExportList.Frm"
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
    Dim ilIndex As Integer
    Dim slDateTime As String
    slDateTime = Format$(gNow(), "m/d/yy") & " " & Format(Now, "hh:mm:ss AMPM")
    sgMessageFile = "" 'reset the global MessageFilename (Log File), this gets set to the full path/filename in various mOpenMsgFile() functions
    
    On Error Resume Next
    If lbcExport.ListIndex < 0 Then
        Exit Sub
    End If
    mCheckGG
    ilIndex = lbcExport.ItemData(lbcExport.ListIndex)
    If ((ilIndex >= 1) And (ilIndex <= 3)) Or ((ilIndex >= 6) And (ilIndex <= 11)) Or (ilIndex = 13) Or ((ilIndex >= 15) And (ilIndex <= 16)) Or ((ilIndex >= 18) And (ilIndex <= 19)) Then
        'System Flag prevents these from Running
        If (igGGFlag = 0) Then
            gAutomationAlertAndLogHandler "** System Flag prevents these from Running! **"
            Exit Sub
        End If
    Else
        'for accounting type reports..  'Report & System Flags prevents these from Running
        If (igGGFlag = 0) And (igRptGGFlag = 0) Then
            gAutomationAlertAndLogHandler "** (accounting type reports) System Flag prevents these from Running! **"
            Exit Sub
        End If
    End If
    
    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE))
    End If
    sgCommandStr = slStr
    Select Case lbcExport.ItemData(lbcExport.ListIndex)
        Case 0  'Accounting
            ExpACC.Show vbModal
        Case 1  'Audio ISCI Title
            ExpISCIXRef.Show vbModal
        Case 2  'Audio MP2
            ExptMP2.Show vbModal
        Case 3  'Automation
            ExptGen.Show vbModal
        Case 4  'Barter Payment
            ExpGPBarter.Show vbModal
        Case 5  'Clearance n Compensation
            If lbcCnC.ListIndex = 0 Then
                ExpCnCAP.Show vbModal
            ElseIf lbcCnC.ListIndex = 1 Then
                ExpCnCNI.Show vbModal
            ElseIf lbcCnC.ListIndex = 2 Then
                ExpCnCSA.Show vbModal
            ElseIf lbcCnC.ListIndex = 3 Then
                ExpCnCSS.Show vbModal
            End If
        Case 6  'Commercial Change
            ExpCmChg.Show vbModal
        Case 7  'Copy Bulk Feed
            ExpBkCpy.Show vbModal
        Case 8  'Corporate Export
            ilShell = Shell(sgExePath & "ExportProj.Exe /IniLoc:" & CurDir$ & " /UserInput", vbMinimizedNoFocus)
        Case 9  'Dallas Feed
            ExpDall.Show vbModal
        Case 10  'Enco
            ExptEnco.Show vbModal
        Case 11  'Get Paid
            ilRet = MsgBox("This feature can be setup to run automatically.  Click Ok to Proceed", vbDefaultButton2 + vbInformation + vbOKCancel, "GetPaid Export")
            If ilRet = vbOK Then
                ilShell = Shell(sgExePath & "GetPaid.Exe /IniLoc:" & CurDir$, vbMinimizedNoFocus)
            End If
        Case 12  'Great Plains G/L
            ExpGP.Show vbModal
        Case 13  'Invoice
            ExpInv.Show vbModal
        Case 14  'Matrix
            ExpMatrix.Show vbModal
        Case 15  'New York Feed
            If lbcNY.ListIndex = 0 Then
                If igTestSystem Then
                    slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE)) & "\ASP"
                Else
                    slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE)) & "\ASP"
                End If
                sgCommandStr = slStr
                ExpNY.Show vbModal
            'ElseIf lbcNY.ListIndex = 1 Then
            '    If igTestSystem Then
            '        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE)) & "\EAS"
            '    Else
            '        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE)) & "\EAS"
            '    End If
            '    sgCommandStr = slStr
            '    ExpNY.Show vbModal
            'ElseIf lbcNY.ListIndex = 2 Then
            ElseIf lbcNY.ListIndex = 1 Then
                If igTestSystem Then
                    slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(str$(CALLNONE)) & "\ESPN"
                Else
                    slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(str$(CALLNONE)) & "\ESPN"
                End If
                sgCommandStr = slStr
                ExpNY.Show vbModal
            End If
        Case 16  'Phoenix Log
            ExpPhnx.Show vbModal
        Case 17  'Revenue
            ExpRevenue.Show vbModal
        Case 18  'Airwave
            ExptAirWave.Show vbModal
        Case 19  'Carts
            ExptCart.Show vbModal
        Case EXP_EFFICIOREV '20
            ExpEfficioRev.Show vbModal
        Case EXP_EFFICIOPROJ '21
            ExpEfficioRev.Show vbModal
        Case EXP_TABLEAU '22
            ExpMatrix.Show vbModal
        Case 23
            ExpStnFd.Show vbModal
        Case EXP_CASH '24                                 '1-5-18 Cash Receipts
            ExpCashOrInv.Show vbModal
        Case Exp_INVREG '25                                 '1-5-18 Invoice Register
            ExpCashOrInv.Show vbModal
        Case EXP_MK '26
            ExpMK.Show vbModal
        Case EXP_RAB '27
            ExpMatrix.Show vbModal              '1-23-20  RAB export using some matrix code
        Case EXP_CUST_REV '28
            ExpMatrix.Show vbModal              '12-9-20  Custom Revenue Export using some matrix code
        Case EXP_ADSERVERBILLDISC '29
            ExpPodBillDisc.Show vbModal
         Case EXP_AUDACYINV '30
            ExpCashOrInv.Show vbModal           'TTP 10205 - 6/21/21 - JW - WO Invoice Export
         Case EXP_AUDACYINVLINE '31
            ExpCntrLine.Show vbModal           '8/17/21 - JW - TTP 10233 - Audacy: line summary export
    End Select
    
    'TTP 10342: Automation Alerting and Logging
    If sgAutomationLogBuffer <> "" Then
        'Theres un-logged messages print to a generic Exports.Log
        If sgMessageFile = "" Then sgMessageFile = sgDBPath & "Messages\" & "Exports.Log"
        gAutomationAlertAndLogHandler "Selected Export returned to Export List / Finished."
    End If
    sgAutomationLogBuffer = ""
    sgMessageFile = "" 'reset the global MessageFilename (Log File), this gets set to the full path/filename in various mOpenMsgFile() functions
    
        
    ''Dan M   need to find if counterpoint date has been changed in traffic.
    'mGetCsiDate
    bmInModalModule = False
    If igExportType > 1 Then        'if auto mode, kill module
        cmcCancel_Click
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
    gMsgBox "traffic exports couldn't read values in " & slFullPath & ".  Form_GotFocus", vbOkOnly, "Problem reading values from file."
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
    'ExportList.Refresh
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
            If smExportName <> "" Then
                tmcAutoRun.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'messages weren't written, make Exports.log and append messages
    If sgAutomationLogBuffer <> "" Then
        'Theres un-logged messages print to a generic Exports.Log
        sgMessageFile = sgDBPath & "Messages\" & "Exports.Log"
        gAutomationAlertAndLogHandler "Terminating Exports.exe"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    Erase tmExportList
    
    ilRet = btrClose(hgAuf)
    btrDestroy hgAuf
    ilRet = btrClose(hgUlf)
    btrDestroy hgUlf
    
    '4/2/11: Add setting and call.
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    Set ExportList = Nothing   'Remove data segment
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


Private Sub lbcExport_Click()
    Dim slStr As String
    
    If lbcExport.ListIndex >= 0 Then
        slStr = lbcExport.Text
        If slStr = "Clearance n Compensation" Then
            lbcCnC.Visible = True
            lbcNY.Visible = False
        ElseIf slStr = "Engineering Feed" Then
            lbcCnC.Visible = False
            lbcNY.Visible = True
        Else
            lbcCnC.Visible = False
            lbcNY.Visible = False
        End If
        
    Else
        lbcCnC.Visible = False
        lbcNY.Visible = False
    End If
End Sub

Private Sub lbcExport_DblClick()
    cmcDone_Click
End Sub

Private Sub lbcExport_GotFocus()
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
    ReDim tgNtrSortInfo(0 To 0) As NTRSORTINFO              'array of vehicles with start/end indices pointing to TGNTRInfo array of SBF records
    ExportList.Height = cmcDone.Top + 3 * cmcDone.Height '/ 3
    If smExportName <> "" Then
        ExportList.Left = -ExportList.Width - 100
    Else
        gCenterStdAlone ExportList
    End If
    'ExportList.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    smScreenCaption = "Export Selection"
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
    lbcExport.Visible = True
    If imTerminate Then
        Exit Sub
    End If
    
    igGGFlag = 1
    imLastHourGGChecked = -1
    
    plcScreen_Paint
'    gCenterModalForm ExportList
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
    
'Dim Exp_MK As Integer
    
    lbcExport.AddItem "Accounting"
    lbcExport.ItemData(lbcExport.NewIndex) = 0
    lbcExport.AddItem "Audio ISCI Title"
    lbcExport.ItemData(lbcExport.NewIndex) = 1
    '7496
    lbcExport.AddItem "Audio " & UCase(Mid(sgAudioExtension, 2))
   ' lbcExport.AddItem "Audio MP2"
    lbcExport.ItemData(lbcExport.NewIndex) = 2
    lbcExport.AddItem "Automation"
    lbcExport.ItemData(lbcExport.NewIndex) = 3
    If (Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) = GREATPLAINS Then
        lbcExport.AddItem "Barter Payment"
        lbcExport.ItemData(lbcExport.NewIndex) = 4
    End If
    lbcExport.AddItem "Clearance n Compensation"
    lbcExport.ItemData(lbcExport.NewIndex) = 5
    lbcExport.AddItem "Commercial Change"
    lbcExport.ItemData(lbcExport.NewIndex) = 6
    lbcExport.AddItem "Copy Bulk Feed"
    lbcExport.ItemData(lbcExport.NewIndex) = 7
    If (Asc(tgSpf.sUsingFeatures) And REVENUEEXPORT) = REVENUEEXPORT Then
        lbcExport.AddItem "Corporate Export"
        lbcExport.ItemData(lbcExport.NewIndex) = 8
    End If
    lbcExport.AddItem "Dallas Feed"
    lbcExport.ItemData(lbcExport.NewIndex) = 9
    lbcExport.AddItem "Enco"
    lbcExport.ItemData(lbcExport.NewIndex) = 10
    If ((Asc(tgSpf.sUsingFeatures6) And GETPAIDEXPORT) = GETPAIDEXPORT) Then
        lbcExport.AddItem "Get Paid"
        lbcExport.ItemData(lbcExport.NewIndex) = 11
    End If
    If (Asc(tgSpf.sUsingFeatures2) And GREATPLAINS) = GREATPLAINS Then
        lbcExport.AddItem "Great Plains G/L"
        lbcExport.ItemData(lbcExport.NewIndex) = 12
    End If
    
    '5-11-17 Invoice export has been disabled; previously coded for Channel One
'    If ((Asc(tgSpf.sUsingFeatures6) And INVEXPORTPARAMETERS) = INVEXPORTPARAMETERS) Then
'        lbcExport.AddItem "Invoice"
'        lbcExport.ItemData(lbcExport.NewIndex) = 13
'    End If

    If igExportType = 4 Then            'background mode, if disallowed catch it in Matrix module
        lbcExport.AddItem "Matrix"
        lbcExport.ItemData(lbcExport.NewIndex) = 14
    Else                                'not background mode, dont show in menu list of interactive/manual mode
        If (Asc(tgSpf.sUsingFeatures) And MATRIXEXPORT) = MATRIXEXPORT Or (Asc(tgSaf(0).sFeatures1) And MATRIXCAL) = MATRIXCAL Then
            lbcExport.AddItem "Matrix"
            lbcExport.ItemData(lbcExport.NewIndex) = 14
        End If
    End If
    lbcExport.AddItem "Engineering Feed"
    lbcExport.ItemData(lbcExport.NewIndex) = 15
    lbcExport.AddItem "Phoenix Log"
    lbcExport.ItemData(lbcExport.NewIndex) = 16
    If ((Asc(tgSpf.sUsingFeatures7) And EXPORTREVENUE) = EXPORTREVENUE) Then
        lbcExport.AddItem "Revenue"
        lbcExport.ItemData(lbcExport.NewIndex) = 17
    End If

    ilRet = gVffRead()
    'For ilVff = LBound(tgVff) To UBound(tgVff) - 1 Step 1
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If tgVff(ilVff).sExportAirWave = "Y" Then
            lbcExport.AddItem "AirWave"
            lbcExport.ItemData(lbcExport.NewIndex) = 18
            Exit For
        End If
    Next ilVff
    lbcExport.AddItem "Carts"
    lbcExport.ItemData(lbcExport.NewIndex) = 19
    ilRet = gObtainMCF()

    If igExportType = 2 Or igExportType = 3 Then        'background mode, efficio revenue or efficio projection.  place the task in menu list to activate in background.
        lbcExport.AddItem "Efficio Revenue"
        lbcExport.ItemData(lbcExport.NewIndex) = EXP_EFFICIOREV
        lbcExport.AddItem "Efficio Projections"
        lbcExport.ItemData(lbcExport.NewIndex) = EXP_EFFICIOPROJ
    Else                                'not background mode, dont show in menu list of interactive/manual mode if disallowed
        If (Asc(tgSaf(0).sFeatures1) And EFFICIOEXPORT) = EFFICIOEXPORT Then
            lbcExport.AddItem "Efficio Revenue"
            lbcExport.ItemData(lbcExport.NewIndex) = EXP_EFFICIOREV
            lbcExport.AddItem "Efficio Projections"
            lbcExport.ItemData(lbcExport.NewIndex) = EXP_EFFICIOPROJ
        End If
    End If
    
    If igExportType = 5 Then            'Tableau  background mode, if disallowed catch it in Matrix module
        lbcExport.AddItem "Tableau"
        lbcExport.ItemData(lbcExport.NewIndex) = EXP_TABLEAU
    Else
        If (Asc(tgSaf(0).sFeatures2) And TABLEAUEXPORT) = TABLEAUEXPORT Or (Asc(tgSaf(0).sFeatures2) And TABLEAUCAL) = TABLEAUCAL Then
            lbcExport.AddItem "Tableau"
            lbcExport.ItemData(lbcExport.NewIndex) = EXP_TABLEAU
        End If
    End If
    
    If tgSpf.sSystemType <> "R" Then
        If tgSpf.sGUseAffFeed = "Y" Then
            lbcExport.AddItem "Station Feed"
            lbcExport.ItemData(lbcExport.NewIndex) = 23
        End If
    End If
    
    '1-5-18 Cash Receipand Invoice Register exports
    lbcExport.AddItem "Cash Receipts"
    lbcExport.ItemData(lbcExport.NewIndex) = EXP_CASH
    lbcExport.AddItem "Invoice Register"
    lbcExport.ItemData(lbcExport.NewIndex) = Exp_INVREG
    'Date: 6/27/2019    Miller Kaplan Billing exports
    'FYM
'    Exp_MK = 26
    lbcExport.AddItem "Miller Kaplan"
    lbcExport.ItemData(lbcExport.NewIndex) = EXP_MK
    
    'exp_RAB = 27
    If igExportType = 6 Then            'RAB  background mode, if disallowed catch it in Matrix module
        lbcExport.AddItem "RAB"
        lbcExport.ItemData(lbcExport.NewIndex) = EXP_RAB
    Else
        If (Asc(tgSaf(0).sFeatures6) And RABCALENDAR) = RABCALENDAR Or (Asc(tgSaf(0).sFeatures7) And RABSTD) = RABSTD Or (Asc(tgSaf(0).sFeatures7) And RABCALSPOTS) = RABCALSPOTS Then
            lbcExport.AddItem "RAB"
            lbcExport.ItemData(lbcExport.NewIndex) = EXP_RAB
        End If
    End If
    
    'TTP 9992
    If (Asc(tgSaf(0).sFeatures7) And CUSTOMEXPORT) = CUSTOMEXPORT Then
        lbcExport.AddItem "Custom Revenue Export"
        lbcExport.ItemData(lbcExport.NewIndex) = EXP_CUST_REV '28
    End If
    
    '3/12/21 - Ad Server Discrepancy Export
    If (Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER Then
        lbcExport.AddItem "Ad Server Billing Discrepancy"
        lbcExport.ItemData(lbcExport.NewIndex) = EXP_ADSERVERBILLDISC '29
    End If
    
    'TTP 10205 - 6/3/21 - Get SPFX Extended Site Features : spfxInvExpFeature - Audacy WO Invoice Export, TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
    If (tgSpfx.iInvExpFeature And INVEXP_AUDACYWO) = INVEXP_AUDACYWO Then
        lbcExport.AddItem "WO Invoice"
        lbcExport.ItemData(lbcExport.NewIndex) = EXP_AUDACYINV
    End If
    
    '8/17/21 - JW - TTP 10233 - Audacy: Contract line export, TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
    If (tgSpfx.iInvExpFeature And INVEXP_AUDACYLINE) = INVEXP_AUDACYLINE Then
         lbcExport.AddItem "Contract Line"
         lbcExport.ItemData(lbcExport.NewIndex) = EXP_AUDACYINVLINE
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
'    Erase tmExportList
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
    Unload ExportList
'    Set ExportList = Nothing   'Remove data segment
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
    
    Dim slDateTime As String
    slDateTime = Format$(gNow(), "m/d/yy") & " " & Format(Now, "hh:mm:ss AMPM")
    
    tmcAutoRun.Enabled = False
    If smExportName = "CustomRevenueExport" Then
        smExportName = "Custom Revenue Export"
    End If
    If smExportName = "AdServerBillDisc" Then
        smExportName = "Ad Server Billing Discrepancy"
    End If
    Me.Left = -Me.Width - 100
    For ilLoop = 0 To lbcExport.ListCount - 1 Step 1
        slStr = lbcExport.List(ilLoop)
        ilPos = InStr(1, smExportName, slStr, vbTextCompare)
        If ilPos = 1 Then
            lbcExport.ListIndex = ilLoop
            If slStr = "Clearance n Compensation" Then
                ilPos = InStr(1, smExportName, "/Sub: AP", vbTextCompare)
                If InStr(1, smExportName, "/Sub: AP", vbTextCompare) > 0 Then
                    lbcCnC.ListIndex = 0
                ElseIf InStr(1, smExportName, "/Sub: NI", vbTextCompare) > 0 Then
                    lbcCnC.ListIndex = 1
                ElseIf InStr(1, smExportName, "/Sub: SA", vbTextCompare) > 0 Then
                    lbcCnC.ListIndex = 2
                Else
                    lbcCnC.ListIndex = 3
                End If
            ElseIf slStr = "Engineering Feed" Then
                'ilPos = InStr(1, smExportName, "/Sub: EAS", vbTextCompare)
                'If ilPos > 0 Then
                '    lbcNY.ListIndex = 1
                'Else
                    ilPos = InStr(1, smExportName, "/Sub: ESPN", vbTextCompare)
                    If ilPos > 0 Then
                        'lbcNY.ListIndex = 2
                        lbcNY.ListIndex = 1
                    Else
                        lbcNY.ListIndex = 0
                    End If
                'End If
            End If
                        
            cmcDone_Click
            mTerminate
            Exit Sub
        End If
    Next ilLoop
    
    If smExportName = "CntrLine" Then
        smExportName = "Contract Line"
        'the Instr Compare cant determine the difference between "WO Invoice" and "Contract Line"
        For ilLoop = 0 To lbcExport.ListCount - 1 Step 1
            slStr = lbcExport.List(ilLoop)
            If smExportName = slStr Then
                ilPos = ilLoop
                lbcExport.ListIndex = ilLoop
                cmcDone_Click
                mTerminate
                Exit Sub
            End If
        Next ilLoop
    End If

    If igExportType > 1 Then 'Startup in Automation mode
        gAutomationAlertAndLogHandler "No Export was Selected!" 'log
    End If
    
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

Private Sub mGetExportName()
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    smExportName = ""
    ilPos = InStr(1, sgCommandStr, "/Export:", 1)
    If ilPos > 0 Then
        'ilPos = InStr(slCommand, ":")
        ilSpace = InStr(ilPos, sgCommandStr, "~")
        If ilSpace = 0 Then
            smExportName = Trim$(Mid$(sgCommandStr, ilPos + 8))
        Else
            smExportName = Trim$(Mid$(sgCommandStr, ilPos + 8, ilSpace - ilPos - 8))
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
    
    If imLastHourGGChecked = Hour(Now) Then
        Exit Sub
    End If
    imLastHourGGChecked = Hour(Now)
    
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


