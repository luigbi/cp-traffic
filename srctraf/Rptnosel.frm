VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptNoSel 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report"
   ClientHeight    =   1485
   ClientLeft      =   240
   ClientTop       =   3840
   ClientWidth     =   9270
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
   ScaleHeight     =   1485
   ScaleWidth      =   9270
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   18
      Top             =   615
      Width           =   2055
   End
   Begin VB.Timer tmcDone 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8565
      Top             =   -120
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
      Left            =   720
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4515
      Width           =   120
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   8835
      Top             =   -210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".Txt"
      Filter          =   "*.Txt|*.Txt|*.Doc|*.Doc|*.Asc|*.Asc"
      FilterIndex     =   1
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Frame frcCopies 
      Caption         =   "Printing"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2070
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.TextBox edcCopies 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "1"
         Top             =   345
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   375
         Width           =   855
      End
   End
   Begin VB.Frame frcFile 
      Caption         =   "Save to File"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2070
      TabIndex        =   9
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1575
         TabIndex        =   14
         Top             =   975
         Width           =   1005
      End
      Begin VB.ComboBox cbcFileType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   780
         TabIndex        =   11
         Top             =   255
         Width           =   2925
      End
      Begin VB.TextBox edcFileName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   780
         TabIndex        =   13
         Top             =   645
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   285
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   675
         Width           =   645
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   19
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   17
      Top             =   150
      Width           =   2805
   End
   Begin VB.Frame frcOutput 
      Caption         =   "Report Destination"
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   1890
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Save to File"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   870
         Width           =   1425
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Print"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   630
         Width           =   855
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Display"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   1455
      End
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
      Left            =   8295
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   -75
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
      Left            =   7860
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   -90
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8850
      Top             =   1035
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "RptNoSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptnosel.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptNoSel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the MultiName Report selection screen code
Option Explicit
Option Compare Text
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imFTSelectedIndex As Integer
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
Dim imFirstTime As Integer
Dim imGenRpt As Integer    'True=Processing Generate Report; False=Process User input
Dim imOutput As Integer
Dim imCopies As Integer
Dim imFile As Integer
'Contract report
'Seven day log
Dim smUserCode As String
'Dim smVefCode As String
Dim smStartDate As String
Dim hmRvf As Integer        'Rvf handle
Dim imFirstActivate As Integer
Dim imTerminate As Integer

Private Sub cbcFileType_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcFileType.Text <> "" Then
            gManLookAhead cbcFileType, imBSMode, imComboBoxIndex
        End If
        imFTSelectedIndex = cbcFileType.ListIndex
        imChgMode = False
    End If
    mSetCommands
End Sub
Private Sub cbcFileType_Click()
    imComboBoxIndex = cbcFileType.ListIndex
    imFTSelectedIndex = cbcFileType.ListIndex
    mSetCommands
End Sub
Private Sub cbcFileType_GotFocus()
    If cbcFileType.Text = "" Then
        cbcFileType.ListIndex = 0
    End If
    imComboBoxIndex = cbcFileType.ListIndex
    gCtrlGotFocus cbcFileType
End Sub
Private Sub cbcFileType_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcFileType_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcFileType.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cmcBrowse_Click()
    cdcSetup.flags = cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
    cdcSetup.fileName = edcFileName.Text
    cdcSetup.InitDir = Left$(sgRptSavePath, Len(sgRptSavePath) - 1)
    cdcSetup.Action = 2    'DLG_FILE_SAVE
    edcFileName.Text = cdcSetup.fileName
    mSetCommands
    gChDrDir        '3-25-03
    'ChDrive Left$(sgCurDir, 2)  'Set the default drive
    'ChDir sgCurDir              'set the default path
End Sub
Private Sub cmcBrowse_GotFocus()
    gCtrlGotFocus cmcBrowse
End Sub
Private Sub cmcCancel_Click()
    If imGenRpt Then
        Exit Sub
    End If
    'If igRptCallType = COLLECTIONSJOB Then
        mTerminate False
    'Else
    '    mTerminate True
    'End If
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
    Dim ilDDFSet As Integer
    Dim ilCopies As Integer
    Dim ilRet As Integer
    Dim slFileName As String
    Dim ilPassNo As Integer
    Dim ilRptNo As Integer

    If imGenRpt Then
        Exit Sub
    End If
    imGenRpt = True
    imOutput = frcOutput.Enabled
    imCopies = frcCopies.Enabled
    'imWhen = frcWhen.Enabled
    imFile = frcFile.Enabled
    frcOutput.Enabled = False
    frcCopies.Enabled = False
    'frcWhen.Enabled = False
    frcFile.Enabled = False
    If igRptCallType = COLLECTIONSJOB Then
        If igRptType = 1 Then
            ilPassNo = 2
        Else
            If tgSpf.iReconcGroupNo = 0 Then
                ilPassNo = 3
            Else
                ilPassNo = 6        '1-20-04  addl 3 reports by vehicle group if entered in spf
            End If
        End If
    Else
        ilPassNo = 1
    End If
    'dan change for multiple reports to be displayed once 11/05/08
    Set ogReport = New CReportHelper
    ogReport.iLastPrintJob = ilPassNo
    For ilRptNo = 1 To ilPassNo Step 1
        igJobRptNo = ilRptNo
        If Not mGenReport(ilRptNo) Then
            frcOutput.Enabled = imOutput
            frcCopies.Enabled = imCopies
            'frcWhen.Enabled = imWhen
            frcFile.Enabled = imFile
            imGenRpt = False
            Exit Sub
        End If
        If ilRptNo >= ogReport.iLastPrintJob Then     'Dan  only output once.
        
            '11-12-09 Site needs prepass to generate all features/options on report
            If igRptCallType = SITELIST Then
                gCreateSite
            End If
        
            If rbcOutput(0).Value Then
                DoEvents            '9-19-02 fix for timing problem to prevent freezing before calling crystal
                igDestination = 0
                'frcCopies.Enabled = False
                'frcFile.Enabled = False
                'frcOutput.Enabled = False
                'frcWhen.Enabled = False
                'cmcCancel.Enabled = False
                'cmcDone.Enabled = False
                'mGenReport
                Report.Show vbModal
            ElseIf rbcOutput(1).Value Then
                ilCopies = Val(edcCopies.Text)
                ilRet = gOutputToPrinter(ilCopies)
            Else
                slFileName = edcFileName.Text
                'ilRet = gOutputToFile(slFileName, imFTSelectedIndex)
                ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '2-27-04
            End If
        End If      'ilrptno = ilastprintjob
    Next ilRptNo
    Set ogReport = Nothing      'end crxi change  11/05/08
    If (igRptCallType = COLLECTIONSJOB) And (igRptType = 0) Then
        'Report.Hide
        ARReconc.Show vbModal
        mTerminate True            '3-16-04 change to true &force return back to caller for re-entrance
   ElseIf igRptCallType = SITELIST Then       '11-12-09 clear out prepass table
        gCRAvrClear
        gIvrClear
        Screen.MousePointer = vbDefault
    End If
    frcOutput.Enabled = imOutput
    frcCopies.Enabled = imCopies
    'frcWhen.Enabled = imWhen
    frcFile.Enabled = imFile
    imGenRpt = False
    'pbcClickFocus.SetFocus
    tmcDone.Enabled = True
    'mTerminate
    Exit Sub

    ilDDFSet = True
    Resume Next
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcList_Click()
    mTerminate True
End Sub
Private Sub cmcSetup_Click()
    'cdcSetup.Flags = cdlPDPrintSetup
    'cdcSetup.Action = 5    'DLG_PRINT
    cdcSetup.flags = cdlPDPrintSetup
    cdcSetup.ShowPrinter
End Sub
Private Sub edcCopies_Change()
    mSetCommands
End Sub
Private Sub edcCopies_GotFocus()
    gCtrlGotFocus edcCopies
End Sub
Private Sub edcCopies_KeyPress(KeyAscii As Integer)
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcFileName_Change()
    mSetCommands
End Sub
Private Sub edcFileName_GotFocus()
    gCtrlGotFocus edcFileName
End Sub
Private Sub edcFileName_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer

    ilPos = InStr(edcFileName.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcFileName.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    If ((KeyAscii <> KEYBACKSPACE) And (KeyAscii <= 32)) Or (KeyAscii = 34) Or (KeyAscii = 39) Or ((KeyAscii >= KEYDOWN) And (KeyAscii <= 45)) Or ((KeyAscii >= 59) And (KeyAscii <= 63)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    Me.KeyPreview = True
    RptNoSel.Refresh
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

Private Sub Form_Load()
    imTerminate = False
    imFirstActivate = True
    mParseCmmdLine
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    imGenRpt = False
    mInitGenReport
    'RptNoSel.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    PECloseEngine
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set RptNoSel = Nothing   'Remove data segment

End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGenReport                      *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain report file name        *
'*                                                     *
'*******************************************************
Private Function mGenReport(ilRptNo As Integer) As Integer
    Dim slSelection As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slStr As String
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilVehicleGroup As Integer   '1-20-04
    Dim slTime As String


    Select Case igRptCallType
        Case VEHICLEGROUPSLIST    'Vehicle Groups
            If Not gOpenPrtJob("MNfvfgrp.Rpt") Then
                mGenReport = False
                Exit Function
            End If
        Case POTENTIALCODESLIST    'Potential Codes
            If Not gOpenPrtJob("MNfpotn.Rpt") Then
                mGenReport = False
                Exit Function
            End If
        Case BUSCATEGORIESLIST    'Business Categories
            If Not gOpenPrtJob("MNfBCat.Rpt") Then
                mGenReport = False
                Exit Function
            End If
        Case DEMOSLIST      'Customized Demos
            If Not gOpenPrtJob("MNfCDemo.Rpt") Then
                mGenReport = False
                Exit Function
            End If
        Case COMPETITORSLIST    'Competitors
            If Not gOpenPrtJob("MNfCmpet.Rpt") Then
                mGenReport = False
                Exit Function
            End If
        Case ITEMBILLINGTYPESLIST    'Item Billing
            If Not gOpenPrtJob("MNfIt.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "MNfIt.Rpt"
        Case INVOICESORTLIST    'Invoice sort
            If Not gOpenPrtJob("MNfIs.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "MNfIs.Rpt"
        Case EXCLUSIONSLIST
            If Not gOpenPrtJob("MNfPe.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "MNfPe.Rpt"
        Case ANNOUNCERNAMESLIST    'Announcer
            If Not gOpenPrtJob("MNfAn.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "MNfAn.Rpt"
        Case GENRESLIST    'Genre
            If Not gOpenPrtJob("MNfGe.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "MNfGe.Rpt"
        Case SALESREGIONSLIST
            If Not gOpenPrtJob("MNfSr.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "MNfSr.Rpt"
        Case SALESSOURCESLIST    'Sales Source
            If Not gOpenPrtJob("MNfSs.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "MNfSs.Rpt"
        Case SALESTEAMSLIST     'Sales Team
            If Not gOpenPrtJob("MNfSt.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "MNfSt.Rpt"
        Case REVENUESETSLIST    'Revenue Sets
            If Not gOpenPrtJob("MNfRb.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "MNfRb.Rpt"
        Case BOILERPLATESLIST
            If Not gOpenPrtJob("Cmf.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "Cmf.Rpt"
        Case MISSEDREASONSLIST    'Missed Reason
            If Not gOpenPrtJob("MNfMr.Rpt") Then
                mGenReport = False
                Exit Function
            End If
           'Report!crcReport.ReportFileName = sgRptPath & "MNfMr.Rpt"
        Case COMPETITIVESLIST    'Competitive
            If Not gOpenPrtJob("MNfCm.Rpt") Then
                mGenReport = False
               Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "MNfCm.Rpt"
        Case FEEDTYPESLIST    'Network
            If Not gOpenPrtJob("MNfNn.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "MNfNn.Rpt"
        Case EVENTTYPESLIST     'Event Type
            If Not gOpenPrtJob("Etf.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "Etf.Rpt"
        Case AVAILNAMESLIST
            If Not gOpenPrtJob("Anf.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "Anf.Rpt"
        Case SALESOFFICESLIST
            If Not gOpenPrtJob("Sof.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "Sof.Rpt"
        Case MEDIADEFINITIONSLIST
            If Not gOpenPrtJob("Mcf.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            '10-11-06  if using media codes by vehicle, show the applicable vehicle with the media definition
            If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) = MEDIACODEBYVEH Then
                If Not gSetFormula("MediaCodeByVeh", "'Y'") Then 'use media codes by vehicle
                    mGenReport = False
                    Exit Function
                End If
            Else
                If Not gSetFormula("MediaCodeByVeh", "'N'") Then 'use media codes by vehicle
                    mGenReport = False
                    Exit Function
                End If
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "Mcf.Rpt"
        Case LOCKBOXESLIST
            If Not gOpenPrtJob("ArfLk.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "ArfLk.Rpt"
        Case EDISERVICESLIST
            If Not gOpenPrtJob("ArfDp.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "ArfDp.Rpt"
        Case TRANSACTIONSLIST
            If Not gOpenPrtJob("MnfTT.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "ArfDp.Rpt"
        Case SITELIST
            'If Not gOpenPrtJob("Spf.Rpt") Then
            If Not gOpenPrtJob("SiteOptions.rpt") Then      '11-13-09
                mGenReport = False
                Exit Function
            End If
       
            gCurrDateTime slDate, slTime, slMonth, slDay, slYear
            slSelection = "{AVR_Quarterly_Avail.avrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
            slSelection = slSelection & " And Round({AVR_Quarterly_Avail.avrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
            If Not gSetSelection(slSelection) Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "Spf.Rpt"
        Case USERLIST
            If Not gOpenPrtJob("Urf.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.SelectionFormula = slSelection
            'Report!crcReport.ReportFileName = sgRptPath & "Urf.Rpt"
        Case RATECARDSJOB
            If Not gOpenPrtJob("Rcf.Rpt") Then
                mGenReport = False
                Exit Function
            End If
            'Report!crcReport.ReportFileName = sgRptPath & "Rcf.Rpt"
        Case LOGSJOB
            If igRptType = 0 Then   'One day log
                If Not gOpenPrtJob("LogOne.Rpt") Then
                    mGenReport = False
                    Exit Function
                End If
                slSelection = "{ODF_One_Day_Log.odfurfCode} = " & smUserCode
                If Not gSetSelection(slSelection) Then
                    mGenReport = False
                    Exit Function
                End If
                'Report!crcReport.SelectionFormula = slSelection
                'Report!crcReport.ReportFileName = sgRptPath & "LogOne.Rpt"
            ElseIf igRptType = 1 Then   'Commercial schedule
                If Not gOpenPrtJob("LogSeven.Rpt") Then
                    mGenReport = False
                    Exit Function
                End If
                'slSelection = gGetSelectionString(igPrtJob)
                slSelection = "{ODF_One_Day_Log.odfurfCode} = " & smUserCode
                If Not gSetSelection(slSelection) Then
                    mGenReport = False
                    Exit Function
                End If
                'slSelection = gGetSelectionString(igPrtJob)
                'Report!crcReport.SelectionFormula = slSelection
                'slSelection = gGetFormulaString(igPrtJob, "Start Date")
                gObtainYearMonthDayStr smStartDate, True, slYear, slMonth, slDay
                If Not gSetFormula("Start Date", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                    mGenReport = False
                    Exit Function
                End If
                'slSelection = gGetFormulaString(igPrtJob, "Start Date")
                'Report!crcReport.Formulas(0) = "Start Date= Date(" & slYear & ", " & slMonth & ", " & slDay & ")"
                'Report!crcReport.ReportFileName = sgRptPath & "LogSeven.Rpt"
            ElseIf igRptType = 2 Then   'Delivery
                If Not gOpenPrtJob("LogDeliv.Rpt") Then
                    mGenReport = False
                    Exit Function
                End If
                slSelection = "{ODF_One_Day_Log.odfurfCode} = " & smUserCode '& " And {ODF_One_Day_Log.odfmnfFeed} = " & smVefCode
                If Not gSetSelection(slSelection) Then
                    mGenReport = False
                    Exit Function
                End If
                'Report!crcReport.SelectionFormula = slSelection
                'Report!crcReport.ReportFileName = sgRptPath & "LogDeliv.Rpt"
            End If
        Case COLLECTIONSJOB
            If igRptType = 1 Then
                If ilRptNo = 1 Then
                    If Not gOpenPrtJob("CreditAg.Rpt") Then
                        mGenReport = False
                        Exit Function
                    End If
                    slSelection = "{AGF_Agencies.agfCreditRestr} <>" & "'N'" & " And " & "{AGF_Agencies.agfCreditRestr} <>" & "'P'"
                    slSelection = slSelection & " And " & "{@Credit Used} <>" & "0"
                    If Not gSetSelection(slSelection) Then
                        mGenReport = False
                        Exit Function
                    End If
                Else
                    If Not gOpenPrtJob("CreditAd.Rpt") Then
                        mGenReport = False
                        Exit Function
                    End If
                    slSelection = "{ADF_Advertisers.adfCreditRestr} <>" & "'N'" & " And " & "{ADF_Advertisers.adfCreditRestr} <>" & "'P'"
                    slSelection = slSelection & " And " & "{@Credit Used} <>" & "0"
                    If Not gSetSelection(slSelection) Then
                        mGenReport = False
                        Exit Function
                    End If
                End If
            Else    'reconciliation reports - 3 phases: 1 = trans entered for period, 2 = period balancing, 3 = future totals
                '1-20-04 each phase has an additional report sorted by vehicle group (major) if entered in spf
                'If ilRptNo = 1 Then
                If ilRptNo < 4 Then   '3 reconcile reports without vehicle group sorting
                    ilVehicleGroup = 0
                Else
                    ilVehicleGroup = tgSpf.iReconcGroupNo
                End If

                If ilRptNo = 1 Or ilRptNo = 4 Then      '1-20-04
                    'If Not gOpenPrtJob("TranSum.Rpt") Then
                    If Not gOpenPrtJob("Recon1.Rpt") Then
                        mGenReport = False
                        Exit Function
                    End If
                    Screen.MousePointer = vbHourglass
                    hmRvf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
                    ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                    If ilRet <> BTRV_ERR_NONE Then
                        mGenReport = False
                        Exit Function
                    End If
                    gReconcTotals hmRvf
                    ilRet = btrClose(hmRvf)
                    btrDestroy hmRvf
                    Screen.MousePointer = vbDefault
                'ElseIf ilRptNo = 2 Then
                ElseIf ilRptNo = 2 Or ilRptNo = 5 Then      '1-20-04
                    'If Not gOpenPrtJob("Reconcil.Rpt") Then
                    If Not gOpenPrtJob("Recon2.Rpt") Then
                        mGenReport = False
                        Exit Function
                    End If
                'ElseIf ilRptNo = 3 Then
                ElseIf ilRptNo = 3 Or ilRptNo = 6 Then      '1-20-04
                    If Not gOpenPrtJob("Recon3.Rpt") Then
                        mGenReport = False
                        Exit Function
                    End If
                End If
                'If ilRptNo = 2 Then
                If ilRptNo = 2 Or ilRptNo = 5 Then      '1-20-04
                    gPDNToStr tgSpf.sRB, 2, slStr
                    'If Not gSetFormula("Balance", "'" & slStr & "'") Then
                    If Not gSetFormula("Last Month Total", slStr) Then
                        mTerminate False
                        Exit Function
                    End If
                    'Total through current month- Rvf scan
                    If Not gSetFormula("Current Month Total", sgThisMonthsClosing) Then
                        mTerminate False
                        Exit Function
                    End If
               'ElseIf ilRptNo = 3 Then
               ElseIf ilRptNo = 3 Or ilRptNo = 6 Then       '1-20-04
                    gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slDate
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("Future Start", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        mGenReport = False
                        Exit Function
                    End If
                    gUnpackDate tgSpf.iRNRP(0), tgSpf.iRNRP(1), slDate
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("Future End", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        mGenReport = False
                        Exit Function
                    End If
                End If
                'If (ilRptNo = 1) Or (ilRptNo = 2) Then
                If ilRptNo = 1 Or ilRptNo = 2 Or ilRptNo = 4 Or ilRptNo = 5 Then        '1-20-04
                    'pass closing start & end dates
                    gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slDate
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("CloseStart", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        mGenReport = False
                        Exit Function
                    End If
                    gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slDate
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    If Not gSetFormula("CloseEnd", "Date(" & slYear & "," & slMonth & "," & slDay & ")") Then
                        mGenReport = False
                        Exit Function
                    End If
                    gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slDate
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    slSelection = "{RVF_Receivables.rvfTranDate} > Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slDate
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    slSelection = slSelection & " And {RVF_Receivables.rvfTranDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                Else
                    gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slDate
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    slSelection = "{RVF_Receivables.rvfTranDate} > Date(" & slYear & "," & slMonth & "," & slDay & ")"
                    gUnpackDate tgSpf.iRNRP(0), tgSpf.iRNRP(1), slDate
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    slSelection = slSelection & " And {RVF_Receivables.rvfTranDate} <= Date(" & slYear & "," & slMonth & "," & slDay & ")"
                End If

                'Which option to sort the reconcile report (with vehicle group or without)
                If Not gSetFormula("SelectedVehGroup", ilVehicleGroup) Then
                    mTerminate False
                    Exit Function
                End If
                slSelection = slSelection & " And {RVF_Receivables.rvfCashTrade} = 'C'"
                If Not gSetSelection(slSelection) Then
                    mGenReport = False
                    Exit Function
                End If
            End If
    End Select
    mGenReport = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitGenReport                  *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize Vehicle report     *
'*                      screen                         *
'*                                                     *
'*******************************************************
Private Sub mInitGenReport()
    Dim ilRet As Integer
    Screen.MousePointer = vbHourglass
    ilRet = PEOpenEngine()
    If ilRet = 0 Then
        MsgBox "Unable to open print engine"
        mTerminate False
        Exit Sub
    End If
    Select Case igRptCallType
        Case VEHICLEGROUPSLIST    'Vehicle Groups
            RptNoSel.Caption = "Vehicle Groups " & RptNoSel.Caption
        Case POTENTIALCODESLIST    'Potential Codes
            RptNoSel.Caption = "Potential Codes " & RptNoSel.Caption
        Case BUSCATEGORIESLIST    'Business Categories
            RptNoSel.Caption = "Business Categories " & RptNoSel.Caption
        Case DEMOSLIST      'Customized Demos
            RptNoSel.Caption = "Demos List " & RptNoSel.Caption
        Case COMPETITORSLIST    'Competitors
            RptNoSel.Caption = "Competitors " & RptNoSel.Caption
        Case ITEMBILLINGTYPESLIST    'Item Billing
            RptNoSel.Caption = "Item Billing Types " & RptNoSel.Caption
        Case INVOICESORTLIST    'Invoice sort
            RptNoSel.Caption = "Invoice Sorts " & RptNoSel.Caption
        Case EXCLUSIONSLIST
            RptNoSel.Caption = "Program Exclusions " & RptNoSel.Caption
        Case ANNOUNCERNAMESLIST    'Announcer
            RptNoSel.Caption = "Announcer Names " & RptNoSel.Caption
        Case GENRESLIST    'Genre
            RptNoSel.Caption = "Genre Names " & RptNoSel.Caption
        Case SALESREGIONSLIST
            RptNoSel.Caption = "Sales Regions " & RptNoSel.Caption
        Case SALESSOURCESLIST    'Sales Source
            RptNoSel.Caption = "Sales Sources " & RptNoSel.Caption
        Case SALESTEAMSLIST    'Sales Team
            RptNoSel.Caption = "Sales Teams " & RptNoSel.Caption
        Case REVENUESETSLIST    'Revenue Sets
            RptNoSel.Caption = "Revenue Sets " & RptNoSel.Caption
        Case BOILERPLATESLIST
            RptNoSel.Caption = "Boilerplate " & RptNoSel.Caption
        Case MISSEDREASONSLIST    'Missed Reason
            RptNoSel.Caption = "Missed Reasons " & RptNoSel.Caption
        Case COMPETITIVESLIST    'Competitive
            RptNoSel.Caption = "Product Protection " & RptNoSel.Caption
        Case FEEDTYPESLIST    'Network
            RptNoSel.Caption = "Feed Types " & RptNoSel.Caption
        Case EVENTTYPESLIST     'Event Type
            RptNoSel.Caption = "Event Types " & RptNoSel.Caption
        Case AVAILNAMESLIST
            RptNoSel.Caption = "Avail Names " & RptNoSel.Caption
        Case SALESOFFICESLIST
            RptNoSel.Caption = "Sales Offices " & RptNoSel.Caption
        Case MEDIADEFINITIONSLIST
            RptNoSel.Caption = "Media Definitions " & RptNoSel.Caption
        Case LOCKBOXESLIST
            RptNoSel.Caption = "Lock Boxes " & RptNoSel.Caption
        Case EDISERVICESLIST
            RptNoSel.Caption = "Agency DP Services " & RptNoSel.Caption
        Case TRANSACTIONSLIST
            RptNoSel.Caption = "Transaction Types " & RptNoSel.Caption
        Case SITELIST
            RptNoSel.Caption = "Site Options " & RptNoSel.Caption
        Case USERLIST
            RptNoSel.Caption = "User Options " & RptNoSel.Caption
        Case RATECARDSJOB
            RptNoSel.Caption = "Rate Cards " & RptNoSel.Caption
        Case COLLECTIONSJOB
            If igRptType = 1 Then
                RptNoSel.Caption = "Credit Status " & RptNoSel.Caption
            Else
                RptNoSel.Caption = "Reconcile Report"   '3-16-04 re-entrant causes duplicate caption
            End If
    End Select
    'cbcWhenDay.AddItem "One Time"
    'cbcWhenDay.AddItem "Every M-F"
    'cbcWhenDay.AddItem "Every M-Sa"
    'cbcWhenDay.AddItem "Every M-Su"
    'cbcWhenDay.AddItem "Every Monday"
    'cbcWhenDay.AddItem "Every Tuesday"
    'cbcWhenDay.AddItem "Every Wednesday"
    'cbcWhenDay.AddItem "Every Thursday"
    'cbcWhenDay.AddItem "Every Friday"
    'cbcWhenDay.AddItem "Every Saturday"
    'cbcWhenDay.AddItem "Every Sunday"
    'cbcWhenDay.AddItem "Cal Month End+1"
    'cbcWhenDay.AddItem "Cal Month End+2"
    'cbcWhenDay.AddItem "Cal Month End+3"
    'cbcWhenDay.AddItem "Cal Month End+4"
    'cbcWhenDay.AddItem "Cal Month End+5"
    'cbcWhenDay.AddItem "Std Month End+1"
    'cbcWhenDay.AddItem "Std Month End+2"
    'cbcWhenDay.AddItem "Std Month End+3"
    'cbcWhenDay.AddItem "Std Month End+4"
    'cbcWhenDay.AddItem "Std Month End+5"
    'cbcWhenDay.ListIndex = 0
    'cbcWhenTime.AddItem "Right Now"
    'cbcWhenTime.AddItem "at 10PM"
    'cbcWhenTime.AddItem "at 12AM"
    'cbcWhenTime.AddItem "at 2AM"
    'cbcWhenTime.AddItem "at 4AM"
    'cbcWhenTime.AddItem "at 6AM"
    'cbcWhenTime.ListIndex = 0

    '10-19-01
    gPopExportTypes cbcFileType
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"

    cbcFileType.ListIndex = 0
    mSetCommands
    gCenterStdAlone RptNoSel
    Screen.MousePointer = vbDefault
'    gCenterModalForm RptNoSel
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
    Dim slRptListCmmd As String
    slCommand = sgCommandStr    'Command$
    ilRet = gParseItem(slCommand, 1, "||", smCommand)
    If (ilRet <> CP_MSG_NONE) Then
        smCommand = slCommand
    End If
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False
    '    imShowHelpMsg = False
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
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpMsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone RptNoSel, slStr, ilTestSystem
    ''ilRet = gParseItem(slCommand, 3, "\", slStr)
    ''igRptCallType = Val(slStr)

    'D.S. 10/09/01 All report names and types used in this module
    '"Agency DP Services"           EDISERVICESLIST
    '"Announcer Names"              ANNOUNCERNAMESLIST
    '"Avail Names"                  AVAILNAMESLIST
    '"Boilerplate"                  BOILERPLATESLIST
    '"Business Categories"          BUSCATEGORIESLIST
    '"Competitors"                  COMPETITORSLIST
    '"Demos List"                   DEMOSLIST
    '"Event Types"                  EVENTTYPESLIST
    '"Feed Types"                   FEEDTYPESLIST
    '"Genre Names"                  GENRESLIST
    '"Invoice Sorts"                INVOICESORTLIST
    '"Item Billing Types"           ITEMBILLINGTYPESLIST
    '"Lock Boxes"                   LOCKBOXESLIST
    '"Media Definitions"            MEDIADEFINITIONSLIST
    '"Missed Reasons"               MISSEDREASONSLIST
    '"Potential Codes"              POTENTIALCODESLIST
    '"Product Protection"           COMPETITIVESLIST
    '"Program Exclusions"           EXCLUSIONSLIST
    '"Reconcile"                    COLLECTIONSJOB
    '"Revenue Sets"                 REVENUESETSLIST
    '"Sales Offices"                SALESOFFICESLIST
    '"Sales Regions"                SALESREGIONSLIST
    '"Sales Sources"                SALESSOURCESLIST
    '"Sales Teams"                  SALESTEAMSLIST
    '"Site Options"                 SITELIST
    '"Transaction Types"            TRANSACTIONSLIST
    '"User Options"                 USERLIST
    '"Vehicle Groups"               VEHICLEGROUPSLIST

    'If igStdAloneMode Then
    '    smSelectedRptName = "Vehicle Groups"
    '    igRptCallType = VEHICLEGROUPSLIST
    '    slCommand = " x\x\x\0\2\10/17/94" '6=Comtempory
    'Else
        ilRet = gParseItem(slCommand, 2, "||", slRptListCmmd)
        ilRet = gParseItem(slRptListCmmd, 2, "\", smSelectedRptName)
        ilRet = gParseItem(slRptListCmmd, 1, "\", slStr)
        igRptCallType = Val(slStr)
    'End If
    If igRptCallType = LOGSJOB Then
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        igRptType = Val(slStr)  '0-One day log; 1=Seven day log; 2=Delivery
        ilRet = gParseItem(slCommand, 5, "\", smUserCode)
        ilRet = gParseItem(slCommand, 6, "\", smStartDate)
    End If
    If igRptCallType = COLLECTIONSJOB Then
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            igRptType = Val(slStr)  '0-Reconcil set; 1=Credit Status
        Else
            igRptType = 0
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/21/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    Dim ilEnable As Integer
    ilEnable = True
    If ilEnable Then
        If rbcOutput(0).Value Then  'Display
            ilEnable = True
        ElseIf rbcOutput(1).Value Then  'Print
            If edcCopies.Text <> "" Then
                ilEnable = True
            Else
                ilEnable = False
            End If
        Else    'Save As
            If (imFTSelectedIndex >= 0) And (edcFileName.Text <> "") Then
                ilEnable = True
            Else
                ilEnable = False
            End If
        End If
    End If
    cmcDone.Enabled = ilEnable
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate(ilFromCancel As Integer)
'
'   mTerminate
'   Where:
'
    If ilFromCancel Then
        igRptReturn = True
    Else
        igRptReturn = False
    End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload RptNoSel
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub rbcOutput_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcOutput(Index).Value
    'End of coded added
    If rbcOutput(Index).Value Then
        Select Case Index
            Case 0  'Display
                frcFile.Enabled = False
                frcFile.Visible = False
                frcCopies.Enabled = False
                frcCopies.Visible = False
                'frcWhen.Enabled = False
                'pbcWhen.Visible = False
                'mSendHelpMsg "Report sent to display"
            Case 1  'Print
                frcFile.Enabled = False
                frcFile.Visible = False
                frcCopies.Enabled = True
                frcCopies.Visible = True
                'frcWhen.Enabled = True
                'pbcWhen.Visible = True
                'mSendHelpMsg "Report sent to printer"
            Case 2  'File
                frcCopies.Enabled = False
                frcCopies.Visible = False
                frcFile.Enabled = True
                frcFile.Visible = True
                'frcWhen.Enabled = True
                'pbcWhen.Visible = True
                'mSendHelpMsg "Report sent to file"
        End Select
    End If
    mSetCommands
End Sub
Private Sub rbcOutput_GotFocus(Index As Integer)
    If imFirstTime Then
        imFirstTime = False
    End If
    gCtrlGotFocus rbcOutput(Index)
End Sub
Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
