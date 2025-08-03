VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelBR 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   720
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
   ScaleHeight     =   5505
   ScaleWidth      =   9270
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   20
      Top             =   615
      Width           =   2055
   End
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6690
      TabIndex        =   24
      Top             =   -90
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Timer tmcDone 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8835
      Top             =   -150
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
      Left            =   8025
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   -75
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   -90
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.FileListBox lbcFileName 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   3075
      Pattern         =   "*.Dal"
      TabIndex        =   18
      Top             =   4935
      Visible         =   0   'False
      Width           =   1125
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
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4245
      Width           =   90
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   1170
      Top             =   4830
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".pdf"
      Filter          =   $"Rptselbr.frx":0000
      FilterIndex     =   2
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
         Left            =   1095
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "1"
         Top             =   315
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   165
         TabIndex        =   7
         Top             =   810
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   165
         TabIndex        =   5
         Top             =   345
         Width           =   855
      End
   End
   Begin VB.Frame frcFile 
      Caption         =   "Save to File"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2070
      TabIndex        =   8
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   960
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
         TabIndex        =   10
         Top             =   270
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
         TabIndex        =   12
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Vehicle Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3690
      Left            =   90
      TabIndex        =   14
      Top             =   1755
      Width           =   9090
      Begin VB.PictureBox pbcSelC 
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
         Height          =   3360
         Left            =   90
         ScaleHeight     =   3360
         ScaleWidth      =   4530
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   255
         Visible         =   0   'False
         Width           =   4530
      End
      Begin VB.PictureBox pbcSelB 
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
         Height          =   990
         Left            =   135
         ScaleHeight     =   990
         ScaleWidth      =   4425
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   330
         Visible         =   0   'False
         Width           =   4425
      End
      Begin VB.PictureBox pbcOption 
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
         Height          =   3420
         Left            =   4605
         ScaleHeight     =   3420
         ScaleWidth      =   4455
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   21
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   19
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
         Top             =   840
         Width           =   1275
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Print"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   555
         Width           =   750
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Display"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   885
      End
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
Attribute VB_Name = "RptSelBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselbr.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelBR.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim imTerminate As Integer
Dim hmCbf As Integer        'Contract prepass file handle
Dim imCbfRecLen As Integer  'CBF record length
Dim tmCbf As CBF            'CBF record image
Dim tmCbfSrchKey As CBFKEY0     'Gen date and time
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
    If igGenRpt Then
        Exit Sub
    End If
    mTerminate False
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcGen_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilSBFFound                                                                            *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilListIndex As Integer
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilDoIt As Integer
    Dim ilFilterBRPass As Integer
    
    If igGenRpt Then
        Exit Sub
    End If
    igGenRpt = True
    igOutput = frcOutput.Enabled
    igCopies = frcCopies.Enabled
    igFile = frcFile.Enabled
    igOption = frcOption.Enabled
    frcOutput.Enabled = False
    frcCopies.Enabled = False
    frcFile.Enabled = False
    frcOption.Enabled = False
    igUsingCrystal = True               'all versions of printed contract uses crystal
    'If (sgInclResearch = "N" And sgInclRates = "N") Then     'force to show at least rates, doesnt make sense to see neither rates nor research
    '    sgInclRates = "Y"
    'End If
'    If igBR_SSLinesExist Then                   'should output be forced because at least 1 line exists for detail/summary
'        If igDetSumBoth = 2 Then               ' both
'            'See if the rates are included for the monthly version
'            If sgInclResearch = "Y" And sgInclRates = "Y" Then
'                ilStartJobNo = 1
'                ilNoJobs = 3
'            ElseIf sgInclResearch = "Y" And sgInclRates = "N" Then  '12-13-05
'                ilStartJobNo = 1
'                ilNoJobs = 3
'            Else
'                ilStartJobNo = 1
'                ilNoJobs = 2
'            End If
'        ElseIf igDetSumBoth = 0 Then            'detail only
'            ilStartJobNo = 1
'            ilNoJobs = 1
'        Else                                     'summary only
'            If sgInclResearch = "Y" And sgInclRates = "Y" Then
'                ilStartJobNo = 2
'                ilNoJobs = 3
'            ElseIf sgInclResearch = "Y" And sgInclRates = "N" Then      '12-13-05 always see 2 summaries if Research included
'                ilStartJobNo = 2
'                ilNoJobs = 3
'            Else
'                ilStartJobNo = 2
'                ilNoJobs = 2
'            End If
'        End If
 'Dan change to multi reports.  11/19/08
'get last job number.  If want detail only, only one report, which is default
        Set ogReport = New CReportHelper
        If igDetSumBoth <> 0 Then
'            ogReport.iLastPrintJob = 4
           ogReport.iLastPrintJob = 5              '12-22-20  add cpm line IDs
        End If
        'always assume all 4 version of the contract should be printed
        ilStartJobNo = 1
'        ilNoJobs = 4
        ilNoJobs = 5                    '12-22-20 add another job: cpm line IDs
        For ilJobs = ilStartJobNo To ilNoJobs Step 1
            ' because loops 4 times even if only 1 report, will cause error
            If ogReport Is Nothing Then
                Exit For
            End If
             igJobRptNo = ilJobs
             ilDoIt = True
             'ilJobs :  1 = detail, 2 = NTR, 3 = Research Summary with or without rates, 4 = Billing summary (12 month $)
             'ilJobs 1-7-21 Added CPM Podcast as job #3,others adjusted :  1 = detail, 2 = NTR, 3=CPM Podcast, 4 = Research Summary with or without rates, 5 = Billing summary (12 month $)
             If ilJobs = 1 Then          'detail
                 'has detail been requested
                 If (igDetSumBoth <> 1) Then           '1 = summary only
                     ilDoIt = True
                 Else
                     ilDoIt = False
                 End If
             ElseIf ilJobs = 2 Then      'ntr
                 'any NTR?
                 If (igDetSumBoth <> 0) And (tgChf.sNTRDefined = "Y") Then            '0 = detail only: user requested summary or both, and NTR defined
                     ilDoIt = True
                 Else
                     ilDoIt = False
                 End If
             ElseIf ilJobs = 3 Then      'CPM
                 'any CPM?
                 If (igDetSumBoth <> 0) And (tgChf.sAdServerDefined = "Y") Then            '0 = detail only: user requested summary or both, and CPM defined
                     ilDoIt = True
                 Else
                     ilDoIt = False
                 End If
'             ElseIf ilJobs = 3 Then        'summary w/research
             ElseIf ilJobs = 4 Then        'summary w/research
                 If igDetSumBoth <> 0 And sgInclResearch = "Y" Then   'user requested summary or both w/ research
                        ilDoIt = True
                    Else
                        ilDoIt = False
                    End If
             Else                        ' or billing summary
                 If igDetSumBoth <> 0 Then             'user requested summary or both, and NTR defined
                        ilDoIt = True
                    Else
                        ilDoIt = False
                    End If
             End If
             Screen.MousePointer = vbHourglass
             If ilDoIt Then
                 If Not gGenReportBr() Then
                      igGenRpt = False
                      frcOutput.Enabled = igOutput
                      frcCopies.Enabled = igCopies
                      frcFile.Enabled = igFile
                      frcOption.Enabled = igOption
                      Exit Sub
                  End If
                  ilRet = gCmcGenBr(ilListIndex, imGenShiftKey)
                  '-1 is a Crystal failure of gSetSelection or gSEtFormula
                  If ilRet = -1 Then
                      igGenRpt = False
                      frcOutput.Enabled = igOutput
                      frcCopies.Enabled = igCopies
                      'frcWhen.Enabled = igWhen
                      frcFile.Enabled = igFile
                      frcOption.Enabled = igOption
                      'frcRptType.Enabled = igReportType
                      'mTerminate
                      'JW - 4/5/22 - possible fix for 10416 - Prevent Error when focusing an invsibile or disabled item
                        If (pbcClickFocus.Enabled) And (pbcClickFocus.Visible) Then
                            pbcClickFocus.SetFocus
                        End If
                      tmcDone.Enabled = True
                      Exit Sub
                  ElseIf ilRet = 0 Then           '0 = invalid input data, stay in
                      igGenRpt = False
                      frcOutput.Enabled = igOutput
                      frcCopies.Enabled = igCopies
                      'frcWhen.Enabled = igWhen
                      frcFile.Enabled = igFile
                      frcOption.Enabled = igOption
                      'frcRptType.Enabled = igReportType
                      Exit Sub
                  ElseIf ilRet = 2 Then           'successful from Bridgereport
                      igGenRpt = False
                      frcOutput.Enabled = igOutput
                      frcCopies.Enabled = igCopies
                      'frcWhen.Enabled = igWhen
                      frcFile.Enabled = igFile
                      frcOption.Enabled = igOption
                      'frcRptType.Enabled = igReportType
                      'JW - 4/5/22 - possible fix for 10416 - Prevent Error when focusing an invsibile or disabled item
                        If (pbcClickFocus.Enabled) And (pbcClickFocus.Visible) Then
                            pbcClickFocus.SetFocus
                        End If
                      tmcDone.Enabled = True
                      Exit Sub
                 End If

                '1 falls thru - successful crystal report


                 'Setup correct logo to print for this vehicles log or CP
                 'Rename existing rptlogo.bmp to vehicles logo; then rename back later
                 'If ilVpfIndex >= 0 Then
                 '    If tgvpf(ilVpfIndex).scplogo <> "   " Then
                         'Rename the original rptlogo.bmp to a saved name, then name the vehicle logo to rptlogo.bmp for crystal reporting
                         'slSaveRptLogoName = Trim$(sgRptPath) & Trim$("rptlogo.bmp")
                         'Name slSaveRptLogoName As Trim$(sgRptPath) & "savelogo.bmp"
                         'slVehicleLogo = Trim$(sgRptPath) & "G" & Trim$(tgvpf(ilVpfIndex).scplogo) & Trim$(".bmp")
                         'Name Trim$(slVehicleLogo) As Trim$(sgRptPath) & Trim$("rptlogo.bmp")
                 '    End If
                 'End If
            'Dan M change to multi report: replace old if statement and 'hijack' output until last report
            ' 6 12 09 Dan M commented out below to remove multi reporting. uncommented if/else below : commented out "if iljobs=..."
                 
                 '2-4-10 comment out, changed method due to merging ntr with air time and different combinations of contract report to print
'                 If (ilJobs <> 2 And Not igBR_SSLinesExist) Or (ilJobs = 2 And tgChf.sNTRDefined <> "Y") Then
''                            ogReport.Reports.Remove (ogReport.CurrentReportName)    'do I have to worry about if report is really there?
'                     ogReport.RemoveReport
'                 End If
                 
                '12-4-10 tests to prevent a blank page from printing
                ilFilterBRPass = False
                If ilJobs = 1 Then      'pass 1 : always detail
                    If igBR_SSLinesExist = True Then       'print the detail if at least 1 sche lines exists in one or more contracts
                        ilFilterBRPass = True
                    End If
                ElseIf ilJobs = 2 Then                               'pass 2: always NTR
                    If tgChf.sNTRDefined = "Y" Then                  'any NTR defined at all?
                        ilFilterBRPass = True
                    End If
                ElseIf ilJobs = 3 Then                               'pass 3: always CPM
                    If tgChf.sAdServerDefined = "Y" Then                  'any Podcast cpm defined at all?
                        ilFilterBRPass = True
                    End If
                
                Else
                    'Research summary or bill summary
                    'if combining air time and NTR and at least one NTR defined, print it
                    'if at least one sch line exists, print it
                    If igBRSumZer Then      'billing summary, ok to show this if no schedule lines exist
                        If (sgInclNTRBillSummary = "Y" And tgChf.sNTRDefined = "Y") Or (sgInclNTRBillSummary = "Y" And tgChf.sAdServerDefined = "Y") Or igBR_SSLinesExist = True Then
                            ilFilterBRPass = True
                        End If
                    Else                    'research summary & rates, only show if at least one air time exists
                        'If (sgInclResearch = "Y" And sgInclRates = "Y" And igBR_SSLinesExist = True) Then
                        '2-22-10 show some form of research whether with or without reates
                        If (sgInclResearch = "Y" And igBR_SSLinesExist = True) Or (sgInclResearch = "Y" And tgChf.sAdServerDefined = "Y") Then
                              ilFilterBRPass = True
                        End If
                    End If
                End If
                'if passed all the contract parameters with NTR and airtime, should this pass be printed
                If (ilFilterBRPass = False) Or (ilDoIt = False) Then
                    ilFilterBRPass = False
                End If
                
                If Not ilFilterBRPass Then
                    ogReport.RemoveReport
                End If

                 'If (ilDoIt = False) Or (ilJobs <> 2 And Not igBR_SSLinesExist) Or (ilJobs = 2 And tgChf.sNTRDefined <> "Y") Then
                 '10-31-07 Do nothing if the flag is set to ignore processing, OR
                 'Do nothing if its the BR (not NTR) and there are no schedule lines to print on detail or summary; OR
                 'Do nothing if its the NTR pass and theres no NTR or summaries not requested
                ' Else
                        

                
                 If ilJobs = ogReport.iLastPrintJob Then
                     If rbcOutput(0).Value Then 'Display
                        'TTP 10549 - Learfield Cloud printing 911, Crystal Crashes, Use Adobe
                        If bgUseAdobe = True And sgReportFilename <> "" Then
                            slFileName = sgReportFilename
                            If slFileName = "" Then slFileName = "BR"
                            '12-16-03 alter filenames based on which contract version (detail, summary notation: up to 4 passes)
                            If InStr(slFileName, ".") = 0 Then  'no extension specified
                                If bgUseAdobe = True Then
                                    slFileName = Trim(slFileName)
                                Else
                                    slFileName = Trim(slFileName) & "-" & Trim$(str(igJobRptNo))
                                End If
                            Else
                                'name already has extension, need to insert the contract version (detail, summary notation: up to 4passes)
                                ilLoop = InStr(slFileName, ".")     'find the period before extension name
                                slStr = Trim$(Mid$(slFileName, 1, ilLoop - 1)) & "-" & Trim$(str(igJobRptNo)) & Trim$(Mid(slFileName, ilLoop))
                                slFileName = slStr
                            End If
                            ilRet = gExportCRW(slFileName, imFTSelectedIndex, False, sgReportTempFolder)   '10-20-01

                        Else
                             igDestination = 0
                             DoEvents            '9-19-02 fix for timing problem to prevent freezing before calling crystal
                             Screen.MousePointer = vbDefault
                            ' If ogReport.ReportCount > 0 Then
                                 Report.Show vbModal
                           '  End If
                        End If
                        
                     ElseIf rbcOutput(1).Value Then 'Print
                         ilCopies = Val(edcCopies.Text)
                         ilRet = gOutputToPrinter(ilCopies)
                     
                     Else 'Save to File
                        slFileName = edcFileName.Text
                         '1/25/10 Dan no longer need to add "-x" after export
                         '12-16-03 alter filenames based on which contract version (detail, summary notation: up to 4 passes)
'                         If InStr(slFileName, ".") = 0 Then  'no extension specified
'                             slFileName = Trim(slFileName) & "-" & Trim$(str(igJobRptNo))
'                         Else
'                             'name already has extension, need to insert the contract version (detail, summary notation: up to 4passes)
'                             ilLoop = InStr(slFileName, ".")     'find the period before extension name
'                             slStr = Trim$(Mid$(slFileName, 1, ilLoop - 1)) & "-" & Trim$(str(igJobRptNo)) & Trim$(Mid(slFileName, ilLoop))
'                             slFileName = slStr
'                         End If
                        ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
                     End If
                 End If
             End If
           ' End If  'did user want to cancel display early?
        Next ilJobs
        Set ogReport = Nothing
        Screen.MousePointer = vbDefault
#If programmatic = 1 Then
        Exit Sub
#End If
    'End If
    '
    'Rename the vehicles logo back to rptlogo.bmp
    'Only rename back to rptlog.bmp if a valid vehicle option found
    'If (ilVpfIndex >= 0) And (tgvpf(ilVpfIndex).scplogo <> "   ") Then
        'Name Trim$(slVehicleLogo) As Trim$(sgRptPath) & "G" & Trim$(tgvpf(ilVpfIndex).scplogo) & Trim$(".bmp")
        'Name Trim$(slSaveRptLogoName) As Trim$(sgRptPath) & Trim$("rptlogo.bmp")
    'End If


    imGenShiftKey = 0

    '12-9-02 for NTR generation only, may or maynot be present
'    ilSBFFound = False
'    For ilLoop = LBound(tgIBSbf) To UBound(tgIBSbf) - 1
'        If tgIBSbf(ilLoop).iStatus >= 0 Then
'            ilSBFFound = True
'            Exit For
'        End If
'    Next ilLoop
'    If ilSBFFound Then
'        igJobRptNo = 4              'force to process ntrs
'        If Not gGenReportBr() Then
'            igGenRpt = False
'            frcOutput.Enabled = igOutput
'            frcCopies.Enabled = igCopies
'            frcFile.Enabled = igFile
'            frcOption.Enabled = igOption
'            Exit Sub
'        End If
'        ilRet = gCmcGenBr(ilListIndex, imGenShiftKey)
'        '-1 is a Crystal failure of gSetSelection or gSEtFormula
'        If ilRet = -1 Then
'            igGenRpt = False
'            frcOutput.Enabled = igOutput
'            frcCopies.Enabled = igCopies
'            'frcWhen.Enabled = igWhen
'            frcFile.Enabled = igFile
'            frcOption.Enabled = igOption
'            'frcRptType.Enabled = igReportType
'            'mTerminate
'            pbcClickFocus.SetFocus
'            tmcDone.Enabled = True
'            Exit Sub
'        ElseIf ilRet = 0 Then           '0 = invalid input data, stay in
'            igGenRpt = False
'            frcOutput.Enabled = igOutput
'            frcCopies.Enabled = igCopies
'            'frcWhen.Enabled = igWhen
'            frcFile.Enabled = igFile
'            frcOption.Enabled = igOption
'            'frcRptType.Enabled = igReportType
'            Exit Sub
'        ElseIf ilRet = 2 Then           'successful from Bridgereport
'            igGenRpt = False
'            frcOutput.Enabled = igOutput
'            frcCopies.Enabled = igCopies
'            'frcWhen.Enabled = igWhen
'            frcFile.Enabled = igFile
'            frcOption.Enabled = igOption
'            'frcRptType.Enabled = igReportType
'            pbcClickFocus.SetFocus
'            tmcDone.Enabled = True
'            Exit Sub
'       End If
'        '1 falls thru - successful crystal report
'        If rbcOutput(0).Value Then
'            DoEvents            '9-13-02 fix for timing problem with Avails report & Spot Business Booked (random problem)
'            igDestination = 0
'            Report.Show vbModal
'        ElseIf rbcOutput(1).Value Then
'            ilCopies = Val(edcCopies.Text)
'            ilRet = gOutputToPrinter(ilCopies)
'        Else
'            slFileName = edcFileName.Text
'            '12-16-03 alter filenames based on which contract version (detail, summary notation: up to 4 passes)
'            If InStr(slFileName, ".") = 0 Then  'no extension specified
'                slFileName = Trim(slFileName) & "-" & Trim$(Str(igJobRptNo))
'            Else
'                'name already has extension, need to insert the contract version (detail, summary notation: up to 4passes)
'                ilLoop = InStr(slFileName, ".")     'find the period before extension name
'                slStr = Trim$(Mid$(slFileName, 1, ilLoop - 1)) & "-" & Trim$(Str(igJobRptNo)) & Trim$(Mid(slFileName, ilLoop))
'                slFileName = slStr
'            End If
'            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
'        End If
'        imGenShiftKey = 0
'    End If

    Screen.MousePointer = vbHourglass
    'gClearCbf
    
    '2-16-13  clear the temporary file with the user ID along with gendate and time
    gCBFClearWithUserID
    
    gClearTxr
    Screen.MousePointer = vbDefault
    igGenRpt = False
    
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
    'frcWhen.Enabled = igWhen
    frcFile.Enabled = igFile
    frcOption.Enabled = igOption
    'JW - 4/5/22 - possible fix for 10416 - Prevent Error when focusing an invsibile or disabled item
    If (pbcClickFocus.Enabled) And (pbcClickFocus.Visible) Then
        pbcClickFocus.SetFocus
    End If
    tmcDone.Enabled = True
    Exit Sub
'cmcGenErr:
'    ilDDFSet = True
'    Resume Next
End Sub
Private Sub cmcGen_GotFocus()
    gCtrlGotFocus cmcGen
End Sub
Private Sub cmcGen_KeyDown(KeyCode As Integer, Shift As Integer)
    imGenShiftKey = Shift
End Sub
Private Sub cmcList_Click()
    If igGenRpt Then
        Exit Sub
    End If
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
    If ((KeyAscii <> KEYBACKSPACE) And (KeyAscii <= 32)) Or (KeyAscii = 34) Or (KeyAscii = 39) Or ((KeyAscii >= KeyDown) And (KeyAscii <= 45)) Or ((KeyAscii >= 59) And (KeyAscii <= 63)) Then
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
End Sub

Private Sub Form_Click()
#If programmatic = 1 Then
    'JW - 4/5/22 - possible fix for 10416 - Prevent Error when focusing an invsibile or disabled item
    If (pbcClickFocus.Enabled) And (pbcClickFocus.Visible) Then
        pbcClickFocus.SetFocus
    End If
#End If
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
    igGenRpt = False
    mParseCmmdLine
#If programmatic = 1 Then
    mInitReport
#End If
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    mInit
    If imTerminate = -99 Then
        Exit Sub
    End If
    If imTerminate Then 'Used for print only
        'mTerminate
        cmcCancel_Click
        Exit Sub
    End If
    'RptSelBR.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Erase tgCSVNameCode
    'Erase tgSellNameCode
    Erase tgRptSelBRSalespersonCode
    Erase tgRptSelBRAgencyCode
    Erase tgRptSelBRAdvertiserCode
    Erase tgRptSelBRNameCode
    Erase tgRptSelBRBudgetCode
    'Erase tgMultiCntrCode
    'Erase tgManyCntCode
    Erase tgRptSelBRDemoCode
    'Erase tgSOCode
    PECloseEngine

    
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set RptSelBR = Nothing   'Remove data segment

End Sub
Private Sub gClearCbf()
'*******************************************************
'*                                                     *
'*      Procedure Name:Clear Prepass file for Printed
'*                  contract
'*                                                     *
'*             Created:04/18/96      By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*                                                     *
'*******************************************************
    Dim ilRet As Integer
    hmCbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCbf)
        btrDestroy hmCbf
        Exit Sub
    End If
    imCbfRecLen = Len(tmCbf)
    tmCbfSrchKey.iGenDate(0) = igNowDate(0)
    tmCbfSrchKey.iGenDate(1) = igNowDate(1)
    'tmCbfSrchKey.iGenTime(0) = igNowTime(0)
    'tmCbfSrchKey.iGenTime(1) = igNowTime(1)
    tmCbfSrchKey.lGenTime = lgNowTime
    ilRet = btrGetGreaterOrEqual(hmCbf, tmCbf, imCbfRecLen, tmCbfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
   ' Do While (ilRet = BTRV_ERR_NONE) And (tmCbf.iGenDate(0) = igNowDate(0)) And (tmCbf.iGenDate(1) = igNowDate(1)) And (tmCbf.iGenTime(0) = igNowTime(0)) And (tmCbf.iGenTime(1) = igNowTime(1))
    Do While (ilRet = BTRV_ERR_NONE) And (tmCbf.iGenDate(0) = igNowDate(0)) And (tmCbf.iGenDate(1) = igNowDate(1)) And (tmCbf.lGenTime = lgNowTime)
        ilRet = btrDelete(hmCbf)
        ilRet = btrGetNext(hmCbf, tmCbf, imCbfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmCbf)
    btrDestroy hmCbf
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*                                                     *
'*******************************************************
Private Sub mInit()
    Dim ilRet As Integer

    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    'Start Crystal report engine
    ilRet = PEOpenEngine()
    If ilRet = 0 Then
        MsgBox "Unable to open print engine"
        mTerminate False
        imTerminate = -99
        Exit Sub
    End If

    'Set options for report generate
    'hdJob = rpcRpt.hJob
    'ilMultiTable = True
    'dummy = LlSetOption(hdJob, LL_OPTION_HELPAVAILABLE, False)
    'ilDummy = LlSetOption(hdJob, LL_OPTION_SORTVARIABLES, True)
    'ilDummy = LlSetOption(hdJob, LL_OPTION_ONLYONETABLE, Not ilMultiTable)

    imAllClicked = False
    imSetAll = True
    'gCenterStdAlone RptSelBR
    'RptSelBR.Move -90, -90, 30, 30      'make form small and out of the way so its not seen
    
    'JW - 4/5/22 - BR form still shows a little box in Upper-left corner after clicking generate
    RptSelBR.Move -630, -630, 30, 30      'make form small and out of the way so its not seen
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitReport                     *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*                                                     *
'*******************************************************
Public Sub mInitReport()
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
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    'gPopExportTypes cbcFileType '2-13-04 already populated and type selected
    pbcSelC.Visible = False
    lbcRptType.Clear

    Screen.MousePointer = vbDefault
    'If (igRptType = 0) Or (igRptType = 1) Or (igRptType = 2) Then
    If igOutputTo = 0 Then
        rbcOutput(0).Value = True
    ElseIf igOutputTo = 1 Then
        rbcOutput(1).Value = True          'always print these automatically generated reports
    'rbcOutput(0).Value = True           'display -- for test purposes only
    Else
    End If
    cmcGen_Click
    imTerminate = True
    Exit Sub
    'End If
    mSetCommands
    Screen.MousePointer = vbDefault
'    gCenterModalForm RptSelBR
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*
'*            Special mParseCmmdLine for "Log" process *
'*            Assumes that only the LOG reports come
'*            thru here
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    slCommand = sgCommandStr    'Command$
    ilRet = gParseItem(slCommand, 1, "||", smCommand)
    If (ilRet <> CP_MSG_NONE) Then
        smCommand = slCommand
    End If
    ''igStdAloneMode defined as "Debug" mode
    'igStdAloneMode = True 'Switch from/to stand alone mode-No DDE
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    sgCallAppName = ""
    '    ilTestSystem = False  'True 'False
    '    'Change these following parameters to debug
    '    slCommGenDate = "8/9/99"
    '    slCommGenTime = "10:06:02A"
    '    slDetSumBoth = "2"    '0 = detail, 1 = summary, 2 = both
    '    slInclRates = "N"
    '    slInclResearch = "N"
    '    slInclSplits      Y/N 2-13-04
    '    slInclNTRBillSummary Y/N 2-2-10
    '    slShowNetOnProps Y/N        2-3-10
    '    slShowProdProt  Y/N          8-25-15
    '    slInclProof = "N"
    '    slDisplayPrint = "0"    '0 = display, 1= print
    '    slExportIndex          'index to type of export
    '    slSaveToFileName       'file name of export
    '    'Mandatory parms are 1st and 2nd :
    '    '1st parm:  function coming from (Logs^Test (or Prod or NoHelp)
    '    '2nd parm:  user name
    '    'parms: Logs^Test (or Prod)\ user name\jobcode(igrptcalltype)\Rnfcode (igrpttype)\usercode\Start Date\#days\STartTime\EndTime\VehCode\Zones\DisplPrint
    '    slCommand = "Logs^Test\Guide\" & slCommGenDate & "\" & slCommGenTime & "\" & slDetSumBoth & "\" & slInclRates & "\" & slInclResearch & "\" & slInclSplits & "\" & slInclNTRBillSummary & "\" & slShowNetOnProps & "\" & slShowProdProt & "\" & slInclProof & "\" & slOutputTo & "\" & slExportIndex & "\|" & slSaveToFileName
    
    '    imShowHelpmsg = False
    'Else
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
    '    imShowHelpmsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpmsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone RptSelBR, slStr, ilTestSystem

    ilRet = gParseItem(slCommand, 3, "\", slStr)                'Gen Date
    gPackDate slStr, igNowDate(0), igNowDate(1)
    ilRet = gParseItem(slCommand, 4, "\", slStr)                'Gen Time
    gPackTime slStr, igNowTime(0), igNowTime(1)
    ilRet = gParseItem(slCommand, 5, "\", slStr)                'Detail, Summary or Both
    igDetSumBoth = Val(Trim$(slStr))
    ilRet = gParseItem(slCommand, 6, "\", sgInclRates)       'Include Rates
    ilRet = gParseItem(slCommand, 7, "\", sgInclResearch)          'Include Research
    ilRet = gParseItem(slCommand, 8, "\", sgInclSplits)          '2-13-04 Include splits
    ilRet = gParseItem(slCommand, 9, "\", sgInclNTRBillSummary)  '2-2-10 Include NTR bill summary
    ilRet = gParseItem(slCommand, 10, "\", sgShowNetOnProps)     '2-3-10 Show net amts on proposals
    
    ilRet = gParseItem(slCommand, 11, "\", sgShowProdProt)     '8-25-15 Show Product Protection catgories

    ilRet = gParseItem(slCommand, 12, "\", sgInclProof)       '2-2-10 was index 9, include Proof
    ilRet = gParseItem(slCommand, 13, "\", slStr)            '2-2-10 was index 10, display or print
    igOutputTo = Val(Trim$(slStr))
    ilRet = gPopExportTypes(cbcFileType)  '2-13-04

    If igOutputTo = 0 Then                          'display
        rbcOutput(0).Value = True
        rbcOutput(1).Value = False
    ElseIf igOutputTo = 1 Then                      'print
        rbcOutput(1).Value = True
        rbcOutput(0).Value = False
    Else
        rbcOutput(0).Value = False
        rbcOutput(1).Value = False
        rbcOutput(2).Value = True
        ilRet = gParseItem(slCommand, 14, "\", slStr)               '2-2-10, was index 11,Save File Index
        imFTSelectedIndex = Val(slStr)
        'ilRet = gParseItem(slCommand, 11, "\", slStr)               'Save File Name
        ilRet = gParseItem(slCommand, 2, "\|", slStr)       'save filename
        edcFileName.Text = Trim$(slStr)
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
    cmcGen.Enabled = ilEnable
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
    Unload RptSelBR
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
                frcCopies.Visible = False   'Print Box
                frcFile.Visible = False     'Save to File Box
                frcCopies.Enabled = False
                'frcWhen.Enabled = False
                'pbcWhen.Visible = False
            Case 1  'Print
                frcFile.Visible = False
                frcFile.Enabled = False
                frcCopies.Enabled = True
                'frcWhen.Enabled = False 'True
                'pbcWhen.Visible = False 'True
                frcCopies.Visible = True
            Case 2  'File
                frcCopies.Visible = False
                frcCopies.Enabled = False
                frcFile.Enabled = True
                'frcWhen.Enabled = False 'True
                'pbcWhen.Visible = False 'True
                frcFile.Visible = True
        End Select
    End If
    mSetCommands
End Sub
Private Sub rbcOutput_GotFocus(Index As Integer)
    If imFirstTime Then
        imFirstTime = False
        mInitReport
        If imTerminate Then 'Used for print only
            'mTerminate
            cmcCancel_Click
            Exit Sub
        End If
    End If
    gCtrlGotFocus rbcOutput(Index)
End Sub
Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    mTerminate False
End Sub
