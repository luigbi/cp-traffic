VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelID 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remote Invoice Worksheet"
   ClientHeight    =   5535
   ClientLeft      =   165
   ClientTop       =   1305
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
   ScaleHeight     =   5535
   ScaleWidth      =   9270
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   19
      Top             =   615
      Width           =   2055
   End
   Begin VB.Timer tmcDone 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6675
      Top             =   -180
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
      Left            =   7215
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   -15
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
      Left            =   7575
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   -60
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ScaleWidth      =   30
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
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
         Left            =   1065
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   7
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   5
         Top             =   360
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
         BackColor       =   &H00C0C0C0&
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
         BackColor       =   &H00C0C0C0&
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
      Caption         =   "Remote Invoice Worksheet"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3795
      Left            =   15
      TabIndex        =   14
      Top             =   1680
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
         Left            =   105
         ScaleHeight     =   3360
         ScaleWidth      =   4455
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CheckBox ckcDiscrepOnly 
            Caption         =   "Discrepancy Only"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   32
            Top             =   825
            Width           =   1920
         End
         Begin VB.TextBox edcContract 
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
            MaxLength       =   8
            TabIndex        =   30
            Top             =   435
            Width           =   975
         End
         Begin VB.TextBox edcSelCFrom1 
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
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   25
            Top             =   45
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox edcSelCFrom 
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
            Left            =   840
            MaxLength       =   3
            TabIndex        =   23
            Top             =   45
            Width           =   615
         End
         Begin MSComDlg.CommonDialog cdcSetup 
            Left            =   1485
            Top             =   2460
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DefaultExt      =   ".Txt"
            Filter          =   "*.Txt|*.Txt|*.Doc|*.Doc|*.Asc|*.Asc"
            FilterIndex     =   1
            FontSize        =   0
            MaxFileSize     =   256
         End
         Begin VB.Label lacContract 
            Appearance      =   0  'Flat
            Caption         =   "Contract #"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   31
            Top             =   510
            Width           =   915
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Month"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   28
            Top             =   120
            Width           =   960
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "Year"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1800
            TabIndex        =   24
            Top             =   120
            Visible         =   0   'False
            Width           =   810
         End
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
         Height          =   3570
         Left            =   4590
         ScaleHeight     =   3570
         ScaleWidth      =   4455
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   4455
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   1
            Left            =   0
            MultiSelect     =   2  'Extended
            TabIndex        =   29
            Top             =   360
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   0
            Left            =   15
            MultiSelect     =   2  'Extended
            TabIndex        =   16
            Top             =   360
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Markets"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   15
            TabIndex        =   17
            Top             =   30
            Width           =   3945
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   20
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   18
      Top             =   105
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
         Width           =   1395
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
         Width           =   900
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   8850
      Top             =   1035
      Width           =   360
   End
End
Attribute VB_Name = "RptSelID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptselid.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************


' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelID.Frm - Invoice Discrepancy Worksheet
'           Compares Contract ordered against spots
'            2-18-02 DH
'
' Release: 4.8
'
' Description:
'   This file contains the Report selection for Contracts screen code
Option Explicit
Option Compare Text
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal X%, ByVal Y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal X%, ByVal Y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
'Vehicle link file- used to obtain start date
'Delivery file- used to obtain start date
'Vehicle conflict file- used to obtain start date
'Spot projection- used to obtain date status
'Library calendar file- used to obtain post log date status
'User- used to obtain discrepancy contract that was currently being processed
'      this is used if the system gos down
'Log
Dim smLogUserCode As String
'Import contract report
'Spot week Dump
Dim imTerminate As Integer
'Dim tmSRec As LPOPREC
'Rate Card
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
Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).hwnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
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
    'mTerminate True
    mTerminate False
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcGen_Click()
    Dim ilRet As Integer
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    If igGenRpt Then
        Exit Sub
    End If
    igGenRpt = True
    igOutput = frcOutput.Enabled
    igCopies = frcCopies.Enabled
    'igWhen = frcWhen.Enabled
    igFile = frcFile.Enabled
    igOption = frcOption.Enabled
    'igReportType = frcRptType.Enabled
    frcOutput.Enabled = False
    frcCopies.Enabled = False
    'frcWhen.Enabled = False
    frcFile.Enabled = False
    frcOption.Enabled = False
    'frcRptType.Enabled = False

    igUsingCrystal = True
    ilNoJobs = 1
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs

        If Not gOpenPrtJob(Trim$("InvDisc.rpt")) Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        ilRet = gcmcgenID(imGenShiftKey, smLogUserCode)
        If ilRet = -1 Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            'mTerminate
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Exit Sub
        ElseIf ilRet = 0 Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If

        Screen.MousePointer = vbHourglass

        gIDSummary
        Screen.MousePointer = vbDefault
        If rbcOutput(0).Value Then
            DoEvents            '9-19-02 fix for timing problem to prevent freezing before calling crystal
            igDestination = 0
            Report.Show vbModal
        ElseIf rbcOutput(1).Value Then
            ilCopies = Val(edcCopies.Text)
            ilRet = gOutputToPrinter(ilCopies)
        Else
            slFileName = edcFileName.Text
            'ilret = gOutputToFile(slFileName, imFTSelectedIndex)
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '2-21-02
        End If
    Next ilJobs
    imGenShiftKey = 0

    Screen.MousePointer = vbHourglass
    gCRGrfClear
    Screen.MousePointer = vbDefault
    igGenRpt = False
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
    frcFile.Enabled = igFile
    frcOption.Enabled = igOption
    pbcClickFocus.SetFocus
    tmcDone.Enabled = True
    Exit Sub
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
    'cdcSetup.flags = cdlPDPrintSetup
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
    If (KeyAscii <= 32) Or (KeyAscii = 34) Or (KeyAscii = 39) Or ((KeyAscii >= KEYDOWN) And (KeyAscii <= 45)) Or ((KeyAscii >= 59) And (KeyAscii <= 63)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcSelCFrom_Change()
Dim ilLen As Integer
    ilLen = Len(edcSelCFrom)
    mSetCommands
End Sub
Private Sub edcSelCFrom_GotFocus()
    gCtrlGotFocus edcSelCFrom
End Sub
Private Sub edcSelCFrom1_Change()
    mSetCommands
End Sub
Private Sub edcSelCFrom1_GotFocus()
    gCtrlGotFocus edcSelCFrom1
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub Form_Load()
    imTerminate = False
    igGenRpt = False
    mParseCmmdLine
    mInit
    If imTerminate Then 'Used for print only
        mTerminate True
        Exit Sub
    End If
    'RptSelID.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    PECloseEngine
    
    Set RptSelID = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcSelection_Click(Index As Integer)
    If Not imAllClicked Then
        If Index = 0 Then           'vehicle list box
            imSetAll = False
            ckcAll.Value = vbUnchecked  '9-12-02 False
            imSetAll = True
        End If
    End If
    mSetCommands
End Sub
Private Sub lbcSelection_GotFocus(Index As Integer)
    gCtrlGotFocus lbcSelection(Index)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:6/16/93       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*            Place focus before populating all lists  *                                                   *
'*******************************************************
Private Sub mInit()
Dim ilRet As Integer
Dim ilLoop As Integer
Dim slStr As String

    Screen.MousePointer = vbHourglass
    'Start Crystal report engine
    ilRet = PEOpenEngine()
    If ilRet = 0 Then
        MsgBox "Unable to open print engine"
        mTerminate False
        Exit Sub
    End If
    'Set options for report generate
    'hdJob = rpcRpt.hJob
    'ilMultiTable = True
    'ilDummy = LlSetOption(hdJob, LL_OPTION_SORTVARIABLES, True)
    'ilDummy = LlSetOption(hdJob, LL_OPTION_ONLYONETABLE, Not ilMultiTable)

    RptSelID.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
    ckcAll.Move 15, 0
    lbcSelection(0).Move 15, ckcAll.Height + 30, 4380

    pbcSelC.Move 90, 255, 4515, 3360

    gCenterStdAlone RptSelID
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
Private Sub mInitReport()
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
    'Setup report output types
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0

    gPopExportTypes cbcFileType     '2-21-02
    pbcSelC.Visible = False
    'lbcRptType.Clear
    'lbcRptType.AddItem smSelectedRptName

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    Screen.MousePointer = vbHourglass

    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    pbcSelC.Visible = True
    mMarketPopID                 'populate the markets for selection
    pbcOption.Visible = True
    lacSelCFrom.Visible = True
    edcSelCFrom.Visible = True
    'edcSelCFrom.Move 1350, edcSelCFrom.Top, 945
    lacSelCFrom.Visible = True
    edcSelCFrom.Visible = True
    lacSelCFrom1.Visible = True
    edcSelCFrom1.Visible = True
    lbcSelection(0).Visible = True  'False     'True                  'advertisers
    ckcAll.Visible = True   'False              'True
    ckcAll.Value = vbUnchecked  '9-12-02 False                 'force all advertisrs
    pbcOption.Visible = True
    pbcOption.Enabled = True

    'If lbcRptType.ListCount > 0 Then
    '    gFindMatch smSelectedRptName, 0, lbcRptType
    '    If gLastFound(lbcRptType) < 0 Then
    '        MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
    '        imTerminate = True
    '        Exit Sub
    '    End If
     '   lbcRptType.ListIndex = gLastFound(lbcRptType)
    'End If
    mSetCommands
    Screen.MousePointer = vbDefault
'    gCenterModalForm RptSel
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMarketPop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Market combobox       *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mMarketPopID()
'
'   mMarketPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim slNameCode As String
    Dim ilSortCode As Integer
    Dim ilLoop As Integer
    Dim llLen As Long
    Dim slStr As String
    Dim slTest As String * 1
    Dim igFoundLoc As Integer
    Dim igFoundNat As Integer
    Dim slStamp As String

    ReDim tlMnf(0 To 0) As MNF
    lbcSelection(0).Clear
    ilSortCode = 0

    slTest = "Y"

    'Gather the sales sources for local/national question.  This applies only to those sites whose
    'sales source origins have combination of local & national
    igFoundLoc = False
    igFoundNat = False
    ilRet = gObtainMnfForType("S", slStamp, tlMnf())
    For ilLoop = LBound(tlMnf) To UBound(tlMnf) - 1
        If tlMnf(ilLoop).iGroupNo = 1 Then
            igFoundLoc = True
        End If
        If tlMnf(ilLoop).iGroupNo = 3 Then      'national
            igFoundNat = True
        End If
    Next ilLoop
    If (igFoundNat) And (igFoundLoc) Then   'local & nation exists, must be  a remote site otherwise there would be only Natl
        slTest = "N"
    End If

    ReDim tgMktCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    ilRet = gObtainMnfForType("H3", slStr, tgMktMnf())
    For ilLoop = LBound(tgMktMnf) To UBound(tgMktMnf) - 1 Step 1
        If Trim$(tgMktMnf(ilLoop).sRPU) <> slTest Then 'get all non-clusters
            slName = Trim$(tgMktMnf(ilLoop).sName)
            slName = slName & "\" & Trim$(str$(tgMktMnf(ilLoop).iCode))
            tgMktCode(ilSortCode).sKey = slName
            If ilSortCode >= UBound(tgMktCode) Then
                ReDim Preserve tgMktCode(0 To UBound(tgMktCode) + 100) As SORTCODE
            End If
            ilSortCode = ilSortCode + 1
        End If
    Next ilLoop
    ReDim Preserve tgMktCode(0 To ilSortCode) As SORTCODE
    If UBound(tgMktCode) - 1 > 0 Then
        ArraySortTyp fnAV(tgMktCode(), 0), UBound(tgMktCode), 0, LenB(tgMktCode(0)), 0, LenB(tgMktCode(0).sKey), 0
    End If
    llLen = 0
    For ilLoop = 0 To UBound(tgMktCode) - 1 Step 1
        slNameCode = tgMktCode(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet = CP_MSG_NONE Then
            slName = Trim$(slName)
            If Not gOkAddStrToListBox(slName, llLen, True) Then
                Exit For
            End If
            lbcSelection(0).AddItem slName  'Add ID to list box
        End If
    Next ilLoop
    Exit Sub

    On Error GoTo 0
    imTerminate = True
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

    slCommand = sgCommandStr
    ilRet = gParseItem(slCommand, 1, "||", smCommand)
    If (ilRet <> CP_MSG_NONE) Then
        smCommand = slCommand
    End If
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False  'True 'False
    '    imShowHelpMsg = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            End
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
    'gInitStdAlone RptSelID, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Remote Invoice Worksheet"
    '    igRptCallType = -1  'unused in standalone exe, CONTRACTSJOB 'SLSPCOMMSJOB   'LOGSJOB 'CONTRACTSJOB 'COPYJOB 'COLLECTIONSJOB'CONTRACTSJOB 'CHFCONVMENU 'PROGRAMMINGJOB 'INVOICESJOB  'ADVERTISERSLIST 'POSTLOGSJOB 'DALLASFEED 'BULKCOPY 'PHOENIXFEED 'CMMLCHG 'EXPORTAFFSPOTS 'BUDGETSJOB 'PROPOSALPROJECTION
    '    igRptType = -1  'unused in standalone exe   '3 'Log '1   'Contract    '0   'Summary '3 Program  '1  links
    'Else
        ilRet = gParseItem(slCommand, 2, "||", slRptListCmmd)
        If (ilRet = CP_MSG_NONE) Then
            ilRet = gParseItem(slRptListCmmd, 2, "\", smSelectedRptName)
            ilRet = gParseItem(slRptListCmmd, 1, "\", slStr)
            igRptCallType = Val(slStr)      'Function ID (what function calling this report if )
        End If
   'End If
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
    Dim ilLoop As Integer

    ilEnable = False
    If (edcSelCFrom.Text <> "") And (edcSelCFrom1.Text <> "") Then
        ilEnable = True
        'atleast one market must be selected
        If ilEnable Then
            If ckcAll.Value <> vbChecked Then
                ilEnable = False
                For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1      'market entry must be selected
                    If lbcSelection(0).Selected(ilLoop) Then
                        ilEnable = True
                        Exit For
                    End If
                Next ilLoop
            End If
        End If
    End If

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
    Unload RptSelID
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
    '    Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    '    Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    '    Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub rbcOutput_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcOutput(Index).Value
    'End of Coded added
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
        'mInitDDE
        mInitReport
        If imTerminate Then 'Used for print only
            mTerminate True
            Exit Sub
        End If
        imFirstTime = False
    End If
    gCtrlGotFocus rbcOutput(Index)
End Sub
Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
