VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form RptSelED 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facility Report Selection"
   ClientHeight    =   5535
   ClientLeft      =   495
   ClientTop       =   1485
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
      TabIndex        =   20
      Top             =   615
      Width           =   2055
   End
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6690
      TabIndex        =   26
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
      TabIndex        =   27
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   -90
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.FileListBox lbcFileName 
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
      Height          =   225
      Left            =   3060
      Pattern         =   "*.Dal"
      TabIndex        =   15
      Top             =   4950
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ListBox lbcSort 
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
      Height          =   225
      Left            =   2010
      Sorted          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   960
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
      Left            =   855
      Top             =   4770
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
      Left            =   2055
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
      Caption         =   "Producer Earned Distribution"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4005
      Left            =   90
      TabIndex        =   14
      Top             =   1425
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
         Height          =   3705
         Left            =   45
         ScaleHeight     =   3705
         ScaleWidth      =   4530
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   4530
         Begin VB.PictureBox plcTotalsBy 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   105
            ScaleHeight     =   255
            ScaleWidth      =   3720
            TabIndex        =   41
            Top             =   1095
            Width           =   3720
            Begin VB.OptionButton rbcTotalsBy 
               Caption         =   "Contract"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   870
               TabIndex        =   43
               Top             =   0
               Value           =   -1  'True
               Width           =   1110
            End
            Begin VB.OptionButton rbcTotalsBy 
               Caption         =   "Advertiser"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1995
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   0
               Width           =   1545
            End
         End
         Begin VB.PictureBox plcDistType 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   3135
            TabIndex        =   40
            Top             =   840
            Width           =   3135
            Begin VB.OptionButton rbcEarnYr 
               Caption         =   "Year"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   360
               TabIndex        =   35
               Top             =   0
               Value           =   -1  'True
               Width           =   690
            End
            Begin VB.OptionButton rbcEarnCnt 
               Caption         =   "Contract"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1080
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   0
               Width           =   1095
            End
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
            Left            =   1185
            MaxLength       =   8
            TabIndex        =   37
            Top             =   1365
            Width           =   975
         End
         Begin VB.TextBox edcQtr 
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
            Left            =   2220
            MaxLength       =   1
            TabIndex        =   25
            Top             =   15
            Width           =   450
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
            Left            =   600
            MaxLength       =   4
            TabIndex        =   24
            Top             =   15
            Width           =   675
         End
         Begin VB.ComboBox cbcSel 
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
            Height          =   315
            Left            =   90
            TabIndex        =   29
            Top             =   2115
            Visible         =   0   'False
            Width           =   4305
         End
         Begin VB.PictureBox plcSelC4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   105
            ScaleHeight     =   480
            ScaleWidth      =   4380
            TabIndex        =   31
            Top             =   360
            Width           =   4380
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Consolidated Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   2400
               TabIndex        =   45
               Top             =   240
               Width           =   1860
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Consolidated  Gross"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   2
               Left            =   330
               TabIndex        =   44
               Top             =   240
               Width           =   2070
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Participant Net"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2400
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   1530
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Participant Gross"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   600
               TabIndex        =   32
               Top             =   0
               Width           =   1860
            End
         End
         Begin VB.Label lacContract 
            Appearance      =   0  'Flat
            Caption         =   "Contract #"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   38
            Top             =   1410
            Width           =   915
         End
         Begin VB.Label lacQtr 
            Appearance      =   0  'Flat
            Caption         =   "Start Qtr"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1440
            TabIndex        =   33
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Year"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   30
            Top             =   60
            Width           =   540
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
         Height          =   3420
         Left            =   4590
         ScaleHeight     =   3420
         ScaleWidth      =   4455
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   4455
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            Index           =   0
            Left            =   90
            MultiSelect     =   2  'Extended
            TabIndex        =   17
            Top             =   300
            Width           =   4380
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Participants"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   0
            Width           =   3945
         End
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
         Width           =   930
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
Attribute VB_Name = "RptSelED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelED.Frm   Earned Distribution report
'       8-28-01
'
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
Dim smSelectedRptName As String 'Passed selected report name
Dim smLogUserCode As String
Dim imTerminate  As Integer
Dim tmIsr As ISR                  'Invoice  prepass file
Dim hmIsr As Integer
Dim tmIsrSrchKey1 As ISRKEY1            'ISR record image
Dim imIsrRecLen As Integer        'Generic record length
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
Private Sub cbcSel_Change()
    mSetCommands
End Sub
Private Sub cbcSel_Click()
    mSetCommands
End Sub
Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    Dim ilIndex As Integer
    ilIndex = lbcRptType.ListIndex
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
    cdcSetup.flags = cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or CdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
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
    Dim ilRet As Integer
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilListIndex As Integer
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    Dim llISR_RecsFound As Long          '3-28-02

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
    ilListIndex = lbcRptType.ListIndex

    ilNoJobs = 1
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        If Not gGenReportED() Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        ilRet = gCmcGenEd(ilListIndex, imGenShiftKey, smLogUserCode)
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
            pbcClickFocus.SetFocus
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
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Exit Sub
       End If
       '1 falls thru - successful crystal report
        Screen.MousePointer = vbHourglass
        gCREarnedDistr
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
            'ilRet = gOutputToFile(slFileName, imFTSelectedIndex)
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '5-2-02

        End If
    Next ilJobs
    imGenShiftKey = 0
    Screen.MousePointer = vbHourglass
    gCRGrfClear

    '3-28-02 if total records written matched the number found on file, overlay with zero_image because
    'it takes too long to clear out this file.  Otherwise, each record will have to be deleted
    'reopen ISR file
    hmIsr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmIsr, "", sgDBPath & "Isr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmIsr)
        btrDestroy hmIsr
        Screen.MousePointer = vbDefault
        igGenRpt = False
        frcOutput.Enabled = igOutput
        frcCopies.Enabled = igCopies
        frcFile.Enabled = igFile
        frcOption.Enabled = igOption
        pbcClickFocus.SetFocus
        tmcDone.Enabled = True
        Exit Sub
    End If

    llISR_RecsFound = btrRecords(hmIsr)  'obtain count in ISR to see if any other user is using it
    ilRet = btrClose(hmIsr)
    btrDestroy hmIsr

    If llISR_RecsFound = lgTotal_ISRRecs Then
        On Error GoTo mFileCopyErr
        FileCopy sgDBPath & "ISR_Blk.Btr", sgDBPath & "ISR.Btr"

    End If
    gISRClear   'clear in case there was an error or the record counts dont match
    Screen.MousePointer = vbDefault
    igGenRpt = False
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
    'frcWhen.Enabled = igWhen
    frcFile.Enabled = igFile
    frcOption.Enabled = igOption
    pbcClickFocus.SetFocus
    tmcDone.Enabled = True
    Exit Sub

mFileCopyErr:
    Resume Next
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
Private Sub edcQtr_Change()
    mSetCommands
End Sub
Private Sub edcSelCFrom_Change()
    mSetCommands
End Sub
Private Sub edcSelCFrom_GotFocus()
    gCtrlGotFocus edcSelCFrom
End Sub
Private Sub edcSelCFrom_KeyPress(KeyAscii As Integer)
    Dim ilListIndex As Integer
    If igRptCallType = COPYJOB Then
        ilListIndex = lbcRptType.ListIndex
        If (tgSpf.sUseCartNo = "N") And (ilListIndex >= 4) Then
            ilListIndex = ilListIndex + 1
        End If
        If ilListIndex = 4 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
    If igRptCallType = BULKCOPY Then
        ilListIndex = lbcRptType.ListIndex
        If ilListIndex = 2 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
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
    Me.Refresh
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
    igGenRpt = False
    mParseCmmdLine
    mInit
    If imTerminate = -99 Then
        Exit Sub
    End If
    If imTerminate Then 'Used for print only
        'mTerminate
        cmcCancel_Click
        Exit Sub
    End If
    'RptSelED.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Reset used instead of Close to cause # Clients on network to be decrement
    'ilRet = btrReset(hgHlf)
    'btrDestroy hgHlf
    'btrStopAppl
    'End
End Sub
Private Sub gISRClear()
'*******************************************************
'*                                                     *
'*   Procedure Name:RemoveISR records with matching    *
'*             date & time created
'*                                                     *
'*            Created:12/03/01       By:D. Hosaka      *
'
'*                                                     *
'*                                                     *
'*******************************************************
    Dim ilRet As Integer
    hmIsr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmIsr, "", sgDBPath & "Isr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmIsr)
        btrDestroy hmIsr
        Exit Sub
    End If
    imIsrRecLen = Len(tmIsr)
    tmIsrSrchKey1.iGenDate(0) = igNowDate(0)
    tmIsrSrchKey1.iGenDate(1) = igNowDate(1)
    tmIsrSrchKey1.iGenTime(0) = igNowTime(0)
    tmIsrSrchKey1.iGenTime(1) = igNowTime(1)
    tmIsrSrchKey1.iType = 0
    tmIsrSrchKey1.lCode = 0

    ilRet = btrGetGreaterOrEqual(hmIsr, tmIsr, imIsrRecLen, tmIsrSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmIsr.iGenDate(0) = igNowDate(0)) And (tmIsr.iGenDate(1) = igNowDate(1)) And (tmIsr.iGenTime(0) = igNowTime(0)) And (tmIsr.iGenTime(1) = igNowTime(1))
        ilRet = btrDelete(hmIsr)
        tmIsrSrchKey1.iGenDate(0) = igNowDate(0)
        tmIsrSrchKey1.iGenDate(1) = igNowDate(1)
        tmIsrSrchKey1.iGenTime(0) = igNowTime(0)
        tmIsrSrchKey1.iGenTime(1) = igNowTime(1)
        tmIsrSrchKey1.iType = 0
        tmIsrSrchKey1.lCode = 0
        ilRet = btrGetGreaterOrEqual(hmIsr, tmIsr, imIsrRecLen, tmIsrSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        'ilRet = btrGetNext(hmIsr, tmISR, imIsrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmIsr)
    btrDestroy hmIsr
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcRptType_Click()
    Dim ilListIndex As Integer
        ilListIndex = lbcRptType.ListIndex
        edcSelCFrom.MaxLength = 10      '8 5/28/99 allow 10 char input m/d/yyyy
        edcSelCFrom.Width = 1170
        pbcSelC.Visible = False
        plcSelC4.Visible = False
        edcSelCFrom.Move 1050, 30
        lacSelCFrom.Move 120, 75
        lbcSelection(0).Visible = True
        ckcAll.Visible = True
        lacSelCFrom.Visible = True
        edcSelCFrom.Left = 1050
        edcSelCFrom.Visible = True
        plcSelC4.Move lacSelCFrom.Left, edcSelCFrom.Top + edcSelCFrom.Height
        rbcSelC4(0).Left = 630
        rbcSelC4(0).Width = 900
        rbcSelC4(1).Left = 1630
        rbcSelC4(1).Width = 1140
        If rbcSelC4(1).Value Then
            rbcSelC4_click 1    ', True
        Else
            rbcSelC4(1).Value = True   'Summary
        End If
        pbcSelC.Visible = True
        pbcOption.Visible = True
        pbcOption.Visible = True
    mSetCommands
End Sub
Private Sub lbcSelection_Click(Index As Integer)
    Dim ilListIndex As Integer
    If Not imAllClicked Then
        ilListIndex = lbcRptType.ListIndex

        imSetAll = False
        ckcAll.Value = vbUnchecked  '12-11-01 False
        imSetAll = True
    Else
        'imSetAll = False
        'ckcAll.Value = False
        'imSetAll = True
    End If
    mSetCommands
End Sub
Private Sub lbcSelection_GotFocus(Index As Integer)
    gCtrlGotFocus lbcSelection(Index)
End Sub
Private Sub lbcSort_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked  '12-11-01 False
        imSetAll = True
    End If
    mSetCommands
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
    Dim slStr As String
    Dim ilLoop As Integer
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
    RptSelED.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"

    imAllClicked = False
    imSetAll = True
    cbcSel.Move 120, 30
    'lacSelCFrom.Move 120, 75, 2235
    'edcSelCFrom.Move 2355, 30
    'plcSelC4.Move 120, edcSelCFrom.Top + edcSelCFrom.Height, 2835

    lbcSelection(0).Move 15, ckcAll.Height + 30, 4380, 3000
    'pbcSelA.Move 90, 255, 4515, 3360
    'pbcSelB.Move 90, 255, 4515, 3360
    pbcSelC.Move 90, 255, 4515, 3360
    gCenterStdAlone RptSelED
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
    Dim ilRet As Integer
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
    gPopExportTypes cbcFileType     '5-2-02

    pbcSelC.Visible = True
    'mMnfPop "H", RptSelEd!lbcSelection(0), tgVehicle(), sgVehicleTag    'Traffic!lbcVehicle         'owners groups

    ilRet = gPopMnfPlusFieldsBox(RptSelED, RptSelED!lbcSelection(0), tgVehicle(), sgVehicleTag, "H1")
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If


    If imTerminate Then
        Exit Sub
    End If
    frcOption.Enabled = True
    'lbcSelection(0).Move 15, ckcAll.Height + 30
    pbcSelC.Visible = True
    pbcOption.Visible = True
    mSetCommands
    Screen.MousePointer = vbDefault
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
    'If StrComp(slcommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False  'True 'False
    '    imShowHelpmsg = False
    'Else
     '   igStdAloneMode = False  'Switch from/to stand alone mode
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
        'imShowHelpmsg = True
        'ilRet = gParseItem(slStr, 3, "^", slHelpSystem)
        'If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
        '    imShowHelpmsg = False
        'End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone RptSelED, slStr, ilTestSystem
    'ilRet = gParseItem(slCommand, 3, "\", slStr)
    'igRptCallType = Val(slStr)
    If igRptCallType <> GENERICBUTTON Then
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        igRptType = Val(slStr)
    Else
        igRptType = 0
    End If
    'If igStdAloneMode Then
        'smSelectedRptName = "Copy Inventory by Advertiser"
    '    smSelectedrptName = "Producer Earned Distribution"
   '     igRptCallType = -1 'COLLECTIONSJOB 'INVOICESJOB 'NYFEED  'COLLECTIONSJOB 'SLSPCOMMSJOB   'LOGSJOB 'COPYJOB 'COLLECTIONSJOB 'CHFCONVMENU 'PROGRAMMINGJOB 'INVOICESJOB  'ADVERTISERSLIST 'POSTLOGSJOB 'DALLASFEED 'BULKCOPY 'PHOENIXFEED 'CMMLCHG 'EXPORTAFFSPOTS 'BUDGETSJOB
    '    igRptType = -1  'unused in standalone exe 'Log     '0   'Summary '3 Program  '1  links
    'Else
        ilRet = gParseItem(slCommand, 2, "||", slRptListCmmd)
        If (ilRet = CP_MSG_NONE) Then
            ilRet = gParseItem(slRptListCmmd, 2, "\", smSelectedRptName)
            ilRet = gParseItem(slRptListCmmd, 1, "\", slStr)
            igRptCallType = Val(slStr)
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
    Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    ilEnable = False
    If (edcSelCFrom.Text <> "" And edcQtr <> "") Then
        ilEnable = True
    End If
    If (RptSelED!ckcAll.Value <> vbChecked) And ilEnable Then     '9-12-02
        ilEnable = False
        For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1
            If lbcSelection(0).Selected(ilLoop) Then
                ilEnable = True
                Exit For
            End If
        Next ilLoop
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
    Erase tgAirNameCode
    Erase tgCSVNameCode
    Erase tgSellNameCode
    Erase tgRptSelSalespersonCode
    Erase tgRptSelAgencyCode
    Erase tgRptSelAdvertiserCode
    Erase tgRptSelNameCode
    'Erase tgRptSelBudgetCode
    'Erase tgMultiCntrCode
    'Erase tgManyCntCode
    'Erase tgRptSelDemoCode
    'Erase tgSOCode
    PECloseEngine
    If ilFromCancel Then
        'If Not igStdAloneMode Then
        '    'Shell back to RptList
         '   lgShellRet = Shell(sgExePath & "RptList.Exe" & " " & smCommand, 1)
        'End If
        igRptReturn = True
    Else
        'igParentRestarted = False
        'If Not igStdAloneMode Then
        '    If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
        '        edcLinkDestHelpMsg.LinkExecute "@" & "Done"
        '    Else
        '        edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
        '        edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
        '        edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
        '        edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
        '        edcLinkDestHelpMsg.LinkExecute "Done"
        '    End If
        '    Do While Not igParentRestarted
        '        DoEvents
        '    Loop
        'End If
        igRptReturn = False
    End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload RptSelED
    Set RptSelED = Nothing   'Remove data segment
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcTotalsBy_Paint()
    plcTotalsBy.CurrentX = 0
    plcTotalsBy.CurrentY = 0
    plcTotalsBy.Print "Totals By"
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
        'mInitDDE
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
Private Sub rbcSelC4_click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelC4(Index).Value
    'End of coded added
Dim ilListIndex As Integer
    ilListIndex = lbcRptType.ListIndex
    If igRptCallType = BUDGETSJOB Then
        If Index = 0 Then
            rbcSelC4(1).Enabled = False
        Else
            rbcSelC4(0).Enabled = False
        End If
    End If
End Sub
Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
Private Sub plcDistType_Paint()
    plcDistType.CurrentX = 0
    plcDistType.CurrentY = 0
    plcDistType.Print "By"
End Sub
Private Sub plcSelC4_Paint()
    plcSelC4.CurrentX = 0
    plcSelC4.CurrentY = 0
    plcSelC4.Print "Show"
End Sub
