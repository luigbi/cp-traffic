VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelALLOC 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revenue Allocation Report Selection"
   ClientHeight    =   5715
   ClientLeft      =   195
   ClientTop       =   1545
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
   ScaleHeight     =   5715
   ScaleWidth      =   9270
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6600
      TabIndex        =   18
      Top             =   615
      Width           =   2055
   End
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   8145
      TabIndex        =   9
      Top             =   -15
      Visible         =   0   'False
      Width           =   1095
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
      TabIndex        =   23
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   -60
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
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   30
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   4440
      Top             =   5040
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
         Left            =   1065
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   8
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
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   15
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
         TabIndex        =   12
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
         TabIndex        =   14
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Revenue Allocation Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   4050
      Left            =   45
      TabIndex        =   16
      Top             =   1530
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
         Height          =   3720
         Left            =   120
         ScaleHeight     =   3720
         ScaleWidth      =   5130
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   5130
         Begin VB.CheckBox ckcShowMissingAudience 
            Caption         =   "Show Stations with Missing Audience"
            Height          =   330
            Left            =   120
            TabIndex        =   46
            Top             =   3310
            Value           =   1  'Checked
            Width           =   3615
         End
         Begin VB.Frame frcUseOrderOrAir 
            Caption         =   "Revenue Allocations by"
            Height          =   615
            Left            =   120
            TabIndex        =   33
            Top             =   2370
            Width           =   3375
            Begin VB.OptionButton rbcOrderAir 
               Caption         =   "Aired Spots"
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   35
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton rbcOrderAir 
               Caption         =   "Ordered Spots"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   34
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.CheckBox ckcExtraFundsToBal 
            Caption         =   "Toss extra pennies into last station "
            Height          =   330
            Left            =   120
            TabIndex        =   32
            Top             =   3070
            Value           =   1  'Checked
            Width           =   3375
         End
         Begin VB.PictureBox plcSortBy 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1395
            Left            =   120
            ScaleHeight     =   1395
            ScaleWidth      =   5340
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   810
            Width           =   5340
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Vehicle Group, Cash/Trade, Station"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   5
               Left            =   720
               TabIndex        =   41
               Top             =   480
               Width           =   4365
            End
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Vehicle Group, Cash/Trade, Market Name, Station"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   4
               Left            =   720
               TabIndex        =   40
               Top             =   240
               Width           =   4935
            End
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Vehicle Group, Cash/Trade, Market Rank, Station"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   3
               Left            =   720
               TabIndex        =   39
               Top             =   0
               Width           =   4485
            End
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Cash/Trade, Station"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   2
               Left            =   720
               TabIndex        =   31
               Top             =   1200
               Value           =   -1  'True
               Width           =   2805
            End
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Cash/Trade, Market Rank, Station"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   720
               TabIndex        =   29
               Top             =   720
               Width           =   3735
            End
            Begin VB.OptionButton rbcSortBy 
               Caption         =   "Cash/Trade, Market Name, Station"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   30
               Top             =   960
               Width           =   3165
            End
         End
         Begin VB.TextBox edcContract 
            Height          =   315
            Left            =   3840
            MaxLength       =   9
            TabIndex        =   37
            Top             =   2610
            Width           =   1215
         End
         Begin VB.TextBox edcYear 
            Height          =   315
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   27
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox edcMonth 
            Height          =   315
            Left            =   840
            MaxLength       =   3
            TabIndex        =   25
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lacStdBdcst 
            Caption         =   "Standard Broadcast-"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label lacContract 
            Caption         =   "Contract #"
            Height          =   255
            Left            =   3840
            TabIndex        =   36
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lacYear 
            Caption         =   "Year"
            Height          =   255
            Left            =   1680
            TabIndex        =   26
            Top             =   390
            Width           =   495
         End
         Begin VB.Label lacMonth 
            Caption         =   "Month"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   390
            Width           =   735
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
         Height          =   3660
         Left            =   5445
         ScaleHeight     =   3660
         ScaleWidth      =   3615
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   3615
         Begin VB.TextBox edcSet1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Height          =   225
            Left            =   0
            TabIndex        =   45
            TabStop         =   0   'False
            Text            =   "Vehicle Group"
            Top             =   165
            Width           =   1215
         End
         Begin VB.ComboBox cbcSet 
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
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   120
            Width           =   1500
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicle Groups"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   15
            TabIndex        =   43
            Top             =   600
            Width           =   2055
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2760
            Index           =   0
            ItemData        =   "RptSelAlloc.frx":0000
            Left            =   0
            List            =   "RptSelAlloc.frx":0002
            MultiSelect     =   2  'Extended
            TabIndex        =   42
            Top             =   870
            Visible         =   0   'False
            Width           =   3525
         End
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
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   6
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
         Width           =   1245
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
Attribute VB_Name = "RptSelALLOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselalloc.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelALLOC.Frm  - Revenue Allocation
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection for Contracts screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
Dim smSortBy As String

Dim smLogUserCode As String
Dim imTerminate As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
'Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)

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
 
        If Not gGenReportAlloc() Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        
        'TTP 10376 - Revenue Allocation report: update to use vehicle groups (doEvents allows buttons to be played with)
        cmcGen.Enabled = False
        cmcList.Enabled = False
        cmcCancel.Enabled = False
        frcOption.Enabled = False
                
        ilRet = gCmcGenAlloc()
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
            cmcGen.Enabled = True
            cmcList.Enabled = True
            cmcCancel.Enabled = True
            frcOption.Enabled = True
            Exit Sub
        ElseIf ilRet = 0 Then   '0 = invalid input data, stay in
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            cmcGen.Enabled = True
            cmcList.Enabled = True
            cmcCancel.Enabled = True
            frcOption.Enabled = True
            Exit Sub
        ElseIf ilRet = 2 Then           'successful return from bridge reports
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            cmcGen.Enabled = True
            cmcList.Enabled = True
            cmcCancel.Enabled = True
            frcOption.Enabled = True
            Exit Sub
        End If
       '1 falls thru - successful crystal report
        Screen.MousePointer = vbHourglass
        ilRet = gCreateRevAlloc()
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
           ' ilRet = gOutputToFile(slFileName, imFTSelectedIndex)
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
        End If
    Next ilJobs
    imGenShiftKey = 0
    'gClearGrf
    gCRGrfClear         '8-20-13 use only 1 common grf clear rtn which changes the way records are removed
    igGenRpt = False
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
    'frcWhen.Enabled = igWhen
    frcFile.Enabled = igFile
    frcOption.Enabled = igOption
    pbcClickFocus.SetFocus
    tmcDone.Enabled = True
    cmcGen.Enabled = True
    cmcList.Enabled = True
    cmcCancel.Enabled = True
    frcOption.Enabled = True
    
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
Private Sub edcMonth_Change()
    mSetCommands
End Sub

Private Sub edcMonth_GotFocus()
    gCtrlGotFocus edcMonth
End Sub

Private Sub edcYear_Change()
    mSetCommands
End Sub

Private Sub edcYear_GotFocus()
    gCtrlGotFocus edcYear
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If cbcSet.ListCount > 0 And cbcSet.ListIndex = -1 Then cbcSet.ListIndex = 0
    Me.KeyPreview = True
    RptSelALLOC.Refresh
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
    If imTerminate Then
        cmcCancel_Click
        Exit Sub
    End If
    mInit
    If imTerminate Then 'Used for print only
        mTerminate True
        Exit Sub
    End If
    'RptSelAL.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    PECloseEngine
    
    Set RptSelALLOC = Nothing   'Remove data segment
    
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
'*             Created:6/16/93       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*            Place focus before populating all lists  *                                                   *
'*******************************************************
Private Sub mInit()
Dim ilRet As Integer
Dim illoop As Integer
Dim slStr As String
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
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

    RptSelALLOC.Caption = smSelectedRptName & " Report"
    'frcOption.Caption = smSelectedRptName & " Selection"
    slStr = Trim$(smSelectedRptName)
    illoop = InStr(slStr, "&")
    If illoop > 0 Then
        slStr = Left$(slStr, illoop - 1) & "&&" & Mid$(slStr, illoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"

    'TTP 10376 - Revenue Allocation report: update to use vehicle groups
    gPopVehicleGroups RptSelALLOC!cbcSet, tgVehicleSets1(), False
    'remove "Participant" entry from this Vehicle Group list
    'remove "" entry from this Vehicle Group list
    For illoop = cbcSet.ListCount - 1 To 0 Step -1
        If cbcSet.List(illoop) = "Participants" Then cbcSet.RemoveItem (illoop)
        'TTP 10398 - Revenue Allocation report: "invalid procedure call or argument" error when launching the report if there's no Vehicle Groups defined for any vehicle
        If cbcSet.ListCount > 0 And Trim(cbcSet.List(illoop)) = "" Then cbcSet.RemoveItem (illoop)
    Next illoop
    'If cbcSet.ListCount > 0 And cbcSet.ListIndex = -1 Then cbcSet.ListIndex = 0
    
    gCenterStdAlone RptSelALLOC
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitReport                     *
'*                                                     *
'*             Created:8/16/00       By:D. Smith       *
'*             Modified:             By:               *
'*                                                     *
'*            Comments: Initialize report screen       *
'*                                                     *
'*******************************************************
Private Sub mInitReport()


    Dim ilDay As Integer
    Dim llDate As Long
    Dim slNowDate As String

    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    gPopExportTypes cbcFileType         '10-20-01


    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60

    pbcSelC.Visible = True
    pbcOption.Visible = False
   
    mSetCommands
    Screen.MousePointer = vbDefault
    'gCenterModalForm RptSel
    lbcSelection(0).Visible = True
    ckcAll.Visible = True
    cbcSet.Visible = True
    edcSet1.Visible = True

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
    '    ilTestSystem = False  'True 'False
    '    imShowHelpmsg = False
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
    '    imShowHelpmsg = True
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpmsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone RptSelALLOC, slStr, ilTestSystem
    'ilRet = gParseItem(slCommand, 3, "\", slStr)
    'igRptCallType = Val(slStr)
    If igRptCallType <> GENERICBUTTON Then
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        igRptType = Val(slStr)
    Else
        igRptType = 0
    End If
    'If igStdAloneMode Then
    '    smSelectedRptName = "Sales Pricing Analysis"
    '    igRptCallType = -1  'unused in standalone exe
    '    igRptType = -1   'unused in standalone exe
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
    ilEnable = False
    If Trim$(edcMonth) <> "" And Trim$(edcYear) <> "" Then
        'if any fields to check before setting the enabled flag, insert here
        ilEnable = True
        If Not (rbcOrderAir(0).Value) And Not (rbcOrderAir(1).Value) Then
            ilEnable = False
        End If
    End If
    'TTP 10376 - Revenue Allocation report: update to use vehicle groups
    If rbcSortBy(3).Value = True Or rbcSortBy(4).Value = True Or rbcSortBy(5).Value = True Then
        'Groupby Vehicle Group options require a Vehicle Group to be selected
        If lbcSelection(0).SelCount = 0 Then
            ilEnable = False
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
    Unload RptSelALLOC
    igManUnload = NO
End Sub

Private Sub lbcSelection_Click(Index As Integer)
    If ckcAll.Value = vbChecked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
    End If
    mSetCommands
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcSortBy_Paint()
    plcSortBy.CurrentX = 0
    plcSortBy.CurrentY = 0
    smSortBy = "Sort by"
    plcSortBy.Print smSortBy

End Sub

Private Sub rbcOrderAir_Click(Index As Integer)
    mSetCommands
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
        mInitReport
        If imTerminate Then 'Used for print only
            mTerminate True
            Exit Sub
        End If
        imFirstTime = False
    End If
    gCtrlGotFocus rbcOutput(Index)
End Sub

Private Sub rbcSortBy_Click(Index As Integer)
    'TTP 10376 - Revenue Allocation report: update to use vehicle groups
    pbcOption.Visible = False
    Select Case Index
        Case 3, 4, 5 'one of the Vehicle Group options
            pbcOption.Visible = True
            lbcSelection(0).Visible = True
            ckcAll.Visible = True
            cbcSet.Visible = True
            edcSet1.Visible = True
    End Select
    mSetCommands
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub

'TTP 10376 - Revenue Allocation report: update to use vehicle groups
Private Sub cbcSet_Click()
Dim ilSetIndex As Integer
Dim ilRet As Integer
Dim illoop As Integer
    If imFirstActivate = True Then
        Exit Sub
    End If
    illoop = cbcSet.ListIndex
    ilSetIndex = gFindVehGroupInx(illoop + 1, tgVehicleSets1())
    If ilSetIndex > 1 Then
        smVehGp5CodeTag = ""
        lbcSelection(0).Clear
        ilRet = gPopMnfPlusFieldsBox(RptSelAA, lbcSelection(0), tgSOCodeAA(), smVehGp5CodeTag, "H" & Trim$(str$(ilSetIndex)))
        If ilSetIndex = 2 Then              'subtotals vehicle sets
            lbcSelection(0).Visible = True
            ckcAll.Caption = "All Sub-totals"
        ElseIf ilSetIndex = 3 Then          'market vehicle sets
            lbcSelection(0).Visible = True
            ckcAll.Caption = "All Markets"
        ElseIf ilSetIndex = 4 Then          'format vehicle sets
            lbcSelection(0).Visible = True
            ckcAll.Caption = "All Formats"
        ElseIf ilSetIndex = 5 Then          'research vehicle sets
            lbcSelection(0).Visible = True
            ckcAll.Caption = "All Research"
        ElseIf ilSetIndex = 6 Then          'Sub-companies vehicle sets
            lbcSelection(0).Visible = True
            ckcAll.Caption = "All Sub-companies"
        End If
        ckcAll.Visible = True
        ckcAll.Value = vbUnchecked
    Else
        lbcSelection(0).Visible = False
        ckcAll.Value = vbUnchecked
        ckcAll.Visible = False
    End If
    mSetCommands
End Sub

Private Sub ckcAll_Click()
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
        imSetAll = True
    End If
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    ilValue = Value
    If imSetAll Then
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).HWnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    mSetCommands
End Sub

Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub
