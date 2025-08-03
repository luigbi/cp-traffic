VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form rptSelFD 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revenue Sets"
   ClientHeight    =   5535
   ClientLeft      =   180
   ClientTop       =   975
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
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6120
      TabIndex        =   40
      Top             =   1440
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   16
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
      TabIndex        =   35
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
      TabIndex        =   33
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
      TabIndex        =   18
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
      Caption         =   "Revenue Sets Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3705
      Left            =   45
      TabIndex        =   14
      Top             =   1770
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
         Left            =   120
         ScaleHeight     =   3360
         ScaleWidth      =   4485
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   210
         Visible         =   0   'False
         Width           =   4485
         Begin VB.PictureBox plcReport 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   3585
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   1170
            Visible         =   0   'False
            Width           =   3585
            Begin VB.CheckBox ckcReport 
               Caption         =   "Delivery"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   0
               TabIndex        =   29
               Top             =   0
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.CheckBox ckcReport 
               Caption         =   "Engineering"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1140
               TabIndex        =   30
               Top             =   0
               Visible         =   0   'False
               Width           =   1440
            End
         End
         Begin VB.PictureBox plcDays 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   2385
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   2385
            Begin VB.CheckBox ckcDay 
               Caption         =   "M-F"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CheckBox ckcDay 
               Caption         =   "Sa"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   750
               TabIndex        =   27
               Top             =   0
               Visible         =   0   'False
               Width           =   630
            End
            Begin VB.CheckBox ckcDay 
               Caption         =   "Su"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1425
               TabIndex        =   28
               Top             =   0
               Visible         =   0   'False
               Width           =   585
            End
         End
         Begin VB.TextBox edcToDate 
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
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   21
            Top             =   120
            Width           =   1035
         End
         Begin VB.TextBox edcToTime 
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
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   25
            Top             =   480
            Width           =   1035
         End
         Begin VB.TextBox edcFromTime 
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
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   24
            Top             =   480
            Width           =   1035
         End
         Begin VB.TextBox edcFromDate 
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
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   20
            Top             =   120
            Width           =   1035
         End
         Begin MSComDlg.CommonDialog cdcSetup 
            Left            =   735
            Top             =   1785
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DefaultExt      =   ".Txt"
            Filter          =   "*.Txt|*.Txt|*.Doc|*.Doc|*.Asc|*.Asc"
            FilterIndex     =   1
            FontSize        =   0
            MaxFileSize     =   256
         End
         Begin VB.Label lacToDate 
            Appearance      =   0  'Flat
            Caption         =   "End"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2640
            TabIndex        =   39
            Top             =   180
            Width           =   525
         End
         Begin VB.Label lacFromDate 
            Appearance      =   0  'Flat
            Caption         =   "Dates- Start"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   31
            Top             =   180
            Width           =   1125
         End
         Begin VB.Label lacToTime 
            Appearance      =   0  'Flat
            Caption         =   "End"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2640
            TabIndex        =   22
            Top             =   525
            Width           =   450
         End
         Begin VB.Label lacFromTime 
            Appearance      =   0  'Flat
            Caption         =   "Times- Start"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   525
            Width           =   1215
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
         Left            =   4605
         ScaleHeight     =   3420
         ScaleWidth      =   4455
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CheckBox ckcAllVehicles 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            MaskColor       =   &H8000000F&
            TabIndex        =   36
            Top             =   1680
            Width           =   1815
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   1
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   37
            Top             =   1965
            Width           =   4260
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   0
            Left            =   135
            MultiSelect     =   2  'Extended
            TabIndex        =   34
            Top             =   300
            Width           =   4260
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Feed Names"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            MaskColor       =   &H8000000F&
            TabIndex        =   32
            Top             =   0
            Width           =   1815
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   17
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   15
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
         Width           =   960
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
Attribute VB_Name = "rptSelFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptselfd.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************


' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: rptSelFD.Frm  Revenue Sets Report
'       4-30-02
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection for the Revenue Sets Reports (pacing)
Option Explicit
Option Compare Text
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal x%, ByVal y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function TextOut% Lib "GDI" (ByVal hDC%, ByVal x%, ByVal y%, ByVal lpString$, ByVal nCount%)
'Rm**Declare Function SetTextAlign% Lib "GDI" (ByVal hDC%, ByVal wFlags%)
'Rm**Declare Function GetTextExtent& Lib "GDI" (ByVal hDC%, ByVal lpString$, ByVal nCount%)
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imSetAllVehicles As Integer
Dim imAllVehClicked As Integer
Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
Dim smLogUserCode As String
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
Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
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

Private Sub ckcAllVehicles_Click()
 'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllVehicles.Value = vbChecked Then
        Value = True
    End If
    'End of Coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    ilValue = Value
    If imSetAllVehicles Then
        imAllVehClicked = True
        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(1).hwnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllVehClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcDay_Click(Index As Integer)
    mSetCommands
End Sub

Private Sub ckcReport_Click(Index As Integer)
    mSetCommands
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
    Dim slReptName As String
    Dim ilListIndex As Integer

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
    Screen.MousePointer = vbHourglass
    ilListIndex = lbcRptType.ListIndex
    igUsingCrystal = True
    ilNoJobs = 1
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs

        If ilListIndex = 0 Then         'Feed Recap report
            slReptName = "FeedRecap.Rpt"
        ElseIf ilListIndex = 1 Then
            slReptName = "FdPledge.rpt"     'feed pldge
        ElseIf ilListIndex = PREFEED_DUMP Then      '5-6-10
            slReptName = "PreFeed.rpt"
        End If
        If Not gOpenPrtJob(slReptName) Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        ilRet = gCmcGenFD()
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

       ' Screen.MousePointer = vbHourglass

        If ilListIndex = 0 Then            'feed recap
            gCrFeedRecap
        ElseIf ilListIndex = 1 Then
            gCrFeedPledge                   'feed pledge
        ElseIf ilListIndex = PREFEED_DUMP Then      '5-6-10
            gCrPreFeed
        End If
        
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

Private Sub edcFromDate_Change()
    mSetCommands
End Sub

Private Sub edcFromDate_GotFocus()
    gCtrlGotFocus edcFromDate
End Sub

Private Sub edcFromTime_Change()
    mSetCommands
End Sub

Private Sub edcFromTime_GotFocus()
    gCtrlGotFocus edcFromTime
    mSetCommands
End Sub

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub

Private Sub edcToDate_Change()
    mSetCommands
End Sub

Private Sub edcToTime_Change()
    mSetCommands
End Sub
Private Sub edcToTime_GotFocus()
    gCtrlGotFocus edcToTime
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
    rptSelFD.Refresh
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
    'rptSelFD.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgClfRV
    Erase tgCffRV

    'Erase imCodes
    PECloseEngine
    
    Set rptSelFD = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcSelection_Click(Index As Integer)
If Index = 0 Then
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
    End If
Else
    If Not imAllVehClicked Then
        imSetAllVehicles = False
        ckcAllVehicles.Value = vbUnchecked
        imSetAllVehicles = True
    End If
End If
    mSetCommands
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

    rptSelFD.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"

    pbcSelC.Visible = False
    'lbcRptType.Clear
    'lbcRptType.AddItem smSelectedRptName


    edcFromTime.Text = "12M"
    edcToTime.Text = "12M"
    imAllClicked = False
    imSetAll = True
    imAllVehClicked = False
    imSetAllVehicles = True
    gCenterStdAlone rptSelFD
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
    Dim ilListIndex As Integer
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
    gPopExportTypes cbcFileType     '5-2-02

    frcOption.Enabled = True

    lbcRptType.AddItem "Feed Recap", 0
    lbcRptType.AddItem "Feed Pledges", 1
    lbcRptType.AddItem "Pre-Feed", PREFEED_DUMP     '5-6-10
    If lbcRptType.ListCount > 0 Then
        gFindMatch smSelectedRptName, 0, lbcRptType
        If gLastFound(lbcRptType) < 0 Then
            MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
            imTerminate = True
            Exit Sub
        End If
        lbcRptType.ListIndex = gLastFound(lbcRptType)
    End If

    ilListIndex = lbcRptType.ListIndex

    lbcSelection(0).Clear
    lbcSelection(1).Clear

    'ilRet = gPopUserVehicleBox(rptSelFD, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(1), tgVehicle(), sgVehicleTag)

    If ilListIndex = 0 Then              'feed recap
        ilRet = gFeedNamesPop(rptSelFD, FEED_BYINSERT + FEED_BYNEEDSCONVERT + FEED_BYCONVERTED, lbcSelection(0), tgRptNameCode(), sgRptNameCodeTag)
        ilRet = gPopUserVehicleBox(rptSelFD, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(1), tgVehicle(), sgVehicleTag)
        edcFromTime.Visible = True
        edcToTime.Visible = True
        lacFromTime.Visible = True
        lacToTime.Visible = True
    ElseIf ilListIndex = 0 Then
        ilRet = gFeedNamesPop(rptSelFD, FEED_BYNEEDSCONVERT, lbcSelection(0), tgRptNameCode(), sgRptNameCodeTag)
        ilRet = gPopUserVehicleBox(rptSelFD, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcSelection(1), tgVehicle(), sgVehicleTag)
        lacFromDate.Caption = "Feed Dates- Start"
        lacToDate.Caption = "End"
        lacFromDate.Move 120, 180, 1725     'feed dates-start
        edcFromDate.Left = 1605            'from date input
        lacToDate.Left = 2880
        edcToDate.Left = 3240

        edcFromTime.Visible = False
        edcToTime.Visible = False
        lacFromTime.Visible = False
        lacToTime.Visible = False
    ElseIf ilListIndex = PREFEED_DUMP Then
        ilRet = gPopUserVehicleBox(rptSelFD, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHAIRING + VEHSPORT + ACTIVEVEH, lbcSelection(1), tgVehicle(), sgVehicleTag)
        lacFromDate.Caption = "Active Date"
        lacFromDate.Move 120, 150, 1320
        edcFromDate.Left = 1320
        lacToDate.Visible = False
        edcToDate.Visible = False
        lacFromTime.Visible = False
        edcFromTime.Visible = False
        lacToTime.Visible = False
        edcToTime.Visible = False
        ckcAll.Visible = False
        lbcSelection(0).Visible = False
        ckcAllVehicles.Move 120, 0
        lbcSelection(1).Move 120, 300, 4260, 3000
        plcDays.Move 120, edcFromDate.Top + edcFromDate.Height + 60
        plcDays.Visible = True
        ckcDay(0).Visible = True
        ckcDay(1).Visible = True
        ckcDay(2).Visible = True
        plcReport.Move 120, plcDays.Top + plcDays.Height + 60
        ckcReport(0).Visible = True
        ckcReport(1).Visible = True
        plcReport.Visible = True
    End If

    pbcSelC.Height = pbcSelC.Height - 60
    pbcSelC.Visible = True
    pbcOption.Visible = True

    pbcOption.Visible = True
    pbcOption.Enabled = True
    mSetCommands
    Screen.MousePointer = vbDefault
'    gCenterModalForm RptSel
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
    'gInitStdAlone rptSelFD, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Revenue Sets"
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
    Dim ilListIndex As Integer
    ilEnable = True
    
    ilListIndex = rptSelFD!lbcRptType.ListIndex
    If ilListIndex = PREFEED_DUMP Then
        If edcFromDate.Text = "" Then
            ilEnable = False
        Else
            If (ckcDay(0).Value = vbUnchecked And ckcDay(1).Value = vbUnchecked And ckcDay(2).Value = vbUnchecked) Or (ckcReport(0).Value = vbUnchecked And ckcReport(1).Value = vbUnchecked) Then
                ilEnable = False
            End If
        End If
    Else
        If (edcFromDate.Text = "") Or (edcToDate.Text = "") Or (edcFromTime.Text = "") Or (edcToTime.Text = "") Then
            ilEnable = False
        Else
            'atleast one feed name must be selected
            If ckcAll.Value = vbUnchecked Then
                ilEnable = False
                For ilLoop = 0 To lbcSelection(0).ListCount - 1 Step 1      'vehicle entry must be selected
                If lbcSelection(0).Selected(ilLoop) Then
                    ilEnable = True
                    Exit For
                End If
                Next ilLoop
            End If
        End If
    End If
        'atleast one vehicle must be selected
        If ckcAllVehicles.Value = vbUnchecked Then
            ilEnable = False
            For ilLoop = 0 To lbcSelection(1).ListCount - 1 Step 1      'vehicle entry must be selected
            If lbcSelection(1).Selected(ilLoop) Then
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

    If ilFromCancel Then
        igRptReturn = True
    Else
        igRptReturn = False
    End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload rptSelFD
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        Traffic!cdcSetup.Action = 6
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

