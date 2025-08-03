VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelAv 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facility Report Selection"
   ClientHeight    =   5535
   ClientLeft      =   180
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
      TabIndex        =   17
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
      TabIndex        =   42
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
      TabIndex        =   31
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
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
      Caption         =   "Vehicle Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3690
      Left            =   45
      TabIndex        =   14
      Top             =   1785
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
         ScaleWidth      =   4365
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   255
         Visible         =   0   'False
         Width           =   4365
         Begin V81TrafficReports.CSI_Calendar CSI_CalFrom 
            Height          =   255
            Left            =   1320
            TabIndex        =   21
            Top             =   30
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Text            =   "9/3/2019"
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BorderStyle     =   1
            CSI_ShowDropDownOnFocus=   0   'False
            CSI_InputBoxBoxAlignment=   0
            CSI_CalBackColor=   16777130
            CSI_CalDateFormat=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CSI_CurDayBackColor=   16777215
            CSI_CurDayForeColor=   51200
            CSI_ForceMondaySelectionOnly=   0   'False
            CSI_AllowBlankDate=   0   'False
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   0
         End
         Begin VB.PictureBox plcGrossNet 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4140
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1650
            Width           =   4140
            Begin VB.OptionButton rbcGrossNet 
               Caption         =   "Gross"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   375
               TabIndex        =   38
               Top             =   0
               Value           =   -1  'True
               Width           =   1035
            End
            Begin VB.OptionButton rbcGrossNet 
               Caption         =   "Net"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   1440
               TabIndex        =   39
               Top             =   15
               Width           =   975
            End
         End
         Begin VB.PictureBox plcSelC3 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4005
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1320
            Width           =   4005
            Begin VB.CheckBox ckcSelCInclude 
               Caption         =   "NTR"
               Height          =   240
               Index           =   0
               Left            =   840
               TabIndex        =   35
               Top             =   0
               Width           =   825
            End
            Begin VB.CheckBox ckcSelCInclude 
               Caption         =   "Hard Cost"
               Height          =   240
               Index           =   1
               Left            =   1800
               TabIndex        =   36
               Top             =   0
               Width           =   1305
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
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   41
            Top             =   1920
            Width           =   945
         End
         Begin VB.PictureBox plcSelC1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4140
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1020
            Width           =   4140
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   2160
               TabIndex        =   33
               Top             =   15
               Value           =   -1  'True
               Width           =   1470
            End
            Begin VB.OptionButton rbcSelCSelect 
               Caption         =   "Corporate"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   735
               TabIndex        =   32
               Top             =   0
               Width           =   1275
            End
         End
         Begin VB.PictureBox plcSelC2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   720
            Width           =   4275
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "Quarter"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1380
               TabIndex        =   30
               Top             =   0
               Value           =   -1  'True
               Width           =   1395
            End
            Begin VB.OptionButton rbcSelCInclude 
               Caption         =   "Week"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   420
               TabIndex        =   29
               Top             =   0
               Width           =   870
            End
         End
         Begin VB.TextBox edcSelCTo1 
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
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   26
            Top             =   360
            Width           =   465
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
            Left            =   3240
            MaxLength       =   3
            TabIndex        =   28
            Top             =   360
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox edcSelCTo 
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
            Left            =   720
            MaxLength       =   10
            TabIndex        =   24
            Top             =   360
            Width           =   690
         End
         Begin VB.Label lacContract 
            Caption         =   "Contract #"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Effective Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   27
            Top             =   60
            Width           =   1125
         End
         Begin VB.Label lacSelCTo1 
            Appearance      =   0  'Flat
            Caption         =   "Qtr"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1560
            TabIndex        =   25
            Top             =   420
            Width           =   525
         End
         Begin VB.Label lacSelCFrom1 
            Appearance      =   0  'Flat
            Caption         =   "# Qtrs"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2640
            TabIndex        =   22
            Top             =   420
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label lacSelCTo 
            Appearance      =   0  'Flat
            Caption         =   "Year"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   420
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
         Left            =   4605
         ScaleHeight     =   3420
         ScaleWidth      =   4455
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   4455
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   2970
            ItemData        =   "Rptselav.frx":0000
            Left            =   0
            List            =   "Rptselav.frx":0002
            MultiSelect     =   2  'Extended
            TabIndex        =   45
            Top             =   360
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Advertisers"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   15
            TabIndex        =   43
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
      TabIndex        =   18
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   16
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
         Width           =   1425
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
Attribute VB_Name = "RptSelAv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselav.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelAv.Frm           Advertiser Variance (Pacing report)
'
' Release: 1.0
'
' Description:
'   This file contains the Report selection for Contracts screen code
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
Dim imCodes() As Integer
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
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcSelection.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection.HWnd, LB_SELITEMRANGE, ilValue, llRg)
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
    Dim slReptName As String
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
        slReptName = "AdvVarQt.Rpt"                         'assume monthly option
        If RptSelAv!rbcSelCInclude(0).Value Then            'Weekly option
            slReptName = "AdvVarWk.Rpt"
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
        ilRet = gCmcGenAv(imGenShiftKey, smLogUserCode)
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

        gCrAdvVariance
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
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
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
    'cdcSetup.Flags = cdlPDPrintSetup
    'cdcSetup.Action = 5    'DLG_PRINT
    cdcSetup.flags = cdlPDPrintSetup
    cdcSetup.ShowPrinter
End Sub

Private Sub CSI_CalFrom_CalendarChanged()
    mSetCommands
End Sub

Private Sub CSI_CalFrom_GotFocus()
    gCtrlGotFocus CSI_CalFrom
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
Private Sub edcSelCFrom1_Change()
    mSetCommands
End Sub
Private Sub edcSelCFrom1_GotFocus()
    gCtrlGotFocus edcSelCFrom1
End Sub
Private Sub edcSelCTo_GotFocus()
    gCtrlGotFocus edcSelCTo
    mSetCommands
End Sub
Private Sub edcSelCTo1_Change()
    mSetCommands
End Sub
Private Sub edcSelCTo1_GotFocus()
    gCtrlGotFocus edcSelCTo1
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
    RptSelAv.Refresh
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
    'RptSelAv.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase tgCSVNameCode
    Erase tgClfAV
    Erase tgCffAV
    Erase imCodes
    PECloseEngine
    
    Set RptSelAv = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcSelection_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
    End If
    mSetCommands
End Sub
Private Sub mAdvtPop(RptForm As Form, lbcSelection As Control)
'
'   mAdvtPop
'   Where:
'       RptForm as Form
'       lbcSelection as control
'
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(RptSelCt, lbcSelection, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(RptForm, lbcSelection, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", RptForm
        On Error GoTo 0
    End If
    Exit Sub
mAdvtPopErr:
 On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'**********************************************************************************
'
'                   mAskCorpOrStd - Ask on Report input screen:
'                           Month  o Corp   o Std
'                   Set the proper properties to this control
'                   Created: DH 9/9/96
'*********************************************************************************
'
'Private Sub mAskCorpOrStd()
'Dim ilRet As Integer
'    'plcSelC1.Caption = "Month"
'    rbcSelCSelect(0).Caption = "Corporate"
'    rbcSelCSelect(0).Left = 660
'    rbcSelCSelect(0).Width = 1140
'    rbcSelCSelect(1).Caption = "Standard"
'    rbcSelCSelect(1).Left = 1840
'    rbcSelCSelect(1).Width = 1200
'    rbcSelCSelect(0).Visible = True
'    rbcSelCSelect(1).Visible = True
'    rbcSelCSelect(1).Value = True
'    If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
'        rbcSelCSelect(0).Enabled = False
'    Else
'        rbcSelCSelect(0).Value = True
'        ilRet = gObtainCorpCal()
'    End If
'    plcSelC1.Visible = True
'End Sub
'
'
'                   mAskEffDate - Ask Effective Date, Start Year
'                                 and Quarter
'
'                   6/7/97
'
'
'Private Sub mAskEffDate()
'    lacSelCFrom.Left = 120
'    edcSelCFrom.Move 1350, edcSelCFrom.Top, 945
'    edcSelCFrom.MaxLength = 10  '8    5/27/99 chged for short form m/d/yyyy
'    lacSelCFrom.Caption = "Effective Date"
'    lacSelCFrom.Top = 75
'    lacSelCFrom.Visible = True
'    edcSelCFrom.Visible = True
'    lacSelCTo.Caption = "Year"
'    lacSelCTo.Visible = True
'    lacSelCTo.Left = 120
'    lacSelCTo.Top = edcSelCFrom.Top + edcSelCFrom.Height + 75
'    lacSelCTo1.Left = 1580
'    lacSelCTo1.Caption = "Quarter"
'    lacSelCTo1.Width = 810
'    lacSelCTo1.Top = edcSelCFrom.Top + edcSelCFrom.Height + 75
'    lacSelCTo1.Visible = True
'    edcSelCTo.Move 600, edcSelCFrom.Top + edcSelCFrom.Height + 30, 600
'    edcSelCTo1.Move 2340, edcSelCFrom.Top + edcSelCFrom.Height + 30, 300
'    edcSelCTo.MaxLength = 4
'    edcSelCTo1.MaxLength = 1
'    edcSelCTo.Visible = True
'    edcSelCTo1.Visible = True
'End Sub
'**********************************************************************************
'
'                   mAskCorpOrStd - Ask on Report input screen:
'                           Month  o Corp   o Std
'                   Set the proper properties to this control
'                   Created: DH 9/9/96
'*********************************************************************************
'
'Private Sub mAskWkOrQtr()
'    'plcSelC2.Caption = "By"
'    rbcSelCInclude(0).Caption = "Week"
'    rbcSelCInclude(0).Left = 360
'    rbcSelCInclude(0).Width = 800
'    rbcSelCInclude(1).Caption = "Quarter"
'    rbcSelCInclude(1).Move 1260, rbcSelCInclude(0).Top, 1140
'    'rbcSelCInclude(1).Width = 1140
'    rbcSelCInclude(1).Value = True
'    rbcSelCInclude(0).Visible = True
'    rbcSelCInclude(1).Visible = True
'    plcSelC2.Visible = True
'End Sub
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

    RptSelAv.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
'    lacSelCFrom.Move 120, 75
'    lacSelCTo.Move 120, 390
'    lacSelCFrom1.Move 2400, 75
'    lacSelCTo1.Move 2400, 390
'    edcSelCFrom.Move 1500, 30
'    edcSelCFrom1.Move 3240, 30
'    edcSelCTo.Move 1500, 345
'    edcSelCTo1.Move 2715, 345
'    plcSelC1.Move 120, 675
    pbcSelC.Move 90, 255, 4515, 3360

    gCenterStdAlone RptSelAv
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
    'Setup report output types
    gPopExportTypes cbcFileType '10-20-01
    'cbcFileType.AddItem "Report"
    'cbcFileType.AddItem "Fixed Column Width"
    'cbcFileType.AddItem "Comma-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated with Quotes"
    'cbcFileType.AddItem "Tab-Separated w/o Quotes"
    'cbcFileType.AddItem "DIF"
    'cbcFileType.ListIndex = 0
    pbcSelC.Visible = False
    'lbcRptType.Clear
    'lbcRptType.AddItem smSelectedRptName

    lbcSelection.Clear
    mAdvtPop RptSelAv, lbcSelection
    lbcSelection.Tag = ""
    Screen.MousePointer = vbHourglass

    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    pbcSelC.Visible = True
    pbcOption.Visible = True
'    mAskEffDate             'ask effective date, year & qtr

'    plcSelC2.Move 120, edcSelCTo.Top + edcSelCTo.Height + 30
'    mAskWkOrQtr
'    plcSelC1.Move 120, plcSelC1.Top + plcSelC1.Height
'    mAskCorpOrStd
    If tgSpf.sRUseCorpCal <> "Y" Then       'if Using corp cal, dfault it; otherwise disable it
        rbcSelCSelect(0).Enabled = False
    Else
        rbcSelCSelect(0).Value = True
        ilRet = gObtainCorpCal()
    End If
    plcSelC1.Visible = True
    plcSelC2.Visible = True
    plcSelC3.Visible = True
    
    lbcSelection.Visible = True                  'show budget name list box (base budget)
    lbcSelection.Move 15, ckcAll.Top + ckcAll.Height, 4380, 3000
    pbcOption.Visible = True
    pbcOption.Enabled = True
    ' Dan M 6-25-08  NTR/Hard Cost options
'    mAddNTRHardCost
    ' Dan M 7-7-08 contract selectivity
'    lacSelC4.Move 120, 30
'    edcSelC4.Left = lacSelC4.Left + lacSelC4.Width
'    edcSelC4.Visible = True
'    plcSelC4.Move 0, plcSelC3.Top + plcSelC3.Height
'    plcSelC4.Visible = True
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
    'gInitStdAlone RptSelAv, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    smSelectedRptName = "Advertiser Variance"
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
'    If (edcSelCFrom.Text <> "") And (edcSelCTo.Text <> "") And (edcSelCTo1.Text <> "") Then
    If (CSI_CalFrom.Text <> "") And (edcSelCTo.Text <> "") And (edcSelCTo1.Text <> "") Then     '9-3-19 use csi cal control vs edit box
        'atleast one advertiser must be selected
        For ilLoop = 0 To lbcSelection.ListCount - 1 Step 1      'vehicle entry must be selected
            If lbcSelection.Selected(ilLoop) Then
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
    Unload RptSelAv
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
        mInitReport
        If imTerminate Then 'Used for print only
            mTerminate True
            Exit Sub
        End If
        imFirstTime = False
    End If
    gCtrlGotFocus rbcOutput(Index)
End Sub
Private Sub rbcSelCInclude_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelCInclude(Index).Value
    'End of coded added
    If Index = 0 Then           'week, ask # qtrs to produce
        edcSelCFrom1.Text = 4
        edcSelCFrom1.Visible = True
        lacSelCFrom1.Visible = True
'        lacSelCFrom1.Caption = "# Qtrs"
'        lacSelCFrom1.Move edcSelCTo1.Left + edcSelCTo1.Width + 240, lacSelCTo1.Top
'        edcSelCFrom1.Move lacSelCFrom1.Left + 600, edcSelCTo1.Top, 300
    Else                        'qtr, always show 4 qtrs, dont ask for # qtrs
        edcSelCFrom1.Text = 4
        edcSelCFrom1.Visible = False
        lacSelCFrom1.Visible = False
    End If
End Sub
Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'Terminate False
End Sub
Private Sub plcSelC1_Paint()
    plcSelC1.CurrentX = 0
    plcSelC1.CurrentY = 0
    plcSelC1.Print "Month"
End Sub
Private Sub plcSelC2_Paint()
    plcSelC2.CurrentX = 0
    plcSelC2.CurrentY = 0
    plcSelC2.Print "By"
End Sub
Private Sub plcSelC3_Paint()
    plcSelC3.CurrentX = 0
    plcSelC3.CurrentY = 0
    plcSelC3.Print "Include"
End Sub
Private Sub plcGrossNet_Paint()
    plcGrossNet.CurrentX = 0
    plcGrossNet.CurrentY = 0
    plcGrossNet.Print "By"
End Sub

'Private Sub mAddNTRHardCost()
'lbcSelCInclude.Caption = "Include "
'lbcSelCInclude.Move 0, 0, 750
'ckcSelCInclude(0).Caption = "NTR"
'ckcSelCInclude(0).Move lbcSelCInclude.Width + lbcSelCInclude.Left, 0, 750
'Load ckcSelCInclude(1)
'ckcSelCInclude(1).Caption = "Hard Cost"
'ckcSelCInclude(1).Move ckcSelCInclude(0).Left + ckcSelCInclude(0).Width, 0, 2000
'plcSelC3.Move 150, plcSelC1.Top + plcSelC1.Height
'lbcSelCInclude.Visible = True
'ckcSelCInclude(0).Visible = True
'ckcSelCInclude(1).Visible = True
'plcSelC3.Visible = True
'End Sub
