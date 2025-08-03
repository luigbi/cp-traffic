VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelEngrLk 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enginerring Missing Links"
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
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   8115
      TabIndex        =   24
      Top             =   90
      Visible         =   0   'False
      Width           =   1095
   End
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
      TabIndex        =   21
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
      TabIndex        =   19
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
      Caption         =   "Engineering Missing Links"
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
         Left            =   120
         ScaleHeight     =   3360
         ScaleWidth      =   4320
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   225
         Visible         =   0   'False
         Width           =   4320
         Begin VB.CheckBox ckcDays 
            Caption         =   "Sun"
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   30
            Top             =   840
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox ckcDays 
            Caption         =   "Sat"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   29
            Top             =   840
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox ckcDays 
            Caption         =   "M-F"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.ListBox lbcEvtNameCode 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   0
            ItemData        =   "RptSelEngrLinks.frx":0000
            Left            =   1200
            List            =   "RptSelEngrLinks.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   3000
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.ListBox lbcFeedCode 
            Appearance      =   0  'Flat
            Height          =   240
            ItemData        =   "RptSelEngrLinks.frx":0004
            Left            =   1800
            List            =   "RptSelEngrLinks.frx":000B
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1800
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.ListBox lbcEvtName 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   0
            ItemData        =   "RptSelEngrLinks.frx":0014
            Left            =   120
            List            =   "RptSelEngrLinks.frx":001B
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   3000
            Visible         =   0   'False
            Width           =   3210
         End
         Begin VB.ListBox lbcEvtType 
            Appearance      =   0  'Flat
            Height          =   240
            ItemData        =   "RptSelEngrLinks.frx":002C
            Left            =   0
            List            =   "RptSelEngrLinks.frx":0033
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   2520
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.ListBox lbcSubfeed 
            Appearance      =   0  'Flat
            Height          =   240
            ItemData        =   "RptSelEngrLinks.frx":0043
            Left            =   0
            List            =   "RptSelEngrLinks.frx":004A
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   2160
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.ListBox lbcFeed 
            Appearance      =   0  'Flat
            Height          =   240
            ItemData        =   "RptSelEngrLinks.frx":005D
            Left            =   0
            List            =   "RptSelEngrLinks.frx":0064
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1800
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.TextBox edcEndDate 
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
            MaxLength       =   10
            TabIndex        =   27
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox ckcDiscrepOnly 
            Caption         =   "Discrepancies Only"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.TextBox edcStartDate 
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
            MaxLength       =   10
            TabIndex        =   25
            Top             =   120
            Width           =   1215
         End
         Begin MSComDlg.CommonDialog cdcSetup 
            Left            =   3480
            Top             =   2880
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DefaultExt      =   ".Txt"
            Filter          =   "*.Txt|*.Txt|*.Doc|*.Doc|*.Asc|*.Asc"
            FilterIndex     =   1
            FontSize        =   0
            MaxFileSize     =   256
         End
         Begin VB.Label lacEndDate 
            Appearance      =   0  'Flat
            Caption         =   "End Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   840
         End
         Begin VB.Label lacStartDate 
            Appearance      =   0  'Flat
            Caption         =   "Start Date"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   120
            Width           =   945
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
            ItemData        =   "RptSelEngrLinks.frx":0072
            Left            =   15
            List            =   "RptSelEngrLinks.frx":0074
            MultiSelect     =   2  'Extended
            TabIndex        =   33
            Top             =   360
            Width           =   4380
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   90
            Width           =   1305
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
      Top             =   105
      Width           =   2805
   End
   Begin VB.Frame frcOutput 
      Caption         =   "Report Destination"
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   30
      TabIndex        =   0
      Top             =   75
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
Attribute VB_Name = "RptSelEngrLk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptSelEngrLk.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************


' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelEngrLk.Frm - Network Program Schedule (Radar Worksheet)
'         5-13-03
'
' Release: 5.1
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
Dim smLogUserCode As String
Dim imTerminate As Integer
Dim imVeh As Integer
Dim bmFirstCallToVpfFind As Boolean
Dim imVpfIndex As Integer
Dim smDateFilter As String
Dim smEndDate As String
Dim imEndDate0 As Integer
Dim imEndDate1 As Integer
Dim imTFNDay As Integer
Dim imVefCode As Integer
Dim tmDEvt() As DELEVT             'Current Event image
Dim imDate0 As Integer
Dim imDate1 As Integer
Dim imDelIndex As Integer
Dim imMnfFeed As Integer
Dim imDateCode As Integer
Dim imNoEnfSaved  As Integer
Dim tmNameCode() As SORTCODE
Dim smNameCodeTag As String
Dim tmEvtTypeCode() As SORTCODE
Dim smEvtTypeCodeTag As String
Dim tmSubFeedCode() As SORTCODE
Dim smSubFeedCodeTag As String
Dim imEnfCode(0 To 20) As Integer   'Save last 20 used. Index zero ignored
Dim smEnfName(0 To 20) As String

Dim hmLvf As Integer            'Log version file handle
Dim tmLvf As LVF                'LVF record image
Dim tmLvfSrchKey As LONGKEY0     'LVF Key 0 image
Dim imLvfRecLen As Integer      'LVF record length

Dim hmLef As Integer
Dim tmLef As LEF
Dim imLefRecLen As Integer
Dim tmLefSrchKey As LEFKEY0

Dim hmEnf As Integer
Dim tmEnf As ENF
Dim imEnfRecLen As Integer
Dim tmEnfSrchKey As INTKEY0

Dim hmLcf As Integer
Dim tmLcf As LCF
Dim imLcfRecLen As Integer
Dim tmLcfSrchKey As LCFKEY0

'Delivery links and Engineering are of the same format
Dim hmDlf As Integer            'Delivery Vehicle link file handle
Dim tmDlf() As DLFLIST                'DLF record image
Dim tmDlfSrchKey As DLFKEY0            'DLF record image
Dim tmDlfSrchKey1 As LONGKEY0
Dim imDlfRecLen As Integer        'VLF record length

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer         'VEF record length

Dim hmCbf As Integer
Dim tmCbf As CBF
Dim imCbfRecLen As Integer

Dim hmMnf As Integer
Dim tmMnf As MNF
Dim imMnfRecLen As Integer
Dim tmMnfSrchKey As INTKEY0







'*******************************************************
'*                                                     *
'*      Procedure Name:mConvAirVeh                     *
'*                                                     *
'*       Created:7-1-19         By:D. Hosaka       *
'*       Modified:              By:                    *
'*                                                     *
'*       Comments: Populate the selection combobox with*
'*       conventional w/feed, sports & airing vehicles *
'*                                                     *
'*                                                     *
'*******************************************************
Private Sub mConvAirVeh()
    Dim ilRet As Integer
    
    ilRet = gPopVehFeedBox(RptSelEngrLk, VEHCONV_W_FEED + VEHAIRING + VEHEXCLUDESPORT + ACTIVEVEH, lbcSelection, tgCSVNameCode(), sgCSVNameCodeTag)
  
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mConvAirVehERr
        gCPErrorMsg ilRet, "mConvAirVeh (gPopVehFeedBox: Vehicle)", RptSelEngrLk
        On Error GoTo 0
    End If
    Exit Sub
mConvAirVehERr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
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
        llRg = CLng(lbcSelection.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection.hWnd, LB_SELITEMRANGE, ilValue, llRg)
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
    Dim ilListIndex As Integer      '8-29-02
    Dim slEffectiveDate As String
    Dim llEDate As Long
    Dim llSDate As Long
    Dim slStr As String
    Dim slSelection As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim slDate As String
    Dim slTime As String
    
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

    igUsingCrystal = True
    ilNoJobs = 1
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs
        
        smDateFilter = Trim$(RptSelEngrLk!edcStartDate.Text)   'Store Effective Date
        gPackDate smDateFilter, imDate0, imDate1

        smEndDate = Trim$(RptSelEngrLk!edcEndDate.Text)   'Store Effective Date
        If (Trim$(smEndDate) = "") Or (Trim$(smEndDate) = "TFN") Then
'            imEndDate0 = 0
'            imEndDate1 = 0
            smEndDate = "12/31/2069"
        End If
            gPackDate smEndDate, imEndDate0, imEndDate1
 '       End If

        'verify validity of date input
        If Not gValidDate(smDateFilter) Then
            mReset
            RptSelEngrLk!edcStartDate.SetFocus
            Exit Sub
        End If
        If Not gValidDate(smEndDate) Then
            mReset
            RptSelEngrLk!edcEndDate.SetFocus
            Exit Sub
        End If
        
        'open the crystal report
        If Not gOpenPrtJob("EngrMissingLinks.rpt") Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Exit Sub
        End If
        
        If smEndDate = "12/31/2069" Then
            slStr = smDateFilter & " - TFN"
        Else
            slStr = smDateFilter & " - " & smEndDate
        End If
    
        'send start/end dates for crystal report heading
         If Not gSetFormula("DatesHdr", "'" & slStr & "'") Then
            MsgBox "Invalid DatesHdr Formula in EngrLinks.rpt; Call Counterpoint"
            mRestoreFields
            Exit Sub
         End If
         
        If CkcDiscrepOnly.Value = vbChecked Then           'discrepancies only
            If Not gSetFormula("DiscrepOnly", "'" & "D" & "'") Then     'Discreps only
                MsgBox "Invalid DiscrepsOnly Formula in EngrLinks.rpt; Call Counterpoint"
                mRestoreFields
                Exit Sub
            End If
                
        Else
            If Not gSetFormula("DiscrepOnly", "'" & "A" & "'") Then     'All
                MsgBox "Invalid DiscrepsOnly Formula in EngrLinks.rpt; Call Counterpoint"
                mRestoreFields
                Exit Sub
            End If
        End If

        gCurrDateTime slDate, slTime, slMonth, slDay, slYear
        slSelection = "{CBF_Contract_BR.cbfGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({CBF_Contract_BR.cbfGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))
        If Not gSetSelection(slSelection) Then
            MsgBox "Invalid Report Selection; RptSelEngrLk (cmcGen)"
            mRestoreFields
            Exit Sub
        End If

        Screen.MousePointer = vbHourglass

        If Not mOpenLinksFiles() Then       'open all the applicable files
            Exit Sub
        End If
        
        mEngrMissingLinksRpt                'generate all the engr links and whats missing

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
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '2-21-02
        End If
    Next ilJobs
    imGenShiftKey = 0

    Screen.MousePointer = vbHourglass
    gCrCbfClear
    Screen.MousePointer = vbDefault
    mRestoreFields
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
Private Sub edcStartDate_Change()
    mSetCommands
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
    'RptSelEngrLk.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    PECloseEngine
    mCloseLinkFiles
    Set RptSelEngrLk = Nothing   'Remove data segment
    
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub lbcSelection_Click()

     If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked  '9-12-02 False
        imSetAll = True
    End If
    mSetCommands
End Sub
Private Sub lbcSelection_GotFocus()
    gCtrlGotFocus lbcSelection
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

    RptSelEngrLk.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    ilLoop = InStr(slStr, "&")
    If ilLoop > 0 Then
        slStr = Left$(slStr, ilLoop - 1) & "&&" & Mid$(slStr, ilLoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
    ckcAll.Move 30, 60
    lbcSelection.Move 15, ckcAll.Height + 90, 4380
    pbcSelC.Move 90, 255, 4515, 3360
    bmFirstCallToVpfFind = True
    gCenterStdAlone RptSelEngrLk
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
    lbcRptType.Clear

    lbcSelection.Clear
    lbcSelection.Tag = ""
    Screen.MousePointer = vbHourglass

    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    pbcSelC.Visible = True
    pbcOption.Visible = True

    ckcAll.Visible = True
    ckcAll.Value = vbUnchecked                 'default to no vehicles selected
    mConvAirVeh


    pbcOption.Visible = True
    pbcSelC.Visible = True
    frcOption.Enabled = True


    If lbcRptType.ListCount > 0 Then
        gFindMatch smSelectedRptName, 0, lbcRptType
        If gLastFound(lbcRptType) < 0 Then
            MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
            imTerminate = True
            Exit Sub
        End If
        lbcRptType.ListIndex = gLastFound(lbcRptType)
    End If
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
    'gInitStdAlone RptSelEngrLk, slStr, ilTestSystem
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
    If (edcStartDate.Text <> "") Then
        ilEnable = True

        If ilEnable Then
            ilEnable = False
            If ckcAll.Value = vbChecked Then
                ilEnable = True
            Else
                For ilLoop = 0 To lbcSelection.ListCount - 1 Step 1      'market entry must be selected
                    If lbcSelection.Selected(ilLoop) Then
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
    Unload RptSelEngrLk
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

Sub mReset()
    igGenRpt = False
    RptSelRD!frcOutput.Enabled = igOutput
    RptSelRD!frcCopies.Enabled = igCopies
    RptSelRD!frcFile.Enabled = igFile
    RptSelRD!frcOption.Enabled = igOption
    Beep
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name: mBuildDlfRec                   *
'*                                                     *
'*             Created:7/27/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build Dlf records images from   *
'*                     old dlf and events for day      *
'*                                                     *
'*******************************************************
'Private Sub mBuildDlfRec()
Private Sub mEngrMissingLinksRpt()
    Dim ilLoop As Integer
    Dim ilVeh As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slDay As String * 1
    Dim ilUpperBound As Integer
    Dim ilStatus As Integer
    Dim ilDEvt As Integer
    Dim slTime As String
    Dim llTime As Long
    Dim llMatchTime As Long
    Dim ilMatchEtf As Integer
    Dim ilMatchEnf As Integer
    Dim ilExcluded As Integer
    Dim tlDlf As DLF
    Dim ilLoopOnDays As Integer                   '1 = mon-fri, 2 = sat , 3 = sun
    Dim ilZone As Integer
    Dim ilFound As Integer
    Dim slName As String
    Dim slDateFilter As String
    Dim blProcessDay As Boolean
    
        'Most of this code has been copied from LinkDLVY code with modifications to produce the Engineering Missing Links report.
        'Some variables may be set and tested that are basically hard coded
        imDlfRecLen = Len(tlDlf)  'Get and save DlF record length
        imDelIndex = 1
        lbcFeed.Clear
        lbcSubfeed.Clear
        imMnfFeed = -1
   
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmCbf.lGenTime = lgNowTime
        tmCbf.iGenDate(0) = igNowDate(0)
        tmCbf.iGenDate(1) = igNowDate(1)

        slDateFilter = smDateFilter
        For ilVeh = 0 To lbcSelection.ListCount - 1 Step 1
            If lbcSelection.Selected(ilVeh) Then
                slNameCode = tgCSVNameCode(ilVeh).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                imVefCode = Val(slCode)
                tmVefSrchKey.iCode = imVefCode
                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                If ilRet = BTRV_ERR_NONE Then
                    If bmFirstCallToVpfFind Then
                        imVpfIndex = gVpfFind(RptSelEngrLk, imVefCode)
                        bmFirstCallToVpfFind = False
                    Else
                        imVpfIndex = gVpfFindIndex(imVefCode)
                    End If

                    For ilLoopOnDays = 1 To 3
                        ReDim tmDEvt(0 To 0) As DELEVT          'init the array of avails for vehicle to process
                        ReDim tmDlf(0 To 0) As DLFLIST
                        smDateFilter = slDateFilter
                        blProcessDay = False
                        If ilLoopOnDays = 1 And RptSelEngrLk.ckcDays(0).Value = vbChecked Then
                            slDay = "0"
                            imDateCode = 0
                            If gWeekDayStr(smDateFilter) <= 4 Then
                                imTFNDay = gWeekDayStr(smDateFilter)
                                smDateFilter = gObtainPrevMonday(smDateFilter)
                            Else
                                smDateFilter = gObtainNextMonday(smDateFilter)
                                imTFNDay = gWeekDayStr(smDateFilter)
                            End If
                            blProcessDay = True
                        ElseIf ilLoopOnDays = 2 And RptSelEngrLk.ckcDays(1).Value = vbChecked Then
                            imDateCode = 6
                            slDay = "6"
                            smDateFilter = gDecOneDay(gObtainNextSunday(smDateFilter))
                            imTFNDay = gWeekDayStr(smDateFilter)
                            blProcessDay = True
                        ElseIf RptSelEngrLk.ckcDays(2).Value = vbChecked Then
                            imDateCode = 7
                            slDay = "7"
                            smDateFilter = gObtainNextSunday(smDateFilter)
                            imTFNDay = gWeekDayStr(smDateFilter)
                            blProcessDay = True
                       End If
                        If blProcessDay Then
                            mReadLcf tmVef.sType
                                   
                            ilDEvt = LBound(tmDEvt)
                            llMatchTime = -1
                            ilMatchEtf = -1
                            ilMatchEnf = -1
                            tmDlfSrchKey.iVefCode = imVefCode
                            tmDlfSrchKey.sAirDay = Trim(str$(imDateCode))
                            tmDlfSrchKey.iStartDate(0) = imDate0  'Year 1/1/1900
                            tmDlfSrchKey.iStartDate(1) = imDate1
                            tmDlfSrchKey.iAirTime(0) = 0
                            tmDlfSrchKey.iAirTime(1) = 25 * 256 'Hour
                            ilRet = btrGetLessOrEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        '                    If imDelIndex = 0 Then
        '                        Do While (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay) And (tlDlf.iMnfFeed <> imMnfFeed)
        '                            ilRet = btrGetPrevious(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        '                        Loop
        '                    End If
                            If (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay) And (((imDelIndex = 1) And (tlDlf.iMnfFeed > 0)) Or ((tlDlf.iMnfFeed = imMnfFeed) And (imDelIndex = 0))) Then
                                'Start at earliest time and merge Lcf
                                tmDlfSrchKey.iVefCode = imVefCode
                                tmDlfSrchKey.sAirDay = slDay
                                tmDlfSrchKey.iStartDate(0) = tlDlf.iStartDate(0)  'Year 1/1/1900
                                tmDlfSrchKey.iStartDate(1) = tlDlf.iStartDate(1)
                                tmDlfSrchKey.iAirTime(0) = 0
                                tmDlfSrchKey.iAirTime(1) = 0
                                ilRet = btrGetGreaterOrEqual(hmDlf, tlDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                Do While (ilRet = BTRV_ERR_NONE) And (tlDlf.iVefCode = imVefCode) And (tlDlf.sAirDay = slDay) And ((imMnfFeed = -1) Or (tlDlf.iMnfFeed = imMnfFeed))
                                    If ((tlDlf.iTermDate(0) = 0) And (tlDlf.iTermDate(1) = 0)) Or ((tlDlf.iTermDate(1) > imDate1) Or ((tlDlf.iTermDate(1) = imDate1) And (tlDlf.iTermDate(0) >= imDate0))) And ((imMnfFeed = -1) Or (tlDlf.iMnfFeed = imMnfFeed)) Then
                                        'Test if time still exist or record should be deleted
                                        ilStatus = 2    'Remove old or events that don't belong
                                        ilExcluded = False
            '                            If imDelOrEngr = 0 Then 'all events- only version one
            '                                If (tlDlf.sCmmlSched = "N") And (tlDlf.iMnfSubFeed = 0) Then
            '                                    ilExcluded = True
            '                                End If
            '                            Else    'Only avails
                                            'always only need avails
                                            If (tlDlf.sFed = "N") And (tlDlf.iMnfSubFeed = 0) Then
                                                ilExcluded = True
                                            End If
            '                            End If
                                        If (tlDlf.iStartDate(1) > imDate1) Or ((tlDlf.iStartDate(1) = imDate1) And (tlDlf.iStartDate(0) > imDate0)) Then
                                            ilExcluded = True
                                        End If
                                        If Not ilExcluded Then
                                            gUnpackTime tlDlf.iAirTime(0), tlDlf.iAirTime(1), "A", "1", slTime
                                            llTime = CLng(gTimeToCurrency(slTime, True))
                                            For ilDEvt = LBound(tmDEvt) To UBound(tmDEvt) - 1 Step 1
                                                If (llTime = tmDEvt(ilDEvt).lTime) And (tlDlf.iEtfCode = tmDEvt(ilDEvt).iEtfCode) And (tlDlf.iEnfCode = tmDEvt(ilDEvt).iEnfCode) Then
                                                    ilStatus = 1
                                                    tmDEvt(ilDEvt).iStatus = 1  'Used
                                                    'mResetDlfRec tmDEvt(ilDEvt), tlDlf
                                                    Exit For
                                                End If
                                            Next ilDEvt
                                        End If
                
                                        ilUpperBound = UBound(tmDlf)
                                        tmDlf(ilUpperBound).DlfRec = tlDlf
                                        tmDlf(ilUpperBound).iStatus = ilStatus
                                        tmDlf(ilUpperBound).lDlfCode = tlDlf.lCode
        '                                mEvtString tmDlf(ilUpperBound)
                                        ilUpperBound = ilUpperBound + 1
                                        ReDim Preserve tmDlf(0 To ilUpperBound) As DLFLIST
            '                            If ilStatus = 2 Then    'Set change so update can be pressed without changing any other field
            '                                imChg = True
            '                            End If
                                    End If
                                    ilRet = btrGetNext(hmDlf, tlDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                Loop
                                'Add any not already defined by delivery links
                                If LBound(tmDEvt) = UBound(tmDEvt) Then             'no lcf found
                                    For ilLoop = LBound(tmDlf) To UBound(tmDlf) - 1 'loop thru all the events to indicate on report that links exist, but no lcf.
                                                                                    'nothing can be done for this, just to ignore investigating.
                                        tmDlf(ilLoop).iStatus = 3                   'links exist, no libraray
                                    Next ilLoop
                                Else
                                    For ilLoop = LBound(tmDEvt) To UBound(tmDEvt) - 1 Step 1
                                        If tmDEvt(ilLoop).iStatus = 0 Then
                                            mMakeDlfRec tmDEvt(ilLoop)
                                            tmDEvt(ilLoop).iStatus = 1
                                        End If
                                    Next ilLoop
                                End If
                            Else    'Build from Lcf
                                For ilLoop = LBound(tmDEvt) To UBound(tmDEvt) - 1 Step 1
                                    mMakeDlfRec tmDEvt(ilLoop)
                                    tmDEvt(ilLoop).iStatus = 1
                                Next ilLoop
                            End If
                            mCreateForOutput
                        End If
                    Next ilLoopOnDays
                End If
            End If
        Next ilVeh

    Exit Sub
    
mBuildDlfRecErr:
    Resume Next
    
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name: mMakeDlfRec                    *
'*                                                     *
'*             Created:7/27/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:make Dlf records images from    *
'*                     events for day                  *
'*                                                     *
'*******************************************************
Private Sub mMakeDlfRec(tlDEvt As DELEVT)
    Dim ilUpperBound As Integer
    Dim clTime As Currency
    Dim slTime As String
    Dim ilZone As Integer
    Dim ilTimeAdj As Integer
    Dim ilDispl As Integer
    Dim ilCreate As Integer
    Dim ilNumVar As Integer
    Dim ilVff As Integer

'    If imDelOrEngr = 0 Then 'Delivery- only create one version for all events types
'        ilNumVar = 1
'    Else    'Engineering-  only create avails (all versions)
        If (tlDEvt.iEtfCode >= 2) And (tlDEvt.iEtfCode <= 9) Then  'Avail
            ilNumVar = 4    'Create all versions
        Else
            Exit Sub   'Ignore all event types except avails
        End If
'    End If
    ilUpperBound = UBound(tmDlf)
    For ilZone = LBound(tgVpf(imVpfIndex).sGZone) To UBound(tgVpf(imVpfIndex).sGZone) Step 1
        If (Trim$(tgVpf(imVpfIndex).sGZone(ilZone)) <> "") And (((tgVpf(imVpfIndex).iGMnfNCode(ilZone) > 0) And (imDelIndex = 1)) Or ((tgVpf(imVpfIndex).iGMnfNCode(ilZone) = imMnfFeed) And (imDelIndex = 0))) Then
            For ilDispl = 1 To ilNumVar Step 1
                ilCreate = False
                Select Case ilDispl
                    Case 1  'Primary
                        ilTimeAdj = 60 * tgVpf(imVpfIndex).iGV1Z(ilZone)
                        ilCreate = True
                    Case 2
                        If tgVpf(imVpfIndex).iGV2Z(ilZone) <> 0 Then
                            ilTimeAdj = 60 * tgVpf(imVpfIndex).iGV2Z(ilZone)
                            ilCreate = True
                        End If
                    Case 3
                        If tgVpf(imVpfIndex).iGV3Z(ilZone) <> 0 Then
                            ilTimeAdj = 60 * tgVpf(imVpfIndex).iGV3Z(ilZone)
                            ilCreate = True
                        End If
                    Case 4
                        If tgVpf(imVpfIndex).iGV4Z(ilZone) <> 0 Then
                            ilTimeAdj = 60 * tgVpf(imVpfIndex).iGV4Z(ilZone)
                            ilCreate = True
                        End If
                End Select
                If ilCreate Then
                    tmDlf(ilUpperBound).DlfRec.iVefCode = imVefCode
                    tmDlf(ilUpperBound).DlfRec.sAirDay = Trim(str$(imDateCode))
                    clTime = tlDEvt.lTime
                    slTime = gCurrencyToTime(clTime)
                    gPackTime slTime, tmDlf(ilUpperBound).DlfRec.iAirTime(0), tmDlf(ilUpperBound).DlfRec.iAirTime(1)
                    clTime = tlDEvt.lTime + 3600 * tgVpf(imVpfIndex).iGLocalAdj(ilZone) + ilTimeAdj
                    slTime = gCurrencyToTime(clTime)
                    gPackTime slTime, tmDlf(ilUpperBound).DlfRec.iLocalTime(0), tmDlf(ilUpperBound).DlfRec.iLocalTime(1)
                    clTime = tlDEvt.lTime + 3600 * tgVpf(imVpfIndex).iGFeedAdj(ilZone) + ilTimeAdj
                    slTime = gCurrencyToTime(clTime)
                    gPackTime slTime, tmDlf(ilUpperBound).DlfRec.iFeedTime(0), tmDlf(ilUpperBound).DlfRec.iFeedTime(1)
                    tmDlf(ilUpperBound).DlfRec.sZone = tgVpf(imVpfIndex).sGZone(ilZone)
                    tmDlf(ilUpperBound).DlfRec.iEtfCode = tlDEvt.iEtfCode
                    tmDlf(ilUpperBound).DlfRec.iEnfCode = tlDEvt.iEnfCode
                    'Scan backwards for matching Vehicle, Local Time, and time zone- if found
                    'use its sProgCode
'                    For ilLoop = ilUpperBound - 1 To LBound(tmDlf) Step -1
'                        If (tmDlf(ilLoop).DlfRec.iVefCode = tmDlf(ilUpperBound).DlfRec.iVefCode) And (tmDlf(ilLoop).DlfRec.sZone = tmDlf(ilUpperBound).DlfRec.sZone) And (tmDlf(ilLoop).DlfRec.iLocalTime(0) = tmDlf(ilUpperBound).DlfRec.iLocalTime(0)) And (tmDlf(ilLoop).DlfRec.iLocalTime(1) = tmDlf(ilUpperBound).DlfRec.iLocalTime(1)) Then
'                            tlDEvt.sProgCode = tmDlf(ilLoop).DlfRec.sProgCode
'                            Exit For
'                        End If
'                    Next ilLoop
'                    If imDelOrEngr = 0 Then
'                        tmDlf(ilUpperBound).DlfRec.sProgCode = tlDEvt.sProgCode
'                    Else
                        tmDlf(ilUpperBound).DlfRec.sProgCode = ""
'                    End If
'                    If imDelOrEngr = 0 Then
'                        tmDlf(ilUpperBound).DlfRec.sCmmlSched = "Y"
'                    Else
                        tmDlf(ilUpperBound).DlfRec.sCmmlSched = "N"
'                    End If
                    'If tgVpf(imVpfIndex).sGCSVer(ilZone) = "A" Then
                    '    tmDlf(ilUpperBound).DlfRec.sCmmlSched = "Y"
                    'Else
                    '    If ilDispl = 1 Then
                    '        tmDlf(ilUpperBound).DlfRec.sCmmlSched = "Y"
                    '    Else
                    '        tmDlf(ilUpperBound).DlfRec.sCmmlSched = "N"
                    '    End If
                    'End If
                    tmDlf(ilUpperBound).DlfRec.iMnfFeed = tgVpf(imVpfIndex).iGMnfNCode(ilZone)
                    tmDlf(ilUpperBound).DlfRec.sBus = tgVpf(imVpfIndex).sGBus(ilZone)
                    tmDlf(ilUpperBound).DlfRec.sSchedule = tgVpf(imVpfIndex).sGSked(ilZone)
                    tmDlf(ilUpperBound).DlfRec.iStartDate(0) = imDate0
                    tmDlf(ilUpperBound).DlfRec.iStartDate(1) = imDate1
                    tmDlf(ilUpperBound).DlfRec.iTermDate(0) = 0
                    tmDlf(ilUpperBound).DlfRec.iTermDate(1) = 0
                    tmDlf(ilUpperBound).DlfRec.iMnfSubFeed = 0
'                    If imDelOrEngr = 0 Then
'                        If (tlDEvt.iEtfCode = 1) Or (tlDEvt.iEtfCode > 13) Then  'Program event type are always set to No
'                            tmDlf(ilUpperBound).DlfRec.sFed = "N"
'                        Else
'                            '5/11/12: Moved vpf.sGFed for delivery to vff
'                            'tmDlf(ilUpperBound).DlfRec.sFed = tgVpf(imVpfIndex).sGFed(ilZone)
'                            For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
'                                If imVefCode = tgVff(ilVff).iVefCode Then
'                                    tmDlf(ilUpperBound).DlfRec.sFed = tgVff(ilVff).sFedDelivery(ilZone)
'                                    Exit For
'                                End If
'                            Next ilVff
'                        End If
'                    Else
                        tmDlf(ilUpperBound).DlfRec.sFed = "Y"
'                    End If
                    'If (tlDEvt.iEtfCode = 1) Or (tlDEvt.iEtfCode > 13) Then  'Program event type are always set to No
                    '    tmDlf(ilUpperBound).DlfRec.sFed = "N"
                    'Else
                    '    tmDlf(ilUpperBound).DlfRec.sFed = tgVpf(imVpfIndex).sGFed(ilZone)
                    'End If
                    tmDlf(ilUpperBound).lDlfCode = 0
                    tmDlf(ilUpperBound).iStatus = 0
'                    mEvtString tmDlf(ilUpperBound)
                    ilUpperBound = ilUpperBound + 1
                    ReDim Preserve tmDlf(0 To ilUpperBound) As DLFLIST
'                    imChg = True
                End If
            Next ilDispl
        End If
    Next ilZone
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadLcf                        *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified: 4/24/94      By:D. Hannifan    *
'*                                                     *
'*            Comments: Read in all events for a date  *
'*                                                     *
'*******************************************************
Private Sub mReadLcf(slCAType As String)
'
'   slCAType(I)- Vehicle type "C"=Conventional; "A"=airing
'
'   tmDEvt (I/O)-contain the log calendar events
'   imVefCode (I)-Vehicle
'   smDateFilter contains the effective date
'


    Dim ilUpper As Integer          'Upperbound of tmDEvt array
    Dim ilSeqNo As Integer          'Sequence number
    Dim ilRet As Integer            'Return from call
    Dim ilIndex As Integer          'List index
    Dim slStartTime As String       'Effective start time
    Dim slStr As String             'Parse string
    Dim slTime As String
    Dim ilDate0 As Integer          'Byte 0 start date
    Dim ilDate1 As Integer          'Byte 1 start date
    ReDim tmDEvt(0 To 0) As DELEVT      'image
    Dim ilFound As Integer          'True=valid avail found
    Dim ilDay As Integer
    Dim slDate As String
    Dim ilType As Integer
    Dim slComment As String
    Dim slXMid As String
    On Error GoTo mReadLcfErr

    ilUpper = UBound(tmDEvt)
    ilType = 0
    ilSeqNo = 1
    gPackDate smDateFilter, ilDate0, ilDate1
    ilDay = gWeekDayStr(smDateFilter)
    If (slCAType = "A") Or (slCAType = "C") Then    'Determine effective date
        ilFound = False
        tmLcfSrchKey.iType = ilType    'On air
        tmLcfSrchKey.sStatus = "C"  'Current
        tmLcfSrchKey.iVefCode = imVefCode
        tmLcfSrchKey.iLogDate(0) = ilDate0
        tmLcfSrchKey.iLogDate(1) = ilDate1
        tmLcfSrchKey.iSeqNo = ilSeqNo
        ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
        Do While (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = imVefCode) And (tmLcf.sStatus = "C")
            gUnpackDate tmLcf.iLogDate(0), tmLcf.iLogDate(1), slDate
            If ilDay <= 4 Then  'Test for Only partial week defined
                If (gWeekDayStr(slDate) >= 0) And (gWeekDayStr(slDate) <= 4) Then
                    ilDate0 = tmLcf.iLogDate(0)
                    ilDate1 = tmLcf.iLogDate(1)
                    ilFound = True
                    Exit Do
                End If
            Else    'Sat or Sun
                If ilDay = gWeekDayStr(slDate) Then
                    ilDate0 = tmLcf.iLogDate(0)
                    ilDate1 = tmLcf.iLogDate(1)
                    ilFound = True
                    Exit Do
                End If
            End If
            ilRet = btrGetNext(hmLcf, tmLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If Not ilFound Then
            'Use TFN
            tmLcfSrchKey.iType = ilType    'On air
            tmLcfSrchKey.sStatus = "C"  'Current
            tmLcfSrchKey.iVefCode = imVefCode
            tmLcfSrchKey.iLogDate(0) = imTFNDay + 1
            tmLcfSrchKey.iLogDate(1) = 0
            tmLcfSrchKey.iSeqNo = ilSeqNo
            ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
            If (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = imVefCode) And (tmLcf.sStatus = "C") Then
                If (tmLcf.iLogDate(0) <= 7) And (tmLcf.iLogDate(1) = 0) Then
                    If imTFNDay + 1 = tmLcf.iLogDate(0) Then
                        ilDate0 = tmLcf.iLogDate(0)
                        ilDate1 = tmLcf.iLogDate(1)
                        ilFound = True
                    End If
                End If
            End If
        End If
        If Not ilFound Then
            Exit Sub
        End If
    End If
    Do
        tmLcfSrchKey.iType = ilType
        tmLcfSrchKey.sStatus = "C"
        tmLcfSrchKey.iVefCode = imVefCode
        tmLcfSrchKey.iLogDate(0) = ilDate0
        tmLcfSrchKey.iLogDate(1) = ilDate1
        tmLcfSrchKey.iSeqNo = ilSeqNo
        ilRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        If ilRet = BTRV_ERR_NONE Then
            ilSeqNo = ilSeqNo + 1
            For ilIndex = LBound(tmLcf.lLvfCode) To UBound(tmLcf.lLvfCode) Step 1
                If tmLcf.lLvfCode(ilIndex) <> 0 Then
                    gUnpackTime tmLcf.iTime(0, ilIndex), tmLcf.iTime(1, ilIndex), "A", "1", slStartTime
                    'Read in Lnf to obtain name and length
                    tmLvfSrchKey.lCode = tmLcf.lLvfCode(ilIndex)
                    ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'Get current record
                    If ilRet = BTRV_ERR_NONE Then
                        'Read in all the event record (Lef)
                        tmLefSrchKey.lLvfCode = tmLcf.lLvfCode(ilIndex)
                        tmLefSrchKey.iStartTime(0) = 0
                        tmLefSrchKey.iStartTime(1) = 0
                        tmLefSrchKey.iSeqNo = 0
                        ilRet = btrGetGreaterOrEqual(hmLef, tmLef, imLefRecLen, tmLefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmLef.lLvfCode = tmLcf.lLvfCode(ilIndex))
                            gUnpackLength tmLef.iStartTime(0), tmLef.iStartTime(1), "3", False, slStr
                            gAddTimeLength slStartTime, slStr, "A", "1", slTime, slXMid
                            tmDEvt(ilUpper).lTime = CLng(gTimeToCurrency(slTime, True))
                            tmDEvt(ilUpper).iEtfCode = tmLef.iEtfCode
                            tmDEvt(ilUpper).iEnfCode = tmLef.iEnfCode
                            tmDEvt(ilUpper).iStatus = 0 'Unused
                            ilFound = False
                            Select Case tmLef.iEtfCode
                                Case 1  'Program
                                Case 2  'Contract Avail
                                    'Use avail comment for progcode if defined
'                                    If mReadCefRec(tmLef.lCefCode) Then
'                                        'If tmCef.iStrLen > 0 Then
'                                        '    slComment = Trim$(Left$(tmCef.sComment, tmCef.iStrLen))
'                                        'End If
'                                        slComment = gStripChr0(tmCef.sComment)
'                                    End If
                                    ilFound = True
                                Case 3
                                    ilFound = True
                                Case 4
                                    ilFound = True
                                Case 5
                                    ilFound = True
                                Case 6  'Cmml Promo
                                    ilFound = True
                                Case 7  'Feed avail
                                    ilFound = True
                                Case 8  'PSA/Promo (Avail)
                                    ilFound = True
                                Case 9
                                    ilFound = True
                                Case 10  'Page eject, Line space 1, 2 or 3
                                Case 11
                                Case 12
                                Case 13
                                Case Else   'Other
'                                    If imDelOrEngr = 0 Then
'                                        ilFound = True
'                                    End If
                            End Select
                            If ilFound Then
'                                tmDEvt(ilUpper).sProgCode = slComment   'Use program comment on all events as prog code
                                ilUpper = ilUpper + 1
                                ReDim Preserve tmDEvt(0 To ilUpper) As DELEVT
                            End If
                            ilRet = btrGetNext(hmLef, tmLef, imLefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                Else
                    ilSeqNo = -1
                    Exit For
                End If
            Next ilIndex
        Else
            ilSeqNo = -1
        End If
    Loop While ilSeqNo > 0
Exit Sub
mReadLcfErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Public Function mOpenLinksFiles() As Boolean
    Dim ilRet As Integer
    ReDim tmDlf(0 To 1) As DLFLIST                'DLF record image
        mOpenLinksFiles = True
        hmCbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mOpenErr
        gBtrvErrorMsg ilRet, "mOpenLinksFiles (btrOpen: cbf.Btr)", RptSelEngrLk
        On Error GoTo 0
        imCbfRecLen = Len(tmCbf)
    
        hmDlf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
        ilRet = btrOpen(hmDlf, "", sgDBPath & "egf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mOpenErr
        gBtrvErrorMsg ilRet, "mOpenLinksFiles (btrOpen: egf.Btr)", RptSelEngrLk
        On Error GoTo 0
        imDlfRecLen = Len(tmDlf(0).DlfRec)
        
        hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mOpenErr
        gBtrvErrorMsg ilRet, "mOpenLinksFiles (btrOpen: Lcf.Btr)", RptSelEngrLk
        On Error GoTo 0
        imLcfRecLen = Len(tmLcf)  'Get and save LCF record length
        
        hmLvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mOpenErr
        gBtrvErrorMsg ilRet, "mOpenLinksFiles (btrOpen: Lvf.Btr)", RptSelEngrLk
        On Error GoTo 0
        imLvfRecLen = Len(tmLvf)  'Get and save LVF record length
        
        hmLef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmLef, "", sgDBPath & "Lef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mOpenErr
        gBtrvErrorMsg ilRet, "mOpenLinksFiles (btrOpen: Lef.Btr)", RptSelEngrLk
        On Error GoTo 0
        imLefRecLen = Len(tmLef)  'Get and save LEF record length
        
        hmEnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmEnf, "", sgDBPath & "Enf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mOpenErr
        gBtrvErrorMsg ilRet, "mOpenLinksFiles (btrOpen: Enf.Btr)", RptSelEngrLk
        On Error GoTo 0
        imEnfRecLen = Len(tmEnf)  'Get and save LEF record length
       
        hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mOpenErr
        gBtrvErrorMsg ilRet, "mOpenLinksFiles (btrOpen: Vef.Btr)", RptSelEngrLk
        On Error GoTo 0
        imVefRecLen = Len(tmVef)  'Get and save Vlf record length
    Exit Function
    
mOpenErr:
      On Error GoTo 0
        mOpenLinksFiles = False
        Exit Function
End Function

Public Sub mCloseLinkFiles()
    Dim ilRet As Integer
    
        btrDestroy hmCbf
        btrDestroy hmDlf
        btrDestroy hmLcf
        btrDestroy hmLvf
        btrDestroy hmLef
        btrDestroy hmVef
        btrDestroy hmEnf
        ilRet = btrClose(hmCbf)
        ilRet = btrClose(hmDlf)
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmLvf)
        ilRet = btrClose(hmLef)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmEnf)
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name: mEvtString                     *
'*                                                     *
'*             Created:7/27/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:make strings within Dlf records *
'*                      images from events for day     *
'*                                                     *
'*******************************************************
Private Sub mEvtString(tlDlf As DLFLIST)
    Dim slTime As String
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    tlDlf.sVehicle = tmVef.sName
    gUnpackTime tlDlf.DlfRec.iAirTime(0), tlDlf.DlfRec.iAirTime(1), "A", "1", slTime
    tlDlf.sFeed = ""
    For ilLoop = 0 To lbcFeedCode.ListCount - 1 Step 1
        slNameCode = lbcFeedCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If tlDlf.DlfRec.iMnfFeed = Val(slCode) Then
            tlDlf.sFeed = Trim$(slName)
            Exit For
        End If
    Next ilLoop
    'Old value if invalid- replace with correct value
    If tlDlf.sFeed = "" Then
        If lbcFeedCode.ListCount = 1 Then
            slNameCode = lbcFeedCode.List(0)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tlDlf.sFeed = Trim$(slName)
            tlDlf.DlfRec.iMnfFeed = Val(slCode)
'            imChg = True
        End If
    End If
    tlDlf.sAirTime = slTime
    gUnpackTime tlDlf.DlfRec.iLocalTime(0), tlDlf.DlfRec.iLocalTime(1), "A", "1", slTime
    tlDlf.sLocalTime = slTime
    gUnpackTime tlDlf.DlfRec.iFeedTime(0), tlDlf.DlfRec.iFeedTime(1), "A", "1", slTime
    tlDlf.sFeedTime = slTime
    tlDlf.sSubfeed = ""
    For ilLoop = 0 To UBound(tmSubFeedCode) - 1 Step 1 'lbcSubFeedCode.ListCount - 1 Step 1
        slNameCode = tmSubFeedCode(ilLoop).sKey 'lbcSubFeedCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If tlDlf.DlfRec.iMnfSubFeed = Val(slCode) Then
            tlDlf.sSubfeed = Trim$(slName)
            Exit For
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tmEvtTypeCode) - 1 Step 1  'lbcEvtTypeCode.ListCount - 1 Step 1
        slNameCode = tmEvtTypeCode(ilLoop).sKey    'lbcEvtTypeCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slName)
        ilRet = gParseItem(slNameCode, 3, "\", slCode)
        If tlDlf.DlfRec.iEtfCode = Val(slCode) Then
            tlDlf.sEventType = Trim$(slName)
            Exit For
        End If
    Next ilLoop
    If tlDlf.DlfRec.iEnfCode > 0 Then
        For ilLoop = 1 To imNoEnfSaved Step 1
            If imEnfCode(ilLoop) = tlDlf.DlfRec.iEnfCode Then
                tlDlf.sEventName = smEnfName(ilLoop)
                Exit Sub
            End If
        Next ilLoop
        tmEnfSrchKey.iCode = tlDlf.DlfRec.iEnfCode
        ilRet = btrGetEqual(hmEnf, tmEnf, imEnfRecLen, tmEnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        If ilRet = BTRV_ERR_NONE Then
            tlDlf.sEventName = Trim$(tmEnf.sName)
            For ilLoop = 19 To 1 Step -1
                imEnfCode(ilLoop + 1) = imEnfCode(ilLoop)
                smEnfName(ilLoop + 1) = smEnfName(ilLoop)
            Next ilLoop
            imEnfCode(1) = tlDlf.DlfRec.iEnfCode
            smEnfName(1) = tlDlf.sEventName
            If imNoEnfSaved < 20 Then
                imNoEnfSaved = imNoEnfSaved + 1
            End If
        Else
            tlDlf.sEventName = ""
        End If
    Else
        tlDlf.sEventName = ""
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSubfeedPop                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Feed list             *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mSubfeedPop()
'
'   mSubfeedPop
'   Where:
'
    Dim ilRet As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    ilRet = gPopMnfPlusFieldsBox(RptSelEngrLk, lbcSubfeed, tmSubFeedCode(), smSubFeedCodeTag, "NOS")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSubfeedPopErr
        gCPErrorMsg ilRet, "mSubfeedPop (gPopMnfPlusFieldsBox)", RptSelEngrLk
        On Error GoTo 0
        lbcSubfeed.AddItem "[None]", 0
    End If
    Exit Sub
mSubfeedPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mETypePop                       *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection event   *
'*                      type box                       *
'*                                                     *
'*******************************************************
Private Sub mEvtTypePop()
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcEvtType.ListIndex
    If ilIndex > 0 Then
        slName = lbcEvtType.List(ilIndex)
    End If
    ilRet = gPopEvtNmByTypeBox(RptSelEngrLk, True, True, lbcEvtType, tmEvtTypeCode(), smEvtTypeCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mEvtTypePopErr
        gCPErrorMsg ilRet, "mEvtTypePop (gIMoveListBox: EvtType)", RptSelEngrLk
        On Error GoTo 0
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcEvtType
            If gLastFound(lbcEvtType) > 0 Then
                lbcEvtType.ListIndex = gLastFound(lbcEvtType)
            Else
                lbcEvtType.ListIndex = -1
            End If
        Else
            lbcEvtType.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mEvtTypePopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Public Sub mCreateForOutput()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    'Prepass fields used
    'CbfGenTime - generation dtime to filter with
    'cbfGenDate - generation date to filter with
    'cbfStatus - status of the avail, 0 = missing, 1 = ok, 2 = avail no longers exists, link can be removed
    'cbfPrefDT - vehicle name (max 40 char)
    'cbfAirWks = 1= M-F, 6 = Sa, 7 = Su
    'cbfTime = Air Time
    'cbfPropOrdTime = Feed Time
    'cbfStartDate = Start date of link
    'cbfEndDate = end date (null or zero indicates TFN)
    'cbfPriceType = Zone(E, c, M, P)
    'cbfMnfGroup = event type mnf code
    'cbfdnfcode = internal event name code
    'cbfBuyer = buses (5 char)
    'cbfResort = schedule (dlfschedule)
    'cbfcontrno = delivery link, can be 0 if missing (not defined).
    
        For ilLoop = 0 To UBound(tmDlf) - 1
            If (CkcDiscrepOnly.Value = vbChecked And tmDlf(ilLoop).iStatus <> 1) Or CkcDiscrepOnly.Value = vbUnchecked Then     'if discreps only, ignore the entries that are flagged as 1; the links are insync with avails
                'gen date and time has been set; field doesnt change
                tmCbf.sStatus = Trim(str$(tmDlf(ilLoop).iStatus))       '0=New; 1=old and retain, 2=old and delete; -1= New but not used
                tmCbf.sPrefDT = tmVef.sName
                tmCbf.iAirWks = Val(tmDlf(ilLoop).DlfRec.sAirDay)       '0 = m-f, 6 = sat, 7 = sun
                   
                tmCbf.iTime(0) = tmDlf(ilLoop).DlfRec.iAirTime(0)         'air time
                tmCbf.iTime(1) = tmDlf(ilLoop).DlfRec.iAirTime(1)         'air time
                tmCbf.iPropOrdTime(0) = tmDlf(ilLoop).DlfRec.iFeedTime(0)         'feed time
                tmCbf.iPropOrdTime(1) = tmDlf(ilLoop).DlfRec.iFeedTime(1)         'feed time
                tmCbf.iStartDate(0) = tmDlf(ilLoop).DlfRec.iStartDate(0)          'start date
                tmCbf.iStartDate(1) = tmDlf(ilLoop).DlfRec.iStartDate(1)
                tmCbf.iEndDate(0) = tmDlf(ilLoop).DlfRec.iTermDate(0)               'end date
                tmCbf.iEndDate(1) = tmDlf(ilLoop).DlfRec.iTermDate(1)
                tmCbf.sPriceType = tmDlf(ilLoop).DlfRec.sZone
                tmCbf.iMnfGroup = tmDlf(ilLoop).DlfRec.iEtfCode
                tmCbf.iDnfCode = tmDlf(ilLoop).DlfRec.iEnfCode
                tmCbf.sBuyer = tmDlf(ilLoop).DlfRec.sBus                'buses
                tmCbf.sResort = tmDlf(ilLoop).DlfRec.sSchedule
                tmCbf.lContrNo = tmDlf(ilLoop).DlfRec.lCode             'links internal recd code
                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            End If
        Next ilLoop
End Sub

Public Sub mRestoreFields()
        igGenRpt = False
        frcOutput.Enabled = igOutput
        frcCopies.Enabled = igCopies
        frcFile.Enabled = igFile
        frcOption.Enabled = igOption
        pbcClickFocus.SetFocus
        tmcDone.Enabled = True
    Exit Sub
End Sub

