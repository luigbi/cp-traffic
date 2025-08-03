VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelRI 
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
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   8115
      TabIndex        =   30
      Top             =   90
      Visible         =   0   'False
      Width           =   1095
   End
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
         Top             =   225
         Visible         =   0   'False
         Width           =   4455
         Begin VB.PictureBox plcMethods 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4290
            TabIndex        =   41
            Top             =   2040
            Visible         =   0   'False
            Width           =   4290
            Begin VB.CheckBox ckcMethod 
               Caption         =   "Year"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   3000
               TabIndex        =   44
               Top             =   0
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.CheckBox ckcMethod 
               Caption         =   "Month"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1920
               TabIndex        =   43
               Top             =   0
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.CheckBox ckcMethod 
               Caption         =   "Week"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   42
               Top             =   0
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   885
            End
         End
         Begin VB.CheckBox ckcInclPreviousPaid 
            Caption         =   "Include previously paid stations"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1680
            Visible         =   0   'False
            Width           =   3060
         End
         Begin VB.ComboBox cbcSet1 
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
            Left            =   1320
            TabIndex        =   39
            Top             =   1320
            Visible         =   0   'False
            Width           =   1500
         End
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
            Height          =   285
            Left            =   120
            TabIndex        =   38
            TabStop         =   0   'False
            Text            =   "Major Set #"
            Top             =   1320
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.PictureBox plcWSSort 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   35
            Top             =   960
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcWSSort 
               Caption         =   "Advertiser"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1800
               TabIndex        =   37
               Top             =   0
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.OptionButton rbcWSSort 
               Caption         =   "Market"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   720
               TabIndex        =   36
               Top             =   0
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   1065
            End
         End
         Begin VB.CheckBox ckcInclRecd 
            Caption         =   "Include Received Affidavits"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   660
            Visible         =   0   'False
            Width           =   4140
         End
         Begin VB.PictureBox plcSelC4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   60
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   31
            Top             =   420
            Visible         =   0   'False
            Width           =   4275
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Detail"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   75
               TabIndex        =   33
               Top             =   0
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.OptionButton rbcSelC4 
               Caption         =   "Summary"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   1050
               TabIndex        =   32
               Top             =   0
               Visible         =   0   'False
               Width           =   1215
            End
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
         Begin VB.Label lacSelCFrom 
            Appearance      =   0  'Flat
            Caption         =   "Month"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   135
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
            Left            =   0
            TabIndex        =   17
            Top             =   90
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
Attribute VB_Name = "RptSelRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptselri.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************


' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelRI.Frm - Remote Invoice Worksheet, name changed to REP Invoice Worksheet 11-17-02
'            2-18-02 DH
'            9-1-02 Delinquent Affidavit Report &
'            Unbillable Invoices implemented
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
Dim hmSbf As Integer        'Special billing
Dim tmSbf As SBF
Dim tmSbfSrchKey1 As POPLCODE
Dim imSbfRecLen As Integer        'SBF record length
Dim hmSpf As Integer        'Site
Dim tmSpf As SPF
Dim imSpfRecLen As Integer        'SPF record length

'*******************************************************
'*                                                     *
'*      Procedure Name:mRepVeh                         *
'*                                                     *
'*             Created:9-3-02       By:D. Hosaka       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box with REP vehicles only     *
'*                                                     *
'*******************************************************
Private Sub mRepVeh(ilIndex As Integer, ilUselbcVehicle As Integer)
    Dim ilRet As Integer
    If ilUselbcVehicle Then
        ilRet = gPopUserVehicleBox(RptSelRI, ACTIVEVEH + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
    Else
          ilRet = gPopUserVehicleBox(RptSelRI, ACTIVEVEH + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER, lbcSelection(ilIndex), tgCSVNameCode(), sgCSVNameCodeTag)     'lbcCSVNameCode)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mRepVehERr
        gCPErrorMsg ilRet, "mRepVeh (gPopUserVehicleBox: Vehicle)", RptSelRI
        On Error GoTo 0
    End If
    Exit Sub
mRepVehERr:
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
Dim ilListIndex As Integer
Dim ilIndex

    ilListIndex = lbcRptType.ListIndex
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        ilIndex = 0
        If ilListIndex = INV_DELINQUENT Or ilListIndex = INV_UNBILLABLE Or ilListIndex = INV_BARTERPAYMENT Then
            ilIndex = 1
        ElseIf ilListIndex = INV_REMOTE_INVWS Then
            If rbcWSSort(1).Value Then
                ilIndex = 1
            End If
        End If
        llRg = CLng(lbcSelection(ilIndex).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(ilIndex).hwnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub

Private Sub ckcInclRecd_Click()
Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex
    If ilListIndex = INV_BARTERPAYMENT Then
        If ckcInclRecd.Value = vbChecked Then
            rbcOutput(0).Enabled = False
            rbcOutput(2).Value = True           'default to PDF; if updating the barter paid flag force to print
            MsgBox "Barter payments for the requested month will be set to PAID.  The report cannot be rerun for the same station & period."
            ckcInclPreviousPaid.Enabled = False
            ckcInclPreviousPaid = False
        Else                                    'not updating the barter paid flag
            rbcOutput(0).Enabled = True
            ckcInclPreviousPaid.Move 120, ckcInclRecd.Top + ckcInclRecd.Height
            ckcInclPreviousPaid.Visible = True
            ckcInclPreviousPaid.Enabled = True
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
    Dim slReptName As String
    Dim ilLoop As Integer
    Dim llSpfRecPos  As Long
    Dim slNowDate As String
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

    '8-29-02
    ilListIndex = lbcRptType.ListIndex
    If ilListIndex = INV_BARTERPAYMENT Then         '11-28-06
        slReptName = "Barter.rpt"
    ElseIf ilListIndex = INV_DELINQUENT Then
        slReptName = "DelinAff.rpt"
    ElseIf ilListIndex = INV_UNBILLABLE Then
        slReptName = "UnbillInv.rpt"
    Else
        slReptName = "PostRep.rpt"
    End If

    igUsingCrystal = True
    ilNoJobs = 1
    ilStartJobNo = 1
    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs

        If Not gOpenPrtJob(Trim$(slReptName)) Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            'frcWhen.Enabled = igWhen
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            'frcRptType.Enabled = igReportType
            Exit Sub
        End If
        ilRet = gcmcgenRI(ilListIndex, imGenShiftKey, smLogUserCode)
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
        If ilListIndex = INV_BARTERPAYMENT Then             '11-28-06 then
            ilRet = gGenBarterPayment(RptSelRI, -1)          'handle to export file (n/a in Barter Payment report)
        '9-29-02
        ElseIf ilListIndex = INV_REMOTE_INVWS Then
            gRIWorkSheet RptSelRI
        ElseIf ilListIndex = INV_DELINQUENT Or ilListIndex = INV_UNBILLABLE Then
            gCheck_Affs RptSelRI     'create Delinque Affidavit report & Unbillable Invoice report
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
            'ilret = gOutputToFile(slFileName, imFTSelectedIndex)
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '2-21-02
        End If
    Next ilJobs
    imGenShiftKey = 0

    Screen.MousePointer = vbHourglass
    If ilListIndex = INV_BARTERPAYMENT Then

        'set the PAID flag if necessary
        're-open SBF file
        If ckcInclRecd.Value = vbChecked Then                'set barter paid flag
            hmSpf = CBtrvTable(TWOHANDLES)
            ilRet = btrOpen(hmSpf, "", sgDBPath & "Spf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmSpf)
                btrDestroy hmSpf
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            imSpfRecLen = Len(tmSpf)

            hmSbf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
            ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrClose(hmSbf)
                btrDestroy hmSbf
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            imSbfRecLen = Len(tmSbf)

            For ilLoop = 0 To UBound(lgPrintedCnts) Step 1
                Do
                    If lgPrintedCnts(ilLoop) <> 0 Then
                        tmSbfSrchKey1.lCode = lgPrintedCnts(ilLoop)
                        imSbfRecLen = Len(tmSbf)
                        ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet = BTRV_ERR_NONE Then
                            tmSbf.sBarterPaid = "Y"
                            ilRet = btrUpdate(hmSbf, tmSbf, imSbfRecLen)
                        End If
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
            Next ilLoop
            ilRet = btrClose(hmSbf)

            'update Date Last Updated Barter in Site
            ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrGetPosition(hmSpf, llSpfRecPos)
                slNowDate = Format$(Now, "m/d/yy")
                Do
                    ilRet = btrGetDirect(hmSpf, tmSpf, imSpfRecLen, llSpfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                     gPackDate slNowDate, tmSpf.iBarterLPD(0), tmSpf.iBarterLPD(1)
                    ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
            End If
            ilRet = btrClose(hmSpf)
        End If
        Erase lgPrintedCnts

    End If

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
    'RptSelRI.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    PECloseEngine
    gChDrDir        '11/2/10: Handle case where printing directly to abode and rerouting abode output to different Drive/Path
    
    Set RptSelRI = Nothing   'Remove data segment
    
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcRptType_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llTodayDate                   slStr                                                   *
'******************************************************************************************

Dim ilListIndex As Integer
Dim ilRet As Integer
    ilListIndex = lbcRptType.ListIndex

    If ilListIndex = INV_BARTERPAYMENT Then         '11-28-06
        ckcAll.Caption = "All Vehicles"
        mRepVeh 1, False
        plcSelC4.Visible = True             'detail/summary
        rbcSelC4(0).Visible = True          'detail
        rbcSelC4(1).Visible = True          'summary
        ckcInclRecd.Caption = "Update Barter Paid"

        ckcInclRecd.Visible = True
        If ckcInclRecd.Value = vbUnchecked Then
            ckcInclRecd_Click
        End If
        plcWSSort.Visible = False
        edcSet1.Visible = False
        cbcSet1.Visible = False
        lbcSelection(1).Visible = True
        lbcSelection(0).Visible = False
        plcMethods.Move 120, ckcInclPreviousPaid.Top + ckcInclPreviousPaid.Height + 30
        plcMethods.Visible = True
        ckcMethod(0).Visible = True
        ckcMethod(1).Visible = True
        ckcMethod(2).Visible = True
        
    ElseIf ilListIndex = INV_DELINQUENT Then
        ckcAll.Caption = "All Vehicles"
        mRepVeh 1, False
        lbcSelection(0).Visible = False
        lbcSelection(1).Visible = True
        ckcInclRecd.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 30 'include Received but not posted
        ckcInclRecd.Visible = True
        gPopVehicleGroups RptSelRI!cbcSet1, tgVehicleSets1(), True
        cbcSet1.ListIndex = 0
        edcSet1.Move 120, ckcInclRecd.Top + ckcInclRecd.Height + 90
        cbcSet1.Move 120 + edcSet1.Width, edcSet1.Top - 30
        cbcSet1.Visible = True
        edcSet1.Visible = True
    ElseIf ilListIndex = INV_UNBILLABLE Then
        ckcAll.Caption = "All Advertisers"
        ilRet = gPopAdvtBox(RptSelRI, lbcSelection(1), tgRptSelAdvertiserCode(), sgRptSelAdvertiserCodeTag)
        lbcSelection(0).Visible = False
        lbcSelection(1).Visible = True
        ckcInclRecd.Move 120, edcSelCFrom.Top + edcSelCFrom.Height + 30 'include Received but not posted
        ckcInclRecd.Visible = True
    Else                            'rep invoice worksheet
        'Detail or Summary
        ilRet = gPopAdvtBox(RptSelRI, lbcSelection(1), tgRptSelAdvertiserCode(), sgRptSelAdvertiserCodeTag)
        plcSelC4.Visible = True
        rbcSelC4(0).Visible = True
        rbcSelC4(1).Visible = True
        plcWSSort.Top = plcSelC4.Top + plcSelC4.Height
        plcWSSort.Visible = True
        rbcWSSort(0).Visible = True
        rbcWSSort(1).Visible = True
    End If

End Sub

Private Sub lbcSelection_Click(Index As Integer)
Dim ilListIndex As Integer

    ilListIndex = lbcRptType.ListIndex

    If Not imAllClicked Then
        If ilListIndex = INV_DELINQUENT Or ilListIndex = INV_UNBILLABLE Or ilListIndex = INV_BARTERPAYMENT Then
            If Index = 1 Then
                imSetAll = False
                ckcAll.Value = vbUnchecked  '9-12-02 False
                imSetAll = True
            End If
        ElseIf ilListIndex = INV_REMOTE_INVWS Then
                imSetAll = False
                ckcAll.Value = vbUnchecked  '9-12-02 False
                imSetAll = True

        Else
            If Index = 0 Then           'vehicle list box
                imSetAll = False
                ckcAll.Value = vbUnchecked  '9-12-02 False
                imSetAll = True
            End If
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

    RptSelRI.Caption = smSelectedRptName & " Report"
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
    lbcSelection(0).Move 15, ckcAll.Height + 90, 4380
    lbcSelection(1).Move lbcSelection(0).Left, lbcSelection(1).Top
    pbcSelC.Move 90, 255, 4515, 3360

    gCenterStdAlone RptSelRI
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
    Dim ilRet As Integer

    gPopExportTypes cbcFileType     '2-21-02
    pbcSelC.Visible = False
    lbcRptType.Clear

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    Screen.MousePointer = vbHourglass

    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    pbcSelC.Visible = True
    mMarketPopRI                  'populate the markets for selection
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
    ckcAll.Value = vbUnchecked                 'force all advertisrs

    lbcRptType.AddItem "Invoice Register,", INV_REGISTER       'not in this module
    lbcRptType.AddItem "View Export,", INV_VIEWEXPORT          'not in this module
    lbcRptType.AddItem "Billing Distribution,", INV_DISTRIBUTE     'not in this module
    'lbcRptType.AddItem "Remote Invoice Worksheet", INV_REMOTE_INVWS
    lbcRptType.AddItem "Rep Invoice Worksheet", INV_REMOTE_INVWS    'name changed 11-17-02
    lbcRptType.AddItem "Delinquent Affidavits", INV_DELINQUENT
    lbcRptType.AddItem "Unbillable Invoices", INV_UNBILLABLE
    lbcRptType.AddItem "Credit/Debit Memo", INV_CREDITMEMO
    lbcRptType.AddItem "Invoice Summary", INV_SUMMARY
    lbcRptType.AddItem "Barter Payments", INV_BARTERPAYMENT

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
'*      Procedure Name:mMarketPop                      *
'*                                                     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Market combobox       *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mMarketPopRI()
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
    lbcSelection(0).Clear
    ilSortCode = 0
    ReDim tgMktCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    ilRet = gObtainMnfForType("H3", slStr, tgMktMnf())
    For ilLoop = LBound(tgMktMnf) To UBound(tgMktMnf) - 1 Step 1
        'If Trim$(tgMktMnf(ilLoop).sRPU) = "Y" Then     '11-13-02 get all markets, not just clusters
            slName = Trim$(tgMktMnf(ilLoop).sName)
            slName = slName & "\" & Trim$(str$(tgMktMnf(ilLoop).iCode))
            tgMktCode(ilSortCode).sKey = slName
            If ilSortCode >= UBound(tgMktCode) Then
                ReDim Preserve tgMktCode(0 To UBound(tgMktCode) + 100) As SORTCODE
            End If
            ilSortCode = ilSortCode + 1
        'End If
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
    'gInitStdAlone RptSelRI, slStr, ilTestSystem
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
    Dim ilListIndex As Integer
    Dim ilIndex

    ilListIndex = lbcRptType.ListIndex

    ilEnable = False
    If (edcSelCFrom.Text <> "") And (edcSelCFrom1.Text <> "") Then
        ilEnable = True

        If ilEnable Then
            ilEnable = False
            If ckcAll.Value = vbChecked Then
                ilEnable = True
            Else
                ilIndex = 0
                If ilListIndex = INV_DELINQUENT Or ilListIndex = INV_UNBILLABLE Or ilListIndex = INV_BARTERPAYMENT Then
                    ilIndex = 1
                ElseIf ilListIndex = INV_REMOTE_INVWS Then     '12-14-06
                    If rbcWSSort(1).Value Then
                        ilIndex = 1
                    End If
                End If
                For ilLoop = 0 To lbcSelection(ilIndex).ListCount - 1 Step 1      'market entry must be selected
                    If lbcSelection(ilIndex).Selected(ilLoop) Then
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
    Unload RptSelRI
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
    '    Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    '    Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    '    Traffic!cdcSetup.Action = 6
    End If
End Sub


Private Sub plcMethods_Paint()
    plcMethods.Cls
    plcMethods.CurrentX = 0
    plcMethods.CurrentY = 0
    plcMethods.Print "Methods"
End Sub

Private Sub plcWSSort_Paint()
    plcWSSort.Cls
    plcWSSort.CurrentX = 0
    plcWSSort.CurrentY = 0
    plcWSSort.Print "Sort by"
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




Private Sub rbcWSSort_Click(Index As Integer)
Dim ilListIndex As Integer
Dim llRg As Long
Dim llRet As Long

    ilListIndex = lbcRptType.ListIndex

    If ilListIndex = INV_REMOTE_INVWS Then
        If Index = 0 Then            'by market
            lbcSelection(0).Visible = True
            lbcSelection(1).Visible = False
            ckcAll.Caption = "All Markets"
             'force all markets to be unselecterd
            llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(0).hwnd, LB_SELITEMRANGE, False, llRg)
        Else
            lbcSelection(0).Visible = False
            lbcSelection(1).Visible = True
            ckcAll.Caption = "All Advertisers"
            ckcAll.Value = vbUnchecked
            'force all markets to be processed
            llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcSelection(0).hwnd, LB_SELITEMRANGE, True, llRg)

        End If
    End If
    mSetCommands
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
    'mTerminate False
End Sub
