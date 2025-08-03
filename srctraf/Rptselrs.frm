VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RptSelRS 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facility Report Selection"
   ClientHeight    =   5640
   ClientLeft      =   1275
   ClientTop       =   1635
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
   ScaleHeight     =   5640
   ScaleWidth      =   9270
   Begin VB.ListBox lbcRptType 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   6240
      TabIndex        =   43
      Top             =   1440
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CheckBox ckcAllBooks 
      Caption         =   "All Book Names"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   4680
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmcList 
      Appearance      =   0  'Flat
      Caption         =   "Return to Report List"
      Height          =   285
      Left            =   6615
      TabIndex        =   30
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
      TabIndex        =   34
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
      TabIndex        =   35
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
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4245
      Width           =   30
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   3240
      Top             =   4920
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
      TabIndex        =   16
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
         TabIndex        =   19
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   20
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   18
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
      TabIndex        =   21
      Top             =   60
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   26
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
         TabIndex        =   23
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
         TabIndex        =   25
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOption 
      Caption         =   "Research Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3810
      Left            =   45
      TabIndex        =   27
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
         Left            =   135
         ScaleHeight     =   3360
         ScaleWidth      =   4410
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   4410
         Begin V81TrafficReports.CSI_Calendar CSI_CalDate 
            Height          =   255
            Left            =   2040
            TabIndex        =   4
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            Text            =   "10/3/2022"
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
            CSI_AllowBlankDate=   -1  'True
            CSI_AllowTFN    =   0   'False
            CSI_DefaultDateType=   0
         End
         Begin VB.CheckBox ckcNewPage 
            Caption         =   "New Page Each Book"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   53
            Top             =   2200
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox lacDemoMsg 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   0
            TabIndex        =   52
            Text            =   "(Primary must match one of the 5 demos selected)"
            Top             =   3120
            Visible         =   0   'False
            Width           =   4335
         End
         Begin VB.ComboBox cbcPrimaryDemo 
            BackColor       =   &H00FFFF80&
            Height          =   330
            ItemData        =   "Rptselrs.frx":0000
            Left            =   1440
            List            =   "Rptselrs.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   50
            Top             =   2880
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox edcTopHowMany 
            BackColor       =   &H00FFFF80&
            Height          =   315
            Left            =   2520
            MaxLength       =   2
            TabIndex        =   48
            Top             =   2400
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ListBox LbcCustDemos 
            CausesValidation=   0   'False
            Height          =   270
            Left            =   45
            Sorted          =   -1  'True
            TabIndex        =   45
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.PictureBox plcSelC5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   4260
            TabIndex        =   41
            Top             =   1320
            Width           =   4260
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Vehicle"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   2175
               TabIndex        =   12
               Top             =   255
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Time"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   720
               TabIndex        =   11
               Top             =   240
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Extra Daypart"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   2175
               TabIndex        =   10
               Top             =   15
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.CheckBox ckcSelC5 
               Caption         =   "Sold Daypart"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   9
               Top             =   0
               Value           =   1  'Checked
               Width           =   1335
            End
         End
         Begin VB.PictureBox plcSelC3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            ScaleHeight     =   240
            ScaleWidth      =   4380
            TabIndex        =   38
            Top             =   960
            Width           =   4380
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Standard"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   14
               Left            =   720
               TabIndex        =   7
               Top             =   0
               Value           =   1  'Checked
               Width           =   1095
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Custom"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   13
               Left            =   1800
               TabIndex        =   8
               Top             =   0
               Width           =   1440
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Cancelled"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   1350
               TabIndex        =   40
               Top             =   195
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox ckcSelC3 
               Caption         =   "Hidden"
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   2445
               TabIndex        =   39
               Top             =   195
               Visible         =   0   'False
               Width           =   900
            End
         End
         Begin VB.PictureBox plcSelC11 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   4275
            TabIndex        =   37
            Top             =   600
            Width           =   4275
            Begin VB.OptionButton rbcSelC11 
               Caption         =   "Vehicle by Default Book"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   1
               Left            =   2040
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   0
               Width           =   2295
            End
            Begin VB.OptionButton rbcSelC11 
               Caption         =   "Book Name"
               ForeColor       =   &H80000008&
               Height          =   210
               Index           =   0
               Left            =   720
               TabIndex        =   5
               Top             =   0
               Value           =   -1  'True
               Width           =   1275
            End
         End
         Begin VB.CheckBox ckcSelC3 
            Caption         =   "Include Qualitative Data"
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   1920
            Width           =   2415
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
            Index           =   0
            Left            =   120
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "Effective Book Date"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lacPrimaryDemo 
            Caption         =   "Primary Demo"
            Height          =   285
            Left            =   0
            TabIndex        =   51
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lacTopHowMany 
            Caption         =   "Top How Many (Blank = All)"
            Height          =   285
            Left            =   0
            TabIndex        =   49
            Top             =   2520
            Visible         =   0   'False
            Width           =   2415
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
         Height          =   3540
         Left            =   4605
         ScaleHeight     =   3540
         ScaleWidth      =   4455
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   4455
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   2
            ItemData        =   "Rptselrs.frx":0004
            Left            =   0
            List            =   "Rptselrs.frx":0006
            MultiSelect     =   1  'Simple
            TabIndex        =   46
            Top             =   2160
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   1
            ItemData        =   "Rptselrs.frx":0008
            Left            =   0
            List            =   "Rptselrs.frx":000A
            MultiSelect     =   2  'Extended
            TabIndex        =   17
            Top             =   2160
            Width           =   4395
         End
         Begin VB.ListBox lbcSelection 
            Appearance      =   0  'Flat
            Height          =   1290
            Index           =   0
            ItemData        =   "Rptselrs.frx":000C
            Left            =   0
            List            =   "Rptselrs.frx":000E
            MultiSelect     =   2  'Extended
            TabIndex        =   14
            Top             =   360
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.CheckBox ckcAll 
            Caption         =   "All Vehicles"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   2385
         End
         Begin VB.Label lacMaxDemoDescr 
            Caption         =   "Select up to 6 demo categories"
            Height          =   255
            Left            =   1680
            TabIndex        =   47
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lacBookNames 
            Caption         =   "Book Names"
            Height          =   255
            Left            =   2880
            TabIndex        =   44
            Top             =   1800
            Visible         =   0   'False
            Width           =   1695
         End
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   285
      Left            =   6615
      TabIndex        =   31
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "Generate Report"
      Height          =   285
      Left            =   6240
      TabIndex        =   29
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
         Width           =   1185
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
Attribute VB_Name = "RptSelRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptselrs.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  smLogUserCode                                                                         *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptSelRS.Frm - Research Report
'
'
' Release: 4.7 10/2000
'
' Description:
'   This file contains the Report selection for Contracts screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imSetAllBooks As Integer 'True=Set list box; False= don't change list box
Dim imAllBooksClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imFTSelectedIndex As Integer
Dim imFirstTime As Integer
Dim imGenShiftKey As Integer    'Ctrl indictes to run MicroHelp reports
Dim smCommand As String 'Used to pass original command back to RptList if cancel pressed
Dim smSelectedRptName As String 'Passed selected report name
Dim imTerminate As Integer
Dim imListIndex As Integer

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

Private Sub cbcPrimaryDemo_Change()
    mSetCommands
End Sub

Private Sub cbcPrimaryDemo_Click()
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
    Dim llRg As Long
    Dim ilValue As Integer
    Dim llRet As Long
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcSelection(0).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(0).HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub

Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub

Private Sub ckcAllBooks_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcAllBooks(0).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim llRg As Long
    Dim ilValue As Integer
    Dim llRet As Long
    ilValue = Value
    If imSetAllBooks Then
        imAllBooksClicked = True
        llRg = CLng(lbcSelection(1).ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection(1).HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllBooksClicked = False
    End If
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

'******************************************************************************************
'           Generate Research Summary report (for 16 or demos)
'           or generate Special Research Summary report which is for
'           the estimates feature
'
Private Sub cmcGen_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilDnfCode                     slWhichRpt                    ilLoop                    *
'*  slNameCode                    slCode                                                  *
'******************************************************************************************
    Dim ilRet As Integer
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilNoJobs As Integer
    Dim ilJobs As Integer
    Dim ilStartJobNo As Integer
    Dim hlDnf As Integer
    Dim tlDnf As DNF

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
    'Research book Name Data File
    hlDnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlDnf)
        btrDestroy hlDnf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    igUsingCrystal = True
    ilNoJobs = 1
    ilStartJobNo = 1
    'imListIndex = lbcRptType.ListIndex

    For ilJobs = ilStartJobNo To ilNoJobs Step 1
        igJobRptNo = ilJobs

        If Not gGenReportRS(imListIndex, hlDnf, tlDnf) Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Exit Sub
        End If
        ilRet = gCmcGenRS(imListIndex)
        If ilRet = -1 Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            pbcClickFocus.SetFocus
            tmcDone.Enabled = True
            Exit Sub
        ElseIf ilRet = 0 Then
            igGenRpt = False
            frcOutput.Enabled = igOutput
            frcCopies.Enabled = igCopies
            frcFile.Enabled = igFile
            frcOption.Enabled = igOption
            Exit Sub
        End If

        Screen.MousePointer = vbHourglass
        If imListIndex = RS_SUMMARY Then
            gResearchReport_New LbcCustDemos
            Screen.MousePointer = vbDefault

        ElseIf imListIndex = RS_SPECIALSUMMARY Then
            gSpecialResearchReport
        ElseIf imListIndex = RS_DEMORANK Then           '9-10-15
            gDemoRankReport
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
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)
        End If

    Next ilJobs
    imGenShiftKey = 0

    igGenRpt = False
    frcOutput.Enabled = igOutput
    frcCopies.Enabled = igCopies
    frcFile.Enabled = igFile
    frcOption.Enabled = igOption
    pbcClickFocus.SetFocus
    tmcDone.Enabled = True
    Screen.MousePointer = vbHourglass
    'Dan 1/24/2011 created and added gClearRsr
    If imListIndex = RS_SUMMARY Then
        gClearRsr
    ElseIf imListIndex = RS_SPECIALSUMMARY Then
        'gClearGrf
        gCRGrfClear         '8-20-13 use only 1 common grf clear rtn which changes the way records are removed
    Else
        'clear temporary file
        gClearRsr
    End If
    Screen.MousePointer = vbDefault
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
    cdcSetup.flags = cdlPDPrintSetup
    cdcSetup.ShowPrinter
End Sub

Private Sub CSI_CalDate_Change()
    Dim slDate As String
    Dim llDate As Long
    Dim ilRet As Integer
    Dim ilLen As Integer
       
    slDate = CSI_CalDate.Text           'retrieve jan thru dec year
    llDate = gDateValue(slDate)
    'populate Research Book Names
    ilRet = mPopBookNameByDate(RptSelRS, 0, 1, 0, RptSelRS!lbcSelection(1), tgBookNameCode(), "0", llDate)
    mSetCommands
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
    '10-18-01 Previously didnt allow backspace
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
    RptSelRS.Refresh
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
    'RptSelRS.Show
    imFirstTime = True
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    mInitDDE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tgCSVNameCode
    PECloseEngine
    Set RptSelRS = Nothing   'Remove data segment
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
'*  ilRet                                                                                 *
'******************************************************************************************
    Dim ilRet As Integer
    Dim ilDemoLoop As Integer

    ckcNewPage.Visible = False
    imListIndex = lbcRptType.ListIndex
    lbcSelection(0).Visible = True                  'show budget name list box (base budget)
    ckcAllBooks(0).Visible = True
    lacBookNames.Visible = False
    If imListIndex = RS_SUMMARY Then
        'ilRet = mPopBookNameByDate(RptSelRS, 0, 0, 0, RptSelRS!lbcSelection(1), tmBNCode(), smBNCodeTag, 0)
        'cbcSet1.ListIndex = 0
        If rbcSelC11(1).Value Then
            lbcSelection(1).Visible = False
        End If
        ckcNewPage.Visible = True               '8-25-16 new page option
        ckcNewPage.Value = vbUnchecked
    Else
        If tgSpf.sDemoEstAllowed = "Y" Then
            If imListIndex = RS_SPECIALSUMMARY Then         'Special Research Summary
                plcSelC11.Visible = False
                plcSelC3.Visible = False
                plcSelC5.Visible = False
                ckcSelC3(0).Visible = False
 
            End If
        'no demo estimates, adjust the index for the report
        Else
            imListIndex = RS_DEMORANK                       'adjust for the missing report since demo estimates not allowed
        End If
        If imListIndex = RS_DEMORANK Then
            ilRet = gPopMnfPlusFieldsBox(RptSelRS, lbcSelection(2), tgRptSelDemoCodeCT(), sgRptSelDemoCodeTagCT, "D")
            cbcPrimaryDemo.Clear
            For ilDemoLoop = 0 To RptSelRS!lbcSelection(2).ListCount - 1 Step 1
'                slNameCode = tgRptSelDemoCodeCT(ilDemoLoop).sKey
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                cbcPrimaryDemo.AddItem Trim$(RptSelRS!lbcSelection(2).List(ilDemoLoop))
            Next ilDemoLoop
            cbcPrimaryDemo.ListIndex = 0
            
            lbcSelection(0).Visible = True          'vehicle list
            ckcAllBooks(0).Visible = False
            lbcSelection(1).Visible = False         'hide books list
            ckcAll.Visible = True                   'all vehicles check box
            lbcSelection(2).Visible = True          'demo list
            lacMaxDemoDescr.Caption = "Select up to 5 demo categories"
            lacMaxDemoDescr.Move 0, lbcSelection(0).Top + lbcSelection(0).Height + 30, 4000
            lacMaxDemoDescr.Visible = True
            lbcSelection(2).Move 0, lacMaxDemoDescr.Top + lacMaxDemoDescr.Height + 30, 4380, 1290
            lbcSelection(2).Visible = True
            edcSet1(0).Visible = False
'            edcDates.Visible = False
            CSI_CalDate.Visible = False
            plcSelC3.Visible = False
            plcSelC5.Visible = False
            plcSelC11.Visible = False
            ckcSelC3(0).Visible = False
            lacTopHowMany.Move 120, 240
            edcTopHowMany.Move 2535, 210
            lacTopHowMany.Visible = True
            edcTopHowMany.Visible = True
            
            lacDemoMsg.Move 120, edcTopHowMany.Top + edcTopHowMany.Height + 60
            lacPrimaryDemo.Move 120, lacDemoMsg.Top + lacDemoMsg.Height + 30
            cbcPrimaryDemo.Move lacPrimaryDemo.Width + 240, lacDemoMsg.Top + lacDemoMsg.Height
            lacPrimaryDemo.Visible = True
            cbcPrimaryDemo.Visible = True
            lacDemoMsg.Visible = True
            pbcSelC.Visible = True
        End If
        
    End If
    mSetCommands
End Sub

Private Sub lbcSelection_Click(Index As Integer)
    If Index = 0 Then
        If Not imAllClicked Then
            imSetAll = False
            ckcAll.Value = vbUnchecked
            imSetAll = True
        End If
    End If
    If Index = 1 Then
        If Not imAllBooksClicked Then
            imSetAllBooks = False
            ckcAllBooks(0).Value = vbUnchecked
            imSetAllBooks = True
        End If
    End If
    If Index = 2 Then           'demos
        If lbcSelection(2).SelCount > 5 Then
            MsgBox "Max 5 Demo Categories Allowed"
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

    RptSelRS.Caption = smSelectedRptName & " Report"
    slStr = Trim$(smSelectedRptName)
    'Handle the apersand in the option box
    illoop = InStr(slStr, "&")
    If illoop > 0 Then
        slStr = Left$(slStr, illoop - 1) & "&&" & Mid$(slStr, illoop + 1)
    End If
    frcOption.Caption = slStr & " Selection"
    imAllClicked = False
    imSetAll = True
    imAllBooksClicked = False
    imSetAllBooks = True

    gCenterStdAlone RptSelRS
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         ilVefCode                     ilSort                    *
'*  ilShow                                                                                *
'******************************************************************************************

    gPopExportTypes cbcFileType '10-19-01
    pbcSelC.Visible = False

    frcOption.Enabled = True
    lbcRptType.AddItem "Research", 0  'Research Summary
    If tgSpf.sDemoEstAllowed = "Y" Then             'using demo estimates - ok to continue with selectivity
        lbcRptType.AddItem "Special Research Summary", 1
        lbcRptType.AddItem "Demo Rank", RS_DEMORANK     '9-10-15 avg aud ranking by demo & vehicle

    Else
        lbcRptType.AddItem "Demo Rank", 1
    End If

    If lbcRptType.ListCount > 0 Then
        gFindMatch smSelectedRptName, 0, lbcRptType
        If gLastFound(lbcRptType) < 0 Then
            If tgSpf.sDemoEstAllowed = "Y" Then             'using demo estimates - ok to continue with selectivity
                MsgBox "Unable to Find Report Name " & smSelectedRptName, vbCritical, "Reports"
            Else
                MsgBox smSelectedRptName & "  - feature disabled", vbCritical, "Reports"
            End If
            imTerminate = True
            Exit Sub
        End If
        lbcRptType.ListIndex = gLastFound(lbcRptType)
    End If

    lbcSelection(0).Clear
    lbcSelection(0).Tag = ""
    mSellConvVirtVehPop 0
    Screen.MousePointer = vbHourglass

    frcOption.Enabled = True
    pbcSelC.Height = pbcSelC.Height - 60
    pbcSelC.Visible = True
    pbcOption.Visible = True

    pbcOption.Enabled = True
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
        
    ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    
    ilRet = gParseItem(slCommand, 2, "||", slRptListCmmd)
    If (ilRet = CP_MSG_NONE) Then
        ilRet = gParseItem(slRptListCmmd, 2, "\", smSelectedRptName)
        ilRet = gParseItem(slRptListCmmd, 1, "\", slStr)
        igRptCallType = Val(slStr)      'Function ID (what function calling this report if )
    End If

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPopBookNameBox                 *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:D. Smith       *
'*                                                     *
'*            Modified to only populate the box with   *
'*            book dates >= llEnteredDate              *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     Book Name and Date              *
'*                                                     *
'*******************************************************
Private Function mPopBookNameByDate(frm As Form, ilVefCode As Integer, ilSort As Integer, ilShow As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, llEnteredDate As Long) As Integer
'
'   ilRet = mPopBookNameBox (MainForm, ilVefCode, ilSort, ilShow, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilVefCode(I)- Vehicle code (zero if ignored, If specified, then check if Drf defined for vehcile)
'       ilSort(I)- 0=Sort by Book Name only; 1= sort by date, then book name
'       ilShow(I)- 0=Book Name only; 1=Book Name followed by Date
'       lbcLocal (I/O)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       llEnteredDate(I) - user entered book date
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'Dnf date/time stamp
    Dim hlDnf As Integer        'Dnf handle
    Dim ilDnfRecLen As Integer     'Record length
    Dim tlDnf As DNF
    Dim hlDrf As Integer        'Dnf handle
    Dim ilDrfRecLen As Integer     'Record length
    Dim tlDrf As DRF
    Dim tlDrfSrchKey As DRFKEY0
    Dim llNoRec As Long         'Number of records in Sof
    Dim slName As String
    Dim slDate As String
    Dim llDate As Long
    Dim slSortDate As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim ilOffSet As Integer
    Dim llLen As Long
    Dim ilAdd As Integer
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    ilPop = True
    slStamp = gFileDateTime(sgDBPath & "Dnf.Btr") & gFileDateTime(sgDBPath & "Drf.Btr") & Trim$(str$(ilVefCode)) & Trim$(str$(ilSort)) & Trim$(str$(ilShow))
    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                mPopBookNameByDate = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    mPopBookNameByDate = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlDnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mPopBookNameBoxErr
        gBtrvErrorMsg ilRet, "mPopBookNameBox (btrOpen):" & "Dnf.Btr", frm
        On Error GoTo 0
        ilDnfRecLen = Len(tlDnf) 'btrRecordLength(hlDnf)  'Get and save record length
        hlDrf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mPopBookNameBoxErr
        gBtrvErrorMsg ilRet, "mPopBookNameBox (btrOpen):" & "Drf.Btr", frm
        On Error GoTo 0
        ilDrfRecLen = Len(tlDrf) 'btrRecordLength(hlDrf)  'Get and save record length
        ilSortCode = 0
        ReDim tlSortCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlDnf)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlDnf) 'Obtain number of records
        btrExtClear hlDnf   'Clear any previous extend operation
        ilRet = btrGetFirst(hlDnf, tlDnf, ilDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hlDnf)
            btrDestroy hlDnf
            ilRet = btrClose(hlDrf)
            btrDestroy hlDrf
            Exit Function
        Else
            On Error GoTo mPopBookNameBoxErr
            gBtrvErrorMsg ilRet, "mPopBookNameBox (btrGetFirst):" & "Dnf.Btr", frm
            On Error GoTo 0
        End If
        Call btrExtSetBounds(hlDnf, llNoRec, -1, "UC", "DNF", "") 'Set extract limits (all records including first)
        ilOffSet = 0
        ilRet = btrExtAddField(hlDnf, ilOffSet, ilDnfRecLen)  'Extract iCode field
        On Error GoTo mPopBookNameBoxErr
        gBtrvErrorMsg ilRet, "mPopBookNameBox (btrExtAddField):" & "Dnf.Btr", frm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlDnf, tlDnf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mPopBookNameBoxErr
            gBtrvErrorMsg ilRet, "mPopBookNameBox (btrExtGetNextExt):" & "Dnf.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlDnf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlDnf, tlDnf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilAdd = False
                If ilVefCode <= 0 Then
                    ilAdd = True
                Else
                    'Test if Drf exist for vehicle specified
                    tlDrfSrchKey.iDnfCode = tlDnf.iCode
                    tlDrfSrchKey.sDemoDataType = "D"    'Demo Data
                    tlDrfSrchKey.iMnfSocEco = 0 'ilMnfSocEco
                    tlDrfSrchKey.iVefCode = ilVefCode
                    tlDrfSrchKey.sInfoType = "D"        'Daypart
                    tlDrfSrchKey.iRdfCode = 0
                    ilRet = btrGetGreaterOrEqual(hlDrf, tlDrf, ilDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    If (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tlDnf.iCode) And (tlDrf.iVefCode = ilVefCode) Then
                        ilAdd = True
                    Else
                        tlDrfSrchKey.iDnfCode = tlDnf.iCode
                        tlDrfSrchKey.sDemoDataType = "D"    'Demo Data
                        tlDrfSrchKey.iMnfSocEco = 0 'ilMnfSocEco
                        tlDrfSrchKey.iVefCode = ilVefCode
                        tlDrfSrchKey.sInfoType = "T"        'Daypart Or Time
                        tlDrfSrchKey.iRdfCode = 0
                        ilRet = btrGetGreaterOrEqual(hlDrf, tlDrf, ilDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        If (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tlDnf.iCode) And (tlDrf.iVefCode = ilVefCode) Then
                            ilAdd = True
                        Else
                            tlDrfSrchKey.iDnfCode = tlDnf.iCode
                            tlDrfSrchKey.sDemoDataType = "D"    'Demo Data
                            tlDrfSrchKey.iMnfSocEco = 0 'ilMnfSocEco
                            tlDrfSrchKey.iVefCode = ilVefCode
                            tlDrfSrchKey.sInfoType = "V"        'Daypart Or Time
                            tlDrfSrchKey.iRdfCode = 0
                            ilRet = btrGetGreaterOrEqual(hlDrf, tlDrf, ilDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            If (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tlDnf.iCode) And (tlDrf.iVefCode = ilVefCode) Then
                                ilAdd = True
                            End If
                        End If
                    End If
                End If
                If ilAdd Then
                    If llEnteredDate <> 0 Then
                        llEnteredDate = llEnteredDate
                    End If
                    If ilSort = 0 Then
                        'Add to list by Book Name Only
                        gUnpackDate tlDnf.iBookDate(0), tlDnf.iBookDate(1), slDate
                        If ilShow = 0 Then  'Show Book Name Only
                            slName = Trim$(tlDnf.sBookName)
                        Else                'Show Date and Book Name
                            slName = Trim$(tlDnf.sBookName) & ": " & slDate
                        End If
                        slName = slName & "\" & Trim$(str$(tlDnf.iCode))
                    Else
                        'Add to list by Date then Book Name
                        gUnpackDateLong tlDnf.iBookDate(0), tlDnf.iBookDate(1), llDate
                        llDate = 99999 - llDate
                        slSortDate = Trim$(str$(llDate))
                        Do While Len(slSortDate) < 5
                            slSortDate = "0" & slSortDate
                        Loop
                        gUnpackDate tlDnf.iBookDate(0), tlDnf.iBookDate(1), slDate
                        If ilShow = 0 Then
                            slName = Trim$(tlDnf.sBookName)
                        Else
                            slName = Trim$(tlDnf.sBookName) & ": " & slDate
                        End If
                        slName = slSortDate & "|" & slName & "\" & Trim$(str$(tlDnf.iCode))
                    End If

                    tlSortCode(ilSortCode).sKey = slName
                    If gDateValue(slDate) >= llEnteredDate Then
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                    End If
                End If
                ilRet = btrExtGetNext(hlDnf, tlDnf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlDnf, tlDnf, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(0 To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - 1 > 0 Then
                ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
            End If
        End If
        On Error GoTo 0
        ilRet = btrClose(hlDnf)
        btrDestroy hlDnf
        ilRet = btrClose(hlDrf)
        btrDestroy hlDrf
    End If
    llLen = 0
    For illoop = 0 To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        If ilSort = 0 Then
            ilRet = gParseItem(slNameCode, 1, "\", slName)
        Else
            ilRet = gParseItem(slNameCode, 2, "|", slName)
            slNameCode = slName
            ilRet = gParseItem(slNameCode, 1, "\", slName)
        End If
        If ilRet <> CP_MSG_NONE Then
            mPopBookNameByDate = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
mPopBookNameBoxErr:
    ilRet = btrClose(hlDnf)
    btrDestroy hlDnf
    ilRet = btrClose(hlDrf)
    btrDestroy hlDrf
    gDbg_HandleError "RptSelRS: mPopBookNameByDate"
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mSellConvVirtVehPop             *
'*                                                     *
'*             Created:6/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mSellConvVirtVehPop(ilIndex As Integer)
    Dim ilRet As Integer
        ilRet = gPopUserVehicleBox(RptSelRS, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHSPORT + ACTIVEVEH, lbcSelection(ilIndex), tgVehicle(), sgVehicleTag)
   
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSellConvVirtVehPopErr
        gCPErrorMsg ilRet, "mSellConvVirtVehPop (gPopUserVehicleBox: Vehicle)", RptSelRS
        On Error GoTo 0
    End If
    Exit Sub
mSellConvVirtVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/21/93       By:D. Smith       *
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
    Dim ilEnableVeh As Integer
    Dim ilEnableBook As Integer
    Dim illoop As Integer
    Dim slPrimaryDemo As String
    Dim llListindex As Long
    Dim slDemos As String
    Dim blFoundMatch As Boolean
    
    ilEnableVeh = False
    ilEnableBook = False

    cmcGen.Enabled = False
    
    If (RptSelRS!lbcRptType.ListIndex = RS_DEMORANK And tgSpf.sDemoEstAllowed = "Y") Or (RptSelRS!lbcRptType.ListIndex = 1 And tgSpf.sDemoEstAllowed = "N") Then    'adjust for report index when demo est not allowed; do not seee Special Research report in list
        If (lbcSelection(2).SelCount > 0 And lbcSelection(2).SelCount <= 5) And lbcSelection(0).SelCount > 0 Then           'max 5 demos allowed, plus at least one vehicle selected
            'see if the selected primary demo matches one of the 6 selected
            blFoundMatch = False
            llListindex = cbcPrimaryDemo.ListIndex
            slPrimaryDemo = cbcPrimaryDemo.List(llListindex)
            'see if one of the 6 selected matches the primary
            For illoop = 0 To RptSelRS!lbcSelection(2).ListCount - 1
                If RptSelRS!lbcSelection(2).Selected(illoop) Then
                    slDemos = Trim$(RptSelRS!lbcSelection(2).List(illoop))
                    If StrComp(Trim$(slPrimaryDemo), Trim$(slDemos), 1) = 0 Then
                        blFoundMatch = True
                        Exit For
                    End If
                End If
            Next illoop
            If blFoundMatch Then
                cmcGen.Enabled = True
            Else
                MsgBox "Primary demo selected does match one of the 5 demos to rank"
            End If
        End If
    Else
        If rbcSelC11(0).Value Then  'Sort by Book - must pick a book and a vehicle before generating
            If Not (ckcAll.Value = vbChecked) Then
                For illoop = 0 To lbcSelection(0).ListCount - 1 Step 1      'vehicle entry must be selected
                    If lbcSelection(0).Selected(illoop) Then
                        ilEnableVeh = True
                        Exit For
                    End If
                Next illoop
            Else
                ilEnableVeh = True
            End If
    
            If Not (ckcAllBooks(0).Value = vbChecked) Then                                        'book must be selected
                For illoop = 0 To lbcSelection(1).ListCount - 1 Step 1
                    If lbcSelection(1).Selected(illoop) Then
                        ilEnableBook = True
                        Exit For
                    End If
                Next illoop
            Else
                ilEnableBook = True
            End If
        Else       'Sort by Veh - must pick a vehicle only before generating
            If rbcSelC11(1).Value Then
            If Not (ckcAll.Value = vbChecked) Then
                For illoop = 0 To lbcSelection(0).ListCount - 1 Step 1  'vehicle entry must be selected
                    If lbcSelection(0).Selected(illoop) Then
                        ilEnableVeh = True
                        Exit For
                    End If
                Next illoop
            Else
                ilEnableVeh = True
            End If
    
            End If
        End If
        If ilEnableVeh And rbcSelC11(1).Value Then     'Sort by Vehicle
            cmcGen.Enabled = True
        Else                                           'Sort by Book
            If (ilEnableBook And rbcSelC11(0).Value) And (ilEnableVeh) Then
                cmcGen.Enabled = True
            End If
        End If
        If edcFileName.Text = "" And rbcOutput(2).Value Then
            cmcGen.Enabled = False
        End If
        If CSI_CalDate.Text = "" Then
            cmcGen.Enabled = False
        End If
    End If
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
    Unload RptSelRS
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
            Case 1  'Print
                frcFile.Visible = False
                frcFile.Enabled = False
                frcCopies.Enabled = True
                frcCopies.Visible = True
            Case 2  'File
                frcCopies.Visible = False
                frcCopies.Enabled = False
                frcFile.Enabled = True
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

Private Sub rbcSelC11_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcSelC11(Index).Value
    'End of coded added
    Dim ilRet As Integer
    If rbcSelC11(0).Value Then
        lbcSelection(1).Visible = True
        ckcAllBooks(0).Visible = True   '9-12-02 vbChecked
        ilRet = mPopBookNameByDate(RptSelRS, 0, 1, 0, RptSelRS!lbcSelection(1), tgBookNameCode(), "0", 0)
    Else
        lbcSelection(1).Visible = False
        lbcSelection(1).Clear
        ckcAllBooks(0).Visible = False  '9-12-02 vbUnchecked
    End If
End Sub

Private Sub tmcDone_Timer()
    tmcDone.Enabled = False
End Sub

Private Sub plcSelC5_Paint()
    plcSelC5.CurrentX = 0
    plcSelC5.CurrentY = 0
    plcSelC5.Print "Show"
End Sub

Private Sub plcSelC3_Paint()
    plcSelC3.CurrentX = 0
    plcSelC3.CurrentY = 0
    plcSelC3.Print "Demos"
End Sub

Private Sub plcSelC11_Paint()
    plcSelC11.CurrentX = 0
    plcSelC11.CurrentY = 0
    plcSelC11.Print "Sort by"
End Sub
