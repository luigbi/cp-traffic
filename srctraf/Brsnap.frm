VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form BrSnap 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proposal/Contract Report"
   ClientHeight    =   5790
   ClientLeft      =   1290
   ClientTop       =   2655
   ClientWidth     =   6825
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
   ScaleHeight     =   5790
   ScaleWidth      =   6825
   Begin VB.Timer tmcDone 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6195
      Top             =   75
   End
   Begin VB.Frame frcCopies 
      Caption         =   "Printing"
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2520
      TabIndex        =   31
      Top             =   120
      Width           =   3900
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   33
         Top             =   825
         Width           =   1260
      End
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
         TabIndex        =   32
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   34
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame frcFile 
      Caption         =   "Save to File"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2520
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   3900
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
         TabIndex        =   28
         Top             =   615
         Width           =   2925
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
         Height          =   315
         Left            =   780
         TabIndex        =   27
         Top             =   270
         Width           =   2925
      End
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   26
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   30
         Top             =   645
         Width           =   645
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   255
         Width           =   615
      End
   End
   Begin VB.Frame frcOutput 
      Caption         =   "Report Destination"
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   360
      TabIndex        =   21
      Top             =   120
      Width           =   1890
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Display"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Print"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   555
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Save to File"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   840
         Width           =   1275
      End
   End
   Begin VB.Frame frcSnapBR 
      Caption         =   "Snapshot Selection"
      Height          =   3795
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   6090
      Begin VB.CheckBox ckcSuppressNTRDetails 
         Caption         =   "Suppress NTR Details"
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
         Height          =   195
         Left            =   3360
         TabIndex        =   44
         Top             =   3435
         Width           =   2325
      End
      Begin VB.CheckBox ckcShowACT1 
         Caption         =   "Show ACT1 Codes and Settings"
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
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   3435
         Value           =   1  'Checked
         Width           =   3120
      End
      Begin VB.CheckBox ckcProdcastInfo 
         Caption         =   "Include Audience Percentages for Podcast"
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
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   3150
         Width           =   4125
      End
      Begin VB.CheckBox ckcShowProdProt 
         Caption         =   "Show Product Protection"
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
         Height          =   195
         Left            =   3120
         TabIndex        =   41
         Top             =   2865
         Width           =   2685
      End
      Begin VB.CheckBox ckcShowNetOnProps 
         Caption         =   "Show Net Amt on Proposals"
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
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   2865
         Value           =   1  'Checked
         Width           =   2805
      End
      Begin VB.CheckBox ckcShowNTRBillSummary 
         Caption         =   "Combine Air Time and NTR/CPM Totals"
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
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   2595
         Value           =   1  'Checked
         Width           =   3765
      End
      Begin VB.PictureBox plcRptType 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   5025
         TabIndex        =   36
         Top             =   270
         Width           =   5025
         Begin VB.OptionButton rbcRptType 
            Caption         =   "Proposal/Contract"
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
            Height          =   195
            Index           =   0
            Left            =   540
            TabIndex        =   38
            Top             =   0
            Value           =   -1  'True
            Width           =   2025
         End
         Begin VB.OptionButton rbcRptType 
            Caption         =   "Order Audit"
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
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   0
            Width           =   1620
         End
      End
      Begin VB.CheckBox ckcShowSplit 
         Caption         =   "Show Slsp Commission Splits on Summary"
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
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   2310
         Value           =   1  'Checked
         Width           =   4320
      End
      Begin VB.CheckBox ckcDiff 
         Caption         =   "Difference Only"
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
         Height          =   195
         Left            =   3960
         TabIndex        =   20
         Top             =   2595
         Width           =   1800
      End
      Begin VB.CheckBox ckcProof 
         Caption         =   "Hidden"
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
         Height          =   195
         Left            =   1230
         TabIndex        =   19
         Tag             =   "Set print flag on for contract"
         Top             =   2070
         Width           =   1140
      End
      Begin VB.CheckBox ckcResearch 
         Caption         =   "Research"
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
         Height          =   195
         Left            =   1230
         TabIndex        =   18
         Tag             =   "Set print flag on for contract"
         Top             =   1815
         Value           =   1  'Checked
         Width           =   1140
      End
      Begin VB.CheckBox ckcRating 
         Caption         =   "Rates"
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
         Height          =   195
         Left            =   1230
         TabIndex        =   17
         Tag             =   "Set print flag on for contract"
         Top             =   1560
         Value           =   1  'Checked
         Width           =   930
      End
      Begin VB.PictureBox plcInclude 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   960
         TabIndex        =   16
         Top             =   1605
         Width           =   960
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
         Left            =   1215
         TabIndex        =   15
         Top             =   1155
         Width           =   2340
      End
      Begin VB.TextBox lacText 
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
         TabIndex        =   14
         Text            =   "Qualitative"
         Top             =   1185
         Width           =   1080
      End
      Begin VB.PictureBox plcDemo 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   3225
         TabIndex        =   11
         Top             =   840
         Width           =   3225
         Begin VB.OptionButton rbcDemo 
            Caption         =   "All"
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
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   540
         End
         Begin VB.OptionButton rbcDemo 
            Caption         =   "Primary"
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
            Height          =   195
            Index           =   0
            Left            =   660
            TabIndex        =   12
            Top             =   0
            Value           =   -1  'True
            Width           =   945
         End
      End
      Begin VB.PictureBox plcShow 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   4215
         TabIndex        =   7
         Top             =   555
         Width           =   4210
         Begin VB.OptionButton rbcShow 
            Caption         =   "Both"
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
            Height          =   195
            Index           =   2
            Left            =   3360
            TabIndex        =   10
            Top             =   0
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton rbcShow 
            Caption         =   "Summary"
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
            Height          =   195
            Index           =   1
            Left            =   2220
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton rbcShow 
            Caption         =   "Schedule Lines"
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
            Height          =   195
            Index           =   0
            Left            =   540
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   1665
         End
      End
      Begin MSComDlg.CommonDialog cdcSetup 
         Left            =   5280
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   ".pdf"
         Filter          =   $"Brsnap.frx":0000
         FilterIndex     =   2
         FontSize        =   0
         MaxFileSize     =   256
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
      Left            =   480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
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
      Left            =   4920
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5400
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
      Left            =   5640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Top             =   5400
      Width           =   1050
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
      Left            =   -30
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1620
      Width           =   120
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "&Generate"
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   5400
      Width           =   1050
   End
End
Attribute VB_Name = "BrSnap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Brsnap.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: BrSnap.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract Snap Shot Print screen code
Option Explicit
Option Compare Text

Dim imFirstActivate As Integer
Dim imFTSelectedIndex As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer

Dim hmCHF As Integer            'Contract BR file handle
Dim tmChf As CHF    'Used when updating only to eliminate conflict
Dim imCHFRecLen As Integer
Dim tmChfSrchKey1 As CHFKEY1  'CHF key record image (contract #)
Dim hmClf As Integer            'Contract BR file handle
Dim imClfRecLen As Integer      'CBF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract BR file handle
Dim imCffRecLen As Integer      'CBF record length
Dim tmCff As CFF
Dim tmTChf As CHF     'for differnces only mode
Dim tmTClf() As CLFLIST 'for differences only mode (saved image of screen)
Dim tmTCff() As CFFLIST 'for differences only mode (saved image of screen)
Dim imTerminate As Integer
Dim imShowDiff As Integer
Dim smSnapshot As String
Dim imFirstFocus As Integer

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
    'dan M 12/01/2011
    gAdjustCDCFilter imFTSelectedIndex, cdcSetup
    cdcSetup.flags = cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
    cdcSetup.fileName = edcFileName.Text
    cdcSetup.InitDir = Left$(sgRptSavePath, Len(sgRptSavePath) - 1)
    cdcSetup.Action = 2    'DLG_FILE_SAVE
    edcFileName.Text = cdcSetup.fileName
    gChDrDir        '3-25-03
    ChDir sgCurDir
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcGen_Click()
    Dim ilLoop As Integer
    Dim slStr As String

    Dim ilRet As Integer
    Dim llCurrCode As Long
    Dim llPrevCode As Long
    Dim llPrevSpots As Long
    Dim llPrevGross As Long
    Dim ilCurrTotalWks As Integer
    Dim ilCurrAirWks As Integer
    Dim slDate As String            'generation date
    Dim slTime As String            'generation time
    Dim slDetSumBoth As String      '0=det, 1 = sum, 2 = both
    Dim slInclRates As String       'Show rates on printed contract Y/N
    Dim slInclResearch As String    'show research data on printed contracts Y/N
    Dim slInclProof As String       'show hidden lines on printed contracts Y/N
    Dim ilChfRecLen As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slOutputTo As String * 1
    Dim slSaveToFileName As String
    Dim slExportIndex As String
    Dim slInclSplits As String      '2-14-04 show slsp comm splits
    Dim slInclNTRBillSummary As String  '2-2-10 show NTR bill summary with air time summary
    Dim slShowNetOnProps As String      '2-3-10
    Dim ilCopies As Integer
    Dim slFileName As String
    Dim ilErr As Integer
    Dim tlSbfList() As SBF
    Dim llStdStartDates(0 To 37) As Long            'standard start dates for bdcst calendar, 4 years, Index 0 ignored
    Dim slShowProdProt As String
    
    Dim tlBR As BRSELECTIONS        'BR options
    Dim ilGenTime(0 To 1) As Integer    '10-30-01
    'TTP 10549 - Learfield Cloud printing 911, Crystal Crashes, Use Adobe
    Dim olWshShell As Object
    Dim slTempPath As String

    'gObtainCorpCal             'corp calendar (if used), has been retrieved in mInitReport , stored in tgMCof
    Screen.MousePointer = vbHourglass

    If rbcRptType(1).Value Then         'order audit
'        If igTestSystem Then
'            slStr = "Traffic^Test^NOHELP\" & sgUserName
'        Else
'            slStr = "Traffic^Prod^NOHELP\" & sgUserName
'        End If
'        slSaveToFileName = ""
'        If rbcOutput(0).Value = True Then
'            slOutputTo = "0"
'        ElseIf rbcOutput(1).Value = True Then
'            slOutputTo = "1"
'        Else
'            slOutputTo = "2"
'            slExportIndex = Trim$(str(imFTSelectedIndex))
'            slSaveToFileName = Trim$(edcFileName)      'save to location
'        End If
'
'        slDate = Format$(gNow(), "m/d/yy")
'        gPackDate slDate, tlBR.iGenDate(0), tlBR.iGenDate(1)
'        igNowDate(0) = tlBR.iGenDate(0)
'        igNowDate(1) = tlBR.iGenDate(1)
'        slTime = Format$(gNow(), "h:mm:ssAM/PM")
'        gPackTime slTime, ilGenTime(0), ilGenTime(1)    '10-30-01
'        gUnpackTimeLong ilGenTime(0), ilGenTime(1), False, lgNowTime
'
'        'parm 1:  Application name^Test or Prod
'        'parm 2:  user name
'        'parm 3:  Generation date
'        'parm 4:  Generation time
'        'Parm 5: Display(0) or Print(1) (always 1)
'        'Parm 6: File export index
'        'Parm 7: filename for export
'
'        slStr = slStr & "\" & slDate & "\" & slTime & "\" & slOutputTo & "\" & slExportIndex & "\|" & slSaveToFileName
'
'        sgCommandStr = slStr
'        RptSelOA.Show vbModal           'Order Audit



        ilRet = PEOpenEngine()
        If ilRet = 0 Then
            MsgBox "Unable to open print engine"
            Exit Sub
        End If

        igGenRpt = True
        igOutput = frcOutput.Enabled
        igCopies = frcCopies.Enabled
        igFile = frcFile.Enabled
        igOption = frcSnapBR.Enabled
        frcOutput.Enabled = False
        frcCopies.Enabled = False
        frcFile.Enabled = False
        frcSnapBR.Enabled = False
        igUsingCrystal = True               'all versions of printed contract uses crystal

        If Not gGenReportOA() Then
             igGenRpt = False
             frcOutput.Enabled = igOutput
             frcCopies.Enabled = igCopies
             frcFile.Enabled = igFile
             frcSnapBR.Enabled = igOption
             Exit Sub
         End If
         ilRet = gCmcGenOA()
         '-1 is a Crystal failure of gSetSelection or gSEtFormula
         If ilRet = -1 Then
             igGenRpt = False
             frcOutput.Enabled = igOutput
             frcCopies.Enabled = igCopies
             'frcWhen.Enabled = igWhen
             frcFile.Enabled = igFile
             frcSnapBR.Enabled = igOption
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
             frcSnapBR.Enabled = igOption
             'frcRptType.Enabled = igReportType
             Exit Sub
         ElseIf ilRet = 2 Then           'successful from Bridgereport
             igGenRpt = False
             frcOutput.Enabled = igOutput
             frcCopies.Enabled = igCopies
             'frcWhen.Enabled = igWhen
             frcFile.Enabled = igFile
             frcSnapBR.Enabled = igOption
             'frcRptType.Enabled = igReportType
             pbcClickFocus.SetFocus
             tmcDone.Enabled = True
             Exit Sub
        End If
        'process NTR from table tgIBSbf built into memory  (in case contract is undone)
        ReDim tlSbfList(0 To 0) As SBF
        'put the records in a common array so that the same report from report module can use routine
        For ilLoop = LBound(tgIBSbf) To UBound(tgIBSbf) - 1
            tlSbfList(UBound(tlSbfList)) = tgIBSbf(ilLoop).SbfRec
            ReDim Preserve tlSbfList(0 To UBound(tlSbfList) + 1) As SBF
        Next ilLoop

        ilErr = gOrderAudit(llStdStartDates(), tlSbfList())          'process contracts
        gOrderAuditWrite llStdStartDates()

        Screen.MousePointer = vbDefault
        If rbcOutput(0).Value Then
            'TTP 10549 - Learfield Cloud printing 911, Crystal Crashes, Use Adobe
            If bgUseAdobe = True And sgReportFilename <> "" Then
                slFileName = edcFileName.Text
                '12-16-03 alter filenames based on which contract version (detail, summary notation: up to 4 passes)
                slFileName = sgReportFilename
                If InStr(slFileName, ".") = 0 Then  'no extension specified
                    slFileName = Trim(slFileName) & "-" & Trim$(str(igJobRptNo))
                Else
                    'name already has extension, need to insert the contract version (detail, summary notation: up to 4passes)
                    ilLoop = InStr(slFileName, ".")     'find the period before extension name
                    slStr = Trim$(Mid$(slFileName, 1, ilLoop - 1)) & "-" & Trim$(str(igJobRptNo)) & Trim$(Mid(slFileName, ilLoop))
                    slFileName = slStr
                End If
                ilRet = gExportCRW(slFileName, imFTSelectedIndex, False, sgReportTempFolder)
            Else
                igDestination = 0
                DoEvents            '9-19-02 fix for timing problem to prevent freezing before calling crystal
                Report.Show vbModal
            End If
        ElseIf rbcOutput(1).Value Then
            ilCopies = Val(edcCopies.Text)
            ilRet = gOutputToPrinter(ilCopies)
        Else
            slFileName = edcFileName.Text
            '12-16-03 alter filenames based on which contract version (detail, summary notation: up to 4 passes)
            If InStr(slFileName, ".") = 0 Then  'no extension specified
                slFileName = Trim(slFileName) & "-" & Trim$(str(igJobRptNo))
            Else
                'name already has extension, need to insert the contract version (detail, summary notation: up to 4passes)
                ilLoop = InStr(slFileName, ".")     'find the period before extension name
                slStr = Trim$(Mid$(slFileName, 1, ilLoop - 1)) & "-" & Trim$(str(igJobRptNo)) & Trim$(Mid(slFileName, ilLoop))
                slFileName = slStr
            End If
            ilRet = gExportCRW(slFileName, imFTSelectedIndex)   '10-20-01
        End If

        Screen.MousePointer = vbHourglass
        gCrCbfClear                    '#1
        
        'gCbfClearBylgNowTime           '#2 9-14-09 another version to clear by millseconds time value

        Screen.MousePointer = vbDefault
        igGenRpt = False
        frcOutput.Enabled = igOutput
        frcCopies.Enabled = igCopies
        'frcWhen.Enabled = igWhen
        frcFile.Enabled = igFile
        frcSnapBR.Enabled = igOption
        pbcClickFocus.SetFocus
        tmcDone.Enabled = True
        Exit Sub
    Else                                'proposals/contr
        If imShowDiff = 2 Then
            ''parm 1:  Application name^Test or Prod
            ''parm 2:  user name
            ''parm 3:  Generation date
            ''parm 4:  Generation time
            ''parm 5:  Summary #
            ''parm 6:  Include Rate fields
            ''Parm 7: Display(0) or Print(1) (always 1)
            'ilUnFmt = UNFMTDEFAULT
            'slDetSumBoth = Trim$(Str$(imDBTotals))
            'If rbcPrice(0).Value Then
            '    slInclRates = "Y"
            'Else
            '    slInclRates = "N"
            'End If
            'slDate = Format$(Now, "m/d/yy")
            'gPackDate slDate, tmCbf.iGenDate(0), tmCbf.iGenDate(1)
            'slTime = Format$(Now, "h:mm:ssAM/PM")
            'gPackTime slTime, tmCbf.iGenTime(0), tmCbf.iGenTime(1)
            'tmCbf.lContrNo = tgChf.lCntrNo
            'If Contract!lbcAdvt.ListIndex >= 1 Then
            '    slNameCode = tgAdvertiser(Contract!lbcAdvt.ListIndex - 1).sKey  'Traffic!lbcAdvertiser.List(lbcAdvt.ListIndex - 1)
            '    ilRet = gParseItem(slNameCode, 1, "\", slName)
            '    tmCbf.sSurvey = slName
            'Else
            '    tmCbf.sSurvey = ""
            'End If
            'tmCbf.sProduct = tgChf.sProduct
            'If Contract!lbcDBSocEco.ListIndex >= 1 Then
            '    tmCbf.sDemos = Trim$(Contract!edcDBDemo.Text) & " " & Trim$(Contract!edcDBSocEco.Text)
            'Else
            '    tmCbf.sDemos = Trim$(Contract!edcDBDemo.Text)
            'End If
            'ilFoundCnt = False
            'Select Case imDBTotals
            '    Case 5
            '        For ilLoop = LBound(smLnSumShow, 2) To UBound(smLnSumShow, 2) - 1 Step 1
            '            If Trim$(smLnSumShow(12, ilLoop)) = "T" Then
            '                ilFoundCnt = True
            '                Exit For
            '            End If
            '        Next ilLoop
            '    Case 6
            '        ilFoundCnt = False
            '    Case 7
            '        ilFoundCnt = False
            '    Case 8
            '        For ilLoop = LBound(smVSumShow, 2) To UBound(smVSumShow, 2) - 1 Step 1
            '            If Trim$(smVSumShow(11, ilLoop)) = "T" Then
            '                ilFoundCnt = True
            '                Exit For
            '            End If
            '        Next ilLoop
            '    Case 9
            '        For ilLoop = LBound(smDPSumShow, 2) To UBound(smDPSumShow, 2) - 1 Step 1
            '            If Trim$(smDPSumShow(11, ilLoop)) = "T" Then
            '                ilFoundCnt = True
            '                Exit For
            '            End If
            '        Next ilLoop
            'End Select
            'llSeqCount = 0
            'Select Case imDBTotals
            '    Case 5
            '        For ilLoop = LBound(smLnSumShow, 2) To UBound(smLnSumShow, 2) - 1 Step 1
            '            If (Not ilFoundCnt) Or ((Trim$(smLnSumShow(12, ilLoop)) = "T") And (ilFoundCnt)) Then
            '                For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
            '                    If Val(smLnSumShow(1, ilLoop)) = tgClf(ilClf).ClfRec.iLine Then
            '                        llSeqCount = llSeqCount + 1
            '                        tmCbf.lLineNo = llSeqCount
            '                        tmCbf.sSortField1 = smLnSave(1, ilClf + 1)
            '                        For ilDay = 0 To 6 Step 1
            '                            ilOvDays(ilDay) = imLnSave(12 + ilDay, ilClf + 1)
            '                        Next ilDay
            '                        tmCbf.sSortField2 = mSetDPShowName(ilClf + 1, -1, ilOvDays())
            '                        tmCbf.lGRP = gStrDecToLong(smLnSumShow(4, ilLoop), 1)
            '                        tmCbf.iPctDist = Val(smLnSumShow(5, ilLoop))
            '                        If tgSpf.sSAudData <> "H" Then
            '                            tmCbf.lGrImp = Val(smLnSumShow(6, ilLoop))
            '                        Else
            '                            tmCbf.lGrImp = gStrDecToLong(smLnSumShow(6, ilLoop), 1)
            '                        End If
            '                        tmCbf.iPctTrade = Val(smLnSumShow(7, ilLoop))
            '                        tmCbf.lCPM = gStrDecToLong(smLnSumShow(8, ilLoop), 2)
            '                        tmCbf.lCPP = Val(smLnSumShow(9, ilLoop))
            '                        slStr = smLnSumShow(10, ilLoop)
            '                        gUnformatStr slStr, ilUnFmt, slValue
            '                        tmCbf.lRate = gStrDecToLong(slValue, 2)
            '                        slStr = smLnSumShow(11, ilLoop)
            '                        gUnformatStr slStr, ilUnFmt, slValue
            '                        tmCbf.lExtra4Byte = gStrDecToLong(slValue, 2)
            '                        ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            '                        Exit For
            '                    End If
            '                Next ilClf
            '            End If
            '        Next ilLoop
            '    Case 6
            '        For ilLoop = LBound(smWkSumShow, 2) To UBound(smWkSumShow, 2) - 1 Step 1
            '            llSeqCount = llSeqCount + 1
            '            tmCbf.lLineNo = llSeqCount
            '            tmCbf.sSortField1 = smWkSumShow(1, ilLoop)
            '            tmCbf.sSortField2 = ""
            '            tmCbf.lGRP = gStrDecToLong(smWkSumShow(2, ilLoop), 1)
            '            tmCbf.iPctDist = Val(smWkSumShow(3, ilLoop))
            '            If tgSpf.sSAudData <> "H" Then
            '                tmCbf.lGrImp = Val(smWkSumShow(4, ilLoop))
            '            Else
            '                tmCbf.lGrImp = gStrDecToLong(smWkSumShow(4, ilLoop), 1)
            '            End If
            '            tmCbf.iPctTrade = Val(smWkSumShow(5, ilLoop))
            '            tmCbf.lCPM = gStrDecToLong(smWkSumShow(6, ilLoop), 2)
            '            tmCbf.lCPP = Val(smWkSumShow(7, ilLoop))
            '            slStr = smWkSumShow(8, ilLoop)
            '            gUnformatStr slStr, ilUnFmt, slValue
            '            tmCbf.lRate = gStrDecToLong(slValue, 2)
            '            slStr = smWkSumShow(9, ilLoop)
            '            gUnformatStr slStr, ilUnFmt, slValue
            '            tmCbf.lExtra4Byte = gStrDecToLong(slValue, 2)
            '            ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            '        Next ilLoop
            '    Case 7
            '        For ilLoop = LBound(smQSumShow, 2) To UBound(smQSumShow, 2) - 1 Step 1
            '            llSeqCount = llSeqCount + 1
            '            tmCbf.lLineNo = llSeqCount
            '            tmCbf.sSortField1 = smQSumShow(1, ilLoop)
            '            tmCbf.sSortField2 = ""
            '            'GRP
            '            tmCbf.lGRP = gStrDecToLong(smQSumShow(2, ilLoop), 1)
            '            '%GRP
            '            tmCbf.iPctDist = Val(smQSumShow(3, ilLoop))
            '            'GrImp
            '            If tgSpf.sSAudData <> "H" Then
            '                tmCbf.lGrImp = Val(smQSumShow(4, ilLoop))
            '            Else
            '                tmCbf.lGrImp = gStrDecToLong(smQSumShow(4, ilLoop), 1)
            '            End If
            '            '% GrImp
            '            tmCbf.iPctTrade = Val(smQSumShow(5, ilLoop))
            '            'CPM
            '            tmCbf.lCPM = gStrDecToLong(smQSumShow(6, ilLoop), 2)
            '            'CPP
            '            tmCbf.lCPP = Val(smQSumShow(7, ilLoop))
            '            'Avg Price
            '            slStr = smQSumShow(8, ilLoop)
            '            gUnformatStr slStr, ilUnFmt, slValue
            '            tmCbf.lRate = gStrDecToLong(slValue, 2)
            '            'Total Cost
            '            slStr = smQSumShow(9, ilLoop)
            '            gUnformatStr slStr, ilUnFmt, slValue
            '            tmCbf.lExtra4Byte = gStrDecToLong(slValue, 2)
            '            ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            '        Next ilLoop
            '    Case 8
            '        For ilLoop = LBound(smVSumShow, 2) To UBound(smVSumShow, 2) - 1 Step 1
            '            If (Not ilFoundCnt) Or ((Trim$(smVSumShow(11, ilLoop)) = "T") And (ilFoundCnt)) Then
            '                llSeqCount = llSeqCount + 1
            '                tmCbf.lLineNo = llSeqCount
            '                tmCbf.sSortField1 = smVSumShow(1, ilLoop)
            '                tmCbf.sSortField2 = ""
            '                'GRP
            '                tmCbf.lGRP = gStrDecToLong(smVSumShow(2, ilLoop), 1)
            '                '%GRP
            '                tmCbf.iPctDist = Val(smVSumShow(3, ilLoop))
            '                'GrImp
            '                If tgSpf.sSAudData <> "H" Then
            '                    tmCbf.lGrImp = Val(smVSumShow(4, ilLoop))
            '                Else
            '                    tmCbf.lGrImp = gStrDecToLong(smVSumShow(4, ilLoop), 1)
            '                End If
            '                '% GrImp
            '                tmCbf.iPctTrade = Val(smVSumShow(5, ilLoop))
            '                'CPM
            '                tmCbf.lCPM = gStrDecToLong(smVSumShow(6, ilLoop), 2)
            '                'CPP
            '                tmCbf.lCPP = Val(smVSumShow(7, ilLoop))
            '                'Avg Price
            '                slStr = smVSumShow(8, ilLoop)
            '                gUnformatStr slStr, ilUnFmt, slValue
            '                tmCbf.lRate = gStrDecToLong(slValue, 2)
            '                'Total Cost
            '                slStr = smVSumShow(9, ilLoop)
            '                gUnformatStr slStr, ilUnFmt, slValue
            '                tmCbf.lExtra4Byte = gStrDecToLong(slValue, 2)
            '                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            '            End If
            '        Next ilLoop
            '    Case 9
            '        For ilLoop = LBound(smDPSumShow, 2) To UBound(smDPSumShow, 2) - 1 Step 1
            '            If (Not ilFoundCnt) Or ((Trim$(smDPSumShow(11, ilLoop)) = "T") And (ilFoundCnt)) Then
            '                llSeqCount = llSeqCount + 1
            '                tmCbf.lLineNo = llSeqCount
            '                tmCbf.sSortField1 = smDPSumShow(1, ilLoop)
            '                tmCbf.sSortField2 = ""
            '                'GRP
            '                tmCbf.lGRP = gStrDecToLong(smDPSumShow(2, ilLoop), 1)
            '                '%GRP
            '                tmCbf.iPctDist = Val(smDPSumShow(3, ilLoop))
            '                'GrImp
            '                If tgSpf.sSAudData <> "H" Then
            '                    tmCbf.lGrImp = Val(smDPSumShow(4, ilLoop))
            '                Else
            '                    tmCbf.lGrImp = gStrDecToLong(smDPSumShow(4, ilLoop), 1)
            '                End If
            '                '% GrImp
            '                tmCbf.iPctTrade = Val(smDPSumShow(5, ilLoop))
            '                'CPM
            '                tmCbf.lCPM = gStrDecToLong(smDPSumShow(6, ilLoop), 2)
            '                'CPP
            '                tmCbf.lCPP = Val(smDPSumShow(7, ilLoop))
            '                'Avg Price
            '                slStr = smDPSumShow(8, ilLoop)
            '                gUnformatStr slStr, ilUnFmt, slValue
            '                tmCbf.lRate = gStrDecToLong(slValue, 2)
            '                'Total Cost
            '                slStr = smDPSumShow(9, ilLoop)
            '                gUnformatStr slStr, ilUnFmt, slValue
            '                tmCbf.lExtra4Byte = gStrDecToLong(slValue, 2)
            '                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            '            End If
            '        Next ilLoop
            'End Select
            'tmCbf.lLineNo = 99999
            'tmCbf.sSortField1 = ""
            'tmCbf.sSortField2 = ""
            ''GRP
            'tmCbf.lGRP = gStrDecToLong(smTSumShow(1), 1)
            ''%GRP
            'tmCbf.iPctDist = Val(smTSumShow(2))
            ''GrImp
            'If tgSpf.sSAudData <> "H" Then
            '    tmCbf.lGrImp = Val(smTSumShow(3))
            'Else
            '    tmCbf.lGrImp = gStrDecToLong(smTSumShow(3), 1)
            'End If
            ''% GrImp
            'tmCbf.iPctTrade = Val(smTSumShow(4))
            ''CPM
            'tmCbf.lCPM = gStrDecToLong(smTSumShow(5), 2)
            ''CPP
            'tmCbf.lCPP = Val(smTSumShow(6))
            ''Avg Price
            'slStr = smTSumShow(7)
            'gUnformatStr slStr, ilUnFmt, slValue
            'tmCbf.lRate = gStrDecToLong(slValue, 2)
            ''Total Cost
            'slStr = smTSumShow(8)
            'gUnformatStr slStr, ilUnFmt, slValue
            'tmCbf.lExtra4Byte = gStrDecToLong(slValue, 2)
            'ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            '
            'If igTestSystem Then
            '    slStr = "Traffic^Test^NOHELP\" & sgUserName
            'Else
            '    slStr = "Traffic^Prod^NOHELP\" & sgUserName
            'End If
            'slStr = slStr & "\" & slDate & "\" & slTime & "\" & slDetSumBoth & "\" & slInclRates & "\1"
            'ilShell = Shell(sgExePath & "RptSelDB.Exe " & slStr, 1)
            'While GetModuleUsage(ilShell) > 0
            '    ilRet = DoEvents()
            'Wend
        Else
            
            tlBR.sSnapshot = smSnapshot     'flag to indicate snapshot rather than from reports
            tlBR.iShowProof = False     'force to not show Proof mode
'            tlBR.iCorpOrStd = True      'forc standard monthly summaries
            tlBR.iCorpOrStd = 0         '12-19-20 0 = std, 1 = cal, 2 = corp; was true/fasle, force standard monthly summaries
            If tgChf.sBillCycle = "C" Then
                tlBR.iCorpOrStd = 1         'Calendar
            End If
            tlBR.iPrintables = False
            tlBR.iDiffOnly = False
            tlBR.iThisCntMod = False
            tlBR.iShowSplits = True     'always assume to show the splits on summaries
            tlBR.iShowNTRBillSummary = True 'assume to show ntr monthly billing with air time
            tlBR.iShowNetAmtOnProps = True  'assume to show net amts on proposals
            tlBR.iShowProdProt = False          '8-25-15 default to exclude showing Prod Protection categories
            If ckcDiff.Value = vbChecked Then                  '10-30-01 presence of code indicates to do differences
                tlBR.iDiffOnly = True
                tlBR.iThisCntMod = True
            End If

             If BrSnap!ckcRating.Value = vbChecked Then '10-30-01
                tlBR.iShowRates = True
             End If
             If BrSnap!ckcResearch.Value = vbChecked Then   '10-30-01
                tlBR.iShowResearch = True
             End If
             slInclProof = "N"
             If BrSnap!ckcProof.Value = vbChecked Then  '10-30-01
                tlBR.iShowProof = True
             End If
             If BrSnap!rbcShow(0).Value Or BrSnap!rbcShow(2).Value Then         'is selection Detail or Both?, if so set include detail
                tlBR.iDetail = True
             End If
             If BrSnap!rbcShow(1).Value Or BrSnap!rbcShow(2).Value Then         'is selection summary or both, if so set include summary
                tlBR.iSummary = True
             End If
             If BrSnap!rbcDemo(0).Value Then          'Primary demos (vs all)
                tlBR.iAllDemos = False
             Else
                tlBR.iAllDemos = True
             End If

             '2-13-04       show slsp comm splits?
             If BrSnap!ckcShowSplit.Value <> vbChecked Then
                tlBR.iShowSplits = False
            End If
            
            '2-2-10 show ntr billing summary with air time bill summary
            If BrSnap!ckcShowNTRBillSummary.Value = vbUnchecked Then
                tlBR.iShowNTRBillSummary = False
            End If
            
            '2-3-10 show net on proposals
            If BrSnap!ckcShowNetOnProps.Value = vbUnchecked Then
                tlBR.iShowNetAmtOnProps = False
            End If
             
            '8-25-15 Product protection
            If BrSnap!ckcShowProdProt.Value = vbChecked Then
                tlBR.iShowProdProt = True
            End If
            'TTP 10382 - Contract report: Option To not show Act1 codes on PDF
            If BrSnap!ckcShowACT1.Value = vbChecked Then
                tlBR.iShowAct1 = True
            End If
             '10-29-03 has a social economic category been selected?  if so, show that instead of the normal research
             tlBR.iSocEcoMnfCode = 0
             If BrSnap!cbcSet1.ListIndex > 0 Then
                slNameCode = tgSocEcoCode(BrSnap!cbcSet1.ListIndex - 1).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slName)
                tlBR.iSocEcoMnfCode = Val(slName)
             End If

             tlBR.iWhichSort = 0         'force to sort by advt, contr#

            llCurrCode = tgChf.lCode
            llPrevCode = 0
            llPrevSpots = 0
            llPrevGross = 0
            ilCurrTotalWks = 0
            ilCurrAirWks = 0

            tlBR.iPropOrOrder = False               'assume order (HOGN)
            If tgChf.sStatus = "W" Or tgChf.sStatus = "D" Or tgChf.sStatus = "C" Or tgChf.sStatus = "I" Then
                tlBR.iPropOrOrder = True
            End If
            If tlBR.sSnapshot = " " Then                    'user is taking a snapshot of a version just viewed
                ilChfRecLen = Len(tmChf)
                tmChfSrchKey1.lCntrNo = tgChf.lCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, ilChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tgChf.lCntrNo) And (tgChf.lCode <> tmChf.lCode) Then
                    tlBR.sSnapshot = "P"            'previous version
                End If
            End If

            If tlBR.iDiffOnly Then                            'differences for Mods
                If tlBR.iShowNTRBillSummary = True Then       'disable showing the NTR billing summary if differences only
                    MsgBox "NTR Billing Summary disabled on Air Time Billing Summary for Difference option"
                    tlBR.iShowNTRBillSummary = False
                    ckcShowNTRBillSummary.Value = vbUnchecked
                End If
                '7/30/19: Contract difference from order/proposal screen
                If tlBR.sSnapshot <> "D" Then
                    ilChfRecLen = Len(tmChf)
                    tmChfSrchKey1.lCntrNo = tgChf.lCntrNo
                    tmChfSrchKey1.iCntRevNo = tgChf.iCntRevNo
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, ilChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tgChf.lCntrNo) And (tmChf.sSchStatus <> "F")
                        ilRet = btrGetNext(hmCHF, tmChf, ilChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tgChf.lCntrNo) And (tmChf.sSchStatus = "F") Then
                        'tmchf has original contract
                        'differences only on this contract required - compare 2 contracts and return differences in the contract
    
                        'Save the current tgchf, tgclf, & tgcff so that when the snapshot returns,
                        'the original image on the screen will be restored since the Differences mode
                        'fakes out a contract for the differences only into the 'tg' arrays
                        tmTChf = tgChf
                        ReDim tmTClf(0 To 0) As CLFLIST
                        For ilLoop = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                            tmTClf(ilLoop) = tgClf(ilLoop)
                            ReDim Preserve tmTClf(0 To UBound(tmTClf) + 1) As CLFLIST
                        Next ilLoop
                        ReDim tmTCff(0 To 0) As CFFLIST
                        For ilLoop = LBound(tgCff) To UBound(tgCff) - 1 Step 1
                            tmTCff(ilLoop) = tgCff(ilLoop)
                            ReDim Preserve tmTCff(0 To UBound(tmTCff) + 1) As CFFLIST
                        Next ilLoop
                        llPrevCode = tmChf.lCode
                        gGenDiff CONTRACTSJOB, hmCHF, hmClf, hmCff, llCurrCode, llPrevCode, llPrevSpots, llPrevGross, ilCurrTotalWks, ilCurrAirWks
                    Else
                        MsgBox "Previous version does not exist", vbOKOnly + vbExclamation, "Traffic"
                        mTerminate
                        Exit Sub
                    End If
                Else
                    '7/30/19: Contract difference from order/proposal screen
                    llCurrCode = lgCurrChfCode
                    llPrevCode = lgPrevChfCode
                    gGenDiff CONTRACTSJOB, hmCHF, hmClf, hmCff, llCurrCode, llPrevCode, llPrevSpots, llPrevGross, ilCurrTotalWks, ilCurrAirWks, tlBR.sSnapshot
                End If
            End If
            'ProcessFlag 0 = open files, then process cnt, 1 = close files and exit
            tlBR.lPrevSpots = llPrevSpots
            tlBR.lPrevGross = llPrevGross
            tlBR.iCurrTotWks = ilCurrTotalWks
            tlBR.iCurrAirWks = ilCurrAirWks

            slDate = Format$(gNow(), "m/d/yy")
            gPackDate slDate, tlBR.iGenDate(0), tlBR.iGenDate(1)
            igNowDate(0) = tlBR.iGenDate(0)
            igNowDate(1) = tlBR.iGenDate(1)
            slTime = Format$(gNow(), "h:mm:ssAM/PM")
            gPackTime slTime, ilGenTime(0), ilGenTime(1)    '10-30-01
            gUnpackTimeLong ilGenTime(0), ilGenTime(1), False, lgNowTime
            
            '9-14-09 no longer using lgNowTime for the time, getting milliseconds in time
            '11-17-11 go back to using nowTime (do not use milliseconds.  Clients are getting blank reports; need to see if this fixes the issue)
            'lgNowTime = timeGetTime

            ilRet = gProcessBR(tlBR, 0, CONTRACTSJOB)             'open files
            If ilRet <> 0 Then
                ilRet = gProcessBR(tlBR, 2, CONTRACTSJOB)       'close files and terminate
                Screen.MousePointer = vbDefault
                mTerminate
                Exit Sub
            End If
            ilRet = gProcessBR(tlBR, 1, CONTRACTSJOB)            'process one contract
            ilRet = gProcessBR(tlBR, 2, CONTRACTSJOB)            'close files
            'Call Report print module
            'If (Not igStdAloneMode) And (imShowHelpMsg) Then
            '    If igTestSystem Then
            '        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode)) & "\" & slStartDate & "\" & Trim$(Str$(tgSel(ilLoop).lEndDate - tgSel(ilLoop).lStartDate + 1)) & "\" & slStartTime & "\" & slEndTime & "\" & Trim$(Str$(tmVef.iCode)) & "\" & Trim$(Str$(tgSel(ilLoop).iZone)) & "\" & slOutput
            '    Else
            '        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode)) & "\" & slStartDate & "\" & Trim$(Str$(tgSel(ilLoop).lEndDate - tgSel(ilLoop).lStartDate + 1)) & "\" & slStartTime & "\" & slEndTime & "\" & Trim$(Str$(tmVef.iCode)) & "\" & Trim$(Str$(tgSel(ilLoop).iZone)) & "\" & slOutput
            '    End If
            'Else
                If igTestSystem Then
                    slStr = "Traffic^Test^NOHELP\" & sgUserName
                Else
                    slStr = "Traffic^Prod^NOHELP\" & sgUserName
                End If
            'End If

            'Set up the command line string to pass to print the contract
            If tlBR.iDetail And tlBR.iSummary Then
                slDetSumBoth = "2"
            ElseIf tlBR.iDetail And Not tlBR.iSummary Then
                slDetSumBoth = "0"              'detail only
            Else
                slDetSumBoth = "1"              'summary only
            End If
            If tlBR.iShowRates Then
                slInclRates = "Y"
            Else
                slInclRates = "N"
            End If
            If tlBR.iShowResearch Then
                slInclResearch = "Y"
            Else
                slInclResearch = "N"
            End If
            If tlBR.iShowProof Then
                slInclProof = "Y"
            Else
                slInclProof = "N"
            End If

            If tlBR.iShowSplits Then     '2-14-04  show slsp splits on summary
                slInclSplits = "Y"
            Else
                slInclSplits = "N"
            End If
            
            If tlBR.iShowNTRBillSummary Then     '2-2-10 show ntr bill summary with air time
                slInclNTRBillSummary = "Y"
            Else
                slInclNTRBillSummary = "N"
            End If
            
            If tlBR.iShowNetAmtOnProps Then     '2-3-10 show net amt on proposal?
                slShowNetOnProps = "Y"
            Else
                slShowNetOnProps = "N"
            End If
            
            If tlBR.iShowProdProt Then     '8-25-15 Show prod protection categories
                slShowProdProt = "Y"
            Else
                slShowProdProt = "N"
            End If
    
            slSaveToFileName = ""
            If rbcOutput(0).Value = True Then
                slOutputTo = "0"
            ElseIf rbcOutput(1).Value = True Then
                slOutputTo = "1"
            Else
                slOutputTo = "2"
                slExportIndex = Trim$(str(imFTSelectedIndex))
                'ilRet = gExportCRW(edcFileName, imFTSelectedIndex)   '10-20-01
                slSaveToFileName = Trim$(edcFileName)      'save to location
            End If
            'parm 1:  Application name^Test or Prod
            'parm 2:  user name
            'parm 3:  Generation date
            'parm 4:  Generation time
            'parm 5:  Detail(0), Summary(1), Both (2)
            'parm 6: Includes Rates Y/N
            'Parm 7: include Research Y/n
            'Parm 7A: include sls splits (y/n) 2-14-04
            'Parm 7B: Include NTR bill summary with air time summary 2-2-10
            'Parm 7C: Show Net Amts on Proposals  2-3-10
            'Parm 7D: show prod prot codes 8-25-15
            'Parm 8: include Proof
            'Parm 9: Display(0) or Print(1) (always 1)
            'Parm 10: File export index
            'Parm 11: filename for export
    
            slStr = slStr & "\" & slDate & "\" & slTime & "\" & slDetSumBoth & "\" & slInclRates & "\" & slInclResearch & "\" & slInclSplits & "\" & slInclNTRBillSummary & "\" & slShowNetOnProps & "\" & slShowProdProt & "\" & slInclProof & "\" & slOutputTo & "\" & slExportIndex & "\|" & slSaveToFileName
            'lgShellRet = Shell(sgExePath & "RptSelBR.Exe " & slStr, 1)
            'While GetModuleUsage(lgShellRet) > 0
            '    ilRet = DoEvents()
            'Wend
            'gShellAndWait BrSnap, sgExePath & "RptSelBR.exe " & slStr, vbNormalFocus
            sgCommandStr = slStr
            
            'TTP 10549 - Learfield Cloud printing 911, Crystal Crashes, Use Adobe
            sgReportFilename = "BR-" & Contract.edcCntrNo.Text
            If sgReportTempFolder = "" Then
                Set olWshShell = CreateObject("Wscript.Shell")
                sgReportTempFolder = olWshShell.ExpandEnvironmentStrings("%localappdata%\Temp")
                If right(sgReportTempFolder, 1) <> "\" Then sgReportTempFolder = sgReportTempFolder & "\"
            End If
            'TTP 10745 - NTR: add option to only show vehicle, billing date, and description on the contract report, and vehicle and description only on invoice reprint
            bgSuppressNTRDetails = False
            If ckcSuppressNTRDetails.Value = vbChecked Then
                bgSuppressNTRDetails = True
            End If
            
            RptSelBR.Show vbModal
            'Restore the original contents of the screen
            If tlBR.iDiffOnly And tlBR.sSnapshot <> "D" Then
                tgChf = tmTChf
                ReDim tgClf(0 To 0) As CLFLIST
                For ilLoop = LBound(tmTClf) To UBound(tmTClf) - 1 Step 1
                    tgClf(ilLoop) = tmTClf(ilLoop)
                    ReDim Preserve tgClf(0 To UBound(tgClf) + 1) As CLFLIST
                Next ilLoop
                ReDim tgCff(0 To 0) As CFFLIST
                For ilLoop = LBound(tmTCff) To UBound(tmTCff) - 1 Step 1
                    tgCff(ilLoop) = tmTCff(ilLoop)
                    ReDim Preserve tgCff(0 To UBound(tgCff) + 1) As CFFLIST
                Next ilLoop
            End If
        End If
    End If              'rbcRptType
    Screen.MousePointer = vbDefault
    mTerminate
End Sub

Private Sub cmcSetup_Click()
    cdcSetup.flags = cdlPDPrintSetup
    cdcSetup.ShowPrinter
End Sub
Private Sub edcCopies_GotFocus()
    gCtrlGotFocus edcCopies
    mSetCommands
End Sub

Private Sub edcFileName_Change()
    mSetCommands
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
    BrSnap.Refresh
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
        'plcBR.Visible = False
        frcSnapBR.Visible = False
        'plcBR.Visible = True
        frcSnapBR.Visible = True
    End If

End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    If imShowDiff = 2 Then
    '    ilRet = btrClose(hmCbf)
    '    btrDestroy hmCbf
    Else
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
    End If
    Erase tmTClf
    Erase tmTCff
    
    Set BrSnap = Nothing   'Remove data segment
    
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imFirstFocus = True
    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    If (imShowDiff = 0) Or (imShowDiff = 1) Then
        'plcBR.Visible = True
        frcSnapBR.Visible = True
        'plcSum.Visible = False     10-31-03
        '10-30-03 use caption of form to show description
        'If tgChf.lCntrNo <> 0 Then
            'plcScreen.Print smPaintCaption
        'Else
            'plcScreen.Print smPaintCaption
        'End If
        If imShowDiff = 0 Then
            ckcDiff.Visible = False   '10-30-01 False
        '7/30/19: Contract difference from order/proposal screen
        ElseIf imShowDiff = 1 And smSnapshot = "D" Then
            ckcDiff.Value = vbChecked
            ckcDiff.Enabled = False
            plcRptType.Enabled = False
            ckcProof.Value = vbChecked
            ckcProof.Enabled = False
        End If
        hmCHF = CBtrvTable(ONEHANDLE)  '(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", BrSnap
        On Error GoTo 0
        imCHFRecLen = Len(tmChf)
        hmClf = CBtrvTable(ONEHANDLE)  '(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", BrSnap
        On Error GoTo 0
        imClfRecLen = Len(tmClf)
        hmCff = CBtrvTable(ONEHANDLE)  '(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", BrSnap
        On Error GoTo 0
        imCffRecLen = Len(tmCff)
        'In BRSnap call values moved to tgChf, tgClf and tgCff
        'ilRet = csiGetRec("CHF", 0, VarPtr(tgChf), LenB(tgChf))
        'ilRet = csiGetAlloc("CLFLIST", ilStartIndex, ilEndIndex)
        'ReDim tgClf(ilStartIndex To ilEndIndex + 1) As CLFLIST
        'For ilLoop = ilStartIndex To ilEndIndex Step 1
        '    ilRet = csiGetRec("CLFLIST", ilLoop, VarPtr(tgClf(ilLoop)), LenB(tgClf(ilLoop)))
        'Next ilLoop
        'ilRet = csiGetAlloc("CFFLIST", ilStartIndex, ilEndIndex)
        'ReDim tgCff(ilStartIndex To ilEndIndex + 1) As CFFLIST
        'For ilLoop = ilStartIndex To ilEndIndex Step 1
        '    ilRet = csiGetRec("CFFLIST", ilLoop, VarPtr(tgCff(ilLoop)), LenB(tgCff(ilLoop)))
        'Next ilLoop
    Else
        'plcBR.Visible = False
        'plcSum.Visible = True
        'If tgChf.lCntrNo <> 0 Then
        '    plcScreen.Caption = "Demo Bar Snapshot for" & Str$(tgChf.lCntrNo)
        'Else
        '    plcScreen.Caption = "Demo Bar Snapshot"
        'End If
        'hmCbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
        'ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo mInitErr
        'gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", Contract
        'On Error GoTo 0
        'imCbfRecLen = Len(tmCbf)
    End If
    ilRet = gSocEcoPop(RptSelBR, cbcSet1)
     ilRet = gPopExportTypes(cbcFileType)       '10-20-01

    'If Not ((Asc(tgSpf.sOptionFields) And &H80) = &H80) Then '8-31-06 not Using Research turn off research check mark
     If Not ((Asc(tgSpf.sOptionFields) And OFRESEARCH) = OFRESEARCH) Then    'turn off research and disable if not using research
        ckcResearch.Value = vbUnchecked
        ckcResearch.Enabled = False
    End If
    
    ckcProdcastInfo.Value = vbUnchecked         '4-18-18 take the default from site, allow user to change
    If ((Asc(tgSaf(0).sFeatures4) And ACT1CODES) = ACT1CODES) Then
         ckcProdcastInfo.Value = vbChecked
    End If
    'if error when populating socio-economic stuff, dont abort
    
    'TTP 10382 - Contract report: Option To not show Act1 codes on PDF
    If ((Asc(tgSaf(0).sFeatures4) And ACT1CODES) = ACT1CODES) Then
        ckcShowACT1.Value = vbChecked
    Else
        ckcShowACT1.Value = vbUnchecked
        ckcShowACT1.Enabled = False
    End If
    

    'BrSnap.Height = cmcGen.Top + 5 * cmcGen.Height / 3
    gCenterStdAlone BrSnap
    'BrSnap.Show
    Screen.MousePointer = vbDefault
    
    'TTP 10549 - Learfield Cloud printing 911, Crystal Crashes, Use Adobe
    If bgUseAdobe Then
        rbcOutput(1).Enabled = False 'Print
        'rbcOutput(2).Enabled = False 'Save to File
        rbcOutput(0).Enabled = True
        rbcOutput(0).Value = True
    End If
    
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    slCommand = sgCommandStr    'Command$
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
    '    ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
    '    If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
    '        imShowHelpMsg = False
    '    End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone BrSnap, slStr, ilTestSystem
    'If igStdAloneMode Then
    '    igCmmCallSource = CALLNONE
    '    imShowDiff = 0
    '    smSnapShot = "C"
    '    sgDBPath = "c:\csi_data\data\"
    '    hmChf = CBtrvTable(ONEHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    '    ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    '    On Error GoTo mParseErr
    '    gBtrvErrorMsg ilRet, "mParse (btrOpen: Chf.Btr)", BrSnap
    '    On Error GoTo 0
    '    imChfRecLen = Len(tmChf)
   '
   '     ReDim tgClf(0 To 0) As CLFLIST
   '     tgClf(0).iStatus = -1 'Not Used
   '     tgClf(0).lRecPos = 0
   '     tgClf(0).iFirstCff = -1
   '     ReDim tgCff(0 To 0) As CFFLIST
   '     tgCff(0).iStatus = -1 'Not Used
   '     tgCff(0).lRecPos = 0
   '     tgCff(0).iNextCff = -1
   '
   '     hmClf = CBtrvTable(ONEHANDLE)  '(ONEHANDLE) 'CBtrvObj()
   '     ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
   '     On Error GoTo mParseErr
   '     gBtrvErrorMsg ilRet, "mParse (btrOpen: Clf.Btr)", BrSnap
   '     On Error GoTo 0
   '     imClfRecLen = Len(tmClf)
   '
   '     hmCff = CBtrvTable(ONEHANDLE)  '(ONEHANDLE) 'CBtrvObj()
   '     ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
   '     On Error GoTo mParseErr
   '     gBtrvErrorMsg ilRet, "mParse (btrOpen: Cff.Btr)", BrSnap
   '     On Error GoTo 0
   '     imCffRecLen = Len(tmCff)
   '     '******* DEBUG change the contract code when debugging
   '     llContrCode = 119
   '     ilRet = gObtainCntr(hmChf, hmClf, hmCff, llContrCode, False, tgChf, tgClf(), tgCff())
   '
   '     ilRet = btrClose(hmChf)
   '     btrDestroy hmChf
   '     ilRet = btrClose(hmClf)
   '     btrDestroy hmClf
   '     ilRet = btrClose(hmCff)
   '     btrDestroy hmCff
   '     'put record into DLL for retrieval
   '     ilRet = csiSetRec("CHF", 0, VarPtr(tgChf), LenB(tgChf))
   '     ilStartIndex = LBound(tgClf)
   '     ilEndIndex = UBound(tgClf) - 1
   '     ilRet = csiSetAlloc("CLFLIST", ilStartIndex, ilEndIndex)
   '     For ilLoop = LBound(tgClf) To UBound(tgClf) - 1 Step 1
   '         ilRet = csiSetRec("CLFLIST", ilLoop, VarPtr(tgClf(ilLoop)), LenB(tgClf(ilLoop)))
   '     Next ilLoop
   '     ilStartIndex = LBound(tgCff)
   '     ilEndIndex = UBound(tgCff) - 1
   '     ilRet = csiSetAlloc("CFFLIST", ilStartIndex, ilEndIndex)
   '     For ilLoop = LBound(tgCff) To UBound(tgCff) - 1 Step 1
   '         ilRet = csiSetRec("CFFLIST", ilLoop, VarPtr(tgCff(ilLoop)), LenB(tgCff(ilLoop)))
   '     Next ilLoop
   '
   ' End If
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igCmmCallSource = Val(slStr)
    If igCmmCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)    'Get call source
        imShowDiff = Val(slStr)
        ilRet = gParseItem(slCommand, 5, "\", slStr)    'Get call source
        smSnapshot = Trim$(slStr)
    End If
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'
    Dim ilRet As Integer
    Screen.MousePointer = vbDefault
    igManUnload = YES
    'Unload Traffic
    Unload BrSnap
    igManUnload = NO
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub plcRptType_Paint()
    plcRptType.CurrentX = 0
    plcRptType.CurrentY = 0
    plcRptType.Print "Type"
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

Private Sub rbcRptType_Click(Index As Integer)
    If Index = 0 Then     'proposals/contract
        If imShowDiff = 0 Then
            ckcDiff.Visible = False   '10-30-01 False
        '7/30/19: Contract difference from order/proposal screen
        ElseIf imShowDiff = 1 And smSnapshot = "D" Then
            ckcDiff.Value = vbChecked
            ckcDiff.Enabled = False
            plcRptType.Enabled = False
            ckcProof.Value = vbChecked
            ckcProof.Enabled = False
        End If
        plcShow.Visible = True
        plcDemo.Visible = True
        lacText.Visible = True
        cbcSet1.Visible = True
        plcInclude.Visible = True
        ckcRating.Visible = True
        ckcResearch.Visible = True
        ckcProof.Visible = True
        ckcShowSplit.Visible = True
        If imShowDiff <> 1 Then
            ckcDiff.Visible = False
        End If
        ckcShowNTRBillSummary.Visible = True
        ckcShowNetOnProps.Visible = True
        ckcShowProdProt.Visible = True
        ckcProdcastInfo.Visible = True
        ckcShowACT1.Visible = True 'TTP 8410, hide Suppress NTR details with Order Audit
        ckcSuppressNTRDetails.Visible = True
    Else
        plcShow.Visible = False
        plcDemo.Visible = False
        lacText.Visible = False
        cbcSet1.Visible = False
        plcInclude.Visible = False
        ckcRating.Visible = False
        ckcResearch.Visible = False
        ckcProof.Visible = False
        ckcShowSplit.Visible = False
        ckcDiff.Visible = False
        ckcShowNTRBillSummary.Visible = False
        ckcShowNetOnProps.Visible = False
        ckcShowProdProt.Visible = False
        ckcProdcastInfo.Visible = False
        ckcShowACT1.Visible = False 'TTP 8410, hide Suppress NTR details with Order Audit
        ckcSuppressNTRDetails.Visible = False
    End If
End Sub

Private Sub rbcShow_GotFocus(Index As Integer)
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
    End If
End Sub

Private Sub plcInclude_Paint()
    plcInclude.CurrentX = 0
    plcInclude.CurrentY = 0
    plcInclude.Print "Include"
End Sub

Private Sub plcShow_Paint()
    plcShow.CurrentX = 0
    plcShow.CurrentY = 0
    plcShow.Print "Show"
End Sub

Private Sub plcDemo_Paint()
    plcDemo.CurrentX = 0
    plcDemo.CurrentY = 0
    plcDemo.Print "Demos"
End Sub

Public Sub mSetCommands()
    Dim ilEnable As Integer
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
    cmcGen.Enabled = ilEnable
End Sub
