VERSION 5.00
Begin VB.Form StnFdUnd 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5385
   ClientLeft      =   150
   ClientTop       =   1680
   ClientWidth     =   9420
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
   ScaleHeight     =   5385
   ScaleWidth      =   9420
   Begin VB.CommandButton cmcResendDate 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5940
      Picture         =   "Stnfdund.frx":0000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   75
      Width           =   195
   End
   Begin VB.TextBox edcResendDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   5010
      MaxLength       =   10
      TabIndex        =   2
      Top             =   75
      Width           =   930
   End
   Begin VB.TextBox edcRunLetter 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   8745
      MaxLength       =   1
      TabIndex        =   5
      Top             =   90
      Width           =   345
   End
   Begin VB.Timer tmcResend 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   8805
      Top             =   4320
   End
   Begin VB.PictureBox plcCalendar 
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
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   5010
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   20
      Top             =   300
      Visible         =   0   'False
      Width           =   1995
      Begin VB.CommandButton cmcCalUp 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalDn 
         Appearance      =   0  'Flat
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Stnfdund.frx":00FA
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   22
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
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
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   25
         Top             =   30
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcRotInfo 
      BackColor       =   &H00FFFF80&
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
      Height          =   1230
      Left            =   1095
      ScaleHeight     =   1170
      ScaleWidth      =   6465
      TabIndex        =   13
      Top             =   1665
      Visible         =   0   'False
      Width           =   6525
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Rotation: #  by user name"
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
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   14
         Top             =   45
         Width           =   6330
      End
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Bulk Feed: Send on xx/xx/xx"
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
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   18
         Top             =   945
         Width           =   6315
      End
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Last Assignment Done: Date xx/xx/xx   Time xx:xx:xxam"
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
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   17
         Top             =   720
         Width           =   6315
      End
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Date Range Assigned To:  Earliest xx/xx/xx   Latest xx/xx/xx"
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
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   16
         Top             =   495
         Width           =   6330
      End
      Begin VB.Label lacRotInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Entered Date xx/xx/xx    Version Date xx/xx/xx   Modified xx Times"
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
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   15
         Top             =   270
         Width           =   6330
      End
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5595
      TabIndex        =   10
      Top             =   4980
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      Top             =   4980
      Width           =   945
   End
   Begin VB.PictureBox plcBulkFeed 
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
      Height          =   3720
      Left            =   45
      ScaleHeight     =   3720
      ScaleWidth      =   9330
      TabIndex        =   6
      Top             =   420
      Width           =   9330
      Begin VB.PictureBox pbcLbcRot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3570
         Left            =   75
         ScaleHeight     =   3570
         ScaleWidth      =   8940
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   90
         Width           =   8940
      End
      Begin VB.ListBox lbcRot 
         Appearance      =   0  'Flat
         Height          =   3600
         ItemData        =   "Stnfdund.frx":2F14
         Left            =   60
         List            =   "Stnfdund.frx":2F16
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   75
         Width           =   8970
      End
      Begin VB.VScrollBar vbcRot 
         Height          =   3615
         LargeChange     =   16
         Left            =   9030
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   60
         Width           =   270
      End
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
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   11
      Top             =   1770
      Width           =   75
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      ScaleHeight     =   240
      ScaleWidth      =   2985
      TabIndex        =   0
      Top             =   -15
      Width           =   2985
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2835
      TabIndex        =   8
      Top             =   4980
      Width           =   945
   End
   Begin VB.CheckBox ckcResendCart 
      Caption         =   "Set Carts Associated with Rotation as Unsent"
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
      Height          =   225
      Left            =   75
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4215
      Width           =   4200
   End
   Begin VB.Label lacResend 
      Appearance      =   0  'Flat
      Caption         =   "Resend Start Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3405
      TabIndex        =   1
      Top             =   60
      Width           =   1560
   End
   Begin VB.Label lacRunLetter 
      Appearance      =   0  'Flat
      Caption         =   "Run Letter (A or B or Blank)"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6405
      TabIndex        =   4
      Top             =   60
      Width           =   2250
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   4860
      Width           =   360
   End
End
Attribute VB_Name = "StnFdUnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Stnfdund.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: StnFdUnd.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Bulk Feed Status screen code
Option Explicit
Option Compare Text
Dim imListFieldRot(1 To 11) As Integer
Dim smScreenCaption As String
'Contract header
Dim tmChfSrchKey As LONGKEY0  'CHF key record image
Dim hmCHF As Integer        'CHF Handle
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChf As CHF
'Rotation header
Dim tmCrfSrchKey1 As CRFKEY1  'CRF key record image
Dim hmCrf As Integer        'CRF Handle
Dim imCrfRecLen As Integer      'CRF record length
Dim tmCrf As CRF
'Short Title
Dim tmSif As SIF            'SIF record image
Dim hmSif As Integer        'SIF Handle
Dim imSifRecLen As Integer      'SIF record length
'Short Title via Contract
Dim tmVsf As VSF            'VSF record image
Dim tmVsfSrchKey As LONGKEY0  'VSF key record image
Dim hmVsf As Integer        'VSF Handle
Dim imVsfRecLen As Integer      'VSF record length
'Instruction
Dim tmCnf As CNF            'CNF record image
Dim tmCnfSrchKey As CNFKEY0  'CNF key record image
Dim hmCnf As Integer        'CNF Handle
Dim imCnfRecLen As Integer      'CNF record length
'Inventory
Dim tmCif As CIF            'CIF record image
Dim hmCif As Integer        'CIF Handle
Dim imCifRecLen As Integer      'CIF record length
'Vehicle
Dim hmVef As Integer
Dim tmVef() As VEF
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer  'VEF record length
'Region Code
Dim tmRaf As RAF            'RAF record image
Dim tmRafSrchKey As LONGKEY0  'MCF key record image
Dim hmRaf As Integer        'MCF Handle
Dim imRafRecLen As Integer      'MCF record length
'Copy Feed
Dim tmCyf As CYF            'CYF record image
Dim tmCyfSrchKey As CYFKEY0  'CYF key record image
Dim hmCyf As Integer        'CYF Handle
Dim imCyfRecLen As Integer      'CYF record length
'Advertiser name
Dim hmAdf As Integer
Dim tmAdf As ADF
Dim tmAdfSrchKey As INTKEY0 'ADF key record image
Dim imAdfRecLen As Integer  'ADF record length
'Avail name
Dim hmAnf As Integer
Dim tmAnf As ADF
Dim tmAnfSrchKey As INTKEY0 'ANF key record image
Dim imAnfRecLen As Integer  'ANF record length
Dim tmSortCrf() As SORTCRF
'Dim tmRec As LPOPREC
Dim imTypeIndex As Integer
Dim imFirstActivate As Integer
Dim imTerminate As Integer
Dim imLastIndex As Integer
Dim imCurrentIndex As Integer
Dim imShiftKey As Integer   'Bit 0=Shift; 1=Ctrl; 2=Alt
Dim imIgnoreVbcChg As Integer
Dim imButton As Integer
Dim imButtonIndex As Integer
Dim imIgnoreRightMove As Integer
Dim lmNowDate As Long
Dim smNowDate As String
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim smRunLetter As String
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcResendDate.SelStart = 0
    edcResendDate.SelLength = Len(edcResendDate.Text)
    edcResendDate.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcResendDate.SelStart = 0
    edcResendDate.SelLength = Len(edcResendDate.Text)
    edcResendDate.SetFocus
End Sub
Private Sub cmcCancel_Click()
    If igBFReturn = -1 Then
        igBFReturn = 0
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcDone_Click()
    Dim slMess As String
    Dim ilRes As Integer
    If cmcUpdate.Enabled Then
        slMess = "Save Changes"
        ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
        If ilRes = vbCancel Then
            Exit Sub
        End If
        If ilRes = vbYes Then
            mSaveRec
            igBFReturn = 1
        End If
    End If
    If igBFReturn = -1 Then
        igBFReturn = 0  'Cancelled
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcResendDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcResendDate.SelStart = 0
    edcResendDate.SelLength = Len(edcResendDate.Text)
    edcResendDate.SetFocus
End Sub
Private Sub cmcResendDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUpdate_Click()
    mSaveRec
    Screen.MousePointer = vbHourglass
    igBFReturn = 1
    imLastIndex = -1
    imShiftKey = 0
    mResendRotPop
    Screen.MousePointer = vbDefault
    mSetCommands
End Sub
Private Sub cmcUpdate_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub edcResendDate_Change()
    Dim slStr As String
    tmcResend.Enabled = False
    ReDim tmSortCrf(0 To 0) As SORTCRF
    lbcRot.Clear
    slStr = edcResendDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    tmcResend.Enabled = True
End Sub
Private Sub edcResendDate_GotFocus()
    'tmcResend.Enabled = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcResendDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcResendDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcResendDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcResendDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcResendDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcResendDate.Text = slDate
            End If
        End If
        edcResendDate.SelStart = 0
        edcResendDate.SelLength = Len(edcResendDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcResendDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcResendDate.Text = slDate
            End If
        End If
        edcResendDate.SelStart = 0
        edcResendDate.SelLength = Len(edcResendDate.Text)
    End If
End Sub

Private Sub edcRunLetter_Change()
    edcResendDate_Change
End Sub

Private Sub edcRunLetter_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcRunLetter_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
'    gShowBranner
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
        plcCalendar.Visible = False
        gFunctionKeyBranch KeyCode
    End If
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
    
    Erase tmSortCrf
    Erase tmVef
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmAnf)
    btrDestroy hmAnf
    ilRet = btrClose(hmRaf)
    btrDestroy hmRaf
    ilRet = btrClose(hmCyf)
    btrDestroy hmCyf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmCnf)
    btrDestroy hmCnf
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    
    Set StnFdUnd = Nothing   'Remove data segment

End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcRot_Click()
    Dim ilStartIndex As Integer
    Dim ilEndIndex As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilValue As Integer
    Dim llRg As Long
    Dim llRet As Long
    Dim ilListIndex As Integer
    Dim ilResetLast As Integer
    If imIgnoreVbcChg Then
        Exit Sub
    End If
    imIgnoreVbcChg = True
    Screen.MousePointer = vbHourglass
    ilResetLast = True
    ilListIndex = imCurrentIndex + vbcRot.Value
    ilStartIndex = vbcRot.Value
    ilEndIndex = ilStartIndex + vbcRot.LargeChange
    If ilEndIndex > UBound(tmSortCrf) - 1 Then
        ilEndIndex = UBound(tmSortCrf) - 1
    End If
    If ((imShiftKey And 1) = 1) And (imLastIndex >= 0) Then
        For ilLoop = 0 To UBound(tmSortCrf) - 1 Step 1
            tmSortCrf(ilLoop).iSelected = False
        Next ilLoop
        If imLastIndex <= ilListIndex Then
            For ilLoop = imLastIndex To ilListIndex Step 1
                tmSortCrf(ilLoop).iSelected = True
            Next ilLoop
        Else
            For ilLoop = ilListIndex To imLastIndex Step 1
                tmSortCrf(ilLoop).iSelected = True
            Next ilLoop
        End If
        ilValue = False
        If UBound(tmSortCrf) < vbcRot.LargeChange + 1 Then
            llRg = CLng(UBound(tmSortCrf) - 1) * &H10000 Or 0
        Else
            llRg = CLng(vbcRot.LargeChange) * &H10000 Or 0
        End If
        llRet = SendMessageByNum(lbcRot.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        ilIndex = 0
        For ilLoop = ilStartIndex To ilEndIndex Step 1
            lbcRot.Selected(ilIndex) = tmSortCrf(ilLoop).iSelected
            ilIndex = ilIndex + 1
        Next ilLoop
        ilResetLast = False
    ElseIf ((imShiftKey And 2) = 2) Then    'Ctrl
        ilIndex = 0
        For ilLoop = ilStartIndex To ilEndIndex Step 1
            tmSortCrf(ilLoop).iSelected = lbcRot.Selected(ilIndex)
            ilIndex = ilIndex + 1
        Next ilLoop
    Else
        For ilLoop = 0 To UBound(tmSortCrf) - 1 Step 1
            tmSortCrf(ilLoop).iSelected = False
        Next ilLoop
        ilIndex = 0
        For ilLoop = ilStartIndex To ilEndIndex Step 1
            tmSortCrf(ilLoop).iSelected = lbcRot.Selected(ilIndex)
            ilIndex = ilIndex + 1
        Next ilLoop
    End If
    If ilResetLast Then
        imLastIndex = ilListIndex
    End If
    pbcLbcRot_Paint
    Screen.MousePointer = vbDefault
    mSetCommands
    imIgnoreVbcChg = False
End Sub
Private Sub lbcRot_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub lbcRot_KeyDown(KeyCode As Integer, Shift As Integer)
    imShiftKey = Shift
End Sub
Private Sub lbcRot_KeyUp(KeyCode As Integer, Shift As Integer)
    imShiftKey = Shift
End Sub
Private Sub lbcRot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilHeight As Integer
    ilHeight = lbcRot.Height \ (vbcRot.LargeChange + 1)
    imCurrentIndex = Y \ fgListHtArial825

    imButton = Button
    If Button = 2 Then  'Right Mouse
        imButtonIndex = imCurrentIndex + vbcRot.Value
        If (imButtonIndex >= 0) And (imButtonIndex <= UBound(tmSortCrf) - 1) Then
            mShowRotInfo
        End If
    End If
End Sub
Private Sub lbcRot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imIgnoreRightMove Then
        Exit Sub
    End If
    If Button = 2 Then
        If (Y < 0) Or (Y > lbcRot.Height) Then
            imButtonIndex = 0
            plcRotInfo.Visible = False
            Exit Sub
        End If
        If (X < 0) Or (X > lbcRot.Width) Then
            imButtonIndex = 0
            plcRotInfo.Visible = False
            Exit Sub
        End If
        If imButtonIndex <> (Y \ fgListHtArial825) + vbcRot.Value Then
            imIgnoreRightMove = True
            imButtonIndex = Y \ fgListHtArial825 + vbcRot.Value
            If (imButtonIndex >= 0) And (imButtonIndex <= UBound(tmSortCrf) - 1) Then
                mShowRotInfo
            Else
                plcRotInfo.Visible = False
            End If
            imIgnoreRightMove = False
        End If
    End If
End Sub
Private Sub lbcRot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        imButtonIndex = -1
        plcRotInfo.Visible = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = edcResendDate.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim slStr As String
    imTerminate = False

    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imFirstActivate = True
    igBFReturn = -1
    imTypeIndex = 0
    imLastIndex = -1
    imShiftKey = 0
    imIgnoreVbcChg = False
    imButton = 0
    imIgnoreRightMove = False
    imBSMode = False
    imBypassFocus = False
    tmAdf.iCode = 0
    tmAnf.iCode = 0
    imListFieldRot(1) = 15
    imListFieldRot(2) = 15 * igAlignCharWidth
    imListFieldRot(3) = 23 * igAlignCharWidth
    imListFieldRot(4) = 33 * igAlignCharWidth
    imListFieldRot(5) = 40 * igAlignCharWidth
    imListFieldRot(6) = 52 * igAlignCharWidth
    imListFieldRot(7) = 67 * igAlignCharWidth
    imListFieldRot(8) = 78 * igAlignCharWidth
    imListFieldRot(9) = 82 * igAlignCharWidth
    imListFieldRot(10) = 92 * igAlignCharWidth
    imListFieldRot(11) = 97 * igAlignCharWidth
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    imCalType = 0   'Standard
    mInitBox
    smScreenCaption = "Resend Rotation and Inventory"
    StnFdUnd.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone StnFdUnd
    'StnFdUnd.Show
    DoEvents
    Screen.MousePointer = vbHourglass
    'imcHelp.Picture = Traffic!imcHelp.Picture
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", StnFdUnd
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)     'Get and save CHF record length

    hmVef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vef.Btr)", StnFdUnd
    On Error GoTo 0
    ReDim tmVef(0 To 0) As VEF
    imVefRecLen = Len(tmVef(0))     'Get and save VEF record length

    hmCrf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Crf.Btr)", StnFdUnd
    On Error GoTo 0
    imCrfRecLen = Len(tmCrf)     'Get and save CRF record length

    hmSif = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sif.Btr)", StnFdUnd
    On Error GoTo 0
    imSifRecLen = Len(tmSif)     'Get and save CRF record length

    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf.Btr)", StnFdUnd
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)     'Get and save CRF record length

    hmCnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cnf.Btr)", StnFdUnd
    On Error GoTo 0
    imCnfRecLen = Len(tmCnf)     'Get and save CRF record length

    hmCif = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", StnFdUnd
    On Error GoTo 0
    imCifRecLen = Len(tmCif)     'Get and save CRF record length

    hmAdf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", StnFdUnd
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)

    hmAnf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Anf.Btr)", StnFdUnd
    On Error GoTo 0
    imAnfRecLen = Len(tmAnf)

    hmCyf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmCyf, "", sgDBPath & "Cyf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cyf.Btr)", StnFdUnd
    On Error GoTo 0
    imCyfRecLen = Len(tmCyf)

    hmRaf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf.Btr)", StnFdUnd
    On Error GoTo 0
    imRafRecLen = Len(tmRaf)
    ilRet = gObtainVef()
    mVehPop
    If imTerminate Then
        Exit Sub
    End If
    'mResendRotPop
    'If imTerminate Then
    '    Exit Sub
    'End If
    'plcCalendar.Move plcSelect.Left + plcSelect.Width - plcCalendar.Left - 60, plcSelect.Top + edcResendDate.Top + edcResendDate.Height
    slStr = Format$(gNow(), "m/d/yy")
    slStr = gObtainNextMonday(slStr)
    'edcResendDate.Text = slStr
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    lacDate.Visible = False
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim ilLoop As Integer
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    'plcCalendar.Move plcSelect.Left + plcSelect.Width - plcCalendar.Width - 60, plcSelect.Top + edcResendDate.Top + edcResendDate.Height
    plcCalendar.Move edcResendDate.Left, edcResendDate.Top + edcResendDate.Height
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mResendRotPop                   *
'*                                                     *
'*             Created:8/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain rotation specifications *
'*                                                     *
'*******************************************************
Private Sub mResendRotPop()
'
'   iRet = mResendRotPop
'   Where:
'
    Dim ilRet As Integer    'Return status
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slNameCode As String
    Dim slName As String
    Dim ilOffSet As Integer
    Dim llRevCntrNo As Long
    Dim slRevCntrNo As String
    Dim llRevRotNo As Long
    Dim slRevRotNo As String
    Dim llCntrNo As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilDay As Integer
    Dim ilExtLen As Integer
    Dim slDate As String
    Dim ilUpper As Integer
    Dim ilVehIndex As Integer
    Dim ilVeh As Integer
    Dim llSifCode As Long
    Dim ilVsf As Integer
    Dim llStartDate As Long
    ReDim tmSortCrf(0 To 0) As SORTCRF
    ilUpper = 0
    lbcRot.Clear
    pbcLbcRot_Paint
    slStartDate = edcResendDate.Text
    If Not gValidDate(slStartDate) Then
        Beep
        edcResendDate.SetFocus
        Exit Sub
    End If
    llStartDate = gDateValue(slStartDate)
    smRunLetter = Trim$(edcRunLetter.Text)
    btrExtClear hmCrf   'Clear any previous extend operation
    ilExtLen = Len(tmCrf)  'Extract operation record size
    If imTypeIndex = 0 Then
        tmCrfSrchKey1.sRotType = "A"
        tmCrfSrchKey1.iEtfCode = 0
        tmCrfSrchKey1.iEnfCode = 0
        tmCrfSrchKey1.iAdfCode = 0
        tmCrfSrchKey1.lChfCode = 0
        tmCrfSrchKey1.lFsfCode = 0
        tmCrfSrchKey1.iVefCode = 0
        tmCrfSrchKey1.iRotNo = 32000
        ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmCrf, llNoRec, -1, "UC", "CRF", "") 'Set extract limits (all records)
        If igSGOrKC = 0 Then
            ilOffSet = gFieldOffset("Crf", "CrfAffFdStatus")
        Else
            ilOffSet = gFieldOffset("Crf", "CrfKCFdStatus")
        End If
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "R", 1)
        On Error GoTo mResendRotPopErr
        gBtrvErrorMsg ilRet, "mResendRotPop (btrExtAddLogicConst):" & "Crf.Btr", StnFdUnd
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "M", 1)
        On Error GoTo mResendRotPopErr
        gBtrvErrorMsg ilRet, "mResendRotPop (btrExtAddLogicConst):" & "Crf.Btr", StnFdUnd
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "S", 1)
        On Error GoTo mResendRotPopErr
        gBtrvErrorMsg ilRet, "mResendRotPop (btrExtAddLogicConst):" & "Crf.Btr", StnFdUnd
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, ByVal "R", 1)
        On Error GoTo mResendRotPopErr
        gBtrvErrorMsg ilRet, "mResendRotPop (btrExtAddLogicConst):" & "Crf.Btr", StnFdUnd
        On Error GoTo 0
        ilOffSet = gFieldOffset("Crf", "CrfRotType")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, ByVal "X", 1)
        On Error GoTo mResendRotPopErr
        gBtrvErrorMsg ilRet, "mResendRotPop (btrExtAddLogicConst):" & "Crf.Btr", StnFdUnd
        On Error GoTo 0
        ilOffSet = 0
        ilRet = btrExtAddField(hmCrf, ilOffSet, ilExtLen)  'Extract start/end time, and days
        On Error GoTo mResendRotPopErr
        gBtrvErrorMsg ilRet, "mResendRotPop (btrExtAddField):" & "Crf.Btr", StnFdUnd
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmClf)    'Extract record
        ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mResendRotPopErr
            gBtrvErrorMsg ilRet, "mResendRotPop (btrExtGetNextExt):" & "Clf.Btr", StnFdUnd
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmClf, tlClfExt, ilExtLen, llRecPos)
            If ilRet = BTRV_ERR_REJECT_COUNT Then
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
            End If
            Do While ilRet = BTRV_ERR_NONE
                'gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slDate
                'If gDateValue(slDate) >= lmNowDate Then
                If igSGOrKC = 0 Then
                    gUnpackDate tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1), slDate
                Else
                    gUnpackDate tmCrf.iKCFdDate(0), tmCrf.iKCFdDate(1), slDate
                End If
                'If gDateValue(slDate) = llStartDate Then
                If (gDateValue(slDate) = llStartDate) And ((smRunLetter = "") Or ((igSGOrKC = 0) And (smRunLetter = tmCrf.sAffXMitChar)) Or ((igSGOrKC = 1) And (smRunLetter = tmCrf.sKCXMitChar))) Then
                    If tmChf.lCode <> tmCrf.lChfCode Then
                        tmChfSrchKey.lCode = tmCrf.lChfCode
                        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        On Error GoTo mResendRotPopErr
                        gBtrvErrorMsg ilRet, "mResendRotPop (btrGetEqual):" & "Chf.Btr", StnFdUnd
                        On Error GoTo 0
                    End If
                    'Include PSA/Promo since cart are sent 11/17/03
                    'If (tmChf.sType <> "S") And (tmChf.sType <> "M") Then
                        llRevCntrNo = 99999999 - tmChf.lCntrNo
                        slRevCntrNo = Trim$(str$(llRevCntrNo))
                        Do While Len(slRevCntrNo) < 8
                            slRevCntrNo = "0" & slRevCntrNo
                        Loop
                        'Scan for vehicle
                        For ilVeh = 0 To UBound(tmVef) - 1 Step 1
                            If tmVef(ilVeh).iCode = tmCrf.iVefCode Then
                                ilVehIndex = ilVeh
                                Exit For
                            End If
                        Next ilVeh
                        If tmAdf.iCode <> tmCrf.iAdfCode Then
                            tmAdfSrchKey.iCode = tmCrf.iAdfCode
                            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo mResendRotPopErr
                            gBtrvErrorMsg ilRet, "mResendRotPop (btrGetEqual):" & "Adf.Btr", StnFdUnd
                            On Error GoTo 0
                        End If
                        slName = tmVef(ilVehIndex).sName
                        llRevRotNo = 99999 - tmCrf.iRotNo
                        slRevRotNo = Trim$(str$(llRevRotNo))
                        Do While Len(slRevRotNo) < 6
                            slRevRotNo = "0" & slRevRotNo
                        Loop
                        tmSortCrf(ilUpper).sKey = tmAdf.sName & "|" & slRevCntrNo & "|" & tmVef(ilVehIndex).sName & "|" & slRevRotNo
                        tmSortCrf(ilUpper).lCntrNo = tmChf.lCntrNo
                        llSifCode = 0
                        If tmChf.lVefCode < 0 Then
                            tmVsfSrchKey.lCode = -tmChf.lVefCode
                            ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            Do While ilRet = BTRV_ERR_NONE
                                For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                                    If tmVsf.iFSCode(ilVsf) = tmCrf.iVefCode Then
                                        If tmVsf.lFSComm(ilVsf) > 0 Then
                                            llSifCode = tmVsf.lFSComm(ilVsf)
                                        End If
                                        Exit For
                                    End If
                                Next ilVsf
                                If llSifCode <> 0 Then
                                    Exit Do
                                End If
                                If tmVsf.lLkVsfCode <= 0 Then
                                    Exit Do
                                End If
                                tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
                                ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                        End If
                        tmSortCrf(ilUpper).sCntrProd = gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf) 'tmChf.sProduct
                        tmSortCrf(ilUpper).sType = tmChf.sType
                        tmSortCrf(ilUpper).lCrfRecPos = llRecPos
                        tmSortCrf(ilUpper).iSelected = False
                        tmSortCrf(ilUpper).tCrf = tmCrf
                        ReDim Preserve tmSortCrf(0 To ilUpper + 1) As SORTCRF
                        ilUpper = ilUpper + 1
                    'End If
                End If
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                If ilRet = BTRV_ERR_REJECT_COUNT Then
                    ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                End If
            Loop
        End If
        If ilUpper > 0 Then
            ArraySortTyp fnAV(tmSortCrf(), 0), ilUpper, 0, LenB(tmSortCrf(0)), 0, LenB(tmSortCrf(0).sKey), 0
        End If
        imLastIndex = -1
        imIgnoreVbcChg = True
        vbcRot.Min = 0
        If ilUpper > vbcRot.LargeChange + 1 Then
            vbcRot.Max = ilUpper - vbcRot.LargeChange - 1
        Else
            vbcRot.Max = 0
        End If
        imIgnoreVbcChg = False
        btrExtClear hmCrf   'Clear any previous extend operation
        For ilLoop = 0 To ilUpper - 1 Step 1
            slNameCode = tmSortCrf(ilLoop).sKey
            tmCrf = tmSortCrf(ilLoop).tCrf
            ilRet = gParseItem(slNameCode, 1, "|", slName)
            slStr = slName & "|"
            If ilRet <> CP_MSG_NONE Then
                slName = "Missing"
            End If
            llCntrNo = tmSortCrf(ilLoop).lCntrNo
            If (tmSortCrf(ilLoop).sType = "S") Or (tmSortCrf(ilLoop).sType = "M") Then
                slStr = slStr & Trim$(str$(llCntrNo)) & "*|"
            Else
                slStr = slStr & Trim$(str$(llCntrNo)) & "|"
            End If
            ilRet = gParseItem(slNameCode, 3, "|", slName)
            If ilRet <> CP_MSG_NONE Then
                slName = "Missing"
            End If
            slStr = slStr & Left$(slName, 10) & "|"
            slStr = slStr & Trim$(tmCrf.sZone) & "|"
            gUnpackDate tmCrf.iStartDate(0), tmCrf.iStartDate(1), slStartDate
            slStartDate = Left$(slStartDate, Len(slStartDate) - 3)
            gUnpackDate tmCrf.iEndDate(0), tmCrf.iEndDate(1), slEndDate
            slEndDate = Left$(slEndDate, Len(slEndDate) - 3)
            slStr = slStr & slStartDate & "-" & slEndDate & "|"
            gUnpackTime tmCrf.iStartTime(0), tmCrf.iStartTime(1), "A", "1", slStartTime
            gUnpackTime tmCrf.iEndTime(0), tmCrf.iEndTime(1), "A", "1", slEndTime
            slStr = slStr & slStartTime & "-" & slEndTime & "|"
            For ilDay = 0 To 6 Step 1
                slStr = slStr & tmCrf.sDay(ilDay) '& "|"
            Next ilDay
            slStr = slStr & "|"
            slStr = slStr & Trim$(str$(tmCrf.iLen)) & "|"
            If (tmCrf.sInOut = "I") Or (tmCrf.sInOut = "O") Then
                If tmAnf.iCode <> tmCrf.ianfCode Then
                    tmAnfSrchKey.iCode = tmCrf.ianfCode
                    ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    On Error GoTo mResendRotPopErr
                    gBtrvErrorMsg ilRet, "mResendRotPop (btrGetEqual):" & "Anf.Btr", StnFdUnd
                    On Error GoTo 0
                End If
                If tmCrf.sInOut = "O" Then
                    slName = "O" & Trim$(tmAnf.sName)
                Else
                    slName = Trim$(tmAnf.sName)
                End If
            Else
                slName = "All avails"
            End If
            slStr = slStr & slName & "|"
            Select Case tmCrf.sRotType
                Case "A"
                    slStr = slStr & "CS " & "|"
                Case "O"
                    slStr = slStr & "OBB" & "|"
                Case "C"
                    slStr = slStr & "CBB" & "|"
                Case "E"
                    slStr = slStr & "ABB" & "|"
                Case Else
                    slStr = slStr & " |"
            End Select
            If tmCrf.lCsfCode > 0 Then
                slStr = slStr & "C"
            Else
                slStr = slStr & " "
            End If
            If lbcRot.ListCount < vbcRot.LargeChange + 1 Then
                lbcRot.AddItem slStr
            End If
            tmSortCrf(ilLoop).sKey = slStr
        Next ilLoop
    End If
    pbcLbcRot_Paint
    Exit Sub
mResendRotPopErr:
    On Error GoTo 0
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Save Transmit Record           *
'*                                                     *
'*******************************************************
Private Sub mSaveRec()
    Dim ilRet As Integer
    Dim ilCrf As Integer
    Dim tlVef As VEF
    Dim ilVef As Integer
    Dim ilIndex As Integer
    Dim llAffFdDate As Long
    Dim llKCFdDate As Long
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim llCyfDate As Long

    Screen.MousePointer = vbHourglass
    slStartDate = edcResendDate.Text
    llStartDate = gDateValue(slStartDate)
    ilRet = btrBeginTrans(hmCrf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
        Exit Sub
    End If
    For ilCrf = 0 To UBound(tmSortCrf) - 1 Step 1
        If tmSortCrf(ilCrf).iSelected Then
            Do
                ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tmSortCrf(ilCrf).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    ilRet = btrAbortTrans(hmCrf)
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                    Exit Sub
                End If
                'tmRec = tmCrf
                'ilRet = gGetByKeyForUpdate("CRF", hmCrf, tmRec)
                'tmCrf = tmRec
                'If ilRet <> BTRV_ERR_NONE Then
                '    ilRet = btrAbortTrans(hmCrf)
                '    Screen.MousePointer = vbDefault
                '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                '    Exit Sub
                'End If
                If igSGOrKC = 0 Then
                    tmCrf.sAffFdStatus = "R" 'Ready
                    'If (tmCrf.iAffFdWk And 32) <> 0 Then
                    '2/14/05:Fixed to handle skipped weeks in the sending
                    'Not fixed if they jump back two weeks which they don't do
                    'Exported 6/6, 6/13 and 6/20.  Ask to resend 6/6
                    If (tmCrf.iAffFdWk And &H3F) <> 0 Then
                        gUnpackDateLong tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1), llAffFdDate
                        gPackDateLong llAffFdDate - 7, tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1)
                        tmCrf.iAffFdWk = (tmCrf.iAffFdWk And &H3F) * 2
                    Else
                        tmCrf.iAffFdDate(0) = 0
                        tmCrf.iAffFdDate(1) = 0
                        tmCrf.iAffFdWk = 0
                    End If
                Else
                    tmCrf.sKCFdStatus = "R" 'Ready
                    'If (tmCrf.iAffFdWk And 32) <> 0 Then
                    '2/14/05:Fixed to handle skipped weeks in the sending
                    'Not fixed if they jump back two weeks which they don't do
                    'Exported 6/6, 6/13 and 6/20.  Ask to resend 6/6
                    If (tmCrf.iKCFdWk And &H3F) <> 0 Then
                        gUnpackDateLong tmCrf.iKCFdDate(0), tmCrf.iKCFdDate(1), llKCFdDate
                        gPackDateLong llKCFdDate - 7, tmCrf.iKCFdDate(0), tmCrf.iKCFdDate(1)
                        tmCrf.iKCFdWk = (tmCrf.iKCFdWk And &H3F) * 2
                    Else
                        tmCrf.iKCFdDate(0) = 0
                        tmCrf.iKCFdDate(1) = 0
                        tmCrf.iKCFdWk = 0
                    End If
                End If
                ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                ilRet = btrAbortTrans(hmCrf)
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                Exit Sub
            End If
            'If ckcResendCart.Value = vbChecked Then
                tmVefSrchKey.iCode = tmCrf.iVefCode
                ilRet = btrGetEqual(hmVef, tlVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    'Remove Cyf
                    tmCnfSrchKey.lCrfCode = tmCrf.lCode
                    tmCnfSrchKey.iInstrNo = 0
                    ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
                        If tlVef.sType = "S" Then
                            ilIndex = gVpfFind(StnFdUnd, tlVef.iCode)
                            If ilIndex >= 0 Then
                                'Remove airing
                                If tgVpf(ilIndex).iSAGroupNo > 0 Then
                                    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                        If tgMVef(ilVef).sType = "A" Then
                                            If tgVpf(ilIndex).iSAGroupNo = tgVpf(gVpfFind(StnFdUnd, tgMVef(ilVef).iCode)).iSAGroupNo Then
                                                Do
                                                    tmCyfSrchKey.lCifCode = tmCnf.lCifCode
                                                    tmCyfSrchKey.iVefCode = tgMVef(ilVef).iCode
                                                    If igSGOrKC = 0 Then
                                                        tmCyfSrchKey.sSource = "S"
                                                    Else
                                                        tmCyfSrchKey.sSource = "K"
                                                    End If
                                                    tmCyfSrchKey.sTimeZone = tmCrf.sZone
                                                    tmCyfSrchKey.lRafCode = tmCrf.lRafCode
                                                    ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                                    If ilRet = BTRV_ERR_NONE Then
                                                        'ilRet = btrDelete(hmCyf)
                                                        If igSGOrKC = 0 Then
                                                            gUnpackDateLong tmCyf.iAffOrigXMitDate(0), tmCyf.iAffOrigXMitDate(1), llCyfDate
                                                            If (ckcResendCart.Value = vbChecked) Or ((tmCrf.sAffXMitChar = tmCyf.sAffOrigXMitChar) And (llCyfDate = llStartDate)) Then
                                                                ilRet = btrDelete(hmCyf)
                                                            Else
                                                                ilRet = BTRV_ERR_NONE
                                                            End If
                                                        Else
                                                            gUnpackDateLong tmCyf.iKCOrigXMitDate(0), tmCyf.iKCOrigXMitDate(1), llCyfDate
                                                            If (ckcResendCart.Value = vbChecked) Or ((tmCrf.sKCXMitChar = tmCyf.sKCOrigXMitChar) And (llCyfDate = llStartDate)) Then
                                                                ilRet = btrDelete(hmCyf)
                                                            Else
                                                                ilRet = BTRV_ERR_NONE
                                                            End If
                                                        End If
                                                    Else
                                                        ilRet = BTRV_ERR_NONE
                                                    End If
                                                Loop While ilRet = BTRV_ERR_CONFLICT
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    ilRet = btrAbortTrans(hmCrf)
                                                    Screen.MousePointer = vbDefault
                                                    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    Next ilVef
                                End If
                            End If
                        Else
                            Do
                                tmCyfSrchKey.lCifCode = tmCnf.lCifCode
                                tmCyfSrchKey.iVefCode = tlVef.iCode
                                If igSGOrKC = 0 Then
                                    tmCyfSrchKey.sSource = "S"
                                Else
                                    tmCyfSrchKey.sSource = "K"
                                End If
                                tmCyfSrchKey.sTimeZone = tmCrf.sZone
                                tmCyfSrchKey.lRafCode = tmCrf.lRafCode
                                ilRet = btrGetEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                If ilRet = BTRV_ERR_NONE Then
                                    'ilRet = btrDelete(hmCyf)
                                    If igSGOrKC = 0 Then
                                        gUnpackDateLong tmCyf.iAffOrigXMitDate(0), tmCyf.iAffOrigXMitDate(1), llCyfDate
                                        If (ckcResendCart.Value = vbChecked) Or ((tmCrf.sAffXMitChar = tmCyf.sAffOrigXMitChar) And (llCyfDate = llStartDate)) Then
                                            ilRet = btrDelete(hmCyf)
                                        Else
                                            ilRet = BTRV_ERR_NONE
                                        End If
                                    Else
                                        gUnpackDateLong tmCyf.iKCOrigXMitDate(0), tmCyf.iKCOrigXMitDate(1), llCyfDate
                                        If (ckcResendCart.Value = vbChecked) Or ((tmCrf.sKCXMitChar = tmCyf.sKCOrigXMitChar) And (llCyfDate = llStartDate)) Then
                                            ilRet = btrDelete(hmCyf)
                                        Else
                                            ilRet = BTRV_ERR_NONE
                                        End If
                                    End If
                                Else
                                    ilRet = BTRV_ERR_NONE
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            If ilRet <> BTRV_ERR_NONE Then
                                ilRet = btrAbortTrans(hmCrf)
                                Screen.MousePointer = vbDefault
                                ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                                Exit Sub
                            End If
                        End If
                        ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
            'End If
        End If
    Next ilCrf
    ilRet = btrEndTrans(hmCrf)
    Screen.MousePointer = vbDefault
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
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
    Dim ilCrf As Integer
    Dim ilOneSelected As Integer
    ilOneSelected = False
    For ilCrf = 0 To UBound(tmSortCrf) - 1 Step 1
        If tmSortCrf(ilCrf).iSelected Then
            ilOneSelected = True
            Exit For
        End If
    Next ilCrf
    If ilOneSelected Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mShowRotInfo                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Show Dulpicate and Combined    *
'*                      rotations                      *
'*                                                     *
'*******************************************************
Private Sub mShowRotInfo()
    Dim slDate As String
    Dim slTime As String
    Dim ilShow As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilButtonIndex As Integer
    Dim ilVsf As Integer
    Dim llSifCode As Long
    ilButtonIndex = imButtonIndex
    If (imButtonIndex < LBound(tmSortCrf)) Or (imButtonIndex >= UBound(tmSortCrf)) Then
        imButtonIndex = -1
        plcRotInfo.Visible = False
        Exit Sub
    End If
    For ilLoop = 0 To 4 Step 1
        lacRotInfo(ilLoop).Caption = ""
    Next ilLoop
    ilShow = 0
    ilRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, tmSortCrf(ilButtonIndex).lCrfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        tmChfSrchKey.lCode = tmCrf.lChfCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If tmAdf.iCode <> tmChf.iAdfCode Then
                tmAdfSrchKey.iCode = tmChf.iAdfCode
                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            End If
            llSifCode = 0
            If tmChf.lVefCode < 0 Then
                tmVsfSrchKey.lCode = -tmChf.lVefCode
                ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                Do While ilRet = BTRV_ERR_NONE
                    For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                        If tmVsf.iFSCode(ilVsf) = tmCrf.iVefCode Then
                            If tmVsf.lFSComm(ilVsf) > 0 Then
                                llSifCode = tmVsf.lFSComm(ilVsf)
                            End If
                            Exit For
                        End If
                    Next ilVsf
                    If llSifCode <> 0 Then
                        Exit Do
                    End If
                    If tmVsf.lLkVsfCode <= 0 Then
                        Exit Do
                    End If
                    tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
                    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
            If (tgSpf.sUseProdSptScr <> "P") Then
                lacRotInfo(0).Caption = "Rotation #:" & str$(tmCrf.iRotNo) & " Product: " & gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf)    'Trim$(tmChf.sProduct)
            Else
                lacRotInfo(0).Caption = "Rotation #:" & str$(tmCrf.iRotNo) & " Short Title: " & gGetProdOrShtTitle(hmSif, llSifCode, tmChf, tmAdf)    'Trim$(tmChf.sProduct)
            End If
        Else
            lacRotInfo(0).Caption = "Rotation #:" & str$(tmCrf.iRotNo)
        End If
        gUnpackDate tmCrf.iEntryDate(0), tmCrf.iEntryDate(1), slDate
        lacRotInfo(1).Caption = "Entered Date: " & slDate
        gUnpackDate tmCrf.iVersionDate(0), tmCrf.iVersionDate(1), slDate
        If slDate <> "" Then
            lacRotInfo(1).Caption = lacRotInfo(1).Caption & "  Revised on: " & slDate & "  Modified" & str$(tmCrf.iNoTimesMod) & " Times"
        End If
        gUnpackDate tmCrf.iEarliestDateAssg(0), tmCrf.iEarliestDateAssg(1), slDate
        If slDate = "" Then
            slDate = "-"
        End If
        lacRotInfo(2).Caption = "Date Range Assigned to:  Earliest " & slDate
        gUnpackDate tmCrf.iLatestDateAssg(0), tmCrf.iLatestDateAssg(1), slDate
        If slDate = "" Then
            slDate = "-"
        End If
        lacRotInfo(2).Caption = lacRotInfo(2).Caption & "  Latest " & slDate
        gUnpackDate tmCrf.iDateAssgDone(0), tmCrf.iDateAssgDone(1), slDate
        If slDate = "" Then
            slDate = "-"
            lacRotInfo(3).Caption = "Last Assignment Done:  Date " & slDate & "  Time -"
        Else
            gUnpackTime tmCrf.iTimeAssgDone(0), tmCrf.iTimeAssgDone(1), "A", "1", slTime
            lacRotInfo(3).Caption = "Last assignment Done:  Date " & slDate & "  Time " & slTime
        End If
        'If tmCrf.sAffFdStatus = "R" Then
        '    lacRotInfo(4).Caption = "Affiliate Feed:  Ready to Send"
        'ElseIf tmCrf.sAffFdStatus = "P" Then
        '    lacRotInfo(4).Caption = "Affiliate Feed:  Suppress "
        'ElseIf (tmCrf.sAffFdStatus = "S") Or (tmCrf.sAffFdStatus = "X") Then
        '    gUnpackDate tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1), slDate
        '    lacRotInfo(4).Caption = "Station Feed:  Sent " & slDate
        'Else
        'End If
        If (Left$(tmCrf.sZone, 1) = "R") And (tmCrf.lRafCode > 0) Then
            tmRafSrchKey.lCode = tmCrf.lRafCode
            ilRet = btrGetEqual(hmRaf, tmRaf, imRafRecLen, tmRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                tmRaf.sName = ""
            End If
            If igSGOrKC = 0 Then
                If tmCrf.sAffFdStatus = "R" Then
                    lacRotInfo(4).Caption = Trim$(tmRaf.sName) & " Station Feed-StarGuide:  Ready to Send"
                ElseIf tmCrf.sAffFdStatus = "P" Then
                    lacRotInfo(4).Caption = Trim$(tmRaf.sName) & " Station Feed-StarGuide:  Suppress "
                ElseIf (tmCrf.sAffFdStatus = "S") Or (tmCrf.sAffFdStatus = "X") Then
                    gUnpackDate tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1), slDate
                    lacRotInfo(4).Caption = Trim$(tmRaf.sName) & " Station Feed-StarGuide: Sent " & slDate
                Else
                    lacRotInfo(4).Caption = Trim$(tmRaf.sName)
                End If
            Else
                If tmCrf.sKCFdStatus = "R" Then
                    lacRotInfo(4).Caption = Trim$(tmRaf.sName) & " Station Feed-KenCast:  Ready to Send"
                ElseIf tmCrf.sKCFdStatus = "P" Then
                    lacRotInfo(4).Caption = Trim$(tmRaf.sName) & " Station Feed-KenCast:  Suppress "
                ElseIf (tmCrf.sKCFdStatus = "S") Or (tmCrf.sKCFdStatus = "X") Then
                    gUnpackDate tmCrf.iKCFdDate(0), tmCrf.iKCFdDate(1), slDate
                    lacRotInfo(4).Caption = Trim$(tmRaf.sName) & " Station Feed-KenCast: Sent " & slDate
                Else
                    lacRotInfo(4).Caption = Trim$(tmRaf.sName)
                End If
            End If
        Else
            If igSGOrKC = 0 Then
                If (tmCrf.sAffFdStatus = "S") Or (tmCrf.sAffFdStatus = "X") Then
                    gUnpackDate tmCrf.iAffFdDate(0), tmCrf.iAffFdDate(1), slDate
                    lacRotInfo(4).Caption = "Station Feed: Sent " & slDate
                End If
            Else
                If (tmCrf.sKCFdStatus = "S") Or (tmCrf.sKCFdStatus = "X") Then
                    gUnpackDate tmCrf.iKCFdDate(0), tmCrf.iKCFdDate(1), slDate
                    lacRotInfo(4).Caption = "Station Feed-KenCast: Sent " & slDate
                End If
            End If
        End If
    End If
    DoEvents
    If (imButtonIndex < LBound(tmSortCrf)) Or (imButtonIndex >= UBound(tmSortCrf)) Then
        imButtonIndex = -1
        plcRotInfo.Visible = False
        Exit Sub
    End If
    plcRotInfo.Visible = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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

    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload StnFdUnd
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the vehicles to generate *
'*                     bulk feed                       *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
'
'   tmVef will contain the Conventional and selling
'
    Dim ilRet As Integer
    Dim ilUpper As Integer
    ReDim tmVef(0 To 0) As VEF
    ilUpper = 0
    ilRet = btrGetFirst(hmVef, tmVef(ilUpper), imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If (tmVef(ilUpper).sType = "C") Or (tmVef(ilUpper).sType = "S") Then
            ilUpper = ilUpper + 1
            ReDim Preserve tmVef(0 To ilUpper) As VEF
        End If
        ilRet = btrGetNext(hmVef, tmVef(ilUpper), imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
End Sub

Private Sub lbcRot_Scroll()
    pbcLbcRot_Paint
End Sub

Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                edcResendDate.Text = Format$(llDate, "m/d/yy")
                edcResendDate.SelStart = 0
                edcResendDate.SelLength = Len(edcResendDate.Text)
                imBypassFocus = True
                edcResendDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcResendDate.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcLbcRot_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilRotEnd As Integer
    Dim ilField As Integer
    Dim slFields(0 To 10) As String
    Dim llFgColor As Long
    Dim llWidth As Long
    Dim ilCol As Integer
    
    ilRotEnd = lbcRot.TopIndex + lbcRot.Height \ fgListHtArial825
    If ilRotEnd > lbcRot.ListCount Then
        ilRotEnd = lbcRot.ListCount
    End If
    If lbcRot.ListCount <= lbcRot.Height \ fgListHtArial825 Then
        llWidth = lbcRot.Width - 30
    Else
        llWidth = lbcRot.Width - igScrollBarWidth - 30
    End If
    pbcLbcRot.Width = llWidth
    pbcLbcRot.Cls
    llFgColor = pbcLbcRot.ForeColor
    pbcLbcRot.FontBold = True
    For ilLoop = lbcRot.TopIndex To ilRotEnd - 1 Step 1
        pbcLbcRot.ForeColor = llFgColor
        If lbcRot.MultiSelect = 0 Then
            If lbcRot.ListIndex = ilLoop Then
                gPaintArea pbcLbcRot, CSng(0), CSng((ilLoop - lbcRot.TopIndex) * fgListHtArial825), CSng(pbcLbcRot.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcRot.ForeColor = vbWhite
            End If
        Else
            If lbcRot.Selected(ilLoop) Then
                gPaintArea pbcLbcRot, CSng(0), CSng((ilLoop - lbcRot.TopIndex) * fgListHtArial825), CSng(pbcLbcRot.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcRot.ForeColor = vbWhite
            End If
        End If
        slStr = lbcRot.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        For ilCol = UBound(slFields) - 1 To LBound(slFields) Step -1
            slFields(ilCol + 1) = slFields(ilCol)
        Next ilCol
        slFields(0) = ""
        If InStr(1, Trim$(slFields(2)), "*", vbTextCompare) > 0 Then
            pbcLbcRot.FontBold = False
            slFields(2) = Left$(slFields(2), Len(Trim$(slFields(2))) - 1)
        Else
            pbcLbcRot.FontBold = True
        End If
        For ilField = 1 To 10 Step 1
            pbcLbcRot.CurrentX = imListFieldRot(ilField)
            pbcLbcRot.CurrentY = (ilLoop - lbcRot.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilField)
            gAdjShowLen pbcLbcRot, slStr, imListFieldRot(ilField + 1) - imListFieldRot(ilField)
            pbcLbcRot.Print slStr
        Next ilField
        pbcLbcRot.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub tmcResend_Timer()
    tmcResend.Enabled = False
    Screen.MousePointer = vbHourglass
    plcCalendar.Visible = False
    mResendRotPop
    Screen.MousePointer = vbDefault
End Sub
Private Sub vbcRot_Change()
    Dim ilStartIndex As Integer
    Dim ilEndIndex As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    If imIgnoreVbcChg Then
        Exit Sub
    End If
    imIgnoreVbcChg = True
    ilStartIndex = vbcRot.Value
    ilEndIndex = ilStartIndex + vbcRot.LargeChange
    If ilEndIndex > UBound(tmSortCrf) - 1 Then
        ilEndIndex = UBound(tmSortCrf) - 1
    End If
    ilValue = False
    If UBound(tmSortCrf) < vbcRot.LargeChange + 1 Then
        llRg = CLng(UBound(tmSortCrf) - 1) * &H10000 Or 0
    Else
        llRg = CLng(vbcRot.LargeChange) * &H10000 Or 0
    End If
    llRet = SendMessageByNum(lbcRot.hwnd, LB_SELITEMRANGE, ilValue, llRg)
    ilIndex = 0
    For ilLoop = ilStartIndex To ilEndIndex Step 1
        lbcRot.List(ilIndex) = tmSortCrf(ilLoop).sKey
        ilIndex = ilIndex + 1
    Next ilLoop
    ilIndex = 0
    For ilLoop = ilStartIndex To ilEndIndex Step 1
        lbcRot.Selected(ilIndex) = tmSortCrf(ilLoop).iSelected
        ilIndex = ilIndex + 1
    Next ilLoop
    pbcLbcRot_Paint
    imIgnoreVbcChg = False
End Sub
Private Sub vbcRot_Scroll()
    vbcRot_Change
End Sub
Private Sub plcBulkFeed_Paint()
    plcBulkFeed.CurrentX = 0
    plcBulkFeed.CurrentY = 0
    plcBulkFeed.Print " "
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub
