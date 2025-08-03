VERSION 5.00
Begin VB.Form AffCopy 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6060
   ClientLeft      =   285
   ClientTop       =   1290
   ClientWidth     =   8835
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
   ScaleHeight     =   6060
   ScaleWidth      =   8835
   Begin VB.PictureBox plcCalendar 
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
      Height          =   1770
      Left            =   6225
      ScaleHeight     =   1710
      ScaleWidth      =   1935
      TabIndex        =   24
      Top             =   4095
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   30
         Picture         =   "AFFCOPY.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   27
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
            TabIndex        =   28
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
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
         Left            =   30
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
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
         Left            =   1620
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
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
         Left            =   315
         TabIndex        =   29
         Top             =   30
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcAffData 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4395
      Left            =   4530
      ScaleHeight     =   4335
      ScaleWidth      =   4125
      TabIndex        =   4
      Top             =   225
      Width           =   4185
      Begin VB.ListBox lbcVehicle 
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
         Height          =   3345
         Left            =   15
         MultiSelect     =   2  'Extended
         TabIndex        =   5
         Top             =   75
         Width           =   4080
      End
      Begin VB.TextBox edcTranDate 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   7
         Top             =   3630
         Width           =   930
      End
      Begin VB.OptionButton rbcZone 
         Caption         =   "EST"
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
         Left            =   45
         TabIndex        =   9
         Top             =   4080
         Width           =   690
      End
      Begin VB.OptionButton rbcZone 
         Caption         =   "CST"
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
         Left            =   735
         TabIndex        =   10
         Top             =   4080
         Width           =   705
      End
      Begin VB.OptionButton rbcZone 
         Caption         =   "MST"
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
         Left            =   1440
         TabIndex        =   11
         Top             =   4080
         Width           =   720
      End
      Begin VB.OptionButton rbcZone 
         Caption         =   "PST"
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
         Index           =   3
         Left            =   2130
         TabIndex        =   12
         Top             =   4080
         Width           =   690
      End
      Begin VB.OptionButton rbcZone 
         Caption         =   "Other or All"
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
         Index           =   4
         Left            =   2805
         TabIndex        =   13
         Top             =   4080
         Width           =   1305
      End
      Begin VB.CommandButton cmcTranDate 
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
         Left            =   2610
         Picture         =   "AFFCOPY.frx":2E1A
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3645
         Width           =   195
      End
      Begin VB.Label lacTranDate 
         Appearance      =   0  'Flat
         Caption         =   "Transmitted Date"
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
         Left            =   135
         TabIndex        =   6
         Top             =   3615
         Width           =   1560
      End
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "&Undo"
      Height          =   285
      Left            =   5220
      TabIndex        =   18
      Top             =   5685
      Width           =   945
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
      Height          =   285
      Left            =   4140
      TabIndex        =   17
      Top             =   5685
      Width           =   945
   End
   Begin VB.ListBox lbcCifCyfCode 
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
      Left            =   5250
      Sorted          =   -1  'True
      TabIndex        =   22
      Top             =   -15
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ListBox lbcInv 
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
      Left            =   6030
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   -60
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Timer tmcClick 
      Interval        =   2000
      Left            =   4530
      Top             =   4965
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6285
      TabIndex        =   19
      Top             =   5685
      Width           =   945
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   3045
      TabIndex        =   16
      Top             =   5685
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   1965
      TabIndex        =   15
      Top             =   5685
      Width           =   945
   End
   Begin VB.PictureBox plcAffCopy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4395
      Left            =   120
      ScaleHeight     =   4335
      ScaleWidth      =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   4260
      Begin VB.ComboBox cbcAdvt 
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
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   4140
      End
      Begin VB.ListBox lbcCifCyf 
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
         Height          =   3735
         Left            =   75
         MultiSelect     =   1  'Simple
         TabIndex        =   3
         Top             =   495
         Width           =   4140
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
      TabIndex        =   20
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
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   -15
      Width           =   1695
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   885
      TabIndex        =   14
      Top             =   5685
      Width           =   945
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   90
      Top             =   5625
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacInvInfo 
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
      Left            =   105
      TabIndex        =   21
      Top             =   4680
      Width           =   6885
   End
End
Attribute VB_Name = "AffCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of AFFCOPY.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: AffCopy.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Affiliate Copy input screen code
Option Explicit
Option Compare Text
Dim tmInvCode() As SORTCODE
Dim smInvCodeTag As String
Dim tmVehicleCode() As SORTCODE
Dim smVehicleCodeTag As String
'Program library dates Field Areas
Dim tmCif As CIF        'Cif record image
Dim tmCifSrchKey0 As LONGKEY0    'Cif key record image
Dim hmCif As Integer    'Copy inventory file handle
Dim imCifRecLen As Integer        'Cif record length
Dim tmCyf As CYF        'Cyf record image
Dim tmCyfSrchKey As CYFKEY0    'Cyf key record image
Dim hmCyf As Integer    'Copy feed file handle
Dim imCyfRecLen As Integer        'CYF record length
Dim lmCyfRecPos As Long
Dim tmAdf As ADF        'Adf record image
Dim hmAdf As Integer    'Advertiser file handle
Dim imAdfRecLen As Integer        'ADF record length
'Copy product/ISCI file
Dim hmCpf As Integer 'Copy product/ISCI file handle
Dim tmCpf As CPF        'CPF record image
Dim tmCpfSrchKey0 As LONGKEY0    'CPF key record image
Dim imCpfRecLen As Integer        'CPF record length
'Dim tmRec As LPOPREC
Dim imStartMode As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imAdvtIndex As Integer  'Index of selected record
Dim imCifCyfIndex As Integer
Dim imVehIndex As Integer  'Index
Dim imTypeIndex As Integer
Dim imBypassFocus As Integer
Dim imComboBoxIndex As Integer
Dim imSelectDelay As Integer    'True=cbcSelect change mode
'Calendar info
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim LCCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Private Sub cbcAdvt_Change()

    If imStartMode Then
        imStartMode = False
        mCbcAdvtChange
        Exit Sub
    End If
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        If cbcAdvt.Text <> "" Then
            gManLookAhead cbcAdvt, imBSMode, imComboBoxIndex
            tmcClick.Enabled = False
            imSelectDelay = True
            tmcClick.Interval = 2000    '2 seconds
            tmcClick.Enabled = True
        End If
    End If
End Sub
Private Sub cbcAdvt_Click()
    cbcAdvt_Change
End Sub
Private Sub cbcAdvt_GotFocus()
    If imTerminate Then
        Exit Sub
    End If
    plcCalendar.Visible = False
    tmcClick.Enabled = False
    imComboBoxIndex = imAdvtIndex
    gCtrlGotFocus cbcAdvt
    Exit Sub
End Sub
Private Sub cbcAdvt_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcAdvt_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcAdvt.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcTranDate.SelStart = 0
    edcTranDate.SelLength = Len(edcTranDate.Text)
    edcTranDate.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcTranDate.SelStart = 0
    edcTranDate.SelLength = Len(edcTranDate.Text)
    edcTranDate.SetFocus
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
    tmcClick.Enabled = False
End Sub
Private Sub cmcDone_Click()
    Dim slMess As String
    Dim ilRes As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slVehNameCode As String
    Dim slVehName As String
    Dim ilRet As Integer
    If cmcUpdate.Enabled Then
        slNameCode = lbcCifCyfCode.List(imCifCyfIndex)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        slVehNameCode = tmVehicleCode(imVehIndex).sKey 'lbcVehicleCode.List(imVehIndex)
        ilRet = gParseItem(slVehNameCode, 1, "\", slVehName)
        ilRet = gParseItem(slVehName, 3, "|", slVehName)
        If imCifCyfIndex > 0 Then
            slMess = "Save Changes to " & Trim$(slName) & " " & slVehName
        Else
            slMess = "Add " & Trim$(slName) & " " & slVehName
        End If
        ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
        If ilRes = vbCancel Then
            Exit Sub
        End If
        If ilRes = vbYes Then
            If Not mOKName() Then
                Exit Sub
            End If
            mSaveRec
        End If
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    plcCalendar.Visible = False
    tmcClick.Enabled = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Screen.MousePointer = vbHourglass
    If lmCyfRecPos <> 0 Then
        Do
            ilRet = btrGetDirect(hmCyf, tmCyf, imCyfRecLen, lmCyfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
                Exit Sub
            End If
            'tmRec = tmCyf
            'ilRet = gGetByKeyForUpdate("CYF", hmCyf, tmRec)
            'tmCyf = tmRec
            'If ilRet <> BTRV_ERR_NONE Then
            '    Screen.MousePointer = vbDefault
            '    ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
            '    Exit Sub
            'End If
            ilRet = btrDelete(hmCyf)
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Erase Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
            Exit Sub
        End If
    End If
    mInvPop
    lacInvInfo.Caption = ""
    lbcVehicle.ListIndex = -1
    imVehIndex = -1
    imCifCyfIndex = -1
    mSetCommands
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmcErase_GotFocus()
    plcCalendar.Visible = False
    tmcClick.Enabled = False
End Sub
Private Sub cmcReport_GotFocus()
    plcCalendar.Visible = False
    tmcClick.Enabled = False
End Sub
Private Sub cmcTranDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcTranDate.SelStart = 0
    edcTranDate.SelLength = Len(edcTranDate.Text)
    edcTranDate.SetFocus
End Sub
Private Sub cmcUndo_Click()
    mReadCifCyf
    mSetCommands
End Sub
Private Sub cmcUndo_GotFocus()
    plcCalendar.Visible = False
    tmcClick.Enabled = False
End Sub
Private Sub cmcUpdate_Click()
    mSaveRec
    mSetCommands
End Sub
Private Sub cmcUpdate_GotFocus()
    plcCalendar.Visible = False
    tmcClick.Enabled = False
End Sub
Private Sub edcTranDate_Change()
    Dim slStr As String
    slStr = edcTranDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    mSetCommands
End Sub
Private Sub edcTranDate_GotFocus()
    plcCalendar.Move plcAffData.Left + edcTranDate.Left, plcAffData.Top + edcTranDate.Top + edcTranDate.Height
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcTranDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcTranDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcTranDate.SelLength <> 0 Then    'avoid deleting two characters
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
Private Sub edcTranDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcTranDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcTranDate.Text = slDate
            End If
        End If
        edcTranDate.SelStart = 0
        edcTranDate.SelLength = Len(edcTranDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcTranDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcTranDate.Text = slDate
            End If
        End If
        edcTranDate.SelStart = 0
        edcTranDate.SelLength = Len(edcTranDate.Text)
    End If
End Sub
Private Sub Form_Activate()
'    gShowBranner
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
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
    
    Erase tmInvCode
    Erase tmVehicleCode
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    btrExtClear hmCyf   'Clear any previous extend operation
    ilRet = btrClose(hmCyf)
    btrDestroy hmCyf
    btrExtClear hmCpf   'Clear any previous extend operation
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    btrExtClear hmCif   'Clear any previous extend operation
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    
    Set AffCopy = Nothing   'Remove data segment
    
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcCifCyf_Click()
    imCifCyfIndex = lbcCifCyf.ListIndex
    mReadCifCyf
    mSetCommands
End Sub
Private Sub lbcCifCyf_GotFocus()
    plcCalendar.Visible = False
    tmcClick.Enabled = False
End Sub
Private Sub lbcVehicle_Click()
    imVehIndex = lbcVehicle.ListIndex
    mSetCommands
End Sub
Private Sub lbcVehicle_GotFocus()
    plcCalendar.Visible = False
    tmcClick.Enabled = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAdvtPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Advertiser list box   *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mAdvtPop()
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = cbcAdvt.ListIndex
    If ilIndex >= 0 Then
        slName = cbcAdvt.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(AffCopy, cbcAdvt, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(AffCopy, cbcAdvt, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", AffCopy
        On Error GoTo 0
'        cbcAdvt.AddItem "[All]", 0  'Force as first item on list
        If ilIndex >= 0 Then
            gFindMatch slName, 0, cbcAdvt
            If gLastFound(cbcAdvt) >= 0 Then
                cbcAdvt.ListIndex = gLastFound(cbcAdvt)
            Else
                cbcAdvt.ListIndex = -1
            End If
        Else
            cbcAdvt.ListIndex = ilIndex
        End If
    End If
    Exit Sub
mAdvtPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    slStr = edcTranDate.Text
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
'*      Procedure Name:mCbcAdvtChange                  *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Process vehicle change         *
'*                                                     *
'*******************************************************
Private Sub mCbcAdvtChange()
    If imChgMode = False Then
        imChgMode = True
        If cbcAdvt.Text <> "" Then
            gManLookAhead cbcAdvt, imBSMode, imComboBoxIndex
        End If
        imAdvtIndex = cbcAdvt.ListIndex
        Screen.MousePointer = vbHourglass  'Wait
        lbcInv.Clear
        Screen.MousePointer = vbHourglass  'Wait
        mInvPop
        'For ilLoop = 0 To 5 Step 1
        '    ckcZone(ilLoop).Value = False
        'Next ilLoop
        Screen.MousePointer = vbDefault    'Default
        mSetCommands
        imChgMode = False
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
    AffCopy.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone AffCopy
    'AffCopy.Show
    DoEvents
    Screen.MousePointer = vbHourglass
    'imcHelp.Picture = Traffic!imcHelp.Picture
    mInitBox
'    mInitDDE

    hmCif = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", AffCopy
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    hmCpf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", AffCopy
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    hmCyf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmCyf, "", sgDBPath & "Cyf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cyf.Btr)", AffCopy
    On Error GoTo 0
    imCyfRecLen = Len(tmCyf)
    hmAdf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", AffCopy
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    imStartMode = True
    imTerminate = False         'terminate if true
    imChgMode = False           'no change made
    imBSMode = False            'back space key
    imCalType = 0               'Standard type
    imSelectDelay = False
    imBypassFocus = False
    imAdvtIndex = -1
    imCifCyfIndex = -1
    imVehIndex = -1
    imTypeIndex = 0
    tmAdf.iCode = 0
    mAdvtPop
    If imTerminate Then
        Exit Sub
    End If
    mVehPop
    If imTerminate Then
        Exit Sub
    End If
    slStr = Format$(gNow(), "m/d/yy")
    slStr = gObtainNextMonday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    Screen.MousePointer = vbDefault
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    
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
'*             Created:7/19/93       By:D. LeVine      *
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
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInvPop                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection Name    *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mInvPop()
    Dim ilRet As Integer
    Dim ilAdvtCode As Integer
    Dim ilISCIProd As Integer
    Dim llCifCode As Long
    Dim slName As String
    Dim slCode As String
    Dim slNameCode As String
    Dim slVehName As String
    Dim slVehCode As String
    Dim slVehNameCode As String
    Dim ilLoop As Integer
    Dim ilVeh As Integer
    Dim llCyfRecPos As Long
    Dim ilInvType As Integer
    ilISCIProd = 0    '&H100=Include Length
    ilInvType = 0   '0=Active only; 6=Active and purged
    lbcCifCyf.Clear
    lbcCifCyfCode.Clear
    If imAdvtIndex < 0 Then
        Exit Sub
    End If
    'If imAdvtIndex = 0 Then
    '    ilAdvtCode = -1
    'Else
        slNameCode = tgAdvertiser(imAdvtIndex).sKey    'Traffic!lbcAdvertiser.List(imAdvtIndex)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilAdvtCode = Val(slCode)
    'End If
    'ilRet = gPopCopyForAdvtBox(AffCopy, ilAdvtCode, ilISCIProd, ilInvType, lbcInv, lbcInvCode)
    ilRet = gPopCopyForAdvtBox(AffCopy, ilAdvtCode, ilISCIProd, ilInvType, lbcInv, tmInvCode(), smInvCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mInvPopErr
        gCPErrorMsg ilRet, "mInvPop (gPopCopyForAdvtBox: AffCopy)", AffCopy
        On Error GoTo 0
    End If
    For ilLoop = 0 To UBound(tmInvCode) - 1 Step 1  'lbcInvCode.ListCount - 1 Step 1
        slNameCode = tmInvCode(ilLoop).sKey    'lbcInvCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        lbcCifCyfCode.AddItem slName & "  " & "\" & slCode & "\0"
        llCifCode = Val(slCode)
        tmCyfSrchKey.lCifCode = llCifCode
        tmCyfSrchKey.iVefCode = 0
        tmCyfSrchKey.sSource = " "
        tmCyfSrchKey.sTimeZone = ""
        tmCyfSrchKey.lRafCode = 0
        ilRet = btrGetGreaterOrEqual(hmCyf, tmCyf, imCyfRecLen, tmCyfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmCyf.lCifCode = llCifCode)
            If (Trim$(tmCyf.sSource) <> "S") And (Trim$(tmCyf.sSource) <> "K") Then
                ilRet = btrGetPosition(hmCyf, llCyfRecPos)
                For ilVeh = 0 To lbcVehicle.ListCount - 1 Step 1
                    slVehNameCode = tmVehicleCode(ilVeh).sKey  'lbcVehicleCode.List(ilVeh)
                    ilRet = gParseItem(slVehNameCode, 1, "\", slVehName)
                    ilRet = gParseItem(slVehName, 3, "|", slVehName)
                    ilRet = gParseItem(slVehNameCode, 2, "\", slVehCode)
                    If Val(slVehCode) = tmCyf.iVefCode Then
                        lbcCifCyfCode.AddItem Trim$(slName) & " " & Trim$(slVehName) & "\" & slCode & "\" & str$(llCyfRecPos)
                        Exit For
                    End If
                Next ilVeh
            End If
            ilRet = btrGetNext(hmCyf, tmCyf, imCyfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Next ilLoop
    For ilLoop = 0 To lbcCifCyfCode.ListCount - 1 Step 1
        slNameCode = lbcCifCyfCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        lbcCifCyf.AddItem slName
    Next ilLoop
    Exit Sub
mInvPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOKName                         *
'*                                                     *
'*             Created:6/1/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test that name is unique        *
'*                                                     *
'*******************************************************
Private Function mOKName()
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slCifCode As String
    Dim slCyfCode As String
    Dim llRecPos As Long
    Dim slVehCode As String
    Dim slVehNameCode As String
    Dim slZone As String
    Dim tlCyf As CYF
    slNameCode = lbcCifCyfCode.List(imCifCyfIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCifCode)
    slVehNameCode = tmVehicleCode(imVehIndex).sKey 'lbcVehicleCode.List(imVehIndex)
    ilRet = gParseItem(slVehNameCode, 2, "\", slVehCode)
    If rbcZone(0).Value Then
        slZone = "EST"
    ElseIf rbcZone(1).Value Then
        slZone = "CST"
    ElseIf rbcZone(2).Value Then
        slZone = "MST"
    ElseIf rbcZone(3).Value Then
        slZone = "PST"
    Else
        slZone = "   "
    End If
    For ilLoop = 0 To lbcCifCyfCode.ListCount - 1 Step 1
        If ilLoop <> imCifCyfIndex Then
            slNameCode = lbcCifCyfCode.List(ilLoop)
            'ilRet = gParseItem(slNameCode, 2, "\", slCifCode)
            ilRet = gParseItem(slNameCode, 3, "\", slCyfCode)
            llRecPos = Val(slCyfCode)
            If llRecPos > 0 Then
                ilRet = btrGetDirect(hmCyf, tlCyf, imCyfRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                If (tlCyf.lCifCode = Val(slCifCode)) And (tlCyf.iVefCode = Val(slVehCode)) And (Trim$(tlCyf.sTimeZone) = Trim$(slZone)) Then
                    Beep
                    ''MsgBox "Affiliate Copy already defined", vbOkOnly + vbExclamation + vbApplicationModal, "Error"
                    gAutomationAlertAndLogHandler "Affiliate Copy already defined", vbOkOnly + vbCritical + vbApplicationModal, "Error"
                    mOKName = False
                    Exit Function
                End If
            End If
        End If
    Next ilLoop
    mOKName = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadCifCyf                     *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read Copy Inventory            *
'*                                                     *
'*******************************************************
Private Sub mReadCifCyf()
    Dim llCifCode As Long
    Dim slCode As String
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slVehName As String
    Dim slVehCode As String
    Dim slVehNameCode As String
    Dim ilVeh As Integer
    lacInvInfo.Caption = ""
    If imAdvtIndex < 0 Then
        Exit Sub
    End If
    If imCifCyfIndex < 0 Then
        Exit Sub
    End If
    slNameCode = lbcCifCyfCode.List(imCifCyfIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    llCifCode = Val(slCode)
    tmCifSrchKey0.lCode = llCifCode
    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        tmCpfSrchKey0.lCode = tmCif.lcpfCode
        ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmCpf.sName = ""
            tmCpf.sISCI = ""
            tmCpf.sCreative = ""
        End If
    Else
        tmCpf.sName = ""
        tmCpf.sISCI = ""
        tmCpf.sCreative = ""
    End If
    lacInvInfo.Caption = Trim$(tmCpf.sISCI) & " " & Trim$(tmCpf.sCreative) & " " & Trim$(tmCpf.sName)
    ilRet = gParseItem(slNameCode, 3, "\", slCode)
    lmCyfRecPos = Val(slCode)
    If lmCyfRecPos = 0 Then
        lbcVehicle.ListIndex = -1
        rbcZone(0).Value = False
        rbcZone(1).Value = False
        rbcZone(2).Value = False
        rbcZone(3).Value = False
        rbcZone(4).Value = True
    Else
        ilRet = btrGetDirect(hmCyf, tmCyf, imCyfRecLen, lmCyfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        lbcVehicle.ListIndex = -1
        For ilVeh = 0 To lbcVehicle.ListCount - 1 Step 1
            slVehNameCode = tmVehicleCode(ilVeh).sKey  'lbcVehicleCode.List(ilVeh)
            ilRet = gParseItem(slVehNameCode, 1, "\", slVehName)
            ilRet = gParseItem(slVehName, 3, "|", slVehName)
            ilRet = gParseItem(slVehNameCode, 2, "\", slVehCode)
            If Val(slVehCode) = tmCyf.iVefCode Then
                lbcVehicle.ListIndex = ilVeh
                Exit For
            End If
        Next ilVeh
        gUnpackDate tmCyf.iFeedDate(0), tmCyf.iFeedDate(1), slDate
        edcTranDate.Text = slDate
        Select Case tmCyf.sTimeZone
            Case "EST"
                rbcZone(0).Value = True
            Case "CST"
                rbcZone(1).Value = True
            Case "MST"
                rbcZone(2).Value = True
            Case "PST"
                rbcZone(3).Value = True
            Case Else
                rbcZone(4).Value = True
        End Select
    End If
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
    Dim slStr As String
    Dim slCode As String
    Dim slNameCode As String
    Dim slVehCode As String
    Dim slVehNameCode As String
    If imAdvtIndex < 0 Then
        Exit Sub
    End If
    If imCifCyfIndex < 0 Then
        Exit Sub
    End If
    If imVehIndex < 0 Then
        Exit Sub
    End If
    If Not mOKName() Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If lmCyfRecPos <> 0 Then
        Do
            'Key is not modifiable- just delete and add
            ilRet = btrGetDirect(hmCyf, tmCyf, imCyfRecLen, lmCyfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
                Exit Sub
            End If
            'tmRec = tmCyf
            'ilRet = gGetByKeyForUpdate("CYF", hmCyf, tmRec)
            'tmCyf = tmRec
            'If ilRet <> BTRV_ERR_NONE Then
            '    Screen.MousePointer = vbDefault
            '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
            '    Exit Sub
            'End If
            ilRet = btrDelete(hmCyf)
        Loop While ilRet = BTRV_ERR_CONFLICT
        If ilRet <> BTRV_ERR_NONE Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
            Exit Sub
        End If
    End If
    slNameCode = lbcCifCyfCode.List(imCifCyfIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    tmCyf.lCifCode = Val(slCode)
    slVehNameCode = tmVehicleCode(imVehIndex).sKey 'lbcVehicleCode.List(imVehIndex)
    ilRet = gParseItem(slVehNameCode, 2, "\", slVehCode)
    tmCyf.iVefCode = Val(slVehCode)
    slStr = edcTranDate.Text
    gPackDate slStr, tmCyf.iFeedDate(0), tmCyf.iFeedDate(1)
    tmCyf.sSource = " "
    If rbcZone(0).Value Then
        tmCyf.sTimeZone = "EST"
    ElseIf rbcZone(1).Value Then
        tmCyf.sTimeZone = "CST"
    ElseIf rbcZone(2).Value Then
        tmCyf.sTimeZone = "MST"
    ElseIf rbcZone(3).Value Then
        tmCyf.sTimeZone = "PST"
    Else
        tmCyf.sTimeZone = "   "
    End If
    tmCyf.lRafCode = 0
    ilRet = btrInsert(hmCyf, tmCyf, imCyfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault
        ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Save")
        Exit Sub
    End If
    mInvPop
    lbcVehicle.ListIndex = -1
    imVehIndex = -1
    imCifCyfIndex = -1
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
    Dim slStr As String
    Dim ilAltered As Integer
    Dim slDate As String
    Dim slZone As String
    Dim slVehCode As String
    Dim slVehNameCode As String
    Dim ilRet As Integer
    ilAltered = False
    'Update button set if all mandatory fields have data and any field altered
    If (imAdvtIndex >= 0) And (imCifCyfIndex >= 0) And (imVehIndex >= 0) And (rbcZone(0).Value Or rbcZone(1).Value Or rbcZone(2).Value Or rbcZone(3).Value Or rbcZone(4).Value) Then
        slStr = edcTranDate.Text
        If gValidDate(slStr) Then
            If lmCyfRecPos = 0 Then
                ilAltered = True
            Else
                slVehNameCode = tmVehicleCode(imVehIndex).sKey 'lbcVehicleCode.List(imVehIndex)
                ilRet = gParseItem(slVehNameCode, 2, "\", slVehCode)
                If Val(slVehCode) <> tmCyf.iVefCode Then
                    ilAltered = True
                End If
                gUnpackDate tmCyf.iFeedDate(0), tmCyf.iFeedDate(1), slDate
                If gDateValue(slDate) <> gDateValue(slStr) Then
                    ilAltered = True
                End If
                If rbcZone(0).Value Then
                    slZone = "EST"
                ElseIf rbcZone(1).Value Then
                    slZone = "CST"
                ElseIf rbcZone(2).Value Then
                    slZone = "MST"
                ElseIf rbcZone(3).Value Then
                    slZone = "PST"
                Else
                    slZone = "   "
                End If
                If Trim$(slZone) <> Trim$(tmCyf.sTimeZone) Then
                    ilAltered = True
                End If
            End If
            If ilAltered Then
                cmcUpdate.Enabled = True
            Else
                cmcUpdate.Enabled = False
            End If
        Else
            cmcUpdate.Enabled = False
        End If
    Else
        cmcUpdate.Enabled = False
    End If
    If (ilAltered) Or (lmCyfRecPos = 0) Then
        cmcErase.Enabled = False
    Else
        cmcErase.Enabled = True
    End If
    If ilAltered Or ((imAdvtIndex >= 0) And (imCifCyfIndex >= 0) And (imVehIndex >= 0) And (lmCyfRecPos = 0)) Then
        cmcUndo.Enabled = True
        cbcAdvt.Enabled = False
        lbcCifCyf.Enabled = False
    Else
        cmcUndo.Enabled = False
        cbcAdvt.Enabled = True
        lbcCifCyf.Enabled = True
    End If
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
    Unload AffCopy
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    Dim llFilter As Long
    llFilter = VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH   ' Airing and conventional vehicles
    'ilRet = gPopUserVehicleBox(AffCopy, ilFilter, lbcVehicle, lbcVehicleCode)
    ilRet = gPopUserVehicleBox(AffCopy, llFilter, lbcVehicle, tmVehicleCode(), smVehicleCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", AffCopy
        On Error GoTo 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
                edcTranDate.Text = Format$(llDate, "m/d/yy")
                edcTranDate.SelStart = 0
                edcTranDate.SelLength = Len(edcTranDate.Text)
                imBypassFocus = True
                edcTranDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcTranDate.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    plcCalendar.Visible = False
    tmcClick.Enabled = False
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub rbcZone_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcZone(Index).Value
    'End of Coded added
    mSetCommands
End Sub
Private Sub rbcZone_GotFocus(Index As Integer)
    plcCalendar.Visible = False
    tmcClick.Enabled = False
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    If imSelectDelay Then
        imSelectDelay = False
        mCbcAdvtChange
    Else
    End If
End Sub
Private Sub plcAffData_Paint()
    plcAffData.CurrentX = 0
    plcAffData.CurrentY = 0
    plcAffData.Print " "
End Sub
Private Sub plcAffCopy_Paint()
    plcAffCopy.CurrentX = 0
    plcAffCopy.CurrentY = 0
    plcAffCopy.Print " "
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Affiliate Copy "
End Sub
