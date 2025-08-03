VERSION 5.00
Begin VB.Form Locks 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3270
   ClientLeft      =   1170
   ClientTop       =   2310
   ClientWidth     =   4665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3270
   ScaleWidth      =   4665
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3105
      TabIndex        =   18
      Top             =   2850
      Width           =   1050
   End
   Begin VB.CommandButton cmcDropDown 
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
      Left            =   1830
      Picture         =   "Locks.frx":0000
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1005
      ScaleHeight     =   210
      ScaleWidth      =   1740
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   795
      Width           =   1740
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   915
      MaxLength       =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   195
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   1410
      Left            =   300
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1335
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "Locks.frx":00FA
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   45
         Width           =   1020
         Begin VB.Image imcTmeOutline 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   330
            Top             =   270
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image imcTmeInv 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   480
            Left            =   360
            Picture         =   "Locks.frx":0DB8
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   2490
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   975
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Locks.frx":10C2
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   14
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
         Left            =   45
         TabIndex        =   10
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
         Left            =   1635
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   11
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox pbcDates 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   1065
      Left            =   570
      Picture         =   "Locks.frx":3EDC
      ScaleHeight     =   1065
      ScaleWidth      =   3525
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   3525
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3585
      Width           =   75
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   75
      ScaleHeight     =   165
      ScaleWidth      =   135
      TabIndex        =   15
      Top             =   3360
      Width           =   135
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   1
      Top             =   225
      Width           =   105
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1890
      TabIndex        =   17
      Top             =   2850
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   690
      TabIndex        =   16
      Top             =   2850
      Width           =   1050
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   75
      ScaleHeight     =   240
      ScaleWidth      =   4290
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   4290
   End
   Begin VB.PictureBox plcDates 
      ForeColor       =   &H00000000&
      Height          =   1155
      Left            =   525
      ScaleHeight     =   1095
      ScaleWidth      =   3570
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   315
      Width           =   3630
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4245
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3615
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3900
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2235
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   2775
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "Locks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Locks.frm on Wed 6/17/09 @ 12:56 PM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Locks.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim tmCtrls(0 To 6)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current Media Box
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim imBypassFocus As Integer
Dim smSave(0 To 4) As String  'Values saved (1=Start date; 2=End date; 1=Start Time; 2=End Time). Index zero ignored
Dim imSave(0 To 2) As Integer   'Index 1: 0=Unlock, 1=Lock; Index 2: 0=Avail, 1=Spots. Index zero ignored
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim hmVef As Integer 'Veficle file handle
Dim tmVef As VEF        'VEF record image
Dim tmVefSrchKey As INTKEY0    'VEF key record image
Dim imVefRecLen As Integer        'VEF record length
Dim imVehCode As Integer
Dim smScreenCaption As String
Dim imUpdateAllowed As Integer    'User can update records

Const LOCKINDEX = 1         'Unlock/Lock control/field
Const AVAILINDEX = 2        'Avail/Spot control/field
Const STARTDATEINDEX = 3    'Start date control/field
Const ENDDATEINDEX = 4      'End date control/field
Const STARTTIMEINDEX = 5    'Start time control/field
Const ENDTIMEINDEX = 6      'End time control/field
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcCancel_Click()
    igLockCallSource = CALLCANCELLED
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
    Dim ilChg As Integer
    Dim ilRes As Integer
    Dim ilLoop As Integer
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    'Ask if it should be updated
    ilChg = False
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        If tmCtrls(ilLoop).iChg Then
            ilChg = True
            Exit For
        End If
    Next ilLoop
    If ilChg Then
        ilRes = MsgBox("Lock/Unlock", vbYesNoCancel + vbQuestion, "Update")
        If ilRes = vbCancel Then
            Exit Sub
        End If
        If ilRes = vbYes Then
            If mLocks() = False Then
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
    End If
    If igLockCallSource <> CALLNONE Then
        igLockCallSource = CALLDONE
        mTerminate
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case STARTDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case ENDDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case STARTTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case ENDTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUpdate_Click()
    Dim ilLoop As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If mLocks() = False Then
        mEnableBox imBoxNo
        Exit Sub
    End If
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Select Case imBoxNo
        Case STARTDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                Exit Sub
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
            tmCtrls(imBoxNo).iChg = True
        Case ENDDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                Exit Sub
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
            tmCtrls(imBoxNo).iChg = True
        Case STARTTIMEINDEX
            tmCtrls(imBoxNo).iChg = True
        Case ENDTIMEINDEX
            tmCtrls(imBoxNo).iChg = True
    End Select
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case STARTDATEINDEX
        Case ENDDATEINDEX
        Case STARTTIMEINDEX
        Case ENDTIMEINDEX
    End Select
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imBoxNo
        Case STARTDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case ENDDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case STARTTIMEINDEX
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                ilFound = False
                For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
                    If KeyAscii = igLegalTime(ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        Case ENDTIMEINDEX
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                ilFound = False
                For ilLoop = LBound(igLegalTime) To UBound(igLegalTime) Step 1
                    If KeyAscii = igLegalTime(ilLoop) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    Beep
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case STARTDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case ENDDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcCalendar.Visible = Not plcCalendar.Visible
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYUP Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case STARTTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
            Case ENDTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imBoxNo
            Case STARTDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case ENDDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Up arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
            Case STARTTIMEINDEX
            Case ENDTIMEINDEX
        End Select
    End If
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If (igWinStatus(SPOTSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        pbcSTab.Enabled = False
        pbcTab.Enabled = False
        imUpdateAllowed = False
        'Show branner
'        edcLinkDestHelpMsg.LinkExecute "BF"
    Else
        pbcSTab.Enabled = True
        pbcTab.Enabled = True
        imUpdateAllowed = True
        'Show branner
'        edcLinkDestHelpMsg.LinkExecute "BT"
    End If
    gShowBranner imUpdateAllowed
    'This loop is required to prevent a timing problem- if calling
    'with sg----- = "", then loss GotFocus to first control
    'without this loop
'    For ilLoop = 1 To 100 Step 1
'        DoEvents
'    Next ilLoop
'    gShowBranner
    Me.KeyPreview = True
    Locks.Refresh
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
        plcTme.Visible = False
        gFunctionKeyBranch KeyCode
        'Removed as causing a loop
        'If imBoxNo > 0 Then
        '    mEnableBox imBoxNo
        'End If
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
    
    Erase smSave
    Erase imSave
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    
    Set Locks = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
    slStr = edcDropDown.Text
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
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case LOCKINDEX
            If (imSave(1) = -1) Then
                imSave(1) = 1 'Lock
                tmCtrls(ilBoxNo).iChg = True
            End If
            pbcType.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcDates, pbcType, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcType_Paint
            pbcType.Visible = True
            pbcType.SetFocus
        Case AVAILINDEX
            If (imSave(2) = -1) Then
                imSave(2) = 0 'Avail
                tmCtrls(ilBoxNo).iChg = True
            End If
            pbcType.Width = tmCtrls(ilBoxNo).fBoxW
            gMoveFormCtrl pbcDates, pbcType, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            pbcType_Paint
            pbcType.Visible = True
            pbcType.SetFocus
        Case STARTDATEINDEX 'Start date
            edcDropDown.Width = tmCtrls(STARTDATEINDEX).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcDates, edcDropDown, tmCtrls(STARTDATEINDEX).fBoxX, tmCtrls(STARTDATEINDEX).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If smSave(1) = "" Then
                slStr = Format$(gNow(), "m/d/yy")
                slStr = Format$(gDateValue(slStr) + 1, "m/d/yy")
            Else
                slStr = smSave(1)
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ENDDATEINDEX 'Start date
            edcDropDown.Width = tmCtrls(ENDDATEINDEX).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcDates, edcDropDown, tmCtrls(ENDDATEINDEX).fBoxX, tmCtrls(ENDDATEINDEX).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            plcCalendar.Move edcDropDown.Left + cmcDropDown.Width + edcDropDown.Width - plcCalendar.Width, edcDropDown.Top + edcDropDown.Height
            If smSave(2) <> "" Then
                slStr = smSave(2)
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
            Else
                slStr = smSave(1)
            End If
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case STARTTIMEINDEX 'Start time
            edcDropDown.Width = tmCtrls(STARTTIMEINDEX).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcDates, edcDropDown, tmCtrls(STARTTIMEINDEX).fBoxX, tmCtrls(STARTTIMEINDEX).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If smSave(3) = "" Then
                edcDropDown.Text = "12M"
            Else
                edcDropDown.Text = smSave(3)
            End If
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ENDTIMEINDEX 'Start time
            edcDropDown.Width = tmCtrls(ENDTIMEINDEX).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcDates, edcDropDown, tmCtrls(ENDTIMEINDEX).fBoxX, tmCtrls(ENDTIMEINDEX).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If smSave(4) = "" Then
                edcDropDown.Text = "12M"
            Else
                edcDropDown.Text = smSave(4)
            End If
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
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
    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    imLBCDCtrls = 1
    imFirstActivate = True
    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    imTerminate = False
    imFirstFocus = True
    imBypassFocus = False
    imSettingValue = False
    imBoxNo = -1 'Initialize current Box to N/A
    imChgMode = False
    imBSMode = False
    imCalType = 0   'Standard
    imSave(1) = 1   'Default to Lock
    imSave(2) = 0   'Default to Avails
    mInitBox
    Locks.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone Locks
    'Locks.Show
    'imcHelp.Picture = Traffic!imcHelp.Picture
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", Locks
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    tmVefSrchKey.iCode = imVehCode
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If (ilRet = BTRV_ERR_NONE) Then
        smScreenCaption = "Lock- " & Trim$(tmVef.sName)
    End If
    plcScreen_Paint
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
    Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
    flTextHeight = pbcDates.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcDates.Move 525, 315, pbcDates.Width + fgPanelAdj, pbcDates.Height + fgPanelAdj
    pbcDates.Move plcDates.Left + fgBevelX, plcDates.Top + fgBevelY
    'Unlock/Lock
    gSetCtrl tmCtrls(LOCKINDEX), 30, 30, 1725, fgBoxStH
    'Avail/Spot
    gSetCtrl tmCtrls(AVAILINDEX), 1770, tmCtrls(LOCKINDEX).fBoxY, 1725, fgBoxStH
    'Start date
    gSetCtrl tmCtrls(STARTDATEINDEX), tmCtrls(LOCKINDEX).fBoxX, tmCtrls(LOCKINDEX).fBoxY + fgStDeltaY, 1725, fgBoxStH
    'End date
    gSetCtrl tmCtrls(ENDDATEINDEX), tmCtrls(AVAILINDEX).fBoxX, tmCtrls(STARTDATEINDEX).fBoxY, 1725, fgBoxStH
    'Start Time
    gSetCtrl tmCtrls(STARTTIMEINDEX), tmCtrls(LOCKINDEX).fBoxX, tmCtrls(STARTDATEINDEX).fBoxY + fgStDeltaY, 1725, fgBoxStH
    'End date
    gSetCtrl tmCtrls(ENDTIMEINDEX), tmCtrls(AVAILINDEX).fBoxX, tmCtrls(STARTTIMEINDEX).fBoxY, 1725, fgBoxStH
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLocks                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Lock avails and spots as       *
'*                      specified by the user          *
'*                                                     *
'*******************************************************
Private Function mLocks() As Integer
    Dim ilVehCode As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim ilAvailLock As Integer
    Dim ilSpotLock As Integer
    If mTestSaveFields() = NO Then
        mLocks = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    ilVehCode = imVehCode
    slStartDate = smSave(1)
    slEndDate = smSave(2)
    slStartTime = smSave(3)
    slEndTime = smSave(4)
    If imSave(2) = 0 Then   'Lock/Unlock Avail
        ilSpotLock = -1
        If imSave(1) = 0 Then
            ilAvailLock = 0 'Unlock
        Else
            ilAvailLock = 1 'Lock
        End If
    Else    'Lock/unlock spots
        ilAvailLock = -1
        If imSave(1) = 0 Then
            ilSpotLock = 0  'Unlock
        Else
            ilSpotLock = 1  'Lock
        End If
    End If
    gSetLockStatus ilVehCode, ilAvailLock, ilSpotLock, slStartDate, slEndDate, 0, slStartTime, slEndTime
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    For ilCount = LBound(smSave) To UBound(smSave) Step 1
        smSave(ilCount) = ""
    Next ilCount
    For ilCount = LBound(imSave) To UBound(imSave) Step 1
        imSave(ilCount) = -1
    Next ilCount
    pbcDates.Cls
    Screen.MousePointer = vbDefault
    mLocks = True
End Function
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
    'gInitStdAlone Locks, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igLockCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igLockCallSource = CALLNONE
    '    imVehCode = 1
    'End If
    If igLockCallSource <> CALLNONE Then  'If from sales office- set name and branch to control
        ilRet = gParseItem(slCommand, 4, "\", slStr)
        If ilRet = CP_MSG_NONE Then
            sgLockName = slStr  'Vehicle code
            imVehCode = Val(slStr)
        Else
            sgLockName = "1"
            imVehCode = Val(slStr)
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case LOCKINDEX
            pbcType.Visible = False  'Set visibility
            If imSave(1) = 0 Then
                slStr = "Unlock"
            ElseIf imSave(1) = 1 Then
                slStr = "Lock"
            Else
                slStr = ""
            End If
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
        Case AVAILINDEX
            pbcType.Visible = False  'Set visibility
            If imSave(2) = 0 Then
                slStr = "Avails"
            ElseIf imSave(2) = 1 Then
                slStr = "Spots"
            Else
                slStr = ""
            End If
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
        Case STARTDATEINDEX 'Start Date
            plcCalendar.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            smSave(1) = slStr
            slStr = gFormatDate(slStr)
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
        Case ENDDATEINDEX 'End Date
            plcCalendar.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            smSave(2) = slStr
            slStr = gFormatDate(slStr)
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
        Case STARTTIMEINDEX
            cmcDropDown.Visible = False
            plcTme.Visible = False
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            smSave(3) = slStr
            If slStr <> "" Then
                slStr = gFormatTime(slStr, "A", "1")
            End If
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
        Case ENDTIMEINDEX
            cmcDropDown.Visible = False
            plcTme.Visible = False
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            smSave(4) = slStr
            If slStr <> "" Then
                slStr = gFormatTime(slStr, "A", "1")
            End If
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
    End Select
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
    sgDoneMsg = Trim$(str$(igLockCallSource)) & "\" & sgLockName
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Locks
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestSaveFields                 *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSaveFields() As Integer
'
'   iRet = mTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    If smSave(1) = "" Then
        ilRes = MsgBox("Start date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = STARTDATEINDEX
        mTestSaveFields = NO
        Exit Function
    Else
        If Not gValidDate(smSave(1)) Then
            ilRes = MsgBox("Start date must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = STARTDATEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If

    If (smSave(2) = "") Then
        ilRes = MsgBox("End date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = ENDDATEINDEX
        mTestSaveFields = NO
        Exit Function
    Else
        If Not gValidDate(smSave(2)) Then
            ilRes = MsgBox("End date must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = ENDDATEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
        If gDateValue(smSave(1)) > gDateValue(smSave(2)) Then
            ilRes = MsgBox("End date must be after start date", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = ENDDATEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If smSave(3) = "" Then
        ilRes = MsgBox("Start time must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = STARTTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    Else
        If Not gValidTime(smSave(3)) Then
            ilRes = MsgBox("Time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = STARTTIMEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If smSave(4) = "" Then
        ilRes = MsgBox("End time must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = ENDTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    Else
        If Not gValidTime(smSave(4)) Then
            ilRes = MsgBox("End time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = ENDTIMEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If gTimeToCurrency(smSave(3), False) > gTimeToCurrency(smSave(4), True) Then
        ilRes = MsgBox("End time must be after start time", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = ENDTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    mTestSaveFields = YES
End Function
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
                edcDropDown.Text = Format$(llDate, "m/d/yy")
                edcDropDown.SelStart = 0
                edcDropDown.SelLength = Len(edcDropDown.Text)
                imBypassFocus = True
                edcDropDown.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcDropDown.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcDates_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
            If (Y >= (tmCtrls(ilBox).fBoxY)) And (Y <= (tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcDates_Paint()
    Dim ilBox As Integer

    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        'gPaintArea pbcDates, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
        pbcDates.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
        pbcDates.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY '- 30 '+ fgBoxInsetY
        pbcDates.Print tmCtrls(ilBox).sShow
    Next ilBox
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            ilBox = 1
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case LOCKINDEX
            mSetShow imBoxNo
            cmcDone.SetFocus
        Case STARTDATEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo - 1
        Case ENDDATEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo - 1
        Case STARTTIMEINDEX 'Time (first control within header)
            slStr = edcDropDown.Text
            If Not gValidTime(slStr) Then
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo - 1
        Case ENDTIMEINDEX 'Time (first control within header)
            slStr = edcDropDown.Text
            If Not gValidTime(slStr) Then
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo - 1
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            ilBox = ENDDATEINDEX
            Exit Sub
        Case STARTDATEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo + 1
        Case ENDDATEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo + 1
        Case STARTTIMEINDEX
            slStr = edcDropDown.Text
            If Not gValidTime(slStr) Then
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo + 1
        Case ENDTIMEINDEX
            slStr = edcDropDown.Text
            If Not gValidTime(slStr) Then
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            mSetShow imBoxNo
            cmcUpdate.SetFocus
        Case Else 'Last control within header
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeInv.Visible = True
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcTme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRowNo As Integer
    Dim ilColNo As Integer
    Dim flX As Single
    Dim flY As Single
    Dim slKey As String
    Dim ilMaxCol As Integer
    imcTmeInv.Visible = False
    flY = fgPadMinY
    For ilRowNo = 1 To 5 Step 1
        If (Y >= flY) And (Y <= flY + fgPadDeltaY) Then
            flX = fgPadMinX
            If ilRowNo = 4 Then
                ilMaxCol = 2
            Else
                ilMaxCol = 3
            End If
            For ilColNo = 1 To ilMaxCol Step 1
                If (X >= flX) And (X <= flX + fgPadDeltaX) Then
                    imcTmeInv.Move flX, flY
                    imcTmeOutline.Move flX - 15, flY - 15
                    imcTmeOutline.Visible = True
                    Select Case ilRowNo
                        Case 1
                            Select Case ilColNo
                                Case 1
                                    slKey = "7"
                                Case 2
                                    slKey = "8"
                                Case 3
                                    slKey = "9"
                            End Select
                        Case 2
                            Select Case ilColNo
                                Case 1
                                    slKey = "4"
                                Case 2
                                    slKey = "5"
                                Case 3
                                    slKey = "6"
                            End Select
                        Case 3
                            Select Case ilColNo
                                Case 1
                                    slKey = "1"
                                Case 2
                                    slKey = "2"
                                Case 3
                                    slKey = "3"
                            End Select
                        Case 4
                            Select Case ilColNo
                                Case 1
                                    slKey = "0"
                                Case 2
                                    slKey = "00"
                            End Select
                        Case 5
                            Select Case ilColNo
                                Case 1
                                    slKey = ":"
                                Case 2
                                    slKey = "AM"
                                Case 3
                                    slKey = "PM"
                            End Select
                    End Select
                    Select Case imBoxNo
                        Case STARTTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                        Case ENDTIMEINDEX
                            imBypassFocus = True    'Don't change select text
                            edcDropDown.SetFocus
                            'SendKeys slKey
                            gSendKeys edcDropDown, slKey
                    End Select
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
End Sub
Private Sub pbcType_GotFocus()
    gCtrlGotFocus ActiveControl
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
End Sub
Private Sub pbcType_KeyPress(KeyAscii As Integer)
    If imBoxNo = LOCKINDEX Then
        If (KeyAscii = Asc("L")) Or (KeyAscii = Asc("l")) Then
            imSave(1) = 1
            tmCtrls(imBoxNo).iChg = True
            pbcType_Paint
        End If
        If (KeyAscii = Asc("U")) Or (KeyAscii = Asc("u")) Then
            imSave(1) = 0
            tmCtrls(imBoxNo).iChg = True
            pbcType_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSave(1) = 0 Then
                imSave(1) = 1
            Else
                imSave(1) = 0
            End If
            tmCtrls(imBoxNo).iChg = True
            pbcType_Paint
        End If
    ElseIf imBoxNo = AVAILINDEX Then
        If (KeyAscii = Asc("A")) Or (KeyAscii = Asc("a")) Then
            imSave(2) = 0
            tmCtrls(imBoxNo).iChg = True
            pbcType_Paint
        End If
        If (KeyAscii = Asc("S")) Or (KeyAscii = Asc("s")) Then
            imSave(2) = 1
            tmCtrls(imBoxNo).iChg = True
            pbcType_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If imSave(2) = 0 Then
                imSave(2) = 1
            Else
                imSave(2) = 0
            End If
            pbcType_Paint
            tmCtrls(imBoxNo).iChg = True
        End If
    End If
End Sub
Private Sub pbcType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imBoxNo = LOCKINDEX Then
        If imSave(1) = 0 Then
            imSave(1) = 1
        Else
            imSave(1) = 0
        End If
        tmCtrls(imBoxNo).iChg = True
        pbcType_Paint
    ElseIf imBoxNo = AVAILINDEX Then
        If imSave(2) = 0 Then
            imSave(2) = 1
        Else
            imSave(2) = 0
        End If
        tmCtrls(imBoxNo).iChg = True
        pbcType_Paint
    End If
End Sub
Private Sub pbcType_Paint()
    pbcType.Cls
    pbcType.CurrentX = fgBoxInsetX
    pbcType.CurrentY = -15 'fgBoxInsetY
    If imBoxNo = LOCKINDEX Then
        If imSave(1) = 0 Then
            pbcType.Print "Unlock"
        ElseIf imSave(1) = 1 Then
            pbcType.Print "Lock"
        End If
    ElseIf imBoxNo = AVAILINDEX Then
        If imSave(2) = 0 Then
            pbcType.Print "Avail"
        ElseIf imSave(2) = 1 Then
            pbcType.Print "Spot"
        End If
    End If
End Sub
Private Sub plcDates_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub
