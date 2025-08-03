VERSION 5.00
Begin VB.Form Rollover 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3180
   ClientLeft      =   1170
   ClientTop       =   2070
   ClientWidth     =   4515
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3180
   ScaleWidth      =   4515
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
      Left            =   750
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1560
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
         Left            =   30
         Picture         =   "RollOver.frx":0000
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   11
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
            Picture         =   "RollOver.frx":0CBE
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
      Left            =   2310
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   870
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
         Picture         =   "RollOver.frx":0FC8
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   16
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
            TabIndex        =   17
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
         TabIndex        =   13
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   315
         TabIndex        =   14
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3585
      Width           =   75
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
      Left            =   2445
      TabIndex        =   9
      Top             =   2730
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
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
      Left            =   825
      TabIndex        =   8
      Top             =   2730
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
      Left            =   60
      ScaleHeight     =   240
      ScaleWidth      =   840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   840
   End
   Begin VB.PictureBox plcDates 
      ForeColor       =   &H00000000&
      Height          =   1605
      Left            =   195
      ScaleHeight     =   1545
      ScaleWidth      =   3930
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   3990
      Begin VB.TextBox edcDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   3
         Top             =   300
         Width           =   930
      End
      Begin VB.CommandButton cmcDate 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2865
         Picture         =   "RollOver.frx":3DE2
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   300
         Width           =   195
      End
      Begin VB.TextBox edcTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   6
         Top             =   765
         Width           =   930
      End
      Begin VB.CommandButton cmcTime 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   5.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2865
         Picture         =   "RollOver.frx":3EDC
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   765
         Width           =   195
      End
      Begin VB.Label lacRolloverTime 
         Appearance      =   0  'Flat
         Caption         =   "Next Blocking Time"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label lacRolloverDate 
         Appearance      =   0  'Flat
         Caption         =   "Next Blocking Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1680
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   2625
      Width           =   360
   End
End
Attribute VB_Name = "Rollover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rollover.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Rollover.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text
Dim tmSpf As SPF        'Spf record image
Dim hmSpf As Integer    'Site Preference file handle
Dim imSpfRecLen As Integer        'SPF record length
Dim tmPjf As PJF        'Pjf record image
Dim hmPjf As Integer    'Projection file handle
Dim tmPjfSrchKey1 As LONGKEY0
Dim imPjfRecLen As Integer        'PJF record length
Dim lmPjfCode() As Long
'Dim tmRec As LPOPREC
'Program library dates Field Areas
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imBypassFocus As Integer
Dim lmNowDate As Long
Dim imPrjStartWk As Integer
Dim imCurYear As Integer
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
    edcDate.SelStart = 0
    edcDate.SelLength = Len(edcDate.Text)
    edcDate.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcDate.SelStart = 0
    edcDate.SelLength = Len(edcDate.Text)
    edcDate.SetFocus
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
    plcTme.Visible = False
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcDate.SelStart = 0
    edcDate.SelLength = Len(edcDate.Text)
    edcDate.SetFocus
End Sub
Private Sub cmcDate_GotFocus()
    plcTme.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_Click()
    Dim slDate As String
    Dim slTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim ilRes As Integer
    Dim ilRet As Integer
    slDate = Trim$(edcDate.Text)
    If slDate = "" Then
        ilRes = MsgBox("Next Rollover date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        edcDate.SetFocus
        Exit Sub
    Else
        If Not gValidDate(slDate) Then
            ilRes = MsgBox("Next Rollover date must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            edcDate.SetFocus
            Exit Sub
        End If
    End If
    slTime = Trim$(edcTime.Text)
    If slTime = "" Then
        ilRes = MsgBox("Next Rollover time must be specified", vbOKOnly + vbExclamation, "Incomplete")
        edcTime.SetFocus
        Exit Sub
    Else
        If Not gValidTime(slTime) Then
            ilRes = MsgBox("Next Rollover time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            edcTime.SetFocus
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    slNowDate = Format$(gNow(), "m/d/yy")
    slNowTime = Format$(gNow(), "h:mm:ssAM/PM")
    ilRet = mRolloverRec(slNowDate, slNowTime)
    If ilRet Then
        ilRet = mNextRollover(slDate, slTime)
    Else
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    plcCalendar.Visible = False
    plcTme.Visible = False
End Sub
Private Sub cmcTime_Click()
    plcTme.Visible = Not plcTme.Visible
    edcTime.SelStart = 0
    edcTime.SelLength = Len(edcTime.Text)
    edcTime.SetFocus
End Sub
Private Sub cmcTime_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDate_Change()
    Dim slStr As String
    slStr = edcDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    lacDate.Visible = True
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcDate_GotFocus()
    plcTme.Visible = False
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDate.SelLength <> 0 Then    'avoid deleting two characters
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
Private Sub edcDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcDate.Text = slDate
            End If
        End If
        edcDate.SelStart = 0
        edcDate.SelLength = Len(edcDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcDate.Text = slDate
            End If
        End If
        edcDate.SelStart = 0
        edcDate.SelLength = Len(edcDate.Text)
    End If
End Sub
Private Sub edcTime_GotFocus()
    plcCalendar.Visible = False
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub
Private Sub edcTime_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcTime_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcTime.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
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
End Sub
Private Sub edcTime_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcTme.Visible = Not plcTme.Visible
        End If
        edcTime.SelStart = 0
        edcTime.SelLength = Len(edcTime.Text)
    End If
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
        plcTme.Visible = False
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase lmPjfCode
    btrExtClear hmPjf   'Clear any previous extend operation
    ilRet = btrClose(hmPjf)
    btrDestroy hmPjf
    btrExtClear hmSpf   'Clear any previous extend operation
    ilRet = btrClose(hmSpf)
    btrDestroy hmSpf
    
    Set Rollover = Nothing   'Remove data segment

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
    slStr = edcDate.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(Str$(Day(llDate)))
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
    Dim slStr As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slStart As String
    Dim ilCurMonth As Integer
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imFirstActivate = True
    imTerminate = False
    imBypassFocus = False
    imSettingValue = False
    imChgMode = False
    imBSMode = False
    imCalType = 0   'Standard
    mInitBox
    hmPjf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmPjf, "", sgDBPath & "Pjf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Pjf.Btr)", Rollover
    On Error GoTo 0
    imPjfRecLen = Len(tmPjf)
    hmSpf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmSpf, "", sgDBPath & "Spf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Spf.Btr)", Rollover
    On Error GoTo 0
    imSpfRecLen = Len(tmSpf)
    Rollover.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone Rollover
    plcCalendar.Move plcDates.Left + edcDate.Left, plcDates.Top + edcDate.Top + edcDate.Height
    plcTme.Move plcDates.Left + edcTime.Left, plcDates.Top + edcTime.Top + edcTime.Height
    'slStr = Format$(gNow(), "m/d/yy")
    'slStr = gIncOneDay(gObtainPrevMonday(slStr))    'Get tuesday of the current week
    gUnpackDate tgSpf.iNROBlockDate(0), tgSpf.iNROBlockDate(1), slStr
    If slStr = "" Then
        slStr = gNow()
    End If
    slStr = gIncOneWeek(slStr)    'Get Next week
    edcDate.Text = slStr
    'gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    'pbcCalendar_Paint   'mBoxCalDate called within paint
    'lacDate.Visible = False
    gUnpackTime tgSpf.iNROBlockTime(0), tgSpf.iNROBlockTime(1), "A", "1", slStr
    edcTime.Text = slStr    '"12PM"
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    slDate = gObtainEndStd(slDate)
    gObtainMonthYear 0, slDate, ilCurMonth, imCurYear
    'If tgSpf.sRUseCorpCal = "Y" Then
    '    slDate = "1/15/" & Trim$(Str$(imCurYear))
    '    slStart = gObtainStartCorp(slDate, True)
    'Else
        slDate = "1/15/" & Trim$(Str$(imCurYear))
        slStart = gObtainStartStd(slDate)
    'End If
    imPrjStartWk = (lmNowDate - gDateValue(slStart)) \ 7 + 1
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
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mNextRollover                   *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mNextRollover(slRODate As String, slROTime As String) As Integer
'
'   iRet = mNextRollover (slRODate, slROTime)
'   Where:
'       slRODate(I)- Rollover date
'       slROTime(I)- Rollover time
'       iRet (O)- True if record updated
'                 False if not read
'
    Dim ilRet As Integer    'Return status

    Do
        ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        gPackDate slRODate, tmSpf.iNROBlockDate(0), tmSpf.iNROBlockDate(1)
        gPackTime slROTime, tmSpf.iNROBlockTime(0), tmSpf.iNROBlockTime(1)
        ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet = BTRV_ERR_NONE Then
        sgSpfStamp = "~"
        gSpfRead
        'gSetUsingTraffic
    Else
        On Error GoTo mNextRolloverErr
        gBtrvErrorMsg ilRet, "cmcSave (btrUpdate)", BlkDate
        On Error GoTo 0
    End If
    mNextRollover = True
    Exit Function
mNextRolloverErr:
    On Error GoTo 0
    mNextRollover = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mRolloverRec                    *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mRolloverRec(slRODate As String, slROTime As String) As Integer
'
'   iRet = mRolloverRec (slRODate, slROTime)
'   Where:
'       slRODate(I)- Rollover date
'       slROTime(I)- Rollover time
'       iRet (O)- True if record updated
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilCRet As Integer
    Dim slStr As String
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilRecOK As Integer
    Dim ilLoop As Integer
    Dim ilWk As Integer
    Dim llGross As Long
    Dim slMsg As String
    Dim llRolloverTime As Long
    Dim llRolloverDate As Long
    Dim llTestDate As Long
    Dim llTestTime As Long
    Dim slDate As String
    Dim slStart As String
    Dim ilNoWks As Integer
    Dim slEnd As String
    Dim ilRollForward As Integer
    Dim ilStartWk As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    'Dim tlPjf As PJF
    ReDim lmPjfCode(0 To 0) As Long
    llRolloverDate = gDateValue(slRODate)
    llRolloverTime = CLng(gTimeToCurrency(slROTime, False))
    btrExtClear hmPjf   'Clear any previous extend operation
    ilExtLen = Len(tmPjf)  'Extract operation record size
    ilRet = btrGetFirst(hmPjf, tmPjf, imPjfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    'ilRet = btrGetFirst(hmPjf, tgPjf1Rec(1).tPjf, imBsfRecLen, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmPjf, llNoRec, -1, "UC", "PJF", "") '"EG") 'Set extract limits (all records)
        tlDateTypeBuff.iDate0 = 0
        tlDateTypeBuff.iDate1 = 0
        ilOffSet = gFieldOffset("Pjf", "PjfRolloverDate")
        ilRet = btrExtAddLogicConst(hmPjf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        On Error GoTo mRolloverErr
        gBtrvErrorMsg ilRet, "mRolloverRec (btrExtAddLogicConst):" & "Pjf.Btr", Rollover
        On Error GoTo 0
        ilRet = btrExtAddField(hmPjf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mRolloverErr
        gBtrvErrorMsg ilRet, "mRolloverRec (btrExtAddField):" & "Pjf.Btr", Rollover
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
        ilRet = btrExtGetNext(hmPjf, tmPjf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mRolloverErr
            gBtrvErrorMsg ilRet, "mRolloverRec (btrExtGetNextExt):" & "Pjf.Btr", Rollover
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmRpf, tlRpfExt, ilExtLen, llRecPos)
            ilExtLen = Len(tmPjf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmPjf, tmPjf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slStr = ""
                ilRecOK = True
                'Test date/Time
                gUnpackDateLong tmPjf.iEffDate(0), tmPjf.iEffDate(1), llTestDate
                gUnpackTimeLong tmPjf.iEffTime(0), tmPjf.iEffTime(1), False, llTestTime
                If (llTestDate < llRolloverDate) Or ((llTestDate = llRolloverDate) And (llTestTime < llRolloverTime)) Then
                    lmPjfCode(UBound(lmPjfCode)) = tmPjf.lCode  'llRecPos
                    ReDim Preserve lmPjfCode(0 To UBound(lmPjfCode) + 1) As Long
                End If
                ilRet = btrExtGetNext(hmPjf, tmPjf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmPjf, tmPjf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gGetSyncDateTime slSyncDate, slSyncTime
    ilRet = btrBeginTrans(hmPjf, 1000)
    If ilRet <> BTRV_ERR_NONE Then
        Screen.MousePointer = vbDefault    'Default
        ilRet = MsgBox("File in Use [Re-press Ok], BeginTran" & Str$(ilRet), vbOKOnly + vbExclamation, "Rollover")
        'imTerminate = True
        mRolloverRec = False
        Exit Function
    End If
    'Set rollover date
    For ilLoop = LBound(lmPjfCode) To UBound(lmPjfCode) - 1 Step 1
        'ilRet = btrGetDirect(hmPjf, tmPjf, imPjfRecLen, lmPjfCode(ilLoop), INDEXKEY0, BTRV_LOCK_NONE)
        tmPjfSrchKey1.lCode = lmPjfCode(ilLoop)
        ilRet = btrGetEqual(hmPjf, tmPjf, imPjfRecLen, tmPjfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmPjf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("File in Use [Re-press Ok], GetDirect" & Str$(ilRet), vbOKOnly + vbExclamation, "Rollover")
            'imTerminate = True
            mRolloverRec = False
            Exit Function
        End If
        'tmRec = tmPjf
        'ilRet = gGetByKeyForUpdate("PJF", hmPjf, tmRec)
        'tmPjf = tmRec
        'If ilRet <> BTRV_ERR_NONE Then
        '    ilCRet = btrAbortTrans(hmPjf)
        '    Screen.MousePointer = vbDefault    'Default
        '    ilRet = MsgBox("File in Use [Re-press Ok], GetByKey" & Str$(ilRet), vbOkOnly + vbExclamation, "Rollover")
        '    'imTerminate = True
        '    mRolloverRec = False
        '    Exit Function
        'End If
        tmPjf.lCode = 0
        gPackDate slRODate, tmPjf.iRolloverDate(0), tmPjf.iRolloverDate(1)
        tmPjf.iRemoteID = tgUrf(0).iRemoteID
        tmPjf.lAutoCode = tmPjf.lCode
        ilRet = btrInsert(hmPjf, tmPjf, imPjfRecLen, INDEXKEY0)
        slMsg = "mRolloverRec (btrInsert: Projection)"
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmPjf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("File in Use [Re-press Ok], Insert" & Str$(ilRet), vbOKOnly + vbExclamation, "Rollover")
            'imTerminate = True
            mRolloverRec = False
            Exit Function
        End If
        tmPjf.lAutoCode = tmPjf.lCode
        tmPjf.iSourceID = 0
        'gPackDate "", tmPjf.iSyncDate(0), tmPjf.iSyncDate(1)
        'gPackTime "", tmPjf.iSyncTime(0), tmPjf.iSyncTime(1)
        gPackDate slSyncDate, tmPjf.iSyncDate(0), tmPjf.iSyncDate(1)
        gPackTime slSyncTime, tmPjf.iSyncTime(0), tmPjf.iSyncTime(1)
'        ilRet = btrUpdate(hmPjf, tmPjf, imPjfRecLen)
        slMsg = "mRolloverRec (btrUpdate: Projection)"
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmPjf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("File in Use [Re-press Ok], Update" & Str$(ilRet), vbOKOnly + vbExclamation, "Rollover")
            'imTerminate = True
            mRolloverRec = False
            Exit Function
        End If
    Next ilLoop
    'Remove change reason
    slDate = "1/15/" & Trim$(Str$(imCurYear))
    If tgSpf.sRUseCorpCal = "Y" Then
        slStart = gObtainStartCorp(slDate, True)
    Else
        slStart = gObtainStartStd(slDate)
    End If
    ilStartWk = 1
    ilRollForward = False
    For ilLoop = 1 To 12 Step 1
        If tgSpf.sRUseCorpCal = "Y" Then
            slEnd = gObtainEndCorp(slStart, True)
        Else
            slEnd = gObtainEndStd(slStart)
        End If
        ilNoWks = (gDateValue(slEnd) - gDateValue(slStart) + 1) \ 7
        If (imPrjStartWk - 1 >= ilStartWk) And (imPrjStartWk - 1 <= ilStartWk + ilNoWks - 1) Then
            If (imPrjStartWk >= ilStartWk) And (imPrjStartWk <= ilStartWk + ilNoWks - 1) Then
                ilRollForward = True
            End If
            Exit For
        End If
        ilStartWk = ilStartWk + ilNoWks
        slStart = gIncOneDay(slEnd)
    Next ilLoop
    For ilLoop = LBound(lmPjfCode) To UBound(lmPjfCode) - 1 Step 1
        'ilRet = btrGetDirect(hmPjf, tmPjf, imPjfRecLen, lmPjfCode(ilLoop), INDEXKEY0, BTRV_LOCK_NONE)
        tmPjfSrchKey1.lCode = lmPjfCode(ilLoop)
        ilRet = btrGetEqual(hmPjf, tmPjf, imPjfRecLen, tmPjfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmPjf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("File in Use [Re-press Ok], GetDirect" & Str$(ilRet), vbOKOnly + vbExclamation, "Rollover")
            'imTerminate = True
            mRolloverRec = False
            Exit Function
        End If
        'tmRec = tmPjf
        'ilRet = gGetByKeyForUpdate("PJF", hmPjf, tmRec)
        'tmPjf = tmRec
        'If ilRet <> BTRV_ERR_NONE Then
        '    ilCRet = btrAbortTrans(hmPjf)
        '    Screen.MousePointer = vbDefault    'Default
        '    ilRet = MsgBox("File in Use [Re-press Ok], GetByKey" & Str$(ilRet), vbOkOnly + vbExclamation, "Rollover")
        '    'imTerminate = True
        '    mRolloverRec = False
        '    Exit Function
        'End If
        tmPjf.lCxfChgR = 0
        If (imPrjStartWk > 0) And (tmPjf.iYear = imCurYear) Then
            llGross = 0
            For ilWk = 0 To imPrjStartWk - 1 Step 1
                llGross = llGross + tmPjf.lGross(ilWk)
                tmPjf.lGross(ilWk) = 0
            Next ilWk
            If ilRollForward Then
                tmPjf.lGross(imPrjStartWk) = llGross + tmPjf.lGross(imPrjStartWk)
            End If
        End If
        If (tmPjf.iYear < imCurYear) Then
            llGross = 0
            For ilWk = 0 To 53 Step 1
                llGross = llGross + tmPjf.lGross(ilWk)
                tmPjf.lGross(ilWk) = 0
            Next ilWk
        End If
        tmPjf.iSourceID = 0
        gPackDate slSyncDate, tmPjf.iSyncDate(0), tmPjf.iSyncDate(1)
        gPackTime slSyncTime, tmPjf.iSyncTime(0), tmPjf.iSyncTime(1)
        ilRet = btrUpdate(hmPjf, tmPjf, imPjfRecLen)
        slMsg = "mRolloverRec (btrUpdate: Projection)"
        If ilRet <> BTRV_ERR_NONE Then
            ilCRet = btrAbortTrans(hmPjf)
            Screen.MousePointer = vbDefault    'Default
            ilRet = MsgBox("File in Use [Re-press Ok], Update" & Str$(ilRet), vbOKOnly + vbExclamation, "Rollover")
            'imTerminate = True
            mRolloverRec = False
            Exit Function
        End If
        'If (tmPjf.iYear < imCurYear) Then
        '    tmPjfSrchKey.iSlfCode = tmPjf.iSlfCode
        '    tmPjfSrchKey.iRolloverDate(0) = 0
        '    tmPjfSrchKey.iRolloverDate(1) = 0
        '    ilRet = btrGetGreaterOrEqual(hmPjf, tlPjf, imPjfRecLen, tmPjfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        '    Do While (ilRet = BTRV_ERR_NONE) And (tmPjf.iSlfCode = tlPjf.iSlfCode)
        '        If tlPjf.iYear = imCurYear Then
        '            If (tmPjf.lChfCode = tlPjf.lChfCode) And (tmPjf.iSofCode = tlPjf.iSofCode) And (tmPjf.iAdfCode = tlPjf.iAdfCode) And (tmPjf.lPrfCode = tlPjf.lPrfCode) And (tmPjf.iMnfDemo = tlPjf.iMnfDemo) And (tmPjf.iVefCode = tlPjf.iVefCode) And (tmPjf.sNewRet = tlPjf.sNewRet) And (tmPjf.iMnfBus = tlPjf.iMnfBus) Then
        '                tmRec = tlPjf
        '                ilRet = gGetByKeyForUpdate("PJF", hmPjf, tmRec)
        '                tlPjf = tmRec
        '                If ilRet <> BTRV_ERR_NONE Then
        '                    ilRet = btrAbortTrans(hmPjf)
        '                    Screen.MousePointer = vbDefault    'Default
        '                    ilRet = MsgBox("Rollover Stopped, Try Later", vbOkOnly + vbExclamation, "Erase")
        '                    imTerminate = True
        '                    mRolloverRec = False
        '                    Exit Function
        '                End If
        '                tlPjf.lGross(1) = llGross
        '                ilRet = btrUpdate(hmPjf, tlPjf, imPjfRecLen)
        '                Exit Do
        '            End If
        '        End If
        '        ilRet = btrGetNext(hmPjf, tlPjf, imPjfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        '    Loop
        'End If
    Next ilLoop
    ilRet = btrEndTrans(hmPjf)
    'mInitBudgetCtrls
    mRolloverRec = True
    Exit Function
mRolloverErr:
    On Error GoTo 0
    mRolloverRec = False
    Exit Function
End Function
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
    Unload Rollover
    igManUnload = NO
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
        slDay = Trim$(Str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                edcDate.Text = Format$(llDate, "m/d/yy")
                edcDate.SelStart = 0
                edcDate.SelLength = Len(edcDate.Text)
                imBypassFocus = True
                edcDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcDate.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(Str$(imCalMonth)) & "/15/" & Trim$(Str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    plcCalendar.Visible = False
    plcTme.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
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
                    imBypassFocus = True    'Don't change select text
                    edcTime.SetFocus
                    'SendKeys slKey
                    gSendKeys edcTime, slKey
                    Exit Sub
                End If
                flX = flX + fgPadDeltaX
            Next ilColNo
        End If
        flY = flY + fgPadDeltaY
    Next ilRowNo
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
    plcScreen.Print "Rollover"
End Sub
