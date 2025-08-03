VERSION 5.00
Begin VB.Form SlspCrte 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4545
   ClientLeft      =   1170
   ClientTop       =   2115
   ClientWidth     =   7080
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
   ScaleHeight     =   4545
   ScaleWidth      =   7080
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   1425
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   840
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
         Picture         =   "Slspcrte.frx":0000
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
      Left            =   3825
      TabIndex        =   11
      Top             =   4230
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "C&reate"
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
      Left            =   2205
      TabIndex        =   10
      Top             =   4230
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
      ScaleWidth      =   1770
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   1770
   End
   Begin VB.PictureBox plcDates 
      ForeColor       =   &H00000000&
      Height          =   3720
      Left            =   180
      ScaleHeight     =   3660
      ScaleWidth      =   6660
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   6720
      Begin VB.TextBox edcPct 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1320
         Width           =   930
      End
      Begin VB.TextBox edcPct 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   6
         Top             =   795
         Width           =   930
      End
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3180
         Left            =   3420
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   285
         Width           =   3105
      End
      Begin VB.TextBox edcDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1230
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
         Left            =   2160
         Picture         =   "Slspcrte.frx":2E1A
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   300
         Width           =   195
      End
      Begin VB.Label lacPct 
         Appearance      =   0  'Flat
         Caption         =   "Under Remnant %"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1305
         Width           =   1560
      End
      Begin VB.Label lacPct 
         Appearance      =   0  'Flat
         Caption         =   "Under Goal %"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   795
         Width           =   1290
      End
      Begin VB.Label lacStartDate 
         Appearance      =   0  'Flat
         Caption         =   "Start Month"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1110
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   345
      Top             =   4125
      Width           =   360
   End
End
Attribute VB_Name = "SlspCrte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Slspcrte.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SlspCrte.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Salesperson Model library dates input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imBypassFocus As Integer
Dim lmNowDate As Long
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
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcDate.SelStart = 0
    edcDate.SelLength = Len(edcDate.Text)
    edcDate.SetFocus
End Sub
Private Sub cmcDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_Click()
    Dim ilRet As Integer
    Dim slDate As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slStr As String
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    slDate = Trim$(edcDate.Text)
    If slDate = "" Then
        ilRet = MsgBox("Start month must be specified", vbOKOnly + vbExclamation, "Incomplete")
        edcDate.SetFocus
        Exit Sub
    Else
        If Not gValidDate(slDate) Then
            ilRet = MsgBox("Start month must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            edcDate.SetFocus
            Exit Sub
        End If
    End If
    slDate = gObtainStartStd(slDate)
    Screen.MousePointer = vbHourglass
    ReDim tgScfAdd(0 To 0) As SCF
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            ilUpper = UBound(tgScfAdd)
            slNameCode = tgUserVehicle(ilLoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
            tgScfAdd(ilUpper).iVefCode = Val(slCode)
            gPackDate slDate, tgScfAdd(ilUpper).iStartDate(0), tgScfAdd(ilUpper).iStartDate(1)
            gPackDate "", tgScfAdd(ilUpper).iEndDate(0), tgScfAdd(ilUpper).iEndDate(1)
            slStr = edcPct(0).Text
            tgScfAdd(ilUpper).iUnderComm = gStrDecToInt(slStr, 2)
            slStr = edcPct(1).Text
            tgScfAdd(ilUpper).iRemUnderComm = gStrDecToInt(slStr, 2)
            ReDim Preserve tgScfAdd(0 To ilUpper + 1) As SCF
        End If
    Next ilLoop
    Screen.MousePointer = vbDefault
    igModReturn = 1
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    plcCalendar.Visible = False
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
Private Sub edcPct_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcPct_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcPct_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim ilKeyAscii As Integer
    Dim ilPos As Integer
    Dim slStr As String
    ilKeyAscii = KeyAscii
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcPct(Index).SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    ilPos = InStr(edcPct(Index).SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcPct(Index).Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcPct(Index).Text
    slStr = Left$(slStr, edcPct(Index).SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcPct(Index).SelStart - edcPct(Index).SelLength)
    If gCompNumberStr(slStr, "100.00") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
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
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SlspCrte = Nothing   'Remove data segment
End Sub

Private Sub imcHelp_Click()
    plcCalendar.Visible = False
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcVehicle_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
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
    Dim slStr As String
    Dim ilRet As Integer
    Dim slDate As String
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imFirstActivate = True
    imTerminate = False
    igModReturn = 0
    imBypassFocus = False
    imSettingValue = False
    imChgMode = False
    imBSMode = False
    imCalType = 0   'Standard
    mInitBox
    SlspCrte.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone SlspCrte
    plcCalendar.Move plcDates.Left + edcDate.Left, plcDates.Top + edcDate.Top + edcDate.Height
    slStr = Format$(gNow(), "m/d/yy")
    slStr = gObtainStartStd(slStr)    'Get tuesday of the current week
    edcDate.Text = slStr
    ilRet = mVefPop()
    'gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    'pbcCalendar_Paint   'mBoxCalDate called within paint
    'lacDate.Visible = False
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    slDate = gObtainEndStd(slDate)
    Exit Sub

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
    Unload SlspCrte
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVefPop                         *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mVefPop() As Integer
    Dim ilRet As Integer
    ilRet = gPopUserVehicleBox(SlspCrte, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + ACTIVEVEH + DORMANTVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", SlspCrte
        On Error GoTo 0
    End If
    Exit Function
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Function
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
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
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
    plcScreen.Print "Commission: Create"
End Sub
