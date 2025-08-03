VERSION 5.00
Begin VB.Form SlspMod 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4650
   ClientLeft      =   1170
   ClientTop       =   2115
   ClientWidth     =   7530
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
   ScaleHeight     =   4650
   ScaleWidth      =   7530
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   555
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   930
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
         Picture         =   "Slspmod.frx":0000
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
         Left            =   315
         TabIndex        =   11
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
      TabIndex        =   15
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
      Left            =   4035
      TabIndex        =   8
      Top             =   4305
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Model"
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
      Left            =   2415
      TabIndex        =   7
      Top             =   4305
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
      ScaleWidth      =   1710
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   1710
   End
   Begin VB.PictureBox plcDates 
      ForeColor       =   &H00000000&
      Height          =   3810
      Left            =   180
      ScaleHeight     =   3750
      ScaleWidth      =   7170
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   7230
      Begin VB.CheckBox ckcAll 
         Caption         =   "All Vehicles"
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
         Height          =   225
         Left            =   4725
         TabIndex        =   16
         Top             =   105
         Width           =   1335
      End
      Begin VB.ListBox lbcSalesperson 
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
         Left            =   2490
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   390
         Width           =   2100
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
         Left            =   4725
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   390
         Width           =   2295
      End
      Begin VB.TextBox edcDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   3
         Top             =   390
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
         Picture         =   "Slspmod.frx":2E1A
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   390
         Width           =   195
      End
      Begin VB.Label lacStartDate 
         Appearance      =   0  'Flat
         Caption         =   "Start Month"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   135
         TabIndex        =   2
         Top             =   390
         Width           =   1110
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   345
      Top             =   4200
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "SlspMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Slspmod.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SlspMod.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Salesperson Model library dates input screen code
Option Explicit
Option Compare Text
Dim tmScf As SCF        'Scf record image
Dim hmScf As Integer    'Projection file handle
Dim tmScfSrchKey1 As INTKEY0    'Scf key record image
Dim imScfRecLen As Integer        'Scf record length
Dim imSlfCode() As Integer
Dim smFillDate As String
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imBypassFocus As Integer
Dim lmNowDate As Long
Dim imBypassAll As Integer
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
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
    If imBypassAll Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilValue = Value
    If lbcVehicle.ListCount > 0 Then
        llRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehicle.hwnd, LB_SELITEMRANGE, ilValue, llRg)
    End If
    Screen.MousePointer = vbDefault
End Sub
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
    Screen.MousePointer = vbHourglass
    ilRet = mReadRec()
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
    lbcSalesperson.Clear
    lbcVehicle.Clear
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
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase imSlfCode
    btrExtClear hmScf   'Clear any previous extend operation
    ilRet = btrClose(hmScf)
    btrDestroy hmScf
    
    Set SlspMod = Nothing   'Remove data segment
    
End Sub

Private Sub imcHelp_Click()
    plcCalendar.Visible = False
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcSalesperson_Click()
    Dim ilRet As Integer
    lbcVehicle.Clear
    ilRet = mVefPop()
End Sub
Private Sub lbcSalesperson_GotFocus()
    Dim ilRet As Integer
    Dim slStr As String
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
    slStr = edcDate.Text
    If gValidDate(slStr) Then
        If smFillDate <> "" Then
            If gDateValue(slStr) <> gDateValue(smFillDate) Then
                ilRet = mSlspPop()
            End If
        Else
            ilRet = mSlspPop()
        End If
    End If
End Sub
Private Sub lbcVehicle_Click()
    imBypassAll = True
    ckcAll.Value = vbUnchecked
    imBypassAll = False
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
    smFillDate = ""
    imBypassAll = False
    mInitBox
    hmScf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmScf, "", sgDBPath & "Scf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Scf.Btr)", SlspMod
    On Error GoTo 0
    imScfRecLen = Len(tmScf)
    SlspMod.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone SlspMod
    plcCalendar.Move plcDates.Left + edcDate.Left, plcDates.Top + edcDate.Top + edcDate.Height
    slStr = Format$(gNow(), "m/d/yy")
    slStr = gObtainStartStd(slStr)    'Get tuesday of the current week
    edcDate.Text = slStr
    'gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    'pbcCalendar_Paint   'mBoxCalDate called within paint
    'lacDate.Visible = False
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    slDate = gObtainEndStd(slDate)
    ilRet = mSlspPop()
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
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec() As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilSlfCode As Integer
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim llEDate As Long
    Dim llRecPos As Long
    Dim ilVef As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    ReDim tgScfAdd(0 To 0) As SCF
    ilUpper = 0
    If lbcSalesperson.ListIndex < 0 Then
        mReadRec = False
        Exit Function
    End If
    slDate = Trim$(edcDate.Text)
    If Not gValidDate(slDate) Then
        mReadRec = False
        Exit Function
    End If
    llDate = gDateValue(slDate)
    btrExtClear hmScf   'Clear any previous extend operation
    ilExtLen = Len(tmScf)  'Extract operation record size
    For ilLoop = LBound(tmSlspCommSalesperson) To UBound(tmSlspCommSalesperson) - 1 Step 1
        slNameCode = tmSlspCommSalesperson(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If StrComp(Trim$(slName), lbcSalesperson.List(lbcSalesperson.ListIndex), 1) = 0 Then
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilSlfCode = Val(slCode)
            Exit For
        End If
    Next ilLoop
    tmScfSrchKey1.iCode = ilSlfCode
    ilRet = btrGetEqual(hmScf, tmScf, imScfRecLen, tmScfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
        ilRet = BTRV_ERR_END_OF_FILE
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mReadRec = False
            Exit Function
        End If
    End If
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmScf, llNoRec, -1, "UC", "SCF", "") '"EG") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("SCF", "SCFSLFCODE") 'GetOffSetForInt(tmScf, tmScf.iSlfCode)
        tlIntTypeBuff.iType = ilSlfCode
        ilRet = btrExtAddLogicConst(hmScf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mReadRec (btrExtAddLogicConst):" & "Scf.Btr", SlspMod
        On Error GoTo 0
        ilRet = btrExtAddField(hmScf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mReadRec (btrExtAddField):" & "Scf.Btr", SlspMod
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
        ilRet = btrExtGetNext(hmScf, tgScfAdd(ilUpper), ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mReadRecErr
            gBtrvErrorMsg ilRet, "mReadRec (btrExtGetNextExt):" & "Scf.Btr", SlspMod
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmRpf, tlRpfExt, ilExtLen, llRecPos)
            ilExtLen = Len(tmScf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmScf, tgScfAdd(ilUpper), ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                gUnpackDateLong tgScfAdd(ilUpper).iEndDate(0), tgScfAdd(ilUpper).iEndDate(1), llEDate
                If (llEDate = 0) Or (llDate <= llEDate) Then
                    'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If tgMVef(ilLoop).iCode = tgScfAdd(ilUpper).iVefCode Then
                        ilLoop = gBinarySearchVef(tgScfAdd(ilUpper).iVefCode)
                        If ilLoop <> -1 Then
                            For ilVef = 0 To lbcVehicle.ListCount - 1 Step 1
                                If lbcVehicle.Selected(ilVef) Then
                                    If StrComp(Trim$(tgMVef(ilLoop).sName), Trim$(lbcVehicle.List(ilVef)), 1) = 0 Then
                                        ilUpper = ilUpper + 1
                                        ReDim Preserve tgScfAdd(0 To ilUpper) As SCF
                                        Exit For
                                    End If
                                End If
                            Next ilVef
                    '        Exit For
                        End If
                    'Next ilLoop
                End If
                ilRet = btrExtGetNext(hmScf, tgScfAdd(ilUpper), ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmScf, tgScfAdd(ilUpper), ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSlspPop                        *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mSlspPop() As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim llEDate As Long
    Dim llRecPos As Long
    Dim ilFound As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilSlfCode As Integer
    Dim ilSlf As Integer
    ReDim imSlfCode(0 To 0) As Integer
    lbcSalesperson.Clear
    lbcVehicle.Clear
    ilUpper = 0
    slDate = Trim$(edcDate.Text)
    If Not gValidDate(slDate) Then
        mSlspPop = False
        Exit Function
    End If
    llDate = gDateValue(slDate)
    btrExtClear hmScf   'Clear any previous extend operation
    ilExtLen = Len(tmScf)  'Extract operation record size
    ilRet = btrGetFirst(hmScf, tmScf, imScfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
        ilRet = BTRV_ERR_END_OF_FILE
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mSlspPop = False
            Exit Function
        End If
    End If
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmScf, llNoRec, -1, "UC", "SCF", "") '"EG") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("SCF", "SCFSLFCODE") 'GetOffSetForInt(tmScf, tmScf.iSlfCode)
        ilRet = btrExtAddField(hmScf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mSlspPopErr
        gBtrvErrorMsg ilRet, "mSlspPop (btrExtAddField):" & "Scf.Btr", SlspMod
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
        ilRet = btrExtGetNext(hmScf, tmScf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mSlspPopErr
            gBtrvErrorMsg ilRet, "mSlspPop (btrExtGetNextExt):" & "Scf.Btr", SlspMod
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmRpf, tlRpfExt, ilExtLen, llRecPos)
            ilExtLen = Len(tmScf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmScf, tmScf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE

                gUnpackDateLong tmScf.iEndDate(0), tmScf.iEndDate(1), llEDate
                If (llEDate = 0) Or (llDate <= llEDate) Then
                    ilFound = False
                    For ilLoop = LBound(imSlfCode) To UBound(imSlfCode) - 1 Step 1
                        If tmScf.iSlfCode = imSlfCode(ilLoop) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        imSlfCode(ilUpper) = tmScf.iSlfCode
                        ilUpper = ilUpper + 1
                        ReDim Preserve imSlfCode(0 To ilUpper) As Integer
                    End If
                End If
                ilRet = btrExtGetNext(hmScf, tmScf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmScf, tmScf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    For ilLoop = LBound(tmSlspCommSalesperson) To UBound(tmSlspCommSalesperson) - 1 Step 1
        slNameCode = tmSlspCommSalesperson(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilSlfCode = Val(slCode)
        For ilSlf = LBound(imSlfCode) To UBound(imSlfCode) - 1 Step 1
            If ilSlfCode = imSlfCode(ilSlf) Then
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                lbcSalesperson.AddItem Trim$(slName)
                Exit For
            End If
        Next ilSlf
    Next ilLoop
    smFillDate = slDate
    mSlspPop = True
    Exit Function
mSlspPopErr:
    On Error GoTo 0
    mSlspPop = False
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
    Unload SlspMod
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
    Dim ilExtLen As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilSlfCode As Integer
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim llEDate As Long
    Dim llRecPos As Long
    Dim ilFound As Integer
    Dim ilVef As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    ReDim imVefCode(0 To 0) As Integer
    lbcVehicle.Clear
    ckcAll.Value = vbUnchecked
    ilUpper = 0
    If lbcSalesperson.ListIndex < 0 Then
        mVefPop = False
        Exit Function
    End If
    slDate = Trim$(edcDate.Text)
    If Not gValidDate(slDate) Then
        mVefPop = False
        Exit Function
    End If
    llDate = gDateValue(slDate)
    btrExtClear hmScf   'Clear any previous extend operation
    ilExtLen = Len(tmScf)  'Extract operation record size
    For ilLoop = LBound(tmSlspCommSalesperson) To UBound(tmSlspCommSalesperson) - 1 Step 1
        slNameCode = tmSlspCommSalesperson(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If StrComp(Trim$(slName), lbcSalesperson.List(lbcSalesperson.ListIndex), 1) = 0 Then
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilSlfCode = Val(slCode)
            Exit For
        End If
    Next ilLoop
    tmScfSrchKey1.iCode = ilSlfCode
    ilRet = btrGetEqual(hmScf, tmScf, imScfRecLen, tmScfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
        ilRet = BTRV_ERR_END_OF_FILE
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mVefPop = False
            Exit Function
        End If
    End If
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmScf, llNoRec, -1, "UC", "SCF", "") '"EG") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("SCF", "SCFSLFCODE") 'GetOffSetForInt(tmScf, tmScf.iSlfCode)
        tlIntTypeBuff.iType = ilSlfCode
        ilRet = btrExtAddLogicConst(hmScf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        On Error GoTo mVefPopErr
        gBtrvErrorMsg ilRet, "mVefPop (btrExtAddLogicConst):" & "Scf.Btr", SlspMod
        On Error GoTo 0
        ilRet = btrExtAddField(hmScf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mVefPopErr
        gBtrvErrorMsg ilRet, "mVefPop (btrExtAddField):" & "Scf.Btr", SlspMod
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
        ilRet = btrExtGetNext(hmScf, tmScf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mVefPopErr
            gBtrvErrorMsg ilRet, "mVefPop (btrExtGetNextExt):" & "Scf.Btr", SlspMod
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmRpf, tlRpfExt, ilExtLen, llRecPos)
            ilExtLen = Len(tmScf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmScf, tmScf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                gUnpackDateLong tmScf.iEndDate(0), tmScf.iEndDate(1), llEDate
                If (llEDate = 0) Or (llDate <= llEDate) Then
                    ilFound = False
                    For ilLoop = LBound(imVefCode) To UBound(imVefCode) - 1 Step 1
                        If tmScf.iVefCode = imVefCode(ilLoop) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        imVefCode(ilUpper) = tmScf.iVefCode
                        ilUpper = ilUpper + 1
                        ReDim Preserve imVefCode(0 To ilUpper) As Integer
                    End If
                End If
                ilRet = btrExtGetNext(hmScf, tmScf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmScf, tmScf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        For ilVef = LBound(imVefCode) To UBound(imVefCode) - 1 Step 1
   '         If tgMVef(ilLoop).iCode = imVefCode(ilVef) Then
            ilLoop = gBinarySearchVef(imVefCode(ilVef))
            If ilLoop <> -1 Then
                lbcVehicle.AddItem Trim$(tgMVef(ilLoop).sName)
   '             Exit For
            End If
        Next ilVef
    'Next ilLoop
    ckcAll.Value = vbChecked
    mVefPop = True
    Exit Function
mVefPopErr:
    On Error GoTo 0
    mVefPop = False
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
    plcScreen.Print "Commission: Model"
End Sub
