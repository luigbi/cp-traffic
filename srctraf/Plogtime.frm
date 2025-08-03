VERSION 5.00
Begin VB.Form PLogTime 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3915
   ClientLeft      =   2295
   ClientTop       =   1995
   ClientWidth     =   3390
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3915
   ScaleWidth      =   3390
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
      Left            =   1305
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         Picture         =   "Plogtime.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   14
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
            TabIndex        =   15
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
         TabIndex        =   18
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   1890
      TabIndex        =   11
      Top             =   3390
      Width           =   945
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
      TabIndex        =   12
      Top             =   1770
      Width           =   75
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -15
      Width           =   1920
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   570
      TabIndex        =   10
      Top             =   3390
      Width           =   945
   End
   Begin VB.PictureBox plcTimes 
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
      Height          =   2865
      Left            =   150
      ScaleHeight     =   2805
      ScaleWidth      =   2970
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   270
      Width           =   3030
      Begin VB.TextBox edcAirTime 
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
         Left            =   1140
         MaxLength       =   11
         TabIndex        =   9
         Top             =   2505
         Width           =   1680
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
         Left            =   2640
         Picture         =   "Plogtime.frx":2E1A
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   90
         Width           =   195
      End
      Begin VB.TextBox edcDate 
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
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   90
         Width           =   1485
      End
      Begin VB.TextBox edcTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Left            =   1140
         MaxLength       =   11
         TabIndex        =   6
         Top             =   540
         Width           =   1680
      End
      Begin VB.ListBox lbcTimes 
         Appearance      =   0  'Flat
         Height          =   1500
         Left            =   1140
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   765
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Air Time"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2490
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Avail Time"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   525
         Width           =   1020
      End
      Begin VB.Label lacAvailDate 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   60
         Width           =   1020
      End
   End
End
Attribute VB_Name = "PLogTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Plogtime.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PLogTime.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imNoAvailTimes As Integer
Dim smAvailTimes() As String
Dim imAnfCode() As Integer
Dim imButtonIndex As Integer
Dim smDefaultTime As String
Dim imDateChgMode As Integer
Dim imBSMode As Integer
Dim imBypassFocus As Integer
Dim imVefCode As Integer
Dim imGameNo As Integer
Dim smSchDate As String
'Spot summary file
Dim hmSsf As Integer        'Spot summary file handle
Dim tmSsf As SSF            'SSF record image
Dim tmSsfSrchKey As SSFKEY0 'SSF key record image
Dim imSsfRecLen As Integer     'SSF record length
Dim tmAvail As AVAILSS
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
    sgAvailTime = ""
    sgAvailDate = ""
    sgPLogAirTime = ""
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
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
    Dim slStr As String
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim llDate As Long
    sgAvailDate = ""
    slStr = edcDate.Text
    If Trim$(slStr) <> "" Then
        If gValidDate(slStr) Then
            llDate = gDateValue(slStr)
            ilFound = False
            For ilLoop = 0 To UBound(tgWkDates) - 1 Step 1
                If tgWkDates(ilLoop).lDate = llDate Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                Beep
                edcDate.SetFocus
                Exit Sub
            End If
            sgAvailDate = slStr
        Else
            Beep
            edcDate.SetFocus
            Exit Sub
        End If
    Else
        Beep
        edcDate.SetFocus
        Exit Sub
    End If
    If igPLogSource <> 1 Then
        sgAvailTime = ""
        slStr = edcTime.Text
        If Trim$(slStr) <> "" Then
            If gValidTime(slStr) Then
                sgAvailTime = slStr
            Else
                Beep
                lbcTimes.SetFocus
                Exit Sub
            End If
        Else
            Beep
            lbcTimes.SetFocus
            Exit Sub
        End If
    End If
    sgPLogAirTime = ""
    slStr = edcAirTime.Text
    If Trim$(slStr) <> "" Then
        If gValidTime(slStr) Then
            sgPLogAirTime = slStr
        Else
            Beep
            edcAirTime.SetFocus
            Exit Sub
        End If
    Else
        Beep
        edcAirTime.SetFocus
        Exit Sub
    End If
    If igPLogSource = 1 Then
        sgAvailTime = sgPLogAirTime
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcAirTime_Change()
    mSetCommands
End Sub

Private Sub edcAirTime_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDate_Change()
    Dim slStr As String
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim ilFound As Integer
    slStr = edcDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        mSetCommands
        Exit Sub
    End If
    If imDateChgMode = False Then
        imDateChgMode = True
        lbcTimes.ListIndex = -1
        llDate = gDateValue(slStr)
        ilFound = False
        For ilLoop = 0 To UBound(tgWkDates) - 1 Step 1
            If tgWkDates(ilLoop).lDate = llDate Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            imDateChgMode = False
            lacDate.Visible = False
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        lacDate.Visible = True
        gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
        pbcCalendar_Paint   'mBoxCalDate called within paint
        smSchDate = edcDate.Text
        mAvailTimePop
        Screen.MousePointer = vbDefault    'Default
        imDateChgMode = False
    End If
    mSetCommands
End Sub
Private Sub edcDate_GotFocus()
    If Not imBypassFocus Then
        gCtrlGotFocus edcDate
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
Private Sub edcTime_Change()
    mSetCommands
End Sub
Private Sub edcTime_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus ActiveControl
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
    PLogTime.Refresh
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

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    'sgDoneMsg = CmdStr
    'igChildDone = True
    'Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        mTerminate
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase imAnfCode
    Erase smAvailTimes
    btrDestroy hmSsf
    
    Set PLogTime = Nothing
    
End Sub
Private Sub lbcTimes_Click()
    If lbcTimes.ListIndex >= 0 Then
        edcTime.Text = lbcTimes.List(lbcTimes.ListIndex)
        cmcDone.Enabled = True
    Else
        edcTime.Text = ""
        cmcDone.Enabled = False
    End If
End Sub
Private Sub lbcTimes_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
    plcCalendar.Visible = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailTimePop                   *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate list box with avail   *
'*                      times                          *
'*                                                     *
'*******************************************************
Private Sub mAvailTimePop()
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim slTime As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim ilAnfCode As Integer
    Dim ilAnf As Integer
    ' Spot summary File
    ReDim smAvailTimes(0 To 0) As String
    ReDim imAnfCode(0 To 0) As Integer
    
    ilUpper = 0
    imNoAvailTimes = ilUpper
    gPackDate smSchDate, ilDate0, ilDate1
    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
    tmSsfSrchKey.iType = imGameNo 'slType-On Air
    tmSsfSrchKey.iVefCode = imVefCode
    tmSsfSrchKey.iDate(0) = ilDate0
    tmSsfSrchKey.iDate(1) = ilDate1
    tmSsfSrchKey.iStartTime(0) = 0
    tmSsfSrchKey.iStartTime(1) = 0
    ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = imGameNo) And (tmSsf.iVefCode = imVefCode) And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1)
        For ilLoop = 1 To tmSsf.iCount Step 1
           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilLoop)
            If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                ilAnfCode = tmAvail.ianfCode
                If (ilAnfCode = igAvailAnfCode) Or (igAvailAnfCode = -1) Then
                    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                    smAvailTimes(ilUpper) = slTime
                    imAnfCode(ilUpper) = ilAnfCode
                    ilUpper = ilUpper + 1
                    ReDim Preserve smAvailTimes(0 To ilUpper) As String
                    ReDim Preserve imAnfCode(0 To ilUpper) As Integer
                End If
            End If
        Next ilLoop
        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
        ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    imNoAvailTimes = ilUpper
    lbcTimes.Clear
    For ilLoop = LBound(smAvailTimes) To UBound(smAvailTimes) - 1 Step 1
        lbcTimes.AddItem smAvailTimes(ilLoop)
        lbcTimes.ItemData(lbcTimes.NewIndex) = -1
        For ilAnf = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1 Step 1
            If tgAvailAnf(ilAnf).iCode = imAnfCode(ilLoop) Then
                lbcTimes.ItemData(lbcTimes.NewIndex) = ilAnf
                Exit For
            End If
        Next ilAnf
    Next ilLoop
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
    Dim ilRet As Integer
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    'mParseCmmdLine
    'If imTerminate Then
    '    Exit Sub
    'End If
    'igStdAloneMode = True
    imVefCode = igAvailVefCode
    imGameNo = igAvailGameNo
    smSchDate = sgAvailDate
    smDefaultTime = sgAvailTime
    imButtonIndex = -1
    imBSMode = False
    imCalType = 0
    imBypassFocus = False
    PLogTime.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone PLogTime
    hmSsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'PLogTime.Show
    Screen.MousePointer = vbHourglass
    mInitBox
    edcDate.Text = smSchDate
    lbcTimes.Clear 'Force population
    mAvailTimePop
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imFirstFocus = True
    gFindMatch smDefaultTime, 0, lbcTimes
    If gLastFound(lbcTimes) >= 0 Then
        lbcTimes.ListIndex = gLastFound(lbcTimes)
        cmcDone.Enabled = True
    Else
        'Exact time must be specified so spot can be associated with an avail
        lbcTimes.ListIndex = -1
        edcTime.Enabled = False
        cmcDone.Enabled = False
    End If
    If imTerminate Then
        Exit Sub
    End If
    edcAirTime.Text = smDefaultTime
    If igPLogSource = 1 Then
        edcTime.Text = smDefaultTime
        edcTime.Visible = False
        Label1.Visible = False
        lbcTimes.Visible = False
    Else
        edcTime.Visible = True
        Label1.Visible = True
        lbcTimes.Visible = True
    End If
'    gCenterModalForm PLogTime
    Screen.MousePointer = vbDefault
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:2/28/94       By:D. LeVine      *
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
    plcCalendar.Move plcTimes.Left + fgBevelX, plcTimes.Top + edcDate.Height + fgBevelY
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

    igManUnload = YES
    Unload PLogTime
    igManUnload = NO
End Sub

Private Sub lbcTimes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilAnf As Integer
    
    If (X < 0) Or (X > lbcTimes.Width) Then
        lbcTimes.ToolTipText = ""
        imButtonIndex = -1
        Exit Sub
    End If
    If imButtonIndex <> (Y \ fgListHtArial825) + lbcTimes.TopIndex Then
        imButtonIndex = Y \ fgListHtArial825 + lbcTimes.TopIndex
        If (imButtonIndex >= 0) And (imButtonIndex <= lbcTimes.ListCount - 1) Then
            ilAnf = lbcTimes.ItemData(imButtonIndex)
            If ilAnf <> -1 Then
                lbcTimes.ToolTipText = Trim$(tgAvailAnf(ilAnf).sName)
            Else
                lbcTimes.ToolTipText = ""
            End If
        Else
            imButtonIndex = -1
            lbcTimes.ToolTipText = ""
        End If
    End If
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
    gPLPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate, tgWkDates()
    mBoxCalDate
End Sub
Private Sub pbcClickFocus_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
    plcCalendar.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        ''Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        ''Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        ''Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Avail Date/Times"
End Sub

Private Sub mSetCommands()
    Dim slStr As String
    cmcDone.Enabled = False
    If igPLogSource = 1 Then
        slStr = edcAirTime.Text
        If Trim$(slStr) <> "" Then
            If gValidTime(slStr) Then
                slStr = edcDate.Text
                If gValidDate(slStr) Then
                    cmcDone.Enabled = True
                End If
            End If
        End If
    Else
        slStr = edcTime.Text
        If Trim$(slStr) <> "" Then
            If gValidTime(slStr) Then
                slStr = edcAirTime.Text
                If Trim$(slStr) <> "" Then
                    If gValidTime(slStr) Then
                        slStr = edcDate.Text
                        If gValidDate(slStr) Then
                            cmcDone.Enabled = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub
