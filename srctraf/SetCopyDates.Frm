VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form SetCopyDates 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5040
   ClientLeft      =   1170
   ClientTop       =   2115
   ClientWidth     =   3870
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
   ScaleHeight     =   5040
   ScaleWidth      =   3870
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   1440
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   855
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
         Picture         =   "SetCopyDates.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   11
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
            TabIndex        =   12
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
         TabIndex        =   8
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
         TabIndex        =   10
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
         TabIndex        =   9
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
      TabIndex        =   13
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
      Left            =   2190
      TabIndex        =   6
      Top             =   4605
      Width           =   1050
   End
   Begin VB.CommandButton cmcSet 
      Appearance      =   0  'Flat
      Caption         =   "Set"
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
      Left            =   585
      TabIndex        =   5
      Top             =   4605
      Width           =   1050
   End
   Begin VB.PictureBox plcDates 
      ForeColor       =   &H00000000&
      Height          =   2865
      Left            =   180
      ScaleHeight     =   2805
      ScaleWidth      =   3420
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   3480
      Begin VB.CheckBox ckcLog 
         Caption         =   "Log Changes to file"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2490
         Width           =   2070
      End
      Begin VB.ListBox lbcMediaCode 
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
         Height          =   1500
         ItemData        =   "SetCopyDates.frx":2E1A
         Left            =   1215
         List            =   "SetCopyDates.frx":2E1C
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   870
         Width           =   1185
      End
      Begin VB.TextBox edcDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   2
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
         Picture         =   "SetCopyDates.frx":2E1E
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   300
         Width           =   195
      End
      Begin VB.Label lacMediaCode 
         Appearance      =   0  'Flat
         Caption         =   "Media Code"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   855
         Width           =   1110
      End
      Begin VB.Label lacStartDate 
         Appearance      =   0  'Flat
         Caption         =   "Active Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1110
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSteps 
      Height          =   1080
      Left            =   180
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3300
      Visible         =   0   'False
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   1905
      _Version        =   393216
      Rows            =   4
      Cols            =   3
      FixedCols       =   0
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label lacSetCopyDates 
      Caption         =   "Set Copy Inventory Dates"
      Height          =   225
      Left            =   75
      TabIndex        =   15
      Top             =   15
      Width           =   2745
   End
End
Attribute VB_Name = "SetCopyDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SetCopyDates.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SetCopyDates.Frm
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
Dim bmSettingDates As Boolean
Dim lmChgCount As Long
Dim lmTotalCount As Long

'Copy Inventory
Dim hmCif As Integer        'Contract file handle
Dim imCifRecLen As Integer  'CHF record length
Dim tmCif As CIF            'CHF record image
Dim tmCifSrchKey0 As LONGKEY0
Dim tmCifSrchKey1 As CIFKEY1

'Copy rotation
Dim hmCrf As Integer        'Copy rotation file handle
Dim tmCrf As CRF            'CRF record image
Dim tmCrfSrchKey1 As CRFKEY1 'CRF key record image
Dim imCrfRecLen As Integer     'CRF record length

'Instruction
Dim tmCnf As CNF            'CNF record image
Dim tmCnfSrchKey As CNFKEY0  'CNF key record image
Dim hmCnf As Integer        'CNF Handle
Dim imCnfRecLen As Integer      'CNF record length

'Instruction
Dim hmCuf As Integer        'CNF Handle

'Media Code
Dim tmMcf As MCF            'MCF record image
Dim tmMcfSrchKey As INTKEY0  'MCF key record image
Dim hmMcf As Integer        'MCF Handle
Dim imMcfRecLen As Integer      'MCF record length
Dim imMcfCode As Integer

'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer

Dim tmCifInfo() As CIFINFO
Dim tmCrfDates() As CRFDATES
Dim tmDateSort() As DATESORT

Private Sub ckcLog_Click()
    mResetControls
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
    If bmSettingDates Then
        imTerminate = True
        igSetCopyTerminated = True
        Exit Sub
    End If
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
Private Sub cmcSet_Click()
    Dim ilRet As Integer
    Dim slDate As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slStr As String
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim slMediaCode As String
    Dim hlRlf As Integer
    Dim lmLock1RecCode As Integer
    Dim slTempDate As String
    
    If bmSettingDates Then
        Exit Sub
    End If
    slDate = Trim$(edcDate.Text)
    If slDate = "" Then
        ilRet = MsgBox("Active Date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        edcDate.SetFocus
        Exit Sub
    Else
        If Not gValidDate(slDate) Then
            ilRet = MsgBox("Start month must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            edcDate.SetFocus
            Exit Sub
        End If
    End If
    ' JD ttp 10815 - It was found if the user enters a date > today, it will cause there to
    ' be blank rot end dates.
    If Format(slDate, "yyyy-mm-dd") > Format(gNow(), "yyyy-mm-dd") Then
        ilRet = MsgBox("Active Date cannot be in the future", vbOKOnly + vbExclamation, "Incomplete")
        edcDate.SetFocus
        Exit Sub
    End If
    
    If lbcMediaCode.ListIndex < 0 Then
        ilRet = MsgBox("Media Code must be selected", vbOKOnly + vbExclamation, "Incomplete")
        Exit Sub
    End If
    slMediaCode = Trim$(lbcMediaCode.List(lbcMediaCode.ListIndex))
    imMcfCode = lbcMediaCode.ItemData(lbcMediaCode.ListIndex)
    Screen.MousePointer = vbHourglass
    grdSteps.Visible = True
    bmSettingDates = True
    
    
    ' JD 08-11-23 added a lock before doing this operation.
    'Record Locks
    On Error GoTo Err1
    hlRlf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlRlf, "", sgDBPath & "Rlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    
    lmLock1RecCode = gCreateLockRec(hlRlf, "Y", "S", 2, True, "")
    If lmLock1RecCode = 0 Then
        ilRet = MsgBox("Unable to set copy dates at this time. Try again in 10 seconds as another user is currently saving.", vbOKOnly + vbInformation, "Block")
        GoTo Finally
    End If
    
    ReDim tmCifInfo(0 To 0) As CIFINFO
    tmCifInfo(0).lFirst = -1
    ReDim tmCrfDates(0 To 0) As CRFDATES
    tmCrfDates(0).lNext = -1
    lmChgCount = 0
    lmTotalCount = 0
    DoEvents
    grdSteps.TextMatrix(1, 2) = "Started"
    gLogMsg "Gathering Copy Inventory for Media Code " & slMediaCode & ": Started", "SetCopyInventoryDates.txt", False
    DoEvents
    gBuildCifArray hmCif, imMcfCode, tmCifInfo()
    DoEvents
    If Not imTerminate Then
        gLogMsg "Gathering Copy Inventory: Completed", "SetCopyInventoryDates.txt", False
        grdSteps.TextMatrix(1, 2) = "Completed"
        grdSteps.TextMatrix(2, 2) = "Started"
        DoEvents
        gLogMsg "Gathering Rotations active as of " & slDate & ": Started", "SetCopyInventoryDates.txt", False
        gBuildRotDates hmCrf, hmCnf, slDate, tmCifInfo(), tmCrfDates()
    End If
    
    If Not imTerminate Then
        gLogMsg "Gathering Rotations: Completed", "SetCopyInventoryDates.txt", False
        grdSteps.TextMatrix(2, 2) = "Completed"
        grdSteps.TextMatrix(3, 2) = "Started"
        DoEvents
        gLogMsg "Create Copy Usage: Started", "SetCopyInventoryDates.txt", False
        gCreateCufDates hmCif, hmCuf, tmCifInfo(), tmCrfDates()
        gLogMsg "Create Copy Usage: Completed", "SetCopyInventoryDates.txt", False
        gLogMsg "Updating Copy Inventory: Started", "SetCopyInventoryDates.txt", False
        If ckcLog.Value = vbChecked Then
            gUpdateCifDates hmCif, tmCifInfo(), tmCrfDates(), lmChgCount, lmTotalCount, True
        Else
            gUpdateCifDates hmCif, tmCifInfo(), tmCrfDates(), lmChgCount, lmTotalCount, False
        End If
    End If
    DoEvents
    If Not imTerminate Then
        gLogMsg "Updating Copy Inventory: Completed", "SetCopyInventoryDates.txt", False
        gLogMsg "Updated " & lmChgCount & " Copy Inventory dates, out of " & lmTotalCount & " pieces of Copy", "SetCopyInventoryDates.txt", False
        grdSteps.TextMatrix(3, 2) = "Completed"
    Else
        gLogMsg "Set Copy Inventory Dates: Terminated by User", "SetCopyInventoryDates.txt", False
    End If
    cmcCancel.Caption = "&Done"
    bmSettingDates = False
    Screen.MousePointer = vbDefault
    cmcSet.Enabled = False

Finally:
    If lmLock1RecCode <> 0 Then
        ilRet = gDeleteLockRec_ByType(hlRlf, "Y", 2)
    End If
    ilRet = btrClose(hlRlf)
    btrDestroy hlRlf
    Exit Sub
    
Err1:
    If lmLock1RecCode <> 0 Then
        GoTo Finally
    End If
End Sub
Private Sub cmcSet_GotFocus()
    plcCalendar.Visible = False
    gCtrlGotFocus cmcSet
End Sub
Private Sub cmcSet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    mResetControls
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
    If (KeyCode = KEYUP) Or (KeyCode = KeyDown) Then
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

Private Sub Form_Terminate()
    Dim ilRet As Integer
    
    On Error Resume Next
    Erase tmCifInfo
    Erase tmCrfDates
    Erase tmDateSort
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmCnf)
    btrDestroy hmCnf
    ilRet = btrClose(hmCuf)
    btrDestroy hmCuf
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    Set SetCopyDates = Nothing   'Remove data segment
End Sub

Private Sub grdSteps_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub lbcMediaCode_Click()
    mResetControls
End Sub

Private Sub lbcMediaCode_GotFocus()
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
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imFirstActivate = True
    imTerminate = False
    igSetCopyTerminated = False
    igModReturn = 0
    imBypassFocus = False
    imSettingValue = False
    imChgMode = False
    imBSMode = False
    bmSettingDates = False
    imCalType = 0   'Standard
    mInitBox
    ReDim tmCifInfo(0 To 0) As CIFINFO
    tmCifInfo(0).lFirst = -1
    ReDim tmCrfDates(0 To 0) As CRFDATES
    tmCrfDates(0).lNext = -1
    ReDim tmDateSort(0 To 0) As DATESORT
    SetCopyDates.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone SetCopyDates
    plcCalendar.Move plcDates.Left + edcDate.Left, plcDates.Top + edcDate.Top + edcDate.Height
    slStr = Format$(gNow(), "m/d/yy")
    slStr = gObtainStartStd(slStr)    'Get tuesday of the current week
    edcDate.Text = slStr
    
    hmCif = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", SetCopyDates
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    
    hmCrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Crf.Btr)", SetCopyDates
    On Error GoTo 0
    imCrfRecLen = Len(tmCrf)
    
    hmCnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cnf.Btr)", SetCopyDates
    On Error GoTo 0
    imCnfRecLen = Len(tmCnf)
    
    hmCuf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmCuf, "", sgDBPath & "Cuf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cuf.Btr)", SetCopyDates
    On Error GoTo 0
    
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: MCF.BTR)", SetCopyDates
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)
    mMediaCopyPop
    mGridColumnWidth
    mGridColumnTitles
    'gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    'pbcCalendar_Paint   'mBoxCalDate called within paint
    'lacDate.Visible = False
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
    Dim ilRet As Integer
    
    Screen.MousePointer = vbDefault
    On Error Resume Next
    
    igManUnload = YES
    Unload SetCopyDates
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMediaCopyPop                   *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Sub mMediaCopyPop()
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilAdd As Integer
    
    lbcMediaCode.Clear
    ilRet = btrGetFirst(hmMcf, tmMcf, imMcfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record
    Do While ilRet = BTRV_ERR_NONE
        ilAdd = True
        If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) = MEDIACODEBYVEH Then
            If tmMcf.iVefCode > 0 Then
                ilVef = gBinarySearchVef(tmMcf.iVefCode)
                If ilVef = -1 Then
                    ilAdd = False
                End If
            End If
        End If
        '12/8/12: Check all media codes
        'If (ilAdd) And (tmMcf.sReuse = "Y") Then
        If (ilAdd) Then
            lbcMediaCode.AddItem Trim$(tmMcf.sName)
            lbcMediaCode.ItemData(lbcMediaCode.NewIndex) = tmMcf.iCode
        End If
        ilRet = btrGetNext(hmMcf, tmMcf, imMcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    Loop
    If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) <> MEDIACODEBYVEH Then
        If (tgSpf.sUseCartNo = "N") Or (tgSpf.sUseCartNo = "B") Then
            lbcMediaCode.AddItem "[ISCI Only]"
            lbcMediaCode.ItemData(lbcMediaCode.NewIndex) = 0
        End If
    End If
    If lbcMediaCode.ListCount = 1 Then
        lbcMediaCode.ListIndex = 0
    End If
    Exit Sub
mMediaCopyPopErr:
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

Private Sub mBuildCifArray()
    Dim ilRet As Integer
    Dim llCif As Long
    
    ReDim tmCifInfo(0 To 0) As CIFINFO
    tmCifInfo(0).lFirst = -1
    
    tmCifSrchKey1.iMcfCode = imMcfCode
    tmCifSrchKey1.sName = ""
    tmCifSrchKey1.sCut = ""
    ilRet = btrGetGreaterOrEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmCif.iMcfCode = imMcfCode)
        If imTerminate Then
            Exit Sub
        End If
        If (tmCif.sPurged = "A") Then
            llCif = UBound(tmCifInfo)
            tmCifInfo(llCif).lCifCode = tmCif.lCode
            tmCifInfo(llCif).iAdfCode = tmCif.iAdfCode
            tmCifInfo(llCif).lFirst = -1
            ReDim Preserve tmCifInfo(0 To llCif + 1) As CIFINFO
        End If
        ilRet = btrGetNext(hmCif, tmCif, imCifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    llCif = UBound(tmCifInfo)
    If llCif > 0 Then
        ArraySortTyp fnAV(tmCifInfo(), 0), UBound(tmCifInfo), 0, LenB(tmCifInfo(0)), 0, -2, 0
    End If

End Sub

Private Sub mBuildRotDates(slActiveDate As String)
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim slType As String
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim llRecPos As Long
    Dim llCrf As Long
    Dim llNext As Long
    Dim llCif As Long
    Dim blFound As Boolean
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record


    btrExtClear hmCrf   'Clear any previous extend operation
    ilExtLen = Len(tmCrf)  'Extract operation record size
    slType = ""
    tmCrfSrchKey1.sRotType = slType
    tmCrfSrchKey1.iEtfCode = 0
    tmCrfSrchKey1.iEnfCode = 0
    tmCrfSrchKey1.iAdfCode = 0
    tmCrfSrchKey1.lChfCode = 0
    tmCrfSrchKey1.lFsfCode = 0
    tmCrfSrchKey1.iVefCode = 0
    tmCrfSrchKey1.iRotNo = 32000
    ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (ilRet <> BTRV_ERR_END_OF_FILE) Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hmCrf, llNoRec, -1, "UC", "CRF", "") 'Set extract limits (all records)
        gPackDate slActiveDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Crf", "CrfEndDate")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        tlCharTypeBuff.sType = "D"
        ilOffSet = gFieldOffset("Crf", "CrfState")
        ilRet = btrExtAddLogicConst(hmCrf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        ilOffSet = 0
        ilRet = btrExtAddField(hmCrf, ilOffSet, ilExtLen)  'Extract start/end time, and days
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        'ilRet = btrExtGetNextExt(hmClf)    'Extract record
        ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            'ilRet = btrExtGetFirst(hmClf, tlClfExt, ilExtLen, llRecPos)
            If ilRet = BTRV_ERR_REJECT_COUNT Then
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
            End If
            Do While ilRet = BTRV_ERR_NONE
                If imTerminate Then
                    Exit Sub
                End If
                tmCnfSrchKey.lCrfCode = tmCrf.lCode
                tmCnfSrchKey.iInstrNo = 0
                ilRet = btrGetGreaterOrEqual(hmCnf, tmCnf, imCnfRecLen, tmCnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmCnf.lCrfCode = tmCrf.lCode)
                    If imTerminate Then
                        Exit Sub
                    End If
                    llCif = mBinarySearchCifInfo(tmCnf.lCifCode)
                    If llCif >= 0 Then
                        blFound = False
                        llNext = tmCifInfo(llCif).lFirst
                        Do While llNext >= 0
                            If tmCrfDates(llNext).lCrfCode = tmCrf.lCode Then
                                blFound = True
                                Exit Do
                            End If
                            llNext = tmCrfDates(llNext).lNext
                        Loop
                        If Not blFound Then
                            llCrf = UBound(tmCrfDates)
                            If tmCifInfo(llCif).lFirst = -1 Then
                                llNext = -1
                            Else
                                llNext = tmCifInfo(llCif).lFirst
                            End If
                            tmCifInfo(llCif).lFirst = llCrf
                            tmCrfDates(llCrf).lCrfCode = tmCrf.lCode
                            gUnpackDateLong tmCrf.iStartDate(0), tmCrf.iStartDate(1), tmCrfDates(llCrf).lStartDate
                            gUnpackDateLong tmCrf.iEndDate(0), tmCrf.iEndDate(1), tmCrfDates(llCrf).lEndDate
                            tmCrfDates(llCrf).iAdfCode = tmCrf.iAdfCode
                            tmCrfDates(llCrf).lNext = llNext
                            ReDim Preserve tmCrfDates(0 To llCrf + 1) As CRFDATES
                            tmCrfDates(UBound(tmCrfDates)).lNext = -1
                        End If
                    End If
                    ilRet = btrGetNext(hmCnf, tmCnf, imCnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                If ilRet = BTRV_ERR_REJECT_COUNT Then
                    ilRet = btrExtGetNext(hmCrf, tmCrf, ilExtLen, llRecPos)
                End If
            Loop
        End If
        btrExtClear hmCrf   'Clear any previous extend operation
    End If
End Sub

Private Function mBinarySearchCifInfo(llCode As Long) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    llMin = LBound(tmCifInfo)
    llMax = UBound(tmCifInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tmCifInfo(llMiddle).lCifCode Then
            'found the match
            mBinarySearchCifInfo = llMiddle
            Exit Function
        ElseIf llCode < tmCifInfo(llMiddle).lCifCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchCifInfo = -1
End Function


Private Sub mUpdateCifDates()
    Dim llCif As Long
    Dim ilRet As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llNext As Long
    Dim llDate As Long
    Dim llRotStartDate As Long
    Dim llRotEndDate As Long
    Dim ilAdfCode As Integer
    Dim blSameAdfCode As Boolean
    
    lmChgCount = 0
    lmTotalCount = 0
    For llCif = 0 To UBound(tmCifInfo) - 1 Step 1
        ReDim tmDateSort(0 To 0) As DATESORT
        ilAdfCode = -1
        blSameAdfCode = True
        llNext = tmCifInfo(llCif).lFirst
        Do While llNext >= 0
            llDate = UBound(tmDateSort)
            tmDateSort(llDate).lStartDate = tmCrfDates(llNext).lStartDate
            tmDateSort(llDate).lEndDate = tmCrfDates(llNext).lEndDate
            tmDateSort(llDate).iAdfCode = tmCrfDates(llNext).iAdfCode
            If ilAdfCode = -1 Then
                ilAdfCode = tmCrfDates(llNext).iAdfCode
            Else
                If ilAdfCode <> tmCrfDates(llNext).iAdfCode Then
                    blSameAdfCode = False
                End If
            End If
            ReDim Preserve tmDateSort(0 To llDate + 1) As DATESORT
            llNext = tmCrfDates(llNext).lNext
        Loop
        llDate = UBound(tmDateSort)
        If llDate > 0 Then
            ArraySortTyp fnAV(tmDateSort(), 0), UBound(tmDateSort), 0, LenB(tmDateSort(0)), 0, -2, 0
        End If
        llStartDate = 0
        llEndDate = 0
        For llDate = 0 To UBound(tmDateSort) - 1 Step 1
            If tmCifInfo(llCif).iAdfCode = tmDateSort(llDate).iAdfCode Then
                If llStartDate = 0 Then
                    llStartDate = tmDateSort(llDate).lStartDate
                    llEndDate = tmDateSort(llDate).lEndDate
                Else
                    'Test for gap
                    If (llEndDate + 1 < tmDateSort(llDate).lStartDate) And (Not blSameAdfCode) Then
                        llStartDate = tmDateSort(llDate).lStartDate
                        llEndDate = tmDateSort(llDate).lEndDate
                    Else
                        If llEndDate < tmDateSort(llDate).lEndDate Then
                            llEndDate = tmDateSort(llDate).lEndDate
                        End If
                    End If
                End If
            End If
        Next llDate
        'Update CIF
        If llStartDate > 0 Then
            lmTotalCount = lmTotalCount + 1
            tmCifSrchKey0.lCode = tmCifInfo(llCif).lCifCode
            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                gUnpackDateLong tmCif.iRotStartDate(0), tmCif.iRotStartDate(1), llRotStartDate
                gUnpackDateLong tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), llRotEndDate
                If (llRotStartDate <> llStartDate) Or (llRotEndDate <> llEndDate) Then
                    lmChgCount = lmChgCount + 1
                    If ckcLog.Value = vbChecked Then
                        gLogMsg "Copy Item " & Trim$(tmCif.sName) & " From Date Range " & Format(llRotStartDate, "m/d/yy") & "-" & Format(llRotEndDate, "m/d/yy") & " to " & Format(llStartDate, "m/d/yy") & "-" & Format(llEndDate, "m/d/yy"), "SetCopyInventoryDates.txt", False
                    End If
                    gPackDateLong llStartDate, tmCif.iRotStartDate(0), tmCif.iRotStartDate(1)
                    gPackDateLong llEndDate, tmCif.iRotEndDate(0), tmCif.iRotEndDate(1)
                    ilRet = btrUpdate(hmCif, tmCif, imCifRecLen)
                End If
            End If
        End If
    Next llCif
End Sub
Private Sub mGridColumnTitles()
    grdSteps.TextMatrix(0, 0) = "Step"
    grdSteps.TextMatrix(1, 0) = "1"
    grdSteps.TextMatrix(2, 0) = "2"
    grdSteps.TextMatrix(3, 0) = "3"
    grdSteps.TextMatrix(0, 1) = "Definition"
    grdSteps.TextMatrix(1, 1) = "Gather Inventory"
    grdSteps.TextMatrix(2, 1) = "Gather Rotations"
    grdSteps.TextMatrix(3, 1) = "Update Inventory"
    grdSteps.TextMatrix(0, 2) = "Status"
    grdSteps.TextMatrix(1, 2) = ""
    grdSteps.TextMatrix(2, 2) = ""
    grdSteps.TextMatrix(3, 2) = ""
End Sub
Private Sub mGridColumnWidth()
    grdSteps.ColWidth(0) = grdSteps.Width * 0.2
    grdSteps.ColWidth(2) = grdSteps.Width * 0.3
    grdSteps.ColWidth(1) = grdSteps.Width - grdSteps.ColWidth(0) - grdSteps.ColWidth(2)
End Sub


Private Sub mResetControls()
    If grdSteps.Visible Then
        grdSteps.TextMatrix(1, 2) = ""
        grdSteps.TextMatrix(2, 2) = ""
        grdSteps.TextMatrix(3, 2) = ""
        grdSteps.Visible = False
        cmcSet.Enabled = True
        cmcCancel.Caption = "&Cancel"
    End If
End Sub
