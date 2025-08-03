VERSION 5.00
Begin VB.Form UndoBkFd 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3150
   ClientLeft      =   1665
   ClientTop       =   2295
   ClientWidth     =   5070
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
   ScaleHeight     =   3150
   ScaleWidth      =   5070
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2850
      TabIndex        =   5
      Top             =   2715
      Width           =   945
   End
   Begin VB.PictureBox plcExport 
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
      Height          =   2370
      Left            =   315
      ScaleHeight     =   2310
      ScaleWidth      =   4350
      TabIndex        =   4
      Top             =   225
      Width           =   4410
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
         Left            =   1905
         ScaleHeight     =   1710
         ScaleWidth      =   1935
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1995
         Begin VB.PictureBox pbcCalendar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ClipControls    =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   1440
            Left            =   30
            Picture         =   "UNDOBKFD.frx":0000
            ScaleHeight     =   1410
            ScaleWidth      =   1845
            TabIndex        =   12
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
               TabIndex        =   13
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
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   1
            Top             =   30
            Width           =   1305
         End
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
         Left            =   3690
         Picture         =   "UNDOBKFD.frx":2E1A
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   150
         Width           =   195
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
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   6
         Top             =   150
         Width           =   930
      End
      Begin VB.Label lacTranDate 
         Appearance      =   0  'Flat
         Caption         =   "Transmission Date to Undo"
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
         TabIndex        =   8
         Top             =   135
         Width           =   2535
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
      TabIndex        =   3
      Top             =   1770
      Width           =   75
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   30
      ScaleHeight     =   240
      ScaleWidth      =   1290
      TabIndex        =   0
      Top             =   -30
      Width           =   1290
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "&Undo"
      Height          =   285
      Left            =   1335
      TabIndex        =   2
      Top             =   2700
      Width           =   945
   End
End
Attribute VB_Name = "UndoBkFd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of UNDOBKFD.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: UndoBkFd.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Cart number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim tmCyf As CYF        'Cyf record image
Dim hmCyf As Integer    'Copy Export file handle
Dim imCyfRecLen As Integer        'Cif record length
Dim tmCrf As CRF        'Crf record image
Dim hmCrf As Integer    'Copy Rotation file handle
Dim imCrfRecLen As Integer        'ADF record length
Dim imTerminate As Integer  'True = terminating task, False= OK
'Dim tmRec As LPOPREC
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imBSMode As Integer     'Backspace flag
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
    igBFReturn = 0
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcTranDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcTranDate.SelStart = 0
    edcTranDate.SelLength = Len(edcTranDate.Text)
    edcTranDate.SetFocus
End Sub
Private Sub cmcTranDate_GotFocus()
    Dim slStr As String
    slStr = edcTranDate.Text
    If gValidDate(slStr) Then
        gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
        pbcCalendar_Paint   'mBoxCalDate called within paint
    Else
        lacDate.Visible = False
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUndo_Click()
    Dim ilRet As Integer
    Dim llRecPos As Long
    Dim ilCRet As Integer
    Dim slUndoDate As String
    Dim ilTranDate0 As Integer
    Dim ilTranDate1 As Integer
    Dim slNowDate As String
    slUndoDate = edcTranDate.Text
    If Not gValidDate(slUndoDate) Then
        Beep
        edcTranDate.SetFocus
        Exit Sub
    End If
    slNowDate = Format$(Now, "m/d/yy")
    If gDateValue(slUndoDate) < gDateValue(slNowDate) Then
        Beep
        MsgBox "Date must be after or equal to today's date", vbOkOnly + vbCritical + vbApplicationModal, "Undo Date"
        edcTranDate.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    gPackDate slUndoDate, ilTranDate0, ilTranDate1
    igBFReturn = 1
    imCyfRecLen = Len(tmCyf)
    ilRet = btrGetFirst(hmCyf, tmCyf, imCyfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If (Trim$(tmCyf.sSource) <> "S") And (Trim$(tmCyf.sSource) <> "K") Then
            If (tmCyf.iFeedDate(0) = ilTranDate0) And (tmCyf.iFeedDate(1) = ilTranDate1) Then
                ilRet = btrGetPosition(hmCyf, llRecPos)
                If ilRet <> BTRV_ERR_NONE Then
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Undo Not Completed, Try Later", vbOkOnly + vbExclamation, "Undo")
                    Exit Sub
                End If
                Do
                    'tmRec = tmCyf
                    'ilRet = gGetByKeyForUpdate("CYF", hmCyf, tmRec)
                    'tmCyf = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    Screen.MousePointer = vbDefault
                    '    ilRet = MsgBox("Undo Not Completed, Try Later", vbOkOnly + vbExclamation, "Undo")
                    '    Exit Sub
                    'End If
                    ilRet = btrDelete(hmCyf)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        ilCRet = btrGetDirect(hmCyf, tmCyf, imCyfRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Undo Not Completed, Try Later", vbOkOnly + vbExclamation, "Undo")
                    Exit Sub
                End If
            End If
        End If
        ilRet = btrGetNext(hmCyf, tmCyf, imCyfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilRet = btrGetFirst(hmCrf, tmCrf, imCrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If Trim$(tmCrf.sZone) <> "R" Then
            If (tmCrf.iFeedDate(0) = ilTranDate0) And (tmCrf.iFeedDate(1) = ilTranDate1) Then
                ilRet = btrGetPosition(hmCrf, llRecPos)
                If ilRet <> BTRV_ERR_NONE Then
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Undo Not Completed, Try Later", vbOkOnly + vbExclamation, "Undo")
                    Exit Sub
                End If
                Do
                    'tmRec = tmCrf
                    'ilRet = gGetByKeyForUpdate("CRF", hmCrf, tmRec)
                    'tmCrf = tmRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    Screen.MousePointer = vbDefault
                    '    ilRet = MsgBox("Undo Not Completed, Try Later", vbOkOnly + vbExclamation, "Undo")
                    '    Exit Sub
                    'End If
                    tmCrf.sFeedStatus = "R" 'Ready
                    ilRet = btrUpdate(hmCrf, tmCrf, imCrfRecLen)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        ilCRet = btrGetDirect(hmCrf, tmCrf, imCrfRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    Screen.MousePointer = vbDefault
                    ilRet = MsgBox("Undo Not Completed, Try Later", vbOkOnly + vbExclamation, "Undo")
                    Exit Sub
                End If
            End If
        End If
        ilRet = btrGetNext(hmCrf, tmCrf, imCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Screen.MousePointer = vbDefault
    mTerminate
End Sub
Private Sub cmcUndo_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub edcTranDate_Change()
    Dim slStr As String
    slStr = edcTranDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        cmcUndo.Enabled = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    cmcUndo.Enabled = True
End Sub
Private Sub edcTranDate_GotFocus()
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
Private Sub Form_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
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

    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    UndoBkFd.Height = cmcUndo.Top + 5 * cmcUndo.Height / 3
    gCenterStdAlone UndoBkFd
    mInitBox
    hmCyf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmCyf, "", sgDBPath & "Cyf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cyf.Btr)", UndoBkFd
    On Error GoTo 0
    imCyfRecLen = Len(tmCyf)
    hmCrf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Crf.Btr)", UndoBkFd
    On Error GoTo 0
    imCrfRecLen = Len(tmCrf)
    imBSMode = False
    imCalType = 0  'Standard
    edcTranDate.Text = Format$(gNow(), "mm/dd/yy")
    'slStr = Format$(Now, "m/d/yy")
    'slStr = gObtainNextMonday(slStr)
    'gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    'pbcCalendar_Paint   'mBoxCalDate called within paint
    'lacDate.Visible = False
    If imTerminate Then
        Exit Sub
    End If
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
    plcCalendar.Move edcTranDate.Left + edcTranDate.Width - plcCalendar.Width, edcTranDate.Top + edcTranDate.Height
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
    Unload UndoBkFd
    igManUnload = NO
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    btrExtClear hmCrf   'Clear any previous extend operation
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    btrExtClear hmCyf   'Clear any previous extend operation
    ilRet = btrClose(hmCyf)
    btrDestroy hmCyf
    
    Set UndoBkFd = Nothing   'Remove data segment
    
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
                'imBypassFocus = True
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
Private Sub plcExport_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Undo Export"
End Sub
