VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form PrgAirInfo 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4125
   ClientLeft      =   885
   ClientTop       =   2415
   ClientWidth     =   9315
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
   ScaleHeight     =   4125
   ScaleWidth      =   9315
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1230
      Top             =   3615
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
      Left            =   2205
      Picture         =   "PrgAirInfo.frx":0000
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   870
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcDay 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6945
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3825
      Visible         =   0   'False
      Width           =   240
      Begin VB.CheckBox ckcDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   0
         TabIndex        =   13
         Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
         Top             =   15
         Width           =   195
      End
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
      Left            =   5490
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1845
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
         Left            =   1620
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   30
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
         Left            =   30
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   30
         Picture         =   "PrgAirInfo.frx":00FA
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
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
            Left            =   525
            TabIndex        =   11
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
         Left            =   315
         TabIndex        =   8
         Top             =   30
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcTme 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   7650
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   1125
      Begin VB.PictureBox pbcTme 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   45
         Picture         =   "PrgAirInfo.frx":2F14
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   15
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
            Picture         =   "PrgAirInfo.frx":3BD2
            Top             =   765
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox pbcSetFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   9195
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   30
      Width           =   45
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   30
      Picture         =   "PrgAirInfo.frx":3EDC
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   675
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   16
      Top             =   3480
      Width           =   45
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   345
      Width           =   30
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1260
      MaxLength       =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5730
      TabIndex        =   19
      Top             =   3720
      Width           =   945
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4005
      TabIndex        =   18
      Top             =   3720
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
      Left            =   9045
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   20
      Top             =   3885
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2235
      TabIndex        =   17
      Top             =   3720
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPrgAirInfo 
      Height          =   3090
      Left            =   210
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   5450
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      GridColor       =   -2147483635
      GridColorFixed  =   -2147483635
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
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
      _Band(0).Cols   =   15
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8625
      Picture         =   "PrgAirInfo.frx":41E6
      Top             =   3570
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Program Names and Air Information"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   8925
   End
End
Attribute VB_Name = "PrgAirInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PrgAirInfo.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PrgAirInfo.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim hmPaf As Integer
Dim smPafStamp As String
Dim tmPaf As PAF        'Rvf record image
Dim tmPopPaf() As PAF
Dim tmPafSrchKey0 As LONGKEY0
Dim imPafRecLen As Integer        'RvF record length
Dim lmDelPaf() As Long

Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imLastSelectRow As Integer
Dim imCtrlKey As Integer
Dim imLastColSorted As Integer
Dim imLastSort As Integer
Dim lmRowSelected As Long
Dim imChg As Integer
Dim imIgnoreScroll As Integer
Dim imFromArrow As Integer

Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim imCtrlVisible As Integer
Dim lmTopRow As Long
Dim imInitNoRows As Integer
Dim imBSMode As Integer
Dim imBypassFocus As Integer

'Calendar info
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer

Const NAMEINDEX = 0
Const STARTDATEINDEX = 1
Const ENDDATEINDEX = 2
Const MOINDEX = 3
Const TUINDEX = 4
Const WEINDEX = 5
Const THINDEX = 6
Const FRINDEX = 7
Const SAINDEX = 8
Const SUINDEX = 9
Const STARTTIMEINDEX = 10
Const ENDTIMEINDEX = 11
Const CHGINDEX = 12
Const PAFCODEINDEX = 13
Const SORTINDEX = 14


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
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDone_Click()
    Dim ilRet As Integer

    If imChg Then
        If MsgBox("Save all changes?", vbYesNo) = vbYes Then
            ilRet = mSaveRec()
            If Not ilRet Then
                Exit Sub
            End If
        End If
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcDropDown_Click()

    Select Case lmEnableCol
        Case NAMEINDEX
        Case STARTDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case ENDDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case MOINDEX
        Case TUINDEX
        Case WEINDEX
        Case THINDEX
        Case FRINDEX
        Case SAINDEX
        Case SUINDEX
        Case STARTTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case ENDTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
    End Select
    grdPrgAirInfo.CellForeColor = vbBlack
End Sub

Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcUpdate_Click()
    Dim ilRet As Integer

    ilRet = mSaveRec()
    If ilRet Then
        mPopulate
    End If
End Sub

Private Sub cmcUpdate_GotFocus()
    mSetShow
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_Change()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilRet                                                   *
'******************************************************************************************
    Dim slStr As String

    Select Case lmEnableCol
        Case NAMEINDEX
        Case STARTDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                Exit Sub
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case ENDDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                Exit Sub
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case MOINDEX
        Case TUINDEX
        Case WEINDEX
        Case THINDEX
        Case FRINDEX
        Case SAINDEX
        Case SUINDEX
        Case STARTTIMEINDEX
        Case ENDTIMEINDEX
    End Select
    grdPrgAirInfo.CellForeColor = vbBlack

End Sub

Private Sub edcDropDown_GotFocus()
    Select Case lmEnableCol
        Case NAMEINDEX
        Case STARTDATEINDEX
        Case ENDDATEINDEX
        Case MOINDEX
        Case TUINDEX
        Case WEINDEX
        Case THINDEX
        Case FRINDEX
        Case SAINDEX
        Case SUINDEX
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
    Dim ilKey As Integer
    Dim ilPos As Integer
    Dim slStr As String
    Dim ilFound As Integer
    Dim ilLoop As Integer

    ilKey = KeyAscii
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If

    Select Case lmEnableCol
        Case NAMEINDEX
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
        Case MOINDEX
        Case TUINDEX
        Case WEINDEX
        Case THINDEX
        Case FRINDEX
        Case SAINDEX
        Case SUINDEX
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
        Select Case lmEnableCol
            Case NAMEINDEX
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
            Case MOINDEX
            Case TUINDEX
            Case WEINDEX
            Case THINDEX
            Case FRINDEX
            Case SAINDEX
            Case SUINDEX
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
        Select Case lmEnableCol
            Case NAMEINDEX
            Case STARTDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Left arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case ENDDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    slDate = edcDropDown.Text
                    If gValidDate(slDate) Then
                        If KeyCode = KEYLEFT Then 'Left arrow
                            slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                        Else
                            slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                        End If
                        gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                        edcDropDown.Text = slDate
                    End If
                End If
            Case MOINDEX
            Case TUINDEX
            Case WEINDEX
            Case THINDEX
            Case FRINDEX
            Case SAINDEX
            Case SUINDEX
            Case STARTTIMEINDEX
            Case ENDTIMEINDEX
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
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
'    gShowBranner
    PrgAirInfo.Refresh
    Me.KeyPreview = True
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
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        'Terminate
        tmcTerminate.Enabled = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    Erase tmPopPaf
    Erase lmDelPaf
    btrExtClear hmPaf   'Clear any previous extend operation
    ilRet = btrClose(hmPaf)
    btrDestroy hmPaf
    Set PrgAirInfo = Nothing   'Remove data segment
End Sub


Private Sub grdPrgAirInfo_EnterCell()
    mSetShow
End Sub

Private Sub grdPrgAirInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lmTopRow = grdPrgAirInfo.TopRow
    grdPrgAirInfo.Redraw = False
End Sub

Private Sub grdPrgAirInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilCol As Integer
    Dim ilRow As Integer

    imIgnoreScroll = False
    If Y < grdPrgAirInfo.RowHeight(0) Then
        grdPrgAirInfo.Col = grdPrgAirInfo.MouseCol
        mSortCol grdPrgAirInfo.Col
        Exit Sub
    End If
    pbcArrow.Visible = False
    ilCol = grdPrgAirInfo.MouseCol
    ilRow = grdPrgAirInfo.MouseRow
    If ilCol < grdPrgAirInfo.FixedCols Then
        grdPrgAirInfo.Redraw = True
        Exit Sub
    End If
    If ilRow < grdPrgAirInfo.FixedRows Then
        grdPrgAirInfo.Redraw = True
        Exit Sub
    End If
    If grdPrgAirInfo.TextMatrix(ilRow, NAMEINDEX) = "" Then
        grdPrgAirInfo.Redraw = False
        Do
            ilRow = ilRow - 1
        Loop While grdPrgAirInfo.TextMatrix(ilRow, NAMEINDEX) = ""
        grdPrgAirInfo.Row = ilRow + 1
        grdPrgAirInfo.Col = NAMEINDEX
        grdPrgAirInfo.Redraw = True
    Else
        grdPrgAirInfo.Row = ilRow
        grdPrgAirInfo.Col = ilCol
    End If
    grdPrgAirInfo.Redraw = True
    lmTopRow = grdPrgAirInfo.TopRow

    mEnableBox
End Sub

Private Sub grdPrgAirInfo_Scroll()
    If imIgnoreScroll Then  'Or igGridIgnoreScroll Then
        imIgnoreScroll = False
        Exit Sub
    End If
    If grdPrgAirInfo.Redraw = False Then
        grdPrgAirInfo.Redraw = True
        If lmTopRow < grdPrgAirInfo.FixedRows Then
            grdPrgAirInfo.TopRow = grdPrgAirInfo.FixedRows
        Else
            grdPrgAirInfo.TopRow = lmTopRow
        End If
        grdPrgAirInfo.Refresh
        grdPrgAirInfo.Redraw = False
    End If
    If (imCtrlVisible) And (grdPrgAirInfo.Row >= grdPrgAirInfo.FixedRows) And (grdPrgAirInfo.Col >= 0) And (grdPrgAirInfo.Col < grdPrgAirInfo.Cols - 1) Then
        If grdPrgAirInfo.RowIsVisible(grdPrgAirInfo.Row) Then
            pbcArrow.Move grdPrgAirInfo.Left - pbcArrow.Width - 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + (grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
            mSetFocus
        Else
            pbcSetFocus.SetFocus
            edcDropDown.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        slNameCode                    slCode                    *
'*                                                                                        *
'******************************************************************************************

'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilVef As Integer

    imFirstActivate = True
    imTerminate = False
    imIgnoreScroll = False
    imFromArrow = False
    imCtrlVisible = False

    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    'mParseCmmdLine
    PrgAirInfo.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone PrgAirInfo
    'PrgAirInfo.Show
    Screen.MousePointer = vbHourglass
    ilVef = gBinarySearchVef(igPrgNameVefCode)
    If ilVef <> -1 Then
        lacScreen.Caption = "Program Names and Air Information: " & Trim(tgMVef(ilVef).sName)
    End If
    If mTestRecLengths() Then
        Screen.MousePointer = vbDefault
        imTerminate = True
        Exit Sub
    End If
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    lmRowSelected = -1
    imChg = False
    imBSMode = False
    imBypassFocus = False

    imFirstFocus = True
    imLastSelectRow = 0
    imCtrlKey = False
    imPafRecLen = Len(tmPaf)
    hmPaf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmPaf, "", sgDBPath & "Paf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Paf.Btr)", PrgAirInfo
    On Error GoTo 0

    mInitBox

    mPopulate
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
    igManUnload = YES
    'Unload Traffic
    Unload PrgAirInfo
    igManUnload = NO
End Sub




Private Sub imcTrash_Click()
    Dim llRow As Long
    Dim ilCol As Integer
    Dim llEnableRow As Long

    If (lmEnableRow >= grdPrgAirInfo.FixedRows) And (lmEnableRow < grdPrgAirInfo.Rows) Then
        llEnableRow = lmEnableRow
        lmDelPaf(UBound(lmDelPaf)) = Val(grdPrgAirInfo.TextMatrix(llEnableRow, PAFCODEINDEX))
        If lmDelPaf(UBound(lmDelPaf)) > 0 Then
            imChg = True
            ReDim Preserve lmDelPaf(0 To UBound(lmDelPaf) + 1) As Long
        End If
        grdPrgAirInfo.Redraw = False
        mSetShow
        For llRow = llEnableRow To grdPrgAirInfo.Rows - 2 Step 1
            For ilCol = NAMEINDEX To SORTINDEX Step 1
                grdPrgAirInfo.TextMatrix(llRow, ilCol) = grdPrgAirInfo.TextMatrix(llRow + 1, ilCol)
                grdPrgAirInfo.TextMatrix(llRow + 1, ilCol) = ""
            Next ilCol
        Next llRow
        pbcClickFocus.SetFocus
        grdPrgAirInfo.Redraw = True
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    If imFirstFocus Then
        imFirstFocus = False
    End If
    If grdPrgAirInfo.Visible Then
        lmRowSelected = -1
        grdPrgAirInfo.Row = 0
        grdPrgAirInfo.Col = PAFCODEINDEX
        mSetCommands
    End If
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub mPopulate()
    Dim ilRet As Integer
    smPafStamp = ""
    ilRet = gObtainPaf(hmPaf, igPrgNameVefCode, smPafStamp, tmPopPaf())
    mMoveRecToCtrl
    ReDim lmDelPaf(0 To 0) As Long
End Sub


Private Sub mSetCommands()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRow                                                                                 *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  cmcEraseErr                                                                           *
'******************************************************************************************

    Dim ilRet As Integer

    If imChg Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    Exit Sub
cmcEraseErr: 'VBC NR
    ilRet = 1
    Resume Next
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
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  flTextHeight                  ilLoop                        ilRow                     *
'*  ilCol                                                                                 *
'******************************************************************************************

'
'   mInitBox
'   Where:
'
    'flTextHeight = pbcDates.TextHeight("1") - 35
    Dim ilLoop As Integer
    

    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    
    mGridLayout
    mGridColumnWidths
    mGridColumns
    'grdPrgAirInfo.Move 180, lacScreen.Top + lacScreen.Height + 120, grdPrgAirInfo.Width, cmcDone.Top - (lacScreen.Top + lacScreen.Height) - 240
    grdPrgAirInfo.Move 180, lacScreen.Top + lacScreen.Height + 120, grdPrgAirInfo.Width, imcTrash.Top - (lacScreen.Top + lacScreen.Height) - 240
    'grdPrgAirInfo.Height = grdPrgAirInfo.RowPos(0) + 14 * grdPrgAirInfo.RowHeight(0) + fgPanelAdj - 15
    'imInitNoRows = (cmcDone.Top - 120 - grdPrgAirInfo.Top) \ fgFlexGridRowH
    imInitNoRows = (imcTrash.Top - 120 - grdPrgAirInfo.Top) \ fgFlexGridRowH
    grdPrgAirInfo.Height = grdPrgAirInfo.RowPos(0) + imInitNoRows * (fgFlexGridRowH) + fgPanelAdj + 15 '- 15
End Sub

Private Sub mGridLayout()
    Dim ilCol As Integer
    Dim ilRow As Integer

    For ilRow = 0 To grdPrgAirInfo.Rows - 1 Step 1
        grdPrgAirInfo.RowHeight(ilRow) = fgFlexGridRowH
    Next ilRow
    For ilCol = 0 To grdPrgAirInfo.Cols - 1 Step 1
        grdPrgAirInfo.ColAlignment(ilCol) = flexAlignLeftCenter
    Next ilCol
End Sub

Private Sub mGridColumns()
    Dim ilCol As Integer

    grdPrgAirInfo.Row = grdPrgAirInfo.FixedRows - 1
    For ilCol = NAMEINDEX To SORTINDEX Step 1
        grdPrgAirInfo.Col = ilCol
        grdPrgAirInfo.CellFontBold = False
        grdPrgAirInfo.CellFontName = "Arial"
        grdPrgAirInfo.CellFontSize = 6.75
        grdPrgAirInfo.CellForeColor = vbBlue
        grdPrgAirInfo.CellBackColor = LIGHTBLUE
    Next ilCol
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, NAMEINDEX) = "Name"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, STARTDATEINDEX) = "Start Date"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, ENDDATEINDEX) = "End Date"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, MOINDEX) = "Mo"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, TUINDEX) = "Tu"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, WEINDEX) = "We"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, THINDEX) = "Th"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, FRINDEX) = "Fr"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, SAINDEX) = "Sa"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, SUINDEX) = "Su"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, STARTTIMEINDEX) = "Start Time"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, ENDTIMEINDEX) = "End Time"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, CHGINDEX) = "Changed"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, PAFCODEINDEX) = "Paf Code"
    grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, SORTINDEX) = "Sort"

End Sub

Private Sub mGridColumnWidths()
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim ilCol As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer

    grdPrgAirInfo.ColWidth(CHGINDEX) = 0
    grdPrgAirInfo.ColWidth(PAFCODEINDEX) = 0
    grdPrgAirInfo.ColWidth(SORTINDEX) = 0
    grdPrgAirInfo.ColWidth(NAMEINDEX) = 0.3 * grdPrgAirInfo.Width
    grdPrgAirInfo.ColWidth(STARTDATEINDEX) = 0.1 * grdPrgAirInfo.Width
    grdPrgAirInfo.ColWidth(ENDDATEINDEX) = 0.1 * grdPrgAirInfo.Width
    grdPrgAirInfo.ColWidth(MOINDEX) = 0.03 * grdPrgAirInfo.Width
    grdPrgAirInfo.ColWidth(TUINDEX) = 0.03 * grdPrgAirInfo.Width
    grdPrgAirInfo.ColWidth(WEINDEX) = 0.03 * grdPrgAirInfo.Width
    grdPrgAirInfo.ColWidth(THINDEX) = 0.03 * grdPrgAirInfo.Width
    grdPrgAirInfo.ColWidth(FRINDEX) = 0.03 * grdPrgAirInfo.Width
    grdPrgAirInfo.ColWidth(SAINDEX) = 0.03 * grdPrgAirInfo.Width
    grdPrgAirInfo.ColWidth(SUINDEX) = 0.03 * grdPrgAirInfo.Width
    grdPrgAirInfo.ColWidth(STARTTIMEINDEX) = 0.1 * grdPrgAirInfo.Width
    grdPrgAirInfo.ColWidth(ENDTIMEINDEX) = 0.1 * grdPrgAirInfo.Width
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdPrgAirInfo.Width
    For ilCol = 0 To grdPrgAirInfo.Cols - 1 Step 1
        llWidth = llWidth + grdPrgAirInfo.ColWidth(ilCol)
        If (grdPrgAirInfo.ColWidth(ilCol) > 15) And (grdPrgAirInfo.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdPrgAirInfo.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdPrgAirInfo.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdPrgAirInfo.Width
            For ilCol = 0 To grdPrgAirInfo.Cols - 1 Step 1
                If (grdPrgAirInfo.ColWidth(ilCol) > 15) And (grdPrgAirInfo.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdPrgAirInfo.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdPrgAirInfo.FixedCols To grdPrgAirInfo.Cols - 1 Step 1
                If grdPrgAirInfo.ColWidth(ilCol) > 15 Then
                    ilColInc = grdPrgAirInfo.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdPrgAirInfo.ColWidth(ilCol) = grdPrgAirInfo.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
End Sub


Private Sub mSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

    For llRow = grdPrgAirInfo.FixedRows To grdPrgAirInfo.Rows - 1 Step 1
        slStr = Trim$(grdPrgAirInfo.TextMatrix(llRow, NAMEINDEX))
        If slStr <> "" Then
            If ilCol = NAMEINDEX Then
                slSort = grdPrgAirInfo.TextMatrix(llRow, NAMEINDEX)
                Do While Len(slSort) < 20
                    slSort = slSort & " "
                Loop
            ElseIf (ilCol = STARTDATEINDEX) Or (ilCol = ENDDATEINDEX) Then
                slSort = Trim$(str$(gDateValue(grdPrgAirInfo.TextMatrix(llRow, ilCol))))
                Do While Len(slSort) < 6
                    slSort = "0" & slSort
                Loop
            ElseIf (ilCol = STARTTIMEINDEX) Then
                slSort = Trim$(str$(gTimeToLong(grdPrgAirInfo.TextMatrix(llRow, ilCol), False)))
                Do While Len(slSort) < 6
                    slSort = slSort & " "
                Loop
            ElseIf (ilCol = ENDTIMEINDEX) Then
                slSort = Trim$(str$(gTimeToLong(grdPrgAirInfo.TextMatrix(llRow, ilCol), True)))
                Do While Len(slSort) < 6
                    slSort = slSort & " "
                Loop
            Else
                slSort = grdPrgAirInfo.TextMatrix(llRow, ilCol)
            End If
            slStr = grdPrgAirInfo.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdPrgAirInfo.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                slRow = Trim$(str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdPrgAirInfo.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastColSorted Then
        imLastColSorted = SORTINDEX
    Else
        imLastColSorted = -1
        imLastSort = -1
    End If
    gGrid_SortByCol grdPrgAirInfo, NAMEINDEX, SORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
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
Private Sub mEnableBox()
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If (grdPrgAirInfo.Row < grdPrgAirInfo.FixedRows) Or (grdPrgAirInfo.Row >= grdPrgAirInfo.Rows) Or (grdPrgAirInfo.Col < grdPrgAirInfo.FixedCols) Or (grdPrgAirInfo.Col >= grdPrgAirInfo.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdPrgAirInfo.Row
    lmEnableCol = grdPrgAirInfo.Col
    pbcArrow.Visible = False
    pbcArrow.Move grdPrgAirInfo.Left - pbcArrow.Width - 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + (grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    imCtrlVisible = True
    Select Case grdPrgAirInfo.Col
        Case NAMEINDEX
            edcDropDown.MaxLength = 20
            slStr = grdPrgAirInfo.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdPrgAirInfo.Row > grdPrgAirInfo.FixedRows Then
                    slStr = grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row - 1, grdPrgAirInfo.Col)
                End If
            End If
            edcDropDown.Text = slStr
        Case STARTDATEINDEX
            edcDropDown.MaxLength = 10
            slStr = grdPrgAirInfo.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdPrgAirInfo.Row > grdPrgAirInfo.FixedRows Then
                    slStr = grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row - 1, grdPrgAirInfo.Col)
                End If
            End If
            If slStr <> "" Then
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
            Else
                slStr = Format$(gNow(), "m/d/yy")
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                slStr = ""
            End If
            edcDropDown.Text = slStr
        Case ENDDATEINDEX
            edcDropDown.MaxLength = 10
            slStr = grdPrgAirInfo.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If (slStr = "") Then
                If grdPrgAirInfo.Row > grdPrgAirInfo.FixedRows Then
                    slStr = grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row - 1, grdPrgAirInfo.Col)
                End If
            End If
            If slStr <> "" Then
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
            Else
                slStr = Format$(gNow(), "m/d/yy")
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                slStr = ""
            End If
            edcDropDown.Text = slStr
        Case MOINDEX
            grdPrgAirInfo.CellFontBold = False
            grdPrgAirInfo.CellFontName = "Monotype Sorts"
            slStr = grdPrgAirInfo.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If slStr = " 4" Then
                ckcDay.Value = vbChecked
            Else
                ckcDay.Value = vbUnchecked
            End If
        Case TUINDEX
            grdPrgAirInfo.CellFontBold = False
            grdPrgAirInfo.CellFontName = "Monotype Sorts"
            slStr = grdPrgAirInfo.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If slStr = " 4" Then
                ckcDay.Value = vbChecked
            Else
                ckcDay.Value = vbUnchecked
            End If
        Case WEINDEX
            grdPrgAirInfo.CellFontBold = False
            grdPrgAirInfo.CellFontName = "Monotype Sorts"
            slStr = grdPrgAirInfo.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If slStr = " 4" Then
                ckcDay.Value = vbChecked
            Else
                ckcDay.Value = vbUnchecked
            End If
        Case THINDEX
            grdPrgAirInfo.CellFontBold = False
            grdPrgAirInfo.CellFontName = "Monotype Sorts"
            slStr = grdPrgAirInfo.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If slStr = " 4" Then
                ckcDay.Value = vbChecked
            Else
                ckcDay.Value = vbUnchecked
            End If
        Case FRINDEX
            grdPrgAirInfo.CellFontBold = False
            grdPrgAirInfo.CellFontName = "Monotype Sorts"
            slStr = grdPrgAirInfo.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If slStr = " 4" Then
                ckcDay.Value = vbChecked
            Else
                ckcDay.Value = vbUnchecked
            End If
        Case SAINDEX
            grdPrgAirInfo.CellFontBold = False
            grdPrgAirInfo.CellFontName = "Monotype Sorts"
            slStr = grdPrgAirInfo.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If slStr = " 4" Then
                ckcDay.Value = vbChecked
            Else
                ckcDay.Value = vbUnchecked
            End If
        Case SUINDEX
            grdPrgAirInfo.CellFontBold = False
            grdPrgAirInfo.CellFontName = "Monotype Sorts"
            slStr = grdPrgAirInfo.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If slStr = " 4" Then
                ckcDay.Value = vbChecked
            Else
                ckcDay.Value = vbUnchecked
            End If
        Case STARTTIMEINDEX
            edcDropDown.MaxLength = 11
            slStr = grdPrgAirInfo.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            edcDropDown.Text = slStr
        Case ENDTIMEINDEX
            edcDropDown.MaxLength = 11
            slStr = grdPrgAirInfo.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            edcDropDown.Text = slStr
    End Select
    mSetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer

    If (grdPrgAirInfo.Row < grdPrgAirInfo.FixedRows) Or (grdPrgAirInfo.Row >= grdPrgAirInfo.Rows) Or (grdPrgAirInfo.Col < grdPrgAirInfo.FixedCols) Or (grdPrgAirInfo.Col >= grdPrgAirInfo.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    pbcArrow.Visible = False
    pbcArrow.Move grdPrgAirInfo.Left - pbcArrow.Width - 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + (grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - pbcArrow.Height) / 2
    pbcArrow.Visible = True
    llColPos = 0
    For ilCol = 0 To grdPrgAirInfo.Col - 1 Step 1
        llColPos = llColPos + grdPrgAirInfo.ColWidth(ilCol)
    Next ilCol
    Select Case grdPrgAirInfo.Col
        Case NAMEINDEX
            edcDropDown.Move grdPrgAirInfo.Left + llColPos + 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + 30, grdPrgAirInfo.ColWidth(grdPrgAirInfo.Col) - 30, grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - 15
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            edcDropDown.SetFocus
        Case STARTDATEINDEX
            edcDropDown.Move grdPrgAirInfo.Left + llColPos + 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + 30, grdPrgAirInfo.ColWidth(grdPrgAirInfo.Col) - 30, grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - 15
            If edcDropDown.Top + edcDropDown.Height + plcCalendar.Height < cmcDone.Top Then
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.Height
            End If
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ENDDATEINDEX
            edcDropDown.Move grdPrgAirInfo.Left + llColPos + 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + 30, grdPrgAirInfo.ColWidth(grdPrgAirInfo.Col) - 30, grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - 15
            If edcDropDown.Top + edcDropDown.Height + plcCalendar.Height < cmcDone.Top Then
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.Height
            End If
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case MOINDEX
            pbcDay.Move grdPrgAirInfo.Left + llColPos + 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + 30, grdPrgAirInfo.ColWidth(grdPrgAirInfo.Col) - 30, grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - 15
            pbcDay.Visible = True
            ckcDay.SetFocus
        Case TUINDEX
            pbcDay.Move grdPrgAirInfo.Left + llColPos + 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + 30, grdPrgAirInfo.ColWidth(grdPrgAirInfo.Col) - 30, grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - 15
            pbcDay.Visible = True
            ckcDay.SetFocus
        Case WEINDEX
            pbcDay.Move grdPrgAirInfo.Left + llColPos + 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + 30, grdPrgAirInfo.ColWidth(grdPrgAirInfo.Col) - 30, grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - 15
            pbcDay.Visible = True
            ckcDay.SetFocus
        Case THINDEX
            pbcDay.Move grdPrgAirInfo.Left + llColPos + 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + 30, grdPrgAirInfo.ColWidth(grdPrgAirInfo.Col) - 30, grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - 15
            pbcDay.Visible = True
            ckcDay.SetFocus
        Case FRINDEX
            pbcDay.Move grdPrgAirInfo.Left + llColPos + 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + 30, grdPrgAirInfo.ColWidth(grdPrgAirInfo.Col) - 30, grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - 15
            pbcDay.Visible = True
            ckcDay.SetFocus
        Case SAINDEX
            pbcDay.Move grdPrgAirInfo.Left + llColPos + 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + 30, grdPrgAirInfo.ColWidth(grdPrgAirInfo.Col) - 30, grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - 15
            pbcDay.Visible = True
            ckcDay.SetFocus
        Case SUINDEX
            pbcDay.Move grdPrgAirInfo.Left + llColPos + 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + 30, grdPrgAirInfo.ColWidth(grdPrgAirInfo.Col) - 30, grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - 15
            pbcDay.Visible = True
            ckcDay.SetFocus
        Case STARTTIMEINDEX
            edcDropDown.Move grdPrgAirInfo.Left + llColPos + 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + 30, grdPrgAirInfo.ColWidth(grdPrgAirInfo.Col) - 30, grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - 15
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ENDTIMEINDEX
            edcDropDown.Move grdPrgAirInfo.Left + llColPos + 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + 30, grdPrgAirInfo.ColWidth(grdPrgAirInfo.Col) - 30, grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - 15
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
    mSetCommands
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
Private Sub mSetShow()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPos                                                                                 *
'******************************************************************************************

'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim ilChg As Integer

    pbcArrow.Visible = False
    ilChg = False
    If (lmEnableRow >= grdPrgAirInfo.FixedRows) And (lmEnableRow < grdPrgAirInfo.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
                edcDropDown.Visible = False  'Set visibility
                slStr = edcDropDown.Text
                If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    ilChg = True
                End If
                grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case STARTDATEINDEX
                edcDropDown.Visible = False  'Set visibility
                plcCalendar.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                If gDateValue(grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol)) <> gDateValue(slStr) Then
                    ilChg = True
                End If
                grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case ENDDATEINDEX
                edcDropDown.Visible = False  'Set visibility
                plcCalendar.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                If gDateValue(grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol)) <> gDateValue(slStr) Then
                    ilChg = True
                End If
                grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case MOINDEX
                pbcDay.Visible = False
                If ckcDay = vbChecked Then
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> " 4" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = " 4"
                Else
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case TUINDEX
                pbcDay.Visible = False
                If ckcDay = vbChecked Then
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> " 4" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = " 4"
                Else
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case WEINDEX
                pbcDay.Visible = False
                If ckcDay = vbChecked Then
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> " 4" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = " 4"
                Else
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case THINDEX
                pbcDay.Visible = False
                If ckcDay = vbChecked Then
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> " 4" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = " 4"
                Else
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case FRINDEX
                pbcDay.Visible = False
                If ckcDay = vbChecked Then
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> " 4" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = " 4"
                Else
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case SAINDEX
                pbcDay.Visible = False
                If ckcDay = vbChecked Then
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> " 4" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = " 4"
                Else
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case SUINDEX
                pbcDay.Visible = False
                If ckcDay = vbChecked Then
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> " 4" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = " 4"
                Else
                    If grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) <> "" Then
                        ilChg = True
                    End If
                    grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = ""
                End If
            Case STARTTIMEINDEX
                edcDropDown.Visible = False  'Set visibility
                plcTme.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                If gTimeToLong(grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol), False) <> gTimeToLong(slStr, False) Then
                    ilChg = True
                End If
                grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = slStr
            Case ENDTIMEINDEX
                edcDropDown.Visible = False  'Set visibility
                plcTme.Visible = False
                cmcDropDown.Visible = False
                slStr = edcDropDown.Text
                If gTimeToLong(grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol), False) <> gTimeToLong(slStr, False) Then
                    ilChg = True
                End If
                grdPrgAirInfo.TextMatrix(lmEnableRow, lmEnableCol) = slStr
        End Select
        If ilChg Then
            grdPrgAirInfo.TextMatrix(lmEnableRow, CHGINDEX) = "Y"
            imChg = ilChg
        End If
    End If
    pbcArrow.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    mSetCommands
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imFromArrow Then
        imFromArrow = False
        mEnableBox
        Exit Sub
    End If
    If imCtrlVisible Then
        mSetShow
        Do
            ilPrev = False
            If grdPrgAirInfo.Col = NAMEINDEX Then
                If grdPrgAirInfo.Row > grdPrgAirInfo.FixedRows Then
                    lmTopRow = -1
                    grdPrgAirInfo.Row = grdPrgAirInfo.Row - 1
                    If Not grdPrgAirInfo.RowIsVisible(grdPrgAirInfo.Row) Then
                        grdPrgAirInfo.TopRow = grdPrgAirInfo.TopRow - 1
                    End If
                    grdPrgAirInfo.Col = ENDTIMEINDEX
                    mEnableBox
                Else
                    cmcCancel.SetFocus
                End If
            Else
                grdPrgAirInfo.Col = grdPrgAirInfo.Col - 1
                'If gColOk(grdPrgAirInfo, grdPrgAirInfo.Row, grdPrgAirInfo.Col) Then
                    mEnableBox
                'Else
                '    ilPrev = True
                'End If
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdPrgAirInfo.TopRow = grdPrgAirInfo.FixedRows
        grdPrgAirInfo.Col = NAMEINDEX
        grdPrgAirInfo.Row = grdPrgAirInfo.FixedRows
        'If gColOk(grdPrgAirInfo, grdPrgAirInfo.Row, grdPrgAirInfo.Col) Then
            mEnableBox
        'Else
        '    cmcCancel.SetFocus
        'End If
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim ilCol As Integer
    Dim ilNext As Integer
    Dim llEnableRow As Long

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        mSetShow
        Do
            ilNext = False
            If grdPrgAirInfo.Col = ENDTIMEINDEX Then
                llRow = grdPrgAirInfo.Rows
                Do
                    llRow = llRow - 1
                Loop While grdPrgAirInfo.TextMatrix(llRow, NAMEINDEX) = ""
                llRow = llRow + 1
                If (grdPrgAirInfo.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdPrgAirInfo.Row = grdPrgAirInfo.Row + 1
                    If Not grdPrgAirInfo.RowIsVisible(grdPrgAirInfo.Row) Or (grdPrgAirInfo.Row - (grdPrgAirInfo.TopRow - grdPrgAirInfo.FixedRows) >= imInitNoRows) Then
                        imIgnoreScroll = True
                        grdPrgAirInfo.TopRow = grdPrgAirInfo.TopRow + 1
                    End If
                    grdPrgAirInfo.Col = NAMEINDEX
                    'grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, CODEINDEX) = 0
                    If Trim$(grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, NAMEINDEX)) <> "" Then
                        'If gColOk(grdPrgAirInfo, grdPrgAirInfo.Row, grdPrgAirInfo.Col) Then
                            mEnableBox
                        'Else
                        '    cmcCancel.SetFocus
                        'End If
                    Else
                        imFromArrow = True
                        pbcArrow.Move grdPrgAirInfo.Left - pbcArrow.Width - 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + (grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    End If
                Else
                    If Trim$(grdPrgAirInfo.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                        lmTopRow = -1
                        If grdPrgAirInfo.Row + 1 >= grdPrgAirInfo.Rows Then
                            grdPrgAirInfo.AddItem ""
                            grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row + 1) = fgFlexGridRowH
                            grdPrgAirInfo.Row = grdPrgAirInfo.Row + 1
                            For ilCol = 0 To grdPrgAirInfo.Cols - 1 Step 1
                                grdPrgAirInfo.ColAlignment(ilCol) = flexAlignLeftCenter
                            Next ilCol
                        Else
                            grdPrgAirInfo.Row = grdPrgAirInfo.Row + 1
                        End If
                        If (Not grdPrgAirInfo.RowIsVisible(grdPrgAirInfo.Row)) Or (grdPrgAirInfo.Row - (grdPrgAirInfo.TopRow - grdPrgAirInfo.FixedRows) >= imInitNoRows) Then
                            imIgnoreScroll = True
                            grdPrgAirInfo.TopRow = grdPrgAirInfo.TopRow + 1
                        End If
                        grdPrgAirInfo.Col = NAMEINDEX
                        grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, PAFCODEINDEX) = 0
                        grdPrgAirInfo.TextMatrix(grdPrgAirInfo.Row, CHGINDEX) = "N"
                        'mEnableBox
                        imFromArrow = True
                        pbcArrow.Move grdPrgAirInfo.Left - pbcArrow.Width - 30, grdPrgAirInfo.Top + grdPrgAirInfo.RowPos(grdPrgAirInfo.Row) + (grdPrgAirInfo.RowHeight(grdPrgAirInfo.Row) - pbcArrow.Height) / 2
                        pbcArrow.Visible = True
                        pbcArrow.SetFocus
                    Else
                        pbcClickFocus.SetFocus
                    End If
                End If
            Else
                grdPrgAirInfo.Col = grdPrgAirInfo.Col + 1
                'If gColOk(grdPrgAirInfo, grdPrgAirInfo.Row, grdPrgAirInfo.Col) Then
                    mEnableBox
                'Else
                '    ilNext = True
                'End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdPrgAirInfo.TopRow = grdPrgAirInfo.FixedRows
        grdPrgAirInfo.Col = NAMEINDEX
        grdPrgAirInfo.Row = grdPrgAirInfo.FixedRows
        'If gColOk(grdPrgAirInfo, grdPrgAirInfo.Row, grdPrgAirInfo.Col) Then
            mEnableBox
        'Else
        '    cmcCancel.SetFocus
        'End If
    End If
End Sub

Private Function mSaveRec() As Integer
    Dim ilRow As Integer
    Dim slMsg As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilError As Integer
    Dim tlPaf As PAF

    ilError = False
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdPrgAirInfo, grdPrgAirInfo, vbHourglass
    For ilRow = grdPrgAirInfo.FixedRows To grdPrgAirInfo.Rows - 1 Step 1
        If mGridFieldsOk(ilRow) = False Then
            ilError = True
        End If
    Next ilRow
    If ilError Then
        gSetMousePointer grdPrgAirInfo, grdPrgAirInfo, vbDefault
        Screen.MousePointer = vbDefault
        Beep
        mSaveRec = False
        Exit Function
    End If
    mMoveCtrlToRec
    ilRet = btrBeginTrans(hmPaf, 1000)
    For ilLoop = 0 To UBound(tmPopPaf) - 1 Step 1
        If tmPopPaf(ilLoop).lCode <= 0 Then
            tmPopPaf(ilLoop).lCode = 0
            ilRet = btrInsert(hmPaf, tmPopPaf(ilLoop), imPafRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert:Program Names)"
            If ilRet <> BTRV_ERR_NONE Then
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, PrgAirInfo
                On Error GoTo 0
            End If
        Else
            Do
                tmPafSrchKey0.lCode = tmPopPaf(ilLoop).lCode
                ilRet = btrGetEqual(hmPaf, tlPaf, imPafRecLen, tmPafSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                ilRet = btrUpdate(hmPaf, tmPopPaf(ilLoop), imPafRecLen)
                slMsg = "mSaveRec (btrUpdate:Program Names)"
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                On Error GoTo mSaveRecErr
                gBtrvErrorMsg ilRet, slMsg, PrgAirInfo
                On Error GoTo 0
            End If
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(lmDelPaf) - 1 Step 1
        Do
            tmPafSrchKey0.lCode = lmDelPaf(ilLoop)
            ilRet = btrGetEqual(hmPaf, tlPaf, imPafRecLen, tmPafSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hmPaf)
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        slMsg = "mSaveRec (btrDelete:Program Names)"
        If ilRet <> BTRV_ERR_NONE Then
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, PrgAirInfo
            On Error GoTo 0
        End If
    Next ilLoop
    imChg = False
    ilRet = btrEndTrans(hmPaf)
    mSaveRec = True
    gSetMousePointer grdPrgAirInfo, grdPrgAirInfo, vbDefault
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    ilRet = btrAbortTrans(hmPaf)
    gSetMousePointer grdPrgAirInfo, grdPrgAirInfo, vbDefault
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer control values to     *
'*                      records                        *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                         ilPaf                                                   *
'******************************************************************************************

    Dim llRow As Long
    Dim slStr As String
    Dim llIndex As Long
    Dim llCode As Long
    Dim llLoop As Long

    ReDim tmPopPaf(0 To 0) As PAF
    For llRow = grdPrgAirInfo.FixedRows To grdPrgAirInfo.Rows - 1 Step 1
        slStr = Trim$(grdPrgAirInfo.TextMatrix(llRow, NAMEINDEX))
        If slStr <> "" Then
            If grdPrgAirInfo.TextMatrix(llRow, CHGINDEX) = "Y" Then
                If grdPrgAirInfo.TextMatrix(llRow, PAFCODEINDEX) <> "" Then
                    llCode = Val(grdPrgAirInfo.TextMatrix(llRow, PAFCODEINDEX))
                Else
                    llCode = 0
                End If
                llIndex = UBound(tmPopPaf)
                tmPopPaf(llIndex).lCode = llCode
                tmPopPaf(llIndex).iVefCode = igPrgNameVefCode
                slStr = Trim$(grdPrgAirInfo.TextMatrix(llRow, NAMEINDEX))
                tmPopPaf(llIndex).sName = slStr
                slStr = Trim$(grdPrgAirInfo.TextMatrix(llRow, STARTDATEINDEX))
                gPackDate slStr, tmPopPaf(llIndex).iStartDate(0), tmPopPaf(llIndex).iStartDate(1)
                slStr = Trim$(grdPrgAirInfo.TextMatrix(llRow, ENDDATEINDEX))
                If slStr = "" Then
                    slStr = "12/31/2069"
                End If
                gPackDate slStr, tmPopPaf(llIndex).iEndDate(0), tmPopPaf(llIndex).iEndDate(1)
                If grdPrgAirInfo.TextMatrix(llRow, MOINDEX) = " 4" Then
                    tmPopPaf(llIndex).sMo = "Y"
                Else
                    tmPopPaf(llIndex).sMo = "N"
                End If
                If grdPrgAirInfo.TextMatrix(llRow, TUINDEX) = " 4" Then
                    tmPopPaf(llIndex).sTu = "Y"
                Else
                    tmPopPaf(llIndex).sTu = "N"
                End If
                If grdPrgAirInfo.TextMatrix(llRow, WEINDEX) = " 4" Then
                    tmPopPaf(llIndex).sWe = "Y"
                Else
                    tmPopPaf(llIndex).sWe = "N"
                End If
                If grdPrgAirInfo.TextMatrix(llRow, THINDEX) = " 4" Then
                    tmPopPaf(llIndex).sTh = "Y"
                Else
                    tmPopPaf(llIndex).sTh = "N"
                End If
                If grdPrgAirInfo.TextMatrix(llRow, FRINDEX) = " 4" Then
                    tmPopPaf(llIndex).sFr = "Y"
                Else
                    tmPopPaf(llIndex).sFr = "N"
                End If
                If grdPrgAirInfo.TextMatrix(llRow, SAINDEX) = " 4" Then
                    tmPopPaf(llIndex).sSa = "Y"
                Else
                    tmPopPaf(llIndex).sSa = "N"
                End If
                If grdPrgAirInfo.TextMatrix(llRow, SUINDEX) = " 4" Then
                    tmPopPaf(llIndex).sSu = "Y"
                Else
                    tmPopPaf(llIndex).sSu = "N"
                End If
                slStr = Trim$(grdPrgAirInfo.TextMatrix(llRow, STARTTIMEINDEX))
                gPackTime slStr, tmPopPaf(llIndex).iStartTime(0), tmPopPaf(llIndex).iStartTime(1)
                slStr = Trim$(grdPrgAirInfo.TextMatrix(llRow, ENDTIMEINDEX))
                gPackTime slStr, tmPopPaf(llIndex).iEndTime(0), tmPopPaf(llIndex).iEndTime(1)
                tmPopPaf(llIndex).iUrfCode = tgUrf(0).iCode
                ReDim Preserve tmPopPaf(0 To UBound(tmPopPaf) + 1) As PAF
            End If
        End If
    Next llRow
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer record values to      *
'*                      controls on the screen         *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRet                                                                                 *
'******************************************************************************************

    Dim llRow As Long
    Dim ilCol As Integer
    Dim llLoop As Long
    Dim slStr As String

    grdPrgAirInfo.Redraw = False
    grdPrgAirInfo.Rows = imInitNoRows
    For llRow = grdPrgAirInfo.FixedRows To grdPrgAirInfo.Rows - 1 Step 1
        grdPrgAirInfo.RowHeight(llRow) = fgFlexGridRowH
        For ilCol = 0 To grdPrgAirInfo.Cols - 1 Step 1
            If ilCol = PAFCODEINDEX Then
                grdPrgAirInfo.TextMatrix(llRow, ilCol) = 0
            Else
                grdPrgAirInfo.TextMatrix(llRow, ilCol) = ""
            End If
        Next ilCol
    Next llRow
    llRow = grdPrgAirInfo.FixedRows

    For llLoop = 0 To UBound(tmPopPaf) - 1 Step 1
        If llRow >= grdPrgAirInfo.Rows Then
            grdPrgAirInfo.AddItem ""
            grdPrgAirInfo.RowHeight(llRow) = fgFlexGridRowH
            grdPrgAirInfo.Row = llRow
            For ilCol = 0 To grdPrgAirInfo.Cols - 1 Step 1
                grdPrgAirInfo.ColAlignment(ilCol) = flexAlignLeftCenter
            Next ilCol
        End If
        grdPrgAirInfo.Row = llRow
        grdPrgAirInfo.TextMatrix(llRow, NAMEINDEX) = Trim$(tmPopPaf(llLoop).sName)
        gUnpackDate tmPopPaf(llLoop).iStartDate(0), tmPopPaf(llLoop).iStartDate(1), slStr
        grdPrgAirInfo.TextMatrix(llRow, STARTDATEINDEX) = slStr
        gUnpackDate tmPopPaf(llLoop).iEndDate(0), tmPopPaf(llLoop).iEndDate(1), slStr
        If gDateValue(slStr) <> gDateValue("12/31/2069") Then
            grdPrgAirInfo.TextMatrix(llRow, ENDDATEINDEX) = slStr
        Else
            grdPrgAirInfo.TextMatrix(llRow, ENDDATEINDEX) = ""
        End If
        grdPrgAirInfo.Col = MOINDEX
        grdPrgAirInfo.CellFontBold = False
        grdPrgAirInfo.CellFontName = "Monotype Sorts"
        If tmPopPaf(llLoop).sMo = "Y" Then
            grdPrgAirInfo.TextMatrix(llRow, MOINDEX) = " 4"
        Else
            grdPrgAirInfo.TextMatrix(llRow, MOINDEX) = ""
        End If
        grdPrgAirInfo.Col = TUINDEX
        grdPrgAirInfo.CellFontBold = False
        grdPrgAirInfo.CellFontName = "Monotype Sorts"
        If tmPopPaf(llLoop).sTu = "Y" Then
            grdPrgAirInfo.TextMatrix(llRow, TUINDEX) = " 4"
        Else
            grdPrgAirInfo.TextMatrix(llRow, TUINDEX) = ""
        End If
        grdPrgAirInfo.Col = WEINDEX
        grdPrgAirInfo.CellFontBold = False
        grdPrgAirInfo.CellFontName = "Monotype Sorts"
        If tmPopPaf(llLoop).sWe = "Y" Then
            grdPrgAirInfo.TextMatrix(llRow, WEINDEX) = " 4"
        Else
            grdPrgAirInfo.TextMatrix(llRow, WEINDEX) = ""
        End If
        grdPrgAirInfo.Col = THINDEX
        grdPrgAirInfo.CellFontBold = False
        grdPrgAirInfo.CellFontName = "Monotype Sorts"
        If tmPopPaf(llLoop).sTh = "Y" Then
            grdPrgAirInfo.TextMatrix(llRow, THINDEX) = " 4"
        Else
            grdPrgAirInfo.TextMatrix(llRow, THINDEX) = ""
        End If
        grdPrgAirInfo.Col = FRINDEX
        grdPrgAirInfo.CellFontBold = False
        grdPrgAirInfo.CellFontName = "Monotype Sorts"
        If tmPopPaf(llLoop).sFr = "Y" Then
            grdPrgAirInfo.TextMatrix(llRow, FRINDEX) = " 4"
        Else
            grdPrgAirInfo.TextMatrix(llRow, FRINDEX) = ""
        End If
        grdPrgAirInfo.Col = SAINDEX
        grdPrgAirInfo.CellFontBold = False
        grdPrgAirInfo.CellFontName = "Monotype Sorts"
        If tmPopPaf(llLoop).sSa = "Y" Then
            grdPrgAirInfo.TextMatrix(llRow, SAINDEX) = " 4"
        Else
            grdPrgAirInfo.TextMatrix(llRow, SAINDEX) = ""
        End If
        grdPrgAirInfo.Col = SUINDEX
        grdPrgAirInfo.CellFontBold = False
        grdPrgAirInfo.CellFontName = "Monotype Sorts"
        If tmPopPaf(llLoop).sSu = "Y" Then
            grdPrgAirInfo.TextMatrix(llRow, SUINDEX) = " 4"
        Else
            grdPrgAirInfo.TextMatrix(llRow, SUINDEX) = ""
        End If
        gUnpackTime tmPopPaf(llLoop).iStartTime(0), tmPopPaf(llLoop).iStartTime(1), "A", "1", slStr
        grdPrgAirInfo.TextMatrix(llRow, STARTTIMEINDEX) = slStr
        gUnpackTime tmPopPaf(llLoop).iEndTime(0), tmPopPaf(llLoop).iEndTime(1), "A", "1", slStr
        grdPrgAirInfo.TextMatrix(llRow, ENDTIMEINDEX) = slStr
        grdPrgAirInfo.TextMatrix(llRow, PAFCODEINDEX) = tmPopPaf(llLoop).lCode
        grdPrgAirInfo.TextMatrix(llRow, CHGINDEX) = "N"
        llRow = llRow + 1
    Next llLoop
    If llRow >= grdPrgAirInfo.Rows Then
        grdPrgAirInfo.AddItem ""
        grdPrgAirInfo.RowHeight(llRow) = fgFlexGridRowH
        grdPrgAirInfo.TextMatrix(llRow, PAFCODEINDEX) = 0
        grdPrgAirInfo.TextMatrix(llRow, CHGINDEX) = "N"
        grdPrgAirInfo.Row = grdPrgAirInfo.Rows - 1
        For ilCol = 0 To grdPrgAirInfo.Cols - 1 Step 1
            grdPrgAirInfo.ColAlignment(ilCol) = flexAlignLeftCenter
        Next ilCol
    End If

    'Remove highlight
    mSortCol STARTTIMEINDEX
    mSortCol NAMEINDEX
    grdPrgAirInfo.Row = 0
    grdPrgAirInfo.Col = PAFCODEINDEX
    grdPrgAirInfo.Redraw = True
    mSetCommands

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGridFieldsOk                   *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mGridFieldsOk(ilRowNo As Integer) As Integer
'
'   iRet = mGridFieldsOk()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilError As Integer
    Dim slStr As String

    ilError = False
    slStr = Trim$(grdPrgAirInfo.TextMatrix(ilRowNo, NAMEINDEX))
    If slStr <> "" Then
        slStartDate = Trim$(grdPrgAirInfo.TextMatrix(ilRowNo, STARTDATEINDEX))
        If slStartDate <> "" Then
            If gValidDate(slStartDate) = False Then
                ilError = True
                grdPrgAirInfo.Row = ilRowNo
                grdPrgAirInfo.Col = STARTDATEINDEX
                grdPrgAirInfo.CellForeColor = vbRed
            End If
        Else
            grdPrgAirInfo.TextMatrix(ilRowNo, STARTDATEINDEX) = "Missing"
            ilError = True
            grdPrgAirInfo.Row = ilRowNo
            grdPrgAirInfo.Col = STARTDATEINDEX
            grdPrgAirInfo.CellForeColor = vbRed
        End If
        slEndDate = Trim$(grdPrgAirInfo.TextMatrix(ilRowNo, ENDDATEINDEX))
        If slEndDate <> "" Then
            If gValidDate(slEndDate) = False Then
                ilError = True
                grdPrgAirInfo.Row = ilRowNo
                grdPrgAirInfo.Col = ENDDATEINDEX
                grdPrgAirInfo.CellForeColor = vbRed
            Else
                If slStartDate <> "" Then
                    If gDateValue(slEndDate) < gDateValue(slStartDate) Then
                        ilError = True
                        grdPrgAirInfo.Row = ilRowNo
                        grdPrgAirInfo.Col = ENDDATEINDEX
                        grdPrgAirInfo.CellForeColor = vbRed
                    End If
                End If
            End If
        End If
        slStartTime = Trim$(grdPrgAirInfo.TextMatrix(ilRowNo, STARTTIMEINDEX))
        If slStartTime <> "" Then
            If gValidTime(slStartTime) = False Then
                ilError = True
                grdPrgAirInfo.Row = ilRowNo
                grdPrgAirInfo.Col = STARTTIMEINDEX
                grdPrgAirInfo.CellForeColor = vbRed
            End If
        End If
        slEndTime = Trim$(grdPrgAirInfo.TextMatrix(ilRowNo, ENDTIMEINDEX))
        If slEndTime <> "" Then
            If gValidTime(slEndTime) = False Then
                ilError = True
                grdPrgAirInfo.Row = ilRowNo
                grdPrgAirInfo.Col = ENDTIMEINDEX
                grdPrgAirInfo.CellForeColor = vbRed
            Else
                If slStartTime <> "" Then
                    If gTimeToLong(slEndTime, True) < gTimeToLong(slStartTime, False) Then
                        ilError = True
                        grdPrgAirInfo.Row = ilRowNo
                        grdPrgAirInfo.Col = ENDTIMEINDEX
                        grdPrgAirInfo.CellForeColor = vbRed
                    End If
                End If
            End If
        End If

    End If
    If ilError Then
        mGridFieldsOk = False
    Else
        mGridFieldsOk = True
    End If
End Function

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
                    Select Case lmEnableCol
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

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestRecLengths                 *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if record lengths match    *
'*                                                     *
'*******************************************************
Private Function mTestRecLengths() As Integer
    Dim ilSizeError As Integer
    Dim ilSize As Integer
    ilSizeError = False
    ilSize = mGetRecLength("Paf.Btr")
    If ilSize <> Len(tmPaf) Then
        If ilSize > 0 Then
            MsgBox "Paf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmPaf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Paf error: " & str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    mTestRecLengths = ilSizeError
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetRecLength                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the record length from   *
'*                     the database                    *
'*                                                     *
'*******************************************************
Private Function mGetRecLength(slFileName As String) As Integer
'
'   ilRecLen = mGetRecLength(slName)
'   Where:
'       slName (I)- Name of the file
'       ilRecLen (O)- record length within the file
'
    Dim hlFile As Integer
    Dim ilRet As Integer
    hlFile = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mGetRecLength = -ilRet
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        Exit Function
    End If
    mGetRecLength = btrRecordLength(hlFile)  'Get and save record length
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
End Function

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    If imTerminate Then
        mTerminate
    End If
End Sub
