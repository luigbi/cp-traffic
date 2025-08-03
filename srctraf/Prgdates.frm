VERSION 5.00
Begin VB.Form PrgDates 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3900
   ClientLeft      =   840
   ClientTop       =   2190
   ClientWidth     =   6045
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
   ScaleHeight     =   3900
   ScaleWidth      =   6045
   Begin VB.PictureBox pbcDay 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   2400
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   210
      Begin VB.CheckBox ckcDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   15
         TabIndex        =   7
         Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
         Top             =   15
         Width           =   180
      End
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   915
      Top             =   3405
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
      Left            =   1200
      Picture         =   "Prgdates.frx":0000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   570
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox pbcViewType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4965
      ScaleHeight     =   180
      ScaleWidth      =   855
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   45
      Width           =   885
   End
   Begin VB.VScrollBar vbcDates 
      Height          =   2550
      LargeChange     =   11
      Left            =   5550
      TabIndex        =   18
      Top             =   360
      Width           =   240
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   255
      MaxLength       =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   795
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
      Left            =   540
      ScaleHeight     =   1380
      ScaleWidth      =   1095
      TabIndex        =   5
      Top             =   1035
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
         Picture         =   "Prgdates.frx":00FA
         ScaleHeight     =   1305
         ScaleWidth      =   1020
         TabIndex        =   9
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
            Picture         =   "Prgdates.frx":0DB8
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
      Left            =   3405
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   10
      Top             =   795
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
         Picture         =   "Prgdates.frx":10C2
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
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   15
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
         Left            =   1635
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   45
      Picture         =   "Prgdates.frx":3EDC
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   615
      Visible         =   0   'False
      Width           =   105
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
      Height          =   2565
      Left            =   240
      Picture         =   "Prgdates.frx":41E6
      ScaleHeight     =   2565
      ScaleWidth      =   5340
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   330
      Width           =   5340
      Begin VB.Label lacFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   23
         Top             =   180
         Visible         =   0   'False
         Width           =   5310
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3585
      Width           =   75
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   45
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   17
      Top             =   3015
      Width           =   45
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   30
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   1
      Top             =   375
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
      Left            =   3060
      TabIndex        =   20
      Top             =   3525
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
      Left            =   1845
      TabIndex        =   19
      Top             =   3525
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
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   15
      Width           =   2475
   End
   Begin VB.PictureBox plcDates 
      ForeColor       =   &H00000000&
      Height          =   2670
      Left            =   180
      ScaleHeight     =   2610
      ScaleWidth      =   5610
      TabIndex        =   2
      Top             =   300
      Width           =   5670
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   3420
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   5175
      Picture         =   "Prgdates.frx":BB58
      Top             =   3270
      Width           =   480
   End
End
Attribute VB_Name = "PrgDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Prgdates.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PrgDates.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Program library dates input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim tmCtrls(0 To 10)  As FIELDAREA
Dim imLBCtrls As Integer
Dim imBoxNo As Integer   'Current Media Box
Dim imRowNo As Integer      'Current row number in Program area (start at 0)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imSettingValue As Integer
Dim imBypassFocus As Integer
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
Dim smShow() As String  'Values shown in date/time area
Dim smSave() As String  'Values saved (1=Start time; 2=Start date; 3=End date)
Dim imSave() As Integer 'Values saved (1=Monday; 2=Tuesday;.. 7=Sunday)
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
'Mouse down

Const LBONE = 1

Const STARTTIMEINDEX = 1    'Start time control/field
Const DAYINDEX = 2          'Day control/field
Const STARTDATEINDEX = 9    'Start date control/field
Const ENDDATEINDEX = 10     'End date control/field
Private Sub ckcDay_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
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
    ReDim tgPrg(0 To 0) As PRGDATE  'Time/Dates
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    lacFrame.Visible = False
    pbcArrow.Visible = False
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDone_Click()
    Dim ilRow As Integer
    For ilRow = LBONE To UBound(smSave, 2) - 1 Step 1
        imRowNo = ilRow
        If mTestSaveFields() = NO Then
            mEnableBox imBoxNo
            Exit Sub
        End If
    Next ilRow
    mXFerCtrlToRec
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    lacFrame.Visible = False
    pbcArrow.Visible = False
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case STARTTIMEINDEX
            plcTme.Visible = Not plcTme.Visible
        Case STARTDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case ENDDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Select Case imBoxNo
        Case STARTTIMEINDEX
        Case STARTDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case ENDDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
    End Select
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case STARTTIMEINDEX
        Case STARTDATEINDEX
        Case ENDDATEINDEX
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
        Case STARTDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case ENDDATEINDEX
            'Disallow TFN for alternate
            If (Len(edcDropDown.Text) = edcDropDown.SelLength) And (igViewType = 0) And (igLibType = 0) Then
                If (KeyAscii = Asc("T")) Or (KeyAscii = Asc("t")) Then
                    edcDropDown.Text = "TFN"
                    edcDropDown.SelStart = 0
                    edcDropDown.SelLength = 3
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
    End Select
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case STARTTIMEINDEX
                If (Shift And vbAltMask) > 0 Then
                    plcTme.Visible = Not plcTme.Visible
                End If
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
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imBoxNo
            Case STARTTIMEINDEX
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
        End Select
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
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Erase smShow
    Erase smSave
    Erase imSave
    Set PrgDates = Nothing   'Remove data segment

End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub imcTrash_Click()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilUpperBound As Integer
    Dim ilRowNo As Integer
    Dim slStr As String
    If (imRowNo < 1) Then
        Exit Sub
    End If
    ilRowNo = imRowNo
    mSetShow imBoxNo
    imBoxNo = -1   '
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
    ilUpperBound = UBound(smSave, 2)
    If ilRowNo = ilUpperBound Then
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            slStr = ""
            gSetShow pbcDates, slStr, tmCtrls(ilLoop)
            smShow(ilLoop, ilRowNo) = tmCtrls(ilLoop).sShow
        Next ilLoop
        pbcDates_Paint
        mInitNew ilRowNo   'Set defaults for extra row
    Else
        For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
            For ilIndex = 1 To UBound(smSave, 1) Step 1
                smSave(ilIndex, ilLoop) = smSave(ilIndex, ilLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(imSave, 1) Step 1
                imSave(ilIndex, ilLoop) = imSave(ilIndex, ilLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(smShow, 1) Step 1
                smShow(ilIndex, ilLoop) = smShow(ilIndex, ilLoop + 1)
            Next ilIndex
        Next ilLoop
        ilUpperBound = UBound(smSave, 2)
        ReDim Preserve smSave(0 To 3, 0 To ilUpperBound - 1) As String
        ReDim Preserve imSave(0 To 7, 0 To ilUpperBound - 1) As Integer
        ReDim Preserve smShow(0 To ENDDATEINDEX, 0 To ilUpperBound - 1) As String
    End If
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcDates.Cls
    pbcDates_Paint
End Sub
Private Sub imcTrash_DragDrop(Source As control, X As Single, Y As Single)
'    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        lacFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
'        lacEvtFrame.DragIcon = IconTraf!imcIconTrash.DragIcon
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub
Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
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
    If (imRowNo < vbcDates.Value) Or (imRowNo >= vbcDates.Value + vbcDates.LargeChange + 1) Then
        mSetShow ilBoxNo
        pbcArrow.Visible = False
        lacFrame.Visible = False
        Exit Sub
    End If
    lacFrame.Visible = False
    lacFrame.Move 0, tmCtrls(STARTTIMEINDEX).fBoxY + (imRowNo - vbcDates.Value) * (fgBoxGridH + 15) - 30
    lacFrame.Visible = True
    pbcArrow.Visible = False
    pbcArrow.Move pbcArrow.Left, plcDates.Top + tmCtrls(STARTTIMEINDEX).fBoxY + (imRowNo - vbcDates.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case STARTTIMEINDEX 'Start time
            edcDropDown.Width = tmCtrls(STARTTIMEINDEX).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcDates, edcDropDown, tmCtrls(STARTTIMEINDEX).fBoxX, tmCtrls(STARTTIMEINDEX).fBoxY + (imRowNo - vbcDates.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcTme.Height < cmcDone.Top Then
                plcTme.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcTme.Move edcDropDown.Left, edcDropDown.Top - plcTme.Height
            End If
            edcDropDown.Text = smSave(1, imRowNo)
            edcDropDown.Visible = True  'Set visibility
            cmcDropDown.Visible = True
            If smSave(1, imRowNo) = "" Then
                plcTme.Visible = True
            End If
            edcDropDown.SetFocus
        Case DAYINDEX To DAYINDEX + 6 'Day index
            gMoveTableCtrl pbcDates, pbcDay, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY + (imRowNo - vbcDates.Value) * (fgBoxGridH + 15)
            If imSave(ilBoxNo - DAYINDEX + 1, imRowNo) = 1 Then
                ckcDay.Value = vbChecked
            Else
                ckcDay.Value = vbUnchecked
            End If
            pbcDay.Visible = True  'Set visibility
            ckcDay.SetFocus
        Case STARTDATEINDEX 'Start date
            edcDropDown.Width = tmCtrls(STARTDATEINDEX).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcDates, edcDropDown, tmCtrls(STARTDATEINDEX).fBoxX, tmCtrls(STARTDATEINDEX).fBoxY + (imRowNo - vbcDates.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcCalendar.Height < cmcDone.Top Then
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                plcCalendar.Move edcDropDown.Left, edcDropDown.Top - plcCalendar.Height
            End If
            If smSave(2, imRowNo) = "" Then
                slStr = gObtainMondayFromToday()
            Else
                slStr = smSave(2, imRowNo)
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            If smSave(2, imRowNo) = "" Then
                pbcCalendar.Visible = True
            End If
            edcDropDown.SetFocus
        Case ENDDATEINDEX 'Start date
            edcDropDown.Width = tmCtrls(ENDDATEINDEX).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcDates, edcDropDown, tmCtrls(ENDDATEINDEX).fBoxX, tmCtrls(ENDDATEINDEX).fBoxY + (imRowNo - vbcDates.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            If edcDropDown.Top + edcDropDown.Height + plcCalendar.Height < cmcDone.Top Then
                plcCalendar.Move edcDropDown.Left + cmcDropDown.Width + edcDropDown.Width - plcCalendar.Width, edcDropDown.Top + edcDropDown.Height
            Else
                plcCalendar.Move edcDropDown.Left + cmcDropDown.Width + edcDropDown.Width - plcCalendar.Width, edcDropDown.Top - plcCalendar.Height
            End If
            If smSave(3, imRowNo) <> "" Then
                slStr = smSave(3, imRowNo)
                gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                pbcCalendar_Paint
            Else
                If (igViewType = 0) And (igLibType = 0) Then  'on Air
                    'Preset calendar
                    If smSave(2, imRowNo) <> "" Then
                        slStr = smSave(2, imRowNo)
                        gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                        pbcCalendar_Paint
                    End If
                    slStr = "TFN"
                Else    'Alternate
                    If smSave(2, imRowNo) <> "" Then
                        slStr = smSave(2, imRowNo)
                        gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                        pbcCalendar_Paint
                    Else
                        'Preset calendar
                        If smSave(2, imRowNo) <> "" Then
                            slStr = smSave(2, imRowNo)
                            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
                            pbcCalendar_Paint
                        End If
                        slStr = ""
                    End If
                End If
            End If
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
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
    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    imLBCDCtrls = 1
    imFirstActivate = True
    imTerminate = False
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    pbcViewType.Enabled = False
    imBypassFocus = False
    imSettingValue = False
    imBoxNo = -1 'Initialize current Box to N/A
    imChgMode = False
    imBSMode = False
    imCalType = 0   'Standard
    mInitBox
    mXFerRecToCtrl
    PrgDates.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone PrgDates
    'gCenterModalForm PrgDates
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = IconTraf!imcHelp.Picture
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
    plcDates.Move 189, 285, pbcDates.Width + vbcDates.Width + fgPanelAdj, pbcDates.Height + fgPanelAdj
    pbcDates.Move plcDates.Left + fgBevelX, plcDates.Top + fgBevelY
    vbcDates.Move pbcDates.Left + pbcDates.Width + 15, pbcDates.Top
    pbcArrow.Move plcDates.Left - pbcArrow.Width - 15   'Set arrow
    'Start Time
    gSetCtrl tmCtrls(STARTTIMEINDEX), 30, 225, 1215, fgBoxGridH
    tmCtrls(STARTTIMEINDEX).iReq = False
    'Days of the week
    For ilLoop = 0 To 6 Step 1
        gSetCtrl tmCtrls(DAYINDEX + ilLoop), 1260 + 225 * (ilLoop), tmCtrls(STARTTIMEINDEX).fBoxY, 210, fgBoxGridH
        tmCtrls(DAYINDEX + ilLoop).iReq = False
    Next ilLoop
    'Start Date
    gSetCtrl tmCtrls(STARTDATEINDEX), 2835, tmCtrls(STARTTIMEINDEX).fBoxY, 1215, fgBoxGridH
    tmCtrls(STARTDATEINDEX).iReq = False
    'End Date
    gSetCtrl tmCtrls(ENDDATEINDEX), 4065, tmCtrls(STARTTIMEINDEX).fBoxY, 1215, fgBoxGridH
    tmCtrls(ENDDATEINDEX).iReq = False
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNew                        *
'*                                                     *
'*             Created:9/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize values              *
'*                                                     *
'*******************************************************
Private Sub mInitNew(ilRowNo As Integer)
    Dim ilLoop As Integer

    smSave(1, ilRowNo) = ""   'Start Time
    smSave(2, ilRowNo) = ""   'Start date
    smSave(3, ilRowNo) = ""   'End date
    For ilLoop = 1 To 7 Step 1
        imSave(ilLoop, ilRowNo) = 0
    Next ilLoop
    For ilLoop = STARTTIMEINDEX To ENDDATEINDEX Step 1
        smShow(ilLoop, ilRowNo) = ""
    Next ilLoop
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
    lacFrame.Visible = False
    pbcArrow.Visible = False
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo 'Branch on box type (control)
        Case STARTTIMEINDEX
            cmcDropDown.Visible = False
            plcTme.Visible = False
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If gValidTime(slStr) Then
                    smSave(1, imRowNo) = slStr
                    slStr = gFormatTime(slStr, "A", "1")
                Else
                    Beep
                    edcDropDown.Text = smSave(1, imRowNo)
                    slStr = smSave(1, imRowNo)
                End If
            End If
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
        Case DAYINDEX To DAYINDEX + 6 'Day
            pbcDay.Visible = False  'Set visibility
            If ckcDay.Value = vbChecked Then
                slStr = "4"
                imSave(ilBoxNo - DAYINDEX + 1, imRowNo) = 1
            Else
                slStr = "  "
                imSave(ilBoxNo - DAYINDEX + 1, imRowNo) = 0
            End If
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
        Case STARTDATEINDEX 'Start Date
            plcCalendar.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            If gValidDate(slStr) Then
                smSave(2, imRowNo) = slStr
                slStr = gFormatDate(slStr)
                gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
                smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
            Else
                Beep
                edcDropDown.Text = smSave(2, imRowNo)
            End If
        Case ENDDATEINDEX 'End Date
            plcCalendar.Visible = False
            cmcDropDown.Visible = False
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            If StrComp(slStr, "TFN", 1) <> 0 Then
                If gValidDate(slStr) Then
                    smSave(3, imRowNo) = slStr
                    slStr = gFormatDate(slStr)
                Else
                    Beep
                    edcDropDown.Text = smSave(3, imRowNo)
                    slStr = smSave(3, imRowNo)
                End If
            Else
                smSave(3, imRowNo) = ""
                slStr = "TFN"
            End If
            gSetShow pbcDates, slStr, tmCtrls(ilBoxNo)
            smShow(ilBoxNo, imRowNo) = tmCtrls(ilBoxNo).sShow
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
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload PrgDates
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
    Dim ilLoop As Integer
    Dim ilOneYes As Integer
    Dim slDate As String
    Dim llStartTime As Long
    If smSave(1, imRowNo) = "" Then
        ilRes = MsgBox("Start time must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = STARTTIMEINDEX
        mTestSaveFields = NO
        Exit Function
    Else
        If Not gValidTime(smSave(1, imRowNo)) Then
            ilRes = MsgBox("Time must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = STARTTIMEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
        llStartTime = CLng(gTimeToCurrency(smSave(1, imRowNo), False))
        If llStartTime + lgLibLength > 86400 Then
            ilRes = MsgBox("Library end time exceeds 12Midnight", vbOKOnly + vbExclamation, "Error")
            imBoxNo = STARTTIMEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If smSave(2, imRowNo) = "" Then
        ilRes = MsgBox("Start date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = STARTDATEINDEX
        mTestSaveFields = NO
        Exit Function
    Else
        If Not gValidDate(smSave(2, imRowNo)) Then
            ilRes = MsgBox("Start date must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = STARTDATEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If lgEarliestDateViaPrg > 0 Then
        If gDateValue(smSave(2, imRowNo)) < lgEarliestDateViaPrg Then
            slDate = Format$(lgEarliestDateViaPrg, "m/d/yy")
            ilRes = MsgBox("Start date must be after " & slDate, vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = STARTDATEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    If (smSave(3, imRowNo) = "") And ((igViewType <> 0) Or (igLibType <> 0)) Then
        ilRes = MsgBox("End date must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = STARTDATEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    If smSave(3, imRowNo) <> "" Then
        If Not gValidDate(smSave(3, imRowNo)) Then
            ilRes = MsgBox("End date must be specified correctly", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = ENDDATEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
        If gDateValue(smSave(2, imRowNo)) > gDateValue(smSave(3, imRowNo)) Then
            ilRes = MsgBox("End date must be after start date", vbOKOnly + vbExclamation, "Incomplete")
            imBoxNo = ENDDATEINDEX
            mTestSaveFields = NO
            Exit Function
        End If
    End If
    ilOneYes = False
    For ilLoop = 1 To 7 Step 1
        If imSave(ilLoop, imRowNo) = 1 Then
            ilOneYes = True
            Exit For
        End If
    Next ilLoop
    If Not ilOneYes Then
        ilRes = MsgBox("One Day must be specified", vbOKOnly + vbExclamation, "Incomplete")
        imBoxNo = DAYINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    'If Selling or Airing (then M-F and/or Sa and/or Su must be selected)
    If (sgVefTypeViaPrg = "S") Or (sgVefTypeViaPrg = "A") Or (sgVefTypeViaPrg = "CF") Then
        ilOneYes = False
        For ilLoop = 1 To 5 Step 1
            If imSave(ilLoop, imRowNo) = 1 Then
                ilOneYes = True
                Exit For
            End If
        Next ilLoop
        If ilOneYes Then
            For ilLoop = 1 To 5 Step 1
                If imSave(ilLoop, imRowNo) <> 1 Then
                    ilRes = MsgBox("Mo-Fr must be specified since one is selected", vbOKOnly + vbExclamation, "Incomplete")
                    imBoxNo = DAYINDEX
                    mTestSaveFields = NO
                    Exit Function
                End If
            Next ilLoop
        End If
    End If
    mTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mXFerCtrlToRec                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer controls to record    *
'*                                                     *
'*******************************************************
Private Sub mXFerCtrlToRec()
'
'   mXFerCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    ReDim tgPrg(0 To UBound(smSave, 2) - 1) As PRGDATE
    For ilLoop = LBONE To UBound(smSave, 2) - 1 Step 1
        tgPrg(ilLoop - 1).sStartTime = smSave(1, ilLoop)
        For ilIndex = 0 To 6 Step 1
            tgPrg(ilLoop - 1).iDay(ilIndex) = imSave(ilIndex + 1, ilLoop)
        Next ilIndex
        tgPrg(ilLoop - 1).sStartDate = smSave(2, ilLoop)
        tgPrg(ilLoop - 1).sEndDate = smSave(3, ilLoop)
    Next ilLoop
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mXFerRecToCtrl                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer record values to      *
'*                      controls on the screen         *
'*                                                     *
'*******************************************************
Private Sub mXFerRecToCtrl()
'
'   mXFerRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slStr As String
    ReDim smSave(0 To 3, 0 To UBound(tgPrg) + 1) As String
    ReDim imSave(0 To 7, 0 To UBound(tgPrg) + 1) As Integer
    ReDim smShow(0 To ENDDATEINDEX, 0 To UBound(tgPrg) + 1) As String
    For ilLoop = 0 To UBound(tgPrg) - 1 Step 1
        smSave(1, ilLoop + 1) = tgPrg(ilLoop).sStartTime
        slStr = gFormatTime(smSave(1, ilLoop + 1), "A", "1")
        gSetShow pbcDates, slStr, tmCtrls(STARTTIMEINDEX)
        smShow(STARTTIMEINDEX, ilLoop + 1) = tmCtrls(STARTTIMEINDEX).sShow
        For ilIndex = 0 To 6 Step 1
            imSave(ilIndex + 1, ilLoop + 1) = tgPrg(ilLoop).iDay(ilIndex)
            If imSave(ilIndex + 1, ilLoop + 1) = 1 Then 'Yes
                smShow(DAYINDEX + ilIndex, ilLoop + 1) = "4"
            Else
                smShow(DAYINDEX + ilIndex, ilLoop + 1) = " "
            End If
        Next ilIndex
        smSave(2, ilLoop + 1) = tgPrg(ilLoop).sStartDate
        slStr = gFormatDate(smSave(2, ilLoop + 1))
        gSetShow pbcDates, slStr, tmCtrls(STARTDATEINDEX)
        smShow(STARTDATEINDEX, ilLoop + 1) = tmCtrls(STARTDATEINDEX).sShow
        smSave(3, ilLoop + 1) = tgPrg(ilLoop).sEndDate
        If smSave(3, ilLoop + 1) <> "" Then
            slStr = gFormatDate(smSave(3, ilLoop + 1))
        Else
            slStr = "TFN"
        End If
        gSetShow pbcDates, slStr, tmCtrls(ENDDATEINDEX)
        smShow(ENDDATEINDEX, ilLoop + 1) = tmCtrls(ENDDATEINDEX).sShow
    Next ilLoop
    vbcDates.Min = LBONE    'LBound(smShow, 2)
    If UBound(smShow, 2) - 1 <= vbcDates.LargeChange Then
        vbcDates.Max = LBONE    'LBound(smShow, 2)
    Else
        vbcDates.Max = UBound(smShow, 2) - vbcDates.LargeChange
    End If
    vbcDates.Value = vbcDates.Min
    mInitNew UBound(smSave, 2)
    Exit Sub
End Sub
Private Sub pbcArrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
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
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    lacFrame.Visible = False
    pbcArrow.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcClickFocus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcDates_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub pbcDates_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    ilCompRow = vbcDates.LargeChange + 1
    If UBound(smSave, 2) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smSave, 2)
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcDates.Value - 1
                    mSetShow imBoxNo
                    imRowNo = ilRowNo
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcDates_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim slFont As String
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer

    ilStartRow = vbcDates.Value  'Top location
    ilEndRow = vbcDates.Value + vbcDates.LargeChange
    If ilEndRow > UBound(smSave, 2) Then
        ilEndRow = UBound(smSave, 2) 'include blank row as it might have data
    End If
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            gPaintArea pbcDates, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
            pbcDates.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            If (ilBox >= DAYINDEX) And (ilBox <= DAYINDEX + 6) Then
                pbcDates.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) + 15 '+ fgBoxInsetY
                slFont = pbcDates.FontName
                pbcDates.FontName = "Monotype Sorts"
                pbcDates.FontBold = False
                pbcDates.Print smShow(ilBox, ilRow)
                pbcDates.FontName = slFont
                pbcDates.FontBold = True
            Else
                pbcDates.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                pbcDates.Print smShow(ilBox, ilRow)
            End If
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imSettingValue = True
            vbcDates.Value = vbcDates.Min
            imSettingValue = False
            If UBound(smSave, 2) <= vbcDates.LargeChange + 1 Then 'was <=
                vbcDates.Max = LBONE    'LBound(smSave, 2)
            Else
                vbcDates.Max = UBound(smSave, 2) - vbcDates.LargeChange ' - 1
            End If
            imRowNo = 1
            If (imRowNo = UBound(smSave, 2)) And (smSave(1, 1) = "") Then
                mInitNew imRowNo
            End If
            imSettingValue = True
            vbcDates.Value = vbcDates.Min
            imSettingValue = False
            ilBox = 1
            imBoxNo = ilBox
            mEnableBox ilBox
            Exit Sub
        Case STARTTIMEINDEX 'Time (first control within header)
            slStr = edcDropDown.Text
            If Not gValidTime(slStr) Then
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            mSetShow imBoxNo
            ilBox = ENDDATEINDEX
'            Do
                If imRowNo <= 1 Then
                    imBoxNo = -1
                    imRowNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imRowNo = imRowNo - 1
                If imRowNo < vbcDates.Value Then
                    imSettingValue = True
                    vbcDates.Value = vbcDates.Value - 1
                    imSettingValue = False
                End If
'            Loop While (smIBSave(1, imIBRowNo) = "I") Or (smIBSave(7, imIBRowNo) = "R") Or (smIBSave(7, imIBRowNo) = "B")
            imBoxNo = ilBox
            mEnableBox ilBox
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
            ilBox = imBoxNo - 1
        Case ENDDATEINDEX
            slStr = edcDropDown.Text
            If (slStr <> "") And (StrComp(slStr, "TFN", 1) <> 0) Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            End If
            ilBox = imBoxNo - 1
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcSTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    Select Case imBoxNo
        Case -1 'Tab from control prior to form area
            imRowNo = UBound(smSave, 2)
            imSettingValue = True
            If imRowNo <= vbcDates.LargeChange + 1 Then
                vbcDates.Value = vbcDates.Min
            Else
                vbcDates.Value = imRowNo - vbcDates.LargeChange
            End If
            imSettingValue = False
            ilBox = 1
            Exit Sub
        Case STARTTIMEINDEX
            If (imRowNo >= UBound(smSave, 2)) And (edcDropDown.Text = "") Then
                mSetShow imBoxNo
                For ilLoop = STARTTIMEINDEX To ENDDATEINDEX Step 1
                    slStr = ""
                    gSetShow pbcDates, slStr, tmCtrls(ilLoop)
                    smShow(ilLoop, imRowNo) = tmCtrls(ilLoop).sShow
                Next ilLoop
                imBoxNo = -1
                imRowNo = -1
                pbcDates_Paint
                cmcDone.SetFocus
                Exit Sub
            End If
            slStr = edcDropDown.Text
            If Not gValidTime(slStr) Then
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imBoxNo + 1
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
            If (slStr <> "") And (StrComp(slStr, "TFN", 1) <> 0) Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            End If
            mSetShow imBoxNo
            If mTestSaveFields() = NO Then
                mEnableBox imBoxNo
                Exit Sub
            End If
            If imRowNo >= UBound(smSave, 2) Then
                ReDim Preserve smSave(0 To 3, 0 To imRowNo + 1) As String
                ReDim Preserve imSave(0 To 7, 0 To imRowNo + 1) As Integer
                ReDim Preserve smShow(0 To ENDDATEINDEX, 0 To imRowNo + 1) As String
                mInitNew imRowNo + 1
                If UBound(smSave, 2) <= vbcDates.LargeChange + 1 Then 'was <=
                    vbcDates.Max = LBONE    'LBound(smSave, 2)
                Else
                    vbcDates.Max = UBound(smSave, 2) - vbcDates.LargeChange ' - 1
                End If
            End If
'            Do
                imRowNo = imRowNo + 1
                If imRowNo > vbcDates.Value + vbcDates.LargeChange Then
                    imSettingValue = True
                    vbcDates.Value = vbcDates.Value + 1
                    imSettingValue = False
                End If
'            Loop While (smIBSave(1, imIBRowNo) = "I") Or (smIBSave(7, imIBRowNo) = "R") Or (smIBSave(7, imIBRowNo) = "B")
            If imRowNo >= UBound(smSave, 2) Then
                imBoxNo = 0
                lacFrame.Move 0, tmCtrls(STARTTIMEINDEX).fBoxY + (imRowNo - vbcDates.Value) * (fgBoxGridH + 15) - 30
                lacFrame.Visible = True
                pbcArrow.Move pbcArrow.Left, plcDates.Top + tmCtrls(STARTTIMEINDEX).fBoxY + (imRowNo - vbcDates.Value) * (fgBoxGridH + 15) + 45
                pbcArrow.Visible = True
                pbcArrow.SetFocus
                Exit Sub
            Else
                ilBox = 1
                imBoxNo = ilBox
                mEnableBox ilBox
            End If
            Exit Sub
        Case Else 'Last control within header
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
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
                    Select Case imBoxNo
                        Case STARTTIMEINDEX
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
Private Sub pbcViewType_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcViewType_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(" ") Then
        If igViewType = 0 Then
            igViewType = 1
        Else
            igViewType = 0
        End If
        pbcViewType_Paint
    End If
End Sub
Private Sub pbcViewType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    If igViewType = 0 Then
        igViewType = 1
    Else
        igViewType = 0
    End If
    pbcViewType_Paint
End Sub
Private Sub pbcViewType_Paint()
    pbcViewType.Cls
    pbcViewType.CurrentX = fgBoxInsetX
    pbcViewType.CurrentY = -15 'fgBoxInsetY
    If igViewType = 0 Then
        pbcViewType.Print "on Air"
    ElseIf igViewType = 1 Then
        pbcViewType.Print "Alternate"
    End If
End Sub
Private Sub plcDates_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcDates_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub tmcDrag_Timer()
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            ilCompRow = vbcDates.LargeChange + 1
            If UBound(smSave, 2) > ilCompRow Then
                ilMaxRow = ilCompRow
            Else
                ilMaxRow = UBound(smSave, 2)
            End If
            For ilRow = 1 To ilMaxRow Step 1
                If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(STARTTIMEINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(STARTTIMEINDEX).fBoxY + tmCtrls(STARTTIMEINDEX).fBoxH)) Then
                    'Only allow deletion of new- might want to be able to delete unbilled
'                    If (smIBSave(7, ilRow + vbcItemBill.Value - 1) = "R") Or (smIBSave(7, ilRow + vbcItemBill.Value - 1) = "B") Then
'                        Beep
'                        Exit Sub
'                    End If
                    mSetShow imBoxNo
                    imBoxNo = -1
                    imRowNo = -1
                    imRowNo = ilRow + vbcDates.Value - 1
                    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
                    lacFrame.Move 0, tmCtrls(STARTTIMEINDEX).fBoxY + (imRowNo - vbcDates.Value) * (fgBoxGridH + 15) - 30
                    'If gInvertArea call then remove visible setting
                    lacFrame.Visible = True
                    pbcArrow.Move pbcArrow.Left, plcDates.Top + tmCtrls(STARTTIMEINDEX).fBoxY + (imRowNo - vbcDates.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    imcTrash.Enabled = True
                    lacFrame.Drag vbBeginDrag
                    lacFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                    Exit Sub
                End If
            Next ilRow
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub vbcDates_Change()
    If imSettingValue Then
        pbcDates.Cls
        pbcDates_Paint
        imSettingValue = False
    Else
        mSetShow imBoxNo
        pbcDates.Cls
        pbcDates_Paint
        mEnableBox imBoxNo
    End If
End Sub
Private Sub vbcDates_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Program Library Event Dates"
End Sub
