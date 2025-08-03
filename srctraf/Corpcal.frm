VERSION 5.00
Begin VB.Form CorpCal 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4185
   ClientLeft      =   1035
   ClientTop       =   2025
   ClientWidth     =   4635
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4185
   ScaleWidth      =   4635
   Begin VB.ComboBox cbcSelect 
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
      Height          =   315
      Left            =   1770
      TabIndex        =   1
      Top             =   210
      Width           =   1800
   End
   Begin VB.CommandButton cmcSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3075
      TabIndex        =   17
      Top             =   3780
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   1755
      TabIndex        =   16
      Top             =   3780
      Width           =   1050
   End
   Begin VB.PictureBox pbcTab 
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
      Height          =   90
      Left            =   0
      ScaleHeight     =   90
      ScaleWidth      =   75
      TabIndex        =   13
      Top             =   930
      Width           =   75
   End
   Begin VB.PictureBox pbcSTab 
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
      Height          =   90
      Left            =   60
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   495
      Width           =   60
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
      Left            =   2010
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1995
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "Corpcal.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   1830
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   255
         Width           =   1860
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
         Width           =   270
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
         TabIndex        =   9
         Top             =   45
         Width           =   1290
      End
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
      Left            =   3285
      Picture         =   "Corpcal.frx":2E1A
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropDown 
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
      Left            =   630
      MaxLength       =   10
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox pbcCorpCal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   2940
      Left            =   600
      Picture         =   "Corpcal.frx":2F14
      ScaleHeight     =   2940
      ScaleWidth      =   2910
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   615
      Width           =   2910
   End
   Begin VB.PictureBox plcCorpCal 
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
      Height          =   3045
      Left            =   540
      ScaleHeight     =   2985
      ScaleWidth      =   2970
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   570
      Width           =   3030
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
      Left            =   2580
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3615
      Width           =   120
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   1755
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   1755
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   450
      TabIndex        =   14
      Top             =   3780
      Width           =   1050
   End
End
Attribute VB_Name = "CorpCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Corpcal.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CorpCal.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract model screen code
Option Explicit
Option Compare Text
Dim tmCtrls(0 To 15)  As FIELDAREA
Dim imLBCtrls As Integer
Dim tmDCtrls(0 To 24)  As FIELDAREA
Dim imLBDCtrls As Integer
Dim smSave(0 To 15) As String   '1=Year; 2=Start Month; 3=Start Date; 4-15=# Weeks/Month. Index zero ignored
Dim imBoxNo As Integer   'Current Avail Name Box
Dim hmCof As Integer
Dim tmCof As COF
Dim tmCofSrchKey As INTKEY0
Dim imCofRecLen As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBypassFocus As Integer
Dim imSelectedIndex As Integer
Dim imBypassSetting As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Const YEARINDEX = 1     'Name control/field
Const STARTMONTHINDEX = 2  'Sustain control/field
Const STARTDATEINDEX = 3  'Sponsorship control/field
Const NOWKS1INDEX = 4
Const STARTDATEWK1INDEX = 1  'Sustain control/field
Const ENDDATEWK1INDEX = 2  'Sponsorship control/field
Private Sub cbcSelect_Change()
    Dim ilLoop As Integer   'For loop control parameter
    Dim ilRet As Integer    'Return status
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    If imChgMode Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        Exit Sub
    End If
    imChgMode = True    'Set change mode to avoid infinite loop
    imBypassSetting = True
    Screen.MousePointer = vbHourglass  'Wait
    ilRet = gOptionLookAhead(cbcSelect, imBSMode, slStr)
    If ilRet = 0 Then
        ilIndex = cbcSelect.ListIndex
        If (ilIndex = 0) And (Not igUpdateOk) Then
            ilRet = 1   'Force clear
        Else
            If Not mReadRec(ilIndex, SETFORREADONLY) Then
                GoTo cbcSelectErr
            End If
        End If
    Else
        If ilRet = 1 Then
            cbcSelect.ListIndex = 0
        End If
        ilRet = 1   'Clear fields as no match name found
    End If
    pbcCorpCal.Cls
    If ilRet = 0 Then
        imSelectedIndex = cbcSelect.ListIndex
        mMoveRecToCtrl
    Else
        imSelectedIndex = 0
        mClearCtrlFields
        If slStr <> "[New]" Then
            edcDropDown.Text = slStr
        End If
    End If
    gSetShow pbcCorpCal, smSave(1), tmCtrls(YEARINDEX)
    gSetShow pbcCorpCal, smSave(2), tmCtrls(STARTMONTHINDEX)
    gSetShow pbcCorpCal, smSave(3), tmCtrls(STARTDATEINDEX)
    For ilLoop = 4 To 15 Step 1
        gSetShow pbcCorpCal, smSave(ilLoop), tmCtrls(NOWKS1INDEX + ilLoop - 4)
    Next ilLoop
    pbcCorpCal_Paint
    Screen.MousePointer = vbDefault
    imChgMode = False
    imBypassSetting = False
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    imChgMode = False
    Exit Sub
End Sub
Private Sub cbcSelect_Click()
    cbcSelect_Change    'Process change as change event is not generated by VB
End Sub
Private Sub cbcSelect_DropDown()
    'mPopulate
    'If imTerminate Then
    '    Exit Sub
    'End If
End Sub
Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    If imTerminate Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    slSvText = cbcSelect.Text
    'mPopulate
    'If imTerminate Then
    '    Exit Sub
    'End If
    gCtrlGotFocus cbcSelect
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields
        If pbcSTab.Enabled Then
            pbcSTab.SetFocus
        Else
            cmcCancel.SetFocus
        End If
        Exit Sub
    End If
    If (slSvText = "") Or (slSvText = "[New]") Then
        cbcSelect.ListIndex = 0
        cbcSelect_Change    'Call change so picture area repainted
    Else
        gFindMatch slSvText, 1, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
'            If (ilSvIndex <> gLastFound(cbcSelect)) Or (ilSvIndex <> cbcSelect.ListIndex) Then
            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Or imPopReqd Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
                cbcSelect_Change    'Call change so picture area repainted
                imPopReqd = False
            End If
        Else
            cbcSelect.ListIndex = 0
            mClearCtrlFields
            cbcSelect_Change    'Call change so picture area repainted
        End If
    End If
End Sub
Private Sub cbcSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcSelect_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcSelect.SelLength <> 0 Then    'avoid deleting two characters
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
    gCtrlGotFocus cmcCancel
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub cmcDone_Click()
    If igUpdateOk Then
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                cmcCancel_Click
                Exit Sub
            End If
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    Dim ilLoop As Integer
    If imBoxNo = -1 Then
        Exit Sub
    End If
    If Not igUpdateOk Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    If Not cmcSave.Enabled Then
        'Cycle to first unanswered mandatory
        For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
            If mSaveTestFields(ilLoop, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                imBoxNo = ilLoop
                mEnableBox imBoxNo
                Exit Sub
            End If
        Next ilLoop
    End If
    gCtrlGotFocus cmcDone
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case YEARINDEX
        Case STARTMONTHINDEX
        Case STARTDATEINDEX
            plcCalendar.Visible = Not plcCalendar.Visible
        Case NOWKS1INDEX To NOWKS1INDEX + 11
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcSave_Click()
    Dim imSvSelectedIndex As Integer
    If Not igUpdateOk Then
        Exit Sub
    End If
    imSvSelectedIndex = imSelectedIndex
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    imBoxNo = -1
    cbcSelect.Enabled = True
    'Must reset display so altered flag is cleared and setcommand will turn select on
    If imSvSelectedIndex <> 0 Then
        cbcSelect.Text = smSave(1)
    Else
        cbcSelect.ListIndex = 0
    End If
    cbcSelect_Change    'Call change so picture area repainted
    mSetCommands
    cbcSelect.SetFocus
End Sub
Private Sub cmcSave_GotFocus()
    gCtrlGotFocus cmcSave
    mSetShow imBoxNo
    imBoxNo = -1
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Select Case imBoxNo
        Case YEARINDEX
        Case STARTMONTHINDEX
        Case STARTDATEINDEX
            slStr = edcDropDown.Text
            If Not gValidDate(slStr) Then
                lacDate.Visible = False
                Exit Sub
            End If
            lacDate.Visible = True
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint   'mBoxCalDate called within paint
        Case NOWKS1INDEX To NOWKS1INDEX + 11
    End Select
    mSetChg imBoxNo
End Sub
Private Sub edcDropDown_GotFocus()
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
    Dim slStr As String
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case imBoxNo
        Case YEARINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case STARTMONTHINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            If gCompNumberStr(slStr, "12") > 0 Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case STARTDATEINDEX
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        Case NOWKS1INDEX To NOWKS1INDEX + 11
            'Filter characters (allow only BackSpace, numbers 0 thru 9
            If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
            slStr = edcDropDown.Text
            slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
            If gCompNumberStr(slStr, "6") > 0 Then
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
            Case YEARINDEX
            Case STARTMONTHINDEX
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
            Case NOWKS1INDEX To NOWKS1INDEX + 11
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case imBoxNo
            Case YEARINDEX
            Case STARTMONTHINDEX
            Case STARTDATEINDEX
                If (Shift And vbAltMask) > 0 Then
                Else
                    'slDate = edcDropDown.Text
                    'If gValidDate(slDate) Then
                    '    If KeyCode = KEYLEFT Then 'Up arrow
                    '        slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                    '    Else
                    '        slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                    '    End If
                    '    gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                    '    edcDropDown.Text = slDate
                    'End If
                End If
            Case NOWKS1INDEX To NOWKS1INDEX + 11
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
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        If cbcSelect.Enabled Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        If ilReSet Then
            cbcSelect.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
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
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear each control on the      *
'*                      screen                         *
'*                                                     *
'*******************************************************
Private Sub mClearCtrlFields()
'
'   mClearCtrlFields
'   Where:
'
    Dim ilLoop As Integer
    imBoxNo = -1
    edcDropDown.Text = ""
    For ilLoop = LBound(smSave) To UBound(smSave) Step 1
        smSave(ilLoop) = ""
    Next ilLoop
    'mMoveCtrlToRec False
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).sShow = ""
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    Dim ilPrevFd As Integer
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case YEARINDEX 'Name
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 4
            gMoveFormCtrl pbcCorpCal, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcDropDown.Text = smSave(1)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case STARTMONTHINDEX 'Name
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 2
            gMoveFormCtrl pbcCorpCal, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            If smSave(2) = "" Then
                ilPrevFd = False
                For ilLoop = LBound(tgMCof) To UBound(tgMCof) Step 1
                    If Val(smSave(1)) = tgMCof(ilLoop).iYear + 1 Then
                        ilPrevFd = True
                        edcDropDown.Text = Trim$(str$(tgMCof(ilLoop).iStartMnthNo))
                        Exit For
                    End If
                Next ilLoop
                If Not ilPrevFd Then
                    edcDropDown.Text = ""
                End If
            Else
                edcDropDown.Text = smSave(2)
            End If
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
        Case STARTDATEINDEX 'Name
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW - cmcDropDown.Width
            edcDropDown.MaxLength = 10
            gMoveFormCtrl pbcCorpCal, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            plcCalendar.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            If smSave(3) = "" Then
                ilPrevFd = False
                For ilLoop = LBound(tgMCof) To UBound(tgMCof) Step 1
                    If Val(smSave(1)) = tgMCof(ilLoop).iYear + 1 Then
                        ilPrevFd = True
                        gUnpackDate tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iEndDate(1, 11), slStr
                        slStr = gIncOneDay(slStr)
                        Exit For
                    End If
                Next ilLoop
                If Not ilPrevFd Then
                    If smSave(2) = "1" Then
                        slStr = "1/15/" & smSave(1)
                    Else
                        slStr = smSave(2) & "/15/" & gSubStr(smSave(1), "1")
                    End If
                    slStr = gObtainStartStd(slStr)
                End If
            Else
                slStr = smSave(3)
            End If
            gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
            pbcCalendar_Paint
            edcDropDown.Text = slStr
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            If smSave(3) = "" Then
                pbcCalendar.Visible = True
            End If
            edcDropDown.SetFocus
        Case NOWKS1INDEX To NOWKS1INDEX + 11
            edcDropDown.Width = tmCtrls(ilBoxNo).fBoxW
            edcDropDown.MaxLength = 1
            gMoveTableCtrl pbcCorpCal, edcDropDown, tmCtrls(ilBoxNo).fBoxX, tmCtrls(ilBoxNo).fBoxY
            edcDropDown.Text = smSave(ilBoxNo)
            edcDropDown.Visible = True  'Set visibility
            edcDropDown.SetFocus
    End Select
    mSetChg ilBoxNo 'set change flag encase the setting of the value didn't cause a change event
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
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
    imLBDCtrls = 1
    imLBCDCtrls = 1
    imFirstActivate = True
    imBypassFocus = False
    imChgMode = False
    imBSMode = False
    imSelectedIndex = -1
    imCalType = 0   'Standard
    imBypassSetting = False
    imPopReqd = False
    imBoxNo = -1
    mInitBox
    CorpCal.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone CorpCal
    hmCof = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCof, "", sgDBPath & "Cof.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", CorpCal
    On Error GoTo 0
    imCofRecLen = Len(tmCof)
'    Traffic!plcHelp.Caption = ""
    cbcSelect.Clear  'Force list to be populated
    mPopulate
    If Not igUpdateOk Then
        pbcCorpCal.Enabled = False
    End If
    If Not imTerminate Then
        cbcSelect.ListIndex = 0
    End If
    'cbcSelect.SetFocus
    Screen.MousePointer = vbDefault
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
    flTextHeight = pbcCorpCal.TextHeight("1") - 35
    'Position panel and picture areas with panel
    cbcSelect.Move 2385, 60
    plcCorpCal.Move 540, 570, pbcCorpCal.Width + fgPanelAdj, pbcCorpCal.Height + fgPanelAdj
    pbcCorpCal.Move plcCorpCal.Left + fgBevelX, plcCorpCal.Top + fgBevelY
    'Year
    gSetCtrl tmCtrls(YEARINDEX), 30, 30, 855, fgBoxStH
    'Start Month
    gSetCtrl tmCtrls(STARTMONTHINDEX), 900, tmCtrls(YEARINDEX).fBoxY, 990, fgBoxStH
    'Start Date
    gSetCtrl tmCtrls(STARTDATEINDEX), 1905, tmCtrls(YEARINDEX).fBoxY, 990, fgBoxStH
    '# of Weeks
    For ilLoop = 0 To 11 Step 1
        'gSetCrtl tmCtrls(NOWKS1INDEX), 30, 600, 855, fgBoxGridH
        gSetCtrl tmCtrls(NOWKS1INDEX + ilLoop), 30, 600 + ilLoop * (fgBoxGridH + 15), 855, fgBoxGridH
        gSetCtrl tmDCtrls(STARTDATEWK1INDEX + 2 * ilLoop), 900, 600 + ilLoop * (fgBoxGridH + 15), 990, fgBoxGridH
        gSetCtrl tmDCtrls(ENDDATEWK1INDEX + 2 * ilLoop), 1905, 600 + ilLoop * (fgBoxGridH + 15), 990, fgBoxGridH
    Next ilLoop
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move control values to record  *
'*                      and set defaults               *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec(ilTestChg As Integer)
'
'   mMoveCtrlToRec iTest
'   Where:
'       iTest (I)- True = only move if field changed
'                  False = move regardless of change state
'
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    If Not ilTestChg Or tmCtrls(YEARINDEX).iChg Then
        tmCof.iYear = Val(smSave(1))
    End If
    If Not ilTestChg Or tmCtrls(STARTMONTHINDEX).iChg Then
        tmCof.iStartMnthNo = Val(smSave(2))
    End If
    For ilLoop = 1 To 12 Step 1
        If ilLoop = 1 Then
            slStartDate = smSave(3)
        Else
            slStartDate = gIncOneDay(slEndDate)
        End If
        slEndDate = Format$(gDateValue(slStartDate) + 7 * Val(smSave(3 + ilLoop)) - 1, "m/d/yy")
        If Not ilTestChg Or tmCtrls(NOWKS1INDEX + ilLoop - 1).iChg Then
            tmCof.iNoWks(ilLoop - 1) = Val(smSave(3 + ilLoop))
            gPackDate slStartDate, tmCof.iStartDate(0, ilLoop - 1), tmCof.iStartDate(1, ilLoop - 1)
            gPackDate slEndDate, tmCof.iEndDate(0, ilLoop - 1), tmCof.iEndDate(1, ilLoop - 1)
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'
'   mMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    smSave(1) = Trim$(str$(tmCof.iYear))
    smSave(2) = Trim$(str$(tmCof.iStartMnthNo))
    'gUnpackDate tmCof.iStartDate(0, 1), tmCof.iStartDate(1, 1), smSave(3)
    gUnpackDate tmCof.iStartDate(0, 0), tmCof.iStartDate(1, 0), smSave(3)
    For ilLoop = 1 To 12 Step 1
        smSave(3 + ilLoop) = Trim$(str$(tmCof.iNoWks(ilLoop - 1)))
    Next ilLoop

    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
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
    Dim slStr As String
    If smSave(1) <> "" Then    'Test name
        slStr = Trim$(smSave(1))
        gFindMatch slStr, 0, cbcSelect    'Determine if name exist
        If gLastFound(cbcSelect) <> -1 Then   'Name found
            If gLastFound(cbcSelect) <> imSelectedIndex Then
                If Trim$(smSave(1)) = cbcSelect.List(gLastFound(cbcSelect)) Then
                    Beep
                    MsgBox "Year already defined, enter a different Year", vbOKOnly + vbExclamation + vbApplicationModal, "Error"
                    smSave(1) = Trim$(str$(tmCof.iYear)) 'Reset text
                    mSetShow imBoxNo
                    mSetChg imBoxNo
                    imBoxNo = 1
                    mEnableBox imBoxNo
                    mOKName = False
                    Exit Function
                End If
            End If
        End If
    End If
    mOKName = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim ilLoop As Integer

    imPopReqd = False
    cbcSelect.Clear
    sgMCofStamp = ""
    ilRet = csiSetStamp("COF", sgMCofStamp)
    ilRet = gObtainCorpCal()
    For ilLoop = UBound(tgMCof) - 1 To LBound(tgMCof) Step -1
        cbcSelect.AddItem Trim$(str$(tgMCof(ilLoop).iYear))
    Next ilLoop
    If igUpdateOk Then
        cbcSelect.AddItem "[New]", 0  'Force as first item on list
    Else
        cbcSelect.AddItem "[None]", 0  'Force as first item on list
    End If
    imPopReqd = True
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec(ilSelectIndex As Integer, ilForUpdate As Integer) As Integer
'
'   iRet = ENmRead(ilSelectIndex)
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status
    Dim ilLoop As Integer
    Dim ilCode As Integer
    slCode = cbcSelect.List(ilSelectIndex)
    ilCode = Val(slCode)
    For ilLoop = UBound(tgMCof) - 1 To LBound(tgMCof) Step -1
        If ilCode = tgMCof(ilLoop).iYear Then
            tmCofSrchKey.iCode = tgMCof(ilLoop).iYear 'CInt(slCode)
            ilRet = btrGetEqual(hmCof, tmCof, imCofRecLen, tmCofSrchKey, INDEXKEY0, BTRV_LOCK_NONE, ilForUpdate)
            On Error GoTo mReadRecErr
            gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", CorpCal
            On Error GoTo 0
            mReadRec = True
            Exit Function
        End If
    Next ilLoop
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRet As Integer
    Dim slMsg As String
    mSetShow imBoxNo
    If mSaveTestFields(TESTALLCTRLS, ALLMANDEFINED + SHOWMSG) = NO Then
        mSaveRec = False
        Exit Function
    End If
    If Not mOKName() Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass  'Wait
    Do  'Loop until record updated or added
        If imSelectedIndex <> 0 Then
            If Not mReadRec(imSelectedIndex, SETFORWRITE) Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Save")
                imTerminate = True
                mSaveRec = False
                Exit Function
            End If
        End If
        mMoveCtrlToRec True
        If imSelectedIndex = 0 Then 'New selected
            ilRet = btrInsert(hmCof, tmCof, imCofRecLen, INDEXKEY0)
            slMsg = "mSaveRec (btrInsert)"
        Else 'Old record-Update
            'ilRet = btrUpdate(hmCof, tmCof, imCofRecLen)
            ilRet = btrDelete(hmCof)    'Delete as Modify is not allowed in case year changed
            slMsg = "mSaveRec (btrDelete)"
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            ilRet = btrInsert(hmCof, tmCof, imCofRecLen, INDEXKEY0)
        End If
    Loop While ilRet = BTRV_ERR_CONFLICT
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, slMsg, CorpCal
    On Error GoTo 0
    mPopulate
    mSaveRec = True
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                      *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if record altered and*
'*                      requires updating              *
'*                                                     *
'*******************************************************
Private Function mSaveRecChg(ilAsk As Integer) As Integer
'
'   iAsk = True
'   iRet = mSaveRecChg(iAsk)
'   Where:
'       iAsk (I)- True = Ask if changed records should be updated;
'                 False= Update record if required without asking user
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRes As Integer
    Dim slMess As String
    Dim ilAltered As Integer
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    If mSaveTestFields(TESTALLCTRLS, ALLMANBLANK + NOMSG) = NO Then
        If ilAltered = YES Then
            If ilAsk Then
                If imSelectedIndex > 0 Then
                    slMess = "Save Changes to " & cbcSelect.List(imSelectedIndex)
                Else
                    slMess = "Add " & smSave(1)
                End If
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    pbcCorpCal_Paint
                    Exit Function
                End If
                If ilRes = vbYes Then
                    ilRes = mSaveRec()
                    mSaveRecChg = ilRes
                    Exit Function
                End If
                If ilRes = vbNo Then
                    cbcSelect.ListIndex = 0
                End If
            Else
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
        End If
    End If
    mSaveRecChg = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveTestFields                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mSaveTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mSaveTestFields(iTest, iState)
'   Where:
'       iTest (I)- Test all controls or control number specified
'       iState (I)- Test one of the following:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory
'                    field blank
'                  ALLMANDEFINED=All mandatory
'                    fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       iRet (O)- True if all mandatory fields blank, False if not all blank
'
'
    Dim ilLoop As Integer
    If (ilCtrlNo = YEARINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSave(1), "", "Year must be specified", tmCtrls(YEARINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = YEARINDEX
            End If
            mSaveTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = STARTMONTHINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSave(2), "", "Start Month must be specified", tmCtrls(STARTMONTHINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = STARTMONTHINDEX
            End If
            mSaveTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = STARTDATEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedStr(smSave(3), "", "Start Date must be specified", tmCtrls(STARTDATEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = STARTDATEINDEX
            End If
            mSaveTestFields = NO
            Exit Function
        End If
        'Test Week day
        If smSave(3) <> "" Then
            If gWeekDayStr(smSave(3)) <> 0 Then
                If gFieldDefinedStr("", "", "Start Date must Start on a Monday", tmCtrls(STARTDATEINDEX).iReq, ilState) = NO Then
                    If ilState = (ALLMANDEFINED + SHOWMSG) Then
                        imBoxNo = STARTDATEINDEX
                    End If
                    mSaveTestFields = NO
                    Exit Function
                End If
            End If
        End If
    End If
    For ilLoop = 1 To 12 Step 1
        If (ilCtrlNo = NOWKS1INDEX + ilLoop - 1) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedStr(smSave(3 + ilLoop), "", "No Weeks must be specified for Month" & str$(ilLoop), tmCtrls(NOWKS1INDEX + ilLoop - 1).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = NOWKS1INDEX + ilLoop - 1
                End If
                mSaveTestFields = NO
                Exit Function
            End If
        End If
    Next ilLoop
    mSaveTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetChg                         *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if value for a       *
'*                      control is different from the  *
'*                      record                         *
'*                                                     *
'*******************************************************
Private Sub mSetChg(ilBoxNo As Integer)
'
'   mSetChg ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be checked
'
    Dim slStr As String
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmCtrls) Then
'        mSetCommands
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case YEARINDEX 'Name
            slStr = Trim$(str$(tmCof.iYear))
            gSetChgFlag slStr, edcDropDown, tmCtrls(ilBoxNo)
        Case STARTMONTHINDEX 'Name
            slStr = Trim$(str$(tmCof.iStartMnthNo))
            gSetChgFlag slStr, edcDropDown, tmCtrls(ilBoxNo)
        Case STARTDATEINDEX 'Name
            slStr = smSave(3)
            'gUnpackDate tmCof.iStartDate(0, 1), tmCof.iStartDate(1, 1), slStr
            gUnpackDate tmCof.iStartDate(0, 0), tmCof.iStartDate(1, 0), slStr
            gSetChgFlag slStr, edcDropDown, tmCtrls(ilBoxNo)
        Case NOWKS1INDEX To NOWKS1INDEX + 11
            'slStr = Trim$(str$(tmCof.iNoWks(ilBoxNo - NOWKS1INDEX + 1)))
            slStr = Trim$(str$(tmCof.iNoWks(ilBoxNo - NOWKS1INDEX)))
            gSetChgFlag slStr, edcDropDown, tmCtrls(ilBoxNo)
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim ilAltered As Integer
    If imBypassSetting Then
        Exit Sub
    End If
    If Not igUpdateOk Then
        Exit Sub
    End If
    ilAltered = gAnyFieldChgd(tmCtrls(), TESTALLCTRLS)
    'Update button set if all mandatory fields have data and any field altered
    If (mSaveTestFields(TESTALLCTRLS, ALLMANDEFINED + NOMSG) = YES) And (ilAltered = YES) Then
        cmcSave.Enabled = True
    Else
        cmcSave.Enabled = False
    End If
    If Not ilAltered Then
        cbcSelect.Enabled = True
    Else
        cbcSelect.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocusx                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case YEARINDEX 'Name
            edcDropDown.SetFocus
        Case STARTMONTHINDEX 'Name
            edcDropDown.SetFocus
        Case STARTDATEINDEX 'Name
            edcDropDown.SetFocus
        Case NOWKS1INDEX To NOWKS1INDEX + 11
            edcDropDown.SetFocus
    End Select
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim ilLoop As Integer   'For loop control parameter
    Dim slStr As String
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If

    Select Case ilBoxNo 'Branch on box type (control)
        Case YEARINDEX 'Name
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            gSetShow pbcCorpCal, slStr, tmCtrls(ilBoxNo)
            If smSave(1) <> slStr Then
                smSave(1) = slStr
                slStr = ""
                smSave(2) = ""
                gSetShow pbcCorpCal, slStr, tmCtrls(STARTMONTHINDEX)
                smSave(3) = ""
                gSetShow pbcCorpCal, slStr, tmCtrls(STARTDATEINDEX)
                For ilLoop = 4 To 15 Step 1
                    smSave(ilLoop) = ""
                    gSetShow pbcCorpCal, slStr, tmCtrls(NOWKS1INDEX + ilLoop - 4)
                    'tmCof.iNoWks(ilLoop - 3) = 0
                    tmCof.iNoWks(ilLoop - 4) = 0
                Next ilLoop
                pbcCorpCal.Cls
            End If
        Case STARTMONTHINDEX 'Name
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            gSetShow pbcCorpCal, slStr, tmCtrls(ilBoxNo)
            If smSave(2) <> slStr Then
                smSave(2) = slStr
                slStr = ""
                smSave(3) = ""
                gSetShow pbcCorpCal, slStr, tmCtrls(STARTDATEINDEX)
                For ilLoop = 4 To 15 Step 1
                    smSave(ilLoop) = ""
                    gSetShow pbcCorpCal, slStr, tmCtrls(NOWKS1INDEX + ilLoop - 4)
                    'tmCof.iNoWks(ilLoop - 3) = 0
                    tmCof.iNoWks(ilLoop - 4) = 0
                Next ilLoop
                pbcCorpCal.Cls
            End If
        Case STARTDATEINDEX
            edcDropDown.Visible = False  'Set visibility
            cmcDropDown.Visible = False
            plcCalendar.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcCorpCal, slStr, tmCtrls(ilBoxNo)
            If smSave(3) <> slStr Then
                smSave(3) = slStr
                slStr = ""
                For ilLoop = 4 To 15 Step 1
                    smSave(ilLoop) = ""
                    gSetShow pbcCorpCal, slStr, tmCtrls(NOWKS1INDEX + ilLoop - 4)
                    'tmCof.iNoWks(ilLoop - 3) = 0
                    tmCof.iNoWks(ilLoop - 4) = 0
                Next ilLoop
                pbcCorpCal.Cls
            End If
        Case NOWKS1INDEX To NOWKS1INDEX + 11
            edcDropDown.Visible = False  'Set visibility
            slStr = edcDropDown.Text
            gSetShow pbcCorpCal, slStr, tmCtrls(ilBoxNo)
            If smSave(ilBoxNo) <> slStr Then
                smSave(ilBoxNo) = slStr
                slStr = ""
                For ilLoop = ilBoxNo + 1 To 15 Step 1
                    smSave(ilLoop) = ""
                    gSetShow pbcCorpCal, slStr, tmCtrls(NOWKS1INDEX + ilLoop - 4)
                    'tmCof.iNoWks(ilLoop - 3) = 0
                    tmCof.iNoWks(ilLoop - 4) = 0
                Next ilLoop
                pbcCorpCal.Cls
            End If
    End Select
    pbcCorpCal_Paint
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
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
    Unload CorpCal
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields(ilCtrlNo As Integer, ilState As Integer) As Integer
'
'   iState = ALLBLANK+NOMSG   'Blank
'   iTest = TESTALLCTRLS
'   iRet = mTestFields(iTest, iState)
'   Where:
'       iTest (I)- Test all controls or control number specified
'       iState (I)- Test one of the following:
'                  ALLBLANK=All fields blank
'                  ALLMANBLANK=All mandatory
'                    field blank
'                  ALLMANDEFINED=All mandatory
'                    fields have data
'                  Plus
'                  NOMSG=No error message shown
'                  SHOWMSG=show error message
'       iRet (O)- True if all mandatory fields blank, False if not all blank
'
'
    Dim ilLoop As Integer
    If (ilCtrlNo = YEARINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcDropDown, "", "Year must be specified", tmCtrls(YEARINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = YEARINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = STARTMONTHINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcDropDown, "", "Start Month must be specified", tmCtrls(STARTMONTHINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = STARTMONTHINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    If (ilCtrlNo = STARTDATEINDEX) Or (ilCtrlNo = TESTALLCTRLS) Then
        If gFieldDefinedCtrl(edcDropDown, "", "Start Date must be specified", tmCtrls(STARTDATEINDEX).iReq, ilState) = NO Then
            If ilState = (ALLMANDEFINED + SHOWMSG) Then
                imBoxNo = STARTDATEINDEX
            End If
            mTestFields = NO
            Exit Function
        End If
    End If
    For ilLoop = 1 To 12 Step 1
        If (ilCtrlNo = NOWKS1INDEX + ilLoop - 1) Or (ilCtrlNo = TESTALLCTRLS) Then
            If gFieldDefinedCtrl(edcDropDown, "", "No Weeks must be specified for Month" & str$(ilLoop), tmCtrls(NOWKS1INDEX + ilLoop - 1).iReq, ilState) = NO Then
                If ilState = (ALLMANDEFINED + SHOWMSG) Then
                    imBoxNo = NOWKS1INDEX + ilLoop - 1
                End If
                mTestFields = NO
                Exit Function
            End If
        End If
    Next ilLoop
    mTestFields = YES
End Function

Private Sub Form_Unload(Cancel As Integer)
    btrDestroy hmCof
    Set CorpCal = Nothing   'Remove data segment
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
End Sub
Private Sub pbcCorpCal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilLoop As Integer
    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If (X >= tmCtrls(ilBox).fBoxX) And (X <= tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW) Then
            If (Y >= tmCtrls(ilBox).fBoxY) And (Y <= tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH) Then
                For ilLoop = YEARINDEX To ilBox - 1 Step 1
                    If smSave(ilLoop) = "" Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                Next ilLoop
                mSetShow imBoxNo
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            End If
        End If
    Next ilBox
    mSetFocus imBoxNo
End Sub
Private Sub pbcCorpCal_Paint()
    Dim ilBox As Integer
    Dim ilBoxNo As Integer
    Dim slStartDate As String
    Dim slEndDate As String

    For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
        If ilBox < NOWKS1INDEX Then
            pbcCorpCal.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcCorpCal.CurrentY = tmCtrls(ilBox).fBoxY + fgBoxInsetY
            pbcCorpCal.Print tmCtrls(ilBox).sShow
        Else
            pbcCorpCal.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcCorpCal.CurrentY = tmCtrls(ilBox).fBoxY - 15 '+ fgBoxInsetY
            pbcCorpCal.Print tmCtrls(ilBox).sShow
            If smSave(ilBox) <> "" Then
                If ilBox = NOWKS1INDEX Then
                    slStartDate = smSave(3)
                Else
                    slStartDate = gIncOneDay(slEndDate)
                End If
                slEndDate = Format$(gDateValue(slStartDate) + 7 * Val(smSave(ilBox)) - 1, "m/d/yy")
                ilBoxNo = 2 * (ilBox - NOWKS1INDEX) + 1
                pbcCorpCal.CurrentX = tmDCtrls(ilBoxNo).fBoxX + fgBoxInsetX
                pbcCorpCal.CurrentY = tmDCtrls(ilBoxNo).fBoxY - 15  '+ fgBoxInsetY
                gSetShow pbcCorpCal, slStartDate, tmDCtrls(ilBoxNo)
                pbcCorpCal.Print tmDCtrls(ilBoxNo).sShow
                ilBoxNo = ilBoxNo + 1
                pbcCorpCal.CurrentX = tmDCtrls(ilBoxNo).fBoxX + fgBoxInsetX
                pbcCorpCal.CurrentY = tmDCtrls(ilBoxNo).fBoxY - 15  '+ fgBoxInsetY
                gSetShow pbcCorpCal, slEndDate, tmDCtrls(ilBoxNo)
                pbcCorpCal.Print tmDCtrls(ilBoxNo).sShow
            End If
        End If
    Next ilBox
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcSTab.hWnd Then
        Exit Sub
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= UBound(tmCtrls)) Then
        If (imBoxNo <> YEARINDEX) Or (Not cbcSelect.Enabled) Then
            If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
                Beep
                mEnableBox imBoxNo
                Exit Sub
            End If
        End If
    End If
    Select Case imBoxNo
        Case -1
            If (imSelectedIndex = 0) And (cbcSelect.Text = "[New]") Then
                ilBox = 1
                mSetCommands
            ElseIf (imSelectedIndex = 0) Then
                mSetChg 1
                imBoxNo = 1
                ilBox = 2
            Else
                mSetChg 1
                ilBox = 4
            End If
        Case 1 'Year (first control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If cbcSelect.Enabled Then
                cbcSelect.SetFocus
                Exit Sub
            End If
            ilBox = 1
        Case Else
            ilBox = imBoxNo - 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    If GetFocus() <> pbcTab.hWnd Then
        Exit Sub
    End If
    If (imBoxNo >= imLBCtrls) And (imBoxNo <= UBound(tmCtrls)) Then
        If mTestFields(imBoxNo, ALLMANDEFINED + NOMSG) = NO Then
            Beep
            mEnableBox imBoxNo
            Exit Sub
        End If
    End If
    Select Case imBoxNo
        Case -1
            ilBox = UBound(tmCtrls)
        Case UBound(tmCtrls) 'Suppress (last control within header)
            mSetShow imBoxNo
            imBoxNo = -1
            If (cmcSave.Enabled) Then
                cmcSave.SetFocus
            Else
                cmcDone.SetFocus
            End If
            Exit Sub
        Case Else
            ilBox = imBoxNo + 1
    End Select
    mSetShow imBoxNo
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub plcCorpCal_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Corporate Calendar"
End Sub
