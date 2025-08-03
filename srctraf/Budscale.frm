VERSION 5.00
Begin VB.Form BudScale 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3765
   ClientLeft      =   1665
   ClientTop       =   1485
   ClientWidth     =   6495
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
   ScaleHeight     =   3765
   ScaleWidth      =   6495
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Vehicles"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      TabIndex        =   19
      Top             =   3045
      Width           =   1425
   End
   Begin VB.PictureBox plcSpec 
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
      Height          =   2805
      Left            =   3240
      ScaleHeight     =   2745
      ScaleWidth      =   2925
      TabIndex        =   3
      Top             =   300
      Width           =   2985
      Begin VB.TextBox edcRound 
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
         Height          =   300
         Left            =   1815
         TabIndex        =   15
         Top             =   2385
         Width           =   795
      End
      Begin VB.ComboBox cbcEndDate 
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
         Height          =   300
         Left            =   1140
         TabIndex        =   7
         Top             =   510
         Width           =   1560
      End
      Begin VB.ComboBox cbcStartDate 
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
         Height          =   300
         Left            =   1140
         TabIndex        =   5
         Top             =   135
         Width           =   1560
      End
      Begin VB.TextBox edcTotalGross 
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
         Height          =   300
         Left            =   1245
         TabIndex        =   12
         Top             =   1560
         Width           =   1560
      End
      Begin VB.TextBox edcIndexValue 
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
         Height          =   300
         Left            =   1245
         TabIndex        =   9
         Top             =   945
         Width           =   795
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   2850
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0FFFF&
         X1              =   120
         X2              =   2850
         Y1              =   2265
         Y2              =   2265
      End
      Begin VB.Label lacRound 
         Appearance      =   0  'Flat
         Caption         =   "Round Budget to"
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
         Left            =   120
         TabIndex        =   14
         Top             =   2430
         Width           =   1560
      End
      Begin VB.Label lacTotal 
         Appearance      =   0  'Flat
         Caption         =   "Total Gross"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1950
         Width           =   2640
      End
      Begin VB.Label lacOr 
         Appearance      =   0  'Flat
         Caption         =   "Or"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   960
         TabIndex        =   10
         Top             =   1290
         Width           =   300
      End
      Begin VB.Label lacTotalGross 
         Appearance      =   0  'Flat
         Caption         =   "Total Gross"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1605
         Width           =   1110
      End
      Begin VB.Label lacEndDate 
         Appearance      =   0  'Flat
         Caption         =   "End Date"
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
         Left            =   120
         TabIndex        =   6
         Top             =   555
         Width           =   960
      End
      Begin VB.Label lacIndexValue 
         Appearance      =   0  'Flat
         Caption         =   "Index Value"
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
         Left            =   135
         TabIndex        =   8
         Top             =   990
         Width           =   1110
      End
      Begin VB.Label lacStartDate 
         Appearance      =   0  'Flat
         Caption         =   "Start Date"
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
         Left            =   120
         TabIndex        =   4
         Top             =   180
         Width           =   960
      End
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1860
      TabIndex        =   16
      Top             =   3345
      Width           =   945
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   75
      ScaleHeight     =   240
      ScaleWidth      =   5640
      TabIndex        =   0
      Top             =   15
      Width           =   5640
   End
   Begin VB.CommandButton cmcApply 
      Appearance      =   0  'Flat
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3495
      TabIndex        =   17
      Top             =   3345
      Width           =   945
   End
   Begin VB.PictureBox plcModel 
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
      Height          =   2670
      Left            =   150
      ScaleHeight     =   2610
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   300
      Width           =   3015
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         Height          =   2550
         Left            =   30
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   30
         Width           =   2895
      End
   End
   Begin VB.Label lacNote 
      Appearance      =   0  'Flat
      Caption         =   "(* Vehicle Scaled)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1875
      TabIndex        =   18
      Top             =   3075
      Width           =   1320
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   165
      Top             =   3360
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "BudScale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Budscale.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: BudScale.Frm
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
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imSDateSelectedIndex As Integer
Dim imEDateSelectedIndex As Integer
Dim imDateComboBoxIndex As Integer
Dim imDateChgMode As Integer
Dim imUpdateAllowed As Integer


Private Sub cbcEndDate_Change()
    Dim ilLoopCount As Integer
    '  imChgMode is used to avoid entering this routine multiple times
    '            if a vehicle selection change occurs during the
    '            processing of a "change"
    If imDateChgMode = False Then
        imDateChgMode = True
        ilLoopCount = 0
        Screen.MousePointer = vbHourglass  'Wait
        Do
            If ilLoopCount > 0 Then
                If cbcEndDate.ListIndex >= 0 Then
                    cbcEndDate.Text = cbcEndDate.List(cbcEndDate.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            If cbcEndDate.Text <> "" Then
                gManLookAhead cbcEndDate, imBSMode, imDateComboBoxIndex
            End If
            imEDateSelectedIndex = cbcEndDate.ListIndex
        Loop While imEDateSelectedIndex <> cbcEndDate.ListIndex
        mComputeTotal True
        Screen.MousePointer = vbDefault    'Default
        imDateChgMode = False
    End If
    mSetCommands
End Sub
Private Sub cbcEndDate_Click()
    imDateComboBoxIndex = cbcEndDate.ListIndex
    cbcEndDate_Change
End Sub
Private Sub cbcEndDate_GotFocus()
    gCtrlGotFocus cbcEndDate
    imDateComboBoxIndex = imEDateSelectedIndex
End Sub
Private Sub cbcEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcEndDate_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcEndDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcStartDate_Change()
    Dim ilLoopCount As Integer
    '  imChgMode is used to avoid entering this routine multiple times
    '            if a vehicle selection change occurs during the
    '            processing of a "change"
    If imDateChgMode = False Then
        imDateChgMode = True
        ilLoopCount = 0
        Screen.MousePointer = vbHourglass  'Wait
        Do
            If ilLoopCount > 0 Then
                If cbcStartDate.ListIndex >= 0 Then
                    cbcStartDate.Text = cbcStartDate.List(cbcStartDate.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            If cbcStartDate.Text <> "" Then
                gManLookAhead cbcStartDate, imBSMode, imDateComboBoxIndex
            End If
            imSDateSelectedIndex = cbcStartDate.ListIndex
        Loop While imSDateSelectedIndex <> cbcStartDate.ListIndex
        mComputeTotal True
        Screen.MousePointer = vbDefault    'Default
        imDateChgMode = False
    End If
    mSetCommands
End Sub
Private Sub cbcStartDate_Click()
    imDateComboBoxIndex = cbcStartDate.ListIndex
    cbcStartDate_Change
End Sub
Private Sub cbcStartDate_GotFocus()
    gCtrlGotFocus cbcStartDate
    imDateComboBoxIndex = imSDateSelectedIndex
End Sub
Private Sub cbcStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcStartDate_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcStartDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
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
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        If lbcVehicle.ListCount > 0 Then       'select all slsp
            llRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcVehicle.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        End If
        imAllClicked = False
        mComputeTotal True
    End If
    mSetCommands
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub
Private Sub cmcApply_Click()
    Dim slIndex As String
    Dim ilLoop As Integer
    Dim ilBvf As Integer
    Dim slVehicle As String
    Dim ilWk As Integer
    Dim slStr As String
    Dim slAsterisk As String
    Dim slChar As String
    Dim slRound As String
    Dim ilByTotal As Integer
    Dim llTotal As Long
    If imEDateSelectedIndex < imSDateSelectedIndex Then
        MsgBox "End Date is prior to Start Date", vbOKOnly + vbExclamation, "Scale"
        cbcEndDate.SetFocus
        Exit Sub
    End If
    If (edcIndexValue.Text = "") And (edcTotalGross.Text = "") Then
        MsgBox "Index Value or Total Gross must be Defined", vbOKOnly + vbExclamation, "Scale"
        edcIndexValue.SetFocus
        Exit Sub
    End If
    If (edcIndexValue.Text <> "") And (edcTotalGross.Text <> "") Then
        MsgBox "Index Value or Total Gross must be Defined, not Both", vbOKOnly + vbExclamation, "Scale"
        edcIndexValue.SetFocus
        Exit Sub
    End If
    ilByTotal = False
    If (edcIndexValue.Text <> "") Then
        slIndex = edcIndexValue.Text
    Else
        slStr = edcTotalGross.Text & ".0000"
        slIndex = gDivStr(slStr, Trim$(Str$(lgTotal)))
        ilByTotal = True
    End If
    slRound = edcRound.Text
    If slRound = "" Then
        slRound = "1"
    End If
    Screen.MousePointer = vbHourglass
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            slVehicle = lbcVehicle.List(ilLoop)
            slChar = Left$(slVehicle, 1)
            slAsterisk = ""
            Do While slChar = "*"
                slAsterisk = slAsterisk & slChar
                slVehicle = Mid$(slVehicle, 2)
                slChar = Left$(slVehicle, 1)
            Loop
            'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
            For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
                If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                    'For ilWk = LBound(tgBvfRec(ilBvf).tBvf.lGross) To UBound(tgBvfRec(ilBvf).tBvf.lGross) Step 1
                    For ilWk = imSDateSelectedIndex + 1 To imEDateSelectedIndex + 1 Step 1
                        If tgBvfRec(ilBvf).tBvf.lGross(ilWk) <> 0 Then
                            slStr = Trim$(Str$(tgBvfRec(ilBvf).tBvf.lGross(ilWk)))
                            slStr = gMulStr(slStr, slIndex)
                            slStr = gRoundStr(slStr, slRound, 0)
                            tgBvfRec(ilBvf).tBvf.lGross(ilWk) = gStrDecToLong(slStr, 0)
                        End If
                    Next ilWk
                End If
            Next ilBvf
            lbcVehicle.List(ilLoop) = "*" & slAsterisk & slVehicle
        End If
    Next ilLoop
    If (ilByTotal) And (slRound = "1") Then
        slStr = edcTotalGross.Text
        llTotal = Val(slStr)
        mComputeTotal False
        If lgTotal <> llTotal Then
            For ilWk = imSDateSelectedIndex + 1 To imEDateSelectedIndex + 1 Step 1
                'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
                For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
                    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
                        If lbcVehicle.Selected(ilLoop) Then
                            slVehicle = lbcVehicle.List(ilLoop)
                            slChar = Left$(slVehicle, 1)
                            slAsterisk = ""
                            Do While slChar = "*"
                                slAsterisk = slAsterisk & slChar
                                slVehicle = Mid$(slVehicle, 2)
                                slChar = Left$(slVehicle, 1)
                            Loop
                            If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                                If tgBvfRec(ilBvf).tBvf.lGross(ilWk) <> 0 Then
                                    If lgTotal > llTotal Then
                                        tgBvfRec(ilBvf).tBvf.lGross(ilWk) = tgBvfRec(ilBvf).tBvf.lGross(ilWk) - 1
                                        lgTotal = lgTotal - 1
                                    ElseIf lgTotal < llTotal Then
                                        tgBvfRec(ilBvf).tBvf.lGross(ilWk) = tgBvfRec(ilBvf).tBvf.lGross(ilWk) + 1
                                        lgTotal = lgTotal + 1
                                    End If
                                End If
                                Exit For
                            End If
                        End If
                    Next ilLoop
                    If lgTotal = llTotal Then
                        Exit For
                    End If
                Next ilBvf
                If lgTotal = llTotal Then
                    Exit For
                End If
            Next ilWk
        End If
    End If
    mComputeTotal True
    edcIndexValue.Text = ""
    edcTotalGross.Text = ""
    igBDReturn = 1
    Screen.MousePointer = vbDefault
    'mTerminate
End Sub
Private Sub cmcApply_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_Click()
    mTerminate
End Sub
Private Sub edcIndexValue_Change()
    edcTotalGross.Text = ""
    mSetCommands
End Sub
Private Sub edcIndexValue_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcIndexValue_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(edcIndexValue.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(edcIndexValue.Text, ".")    'Disallow multi-decimal points
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
    slStr = edcIndexValue.Text
    slStr = Left$(slStr, edcIndexValue.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcIndexValue.SelStart - edcIndexValue.SelLength)
    If gCompNumberStr(slStr, "100.0000") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcRound_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcRound_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcRound.Text
    slStr = Left$(slStr, edcRound.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcRound.SelStart - edcRound.SelLength)
    If gCompNumberStr(slStr, "1000000") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcTotalGross_Change()
    edcIndexValue.Text = ""
    mSetCommands
End Sub
Private Sub edcTotalGross_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcTotalGross_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcTotalGross.Text
    slStr = Left$(slStr, edcTotalGross.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcTotalGross.SelStart - edcTotalGross.SelLength)
    If gCompNumberStr(slStr, "2000000000") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        'gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    'If (igWinStatus(RATECARDSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
    '    lbcBudget.Enabled = False
    'Else
    '    lbcBudget.Enabled = True
    'End If
'    gShowBranner
    If (igWinStatus(BUDGETSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    'gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    BudScale.Refresh
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set BudScale = Nothing   'Remove data segment
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcVehicle_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
        mComputeTotal True
    End If
    mSetCommands
End Sub
Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mComputeTotal                   *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute buget totals           *
'*                                                     *
'*******************************************************
Private Sub mComputeTotal(ilDisplayTotal As Integer)
    Dim ilBvf As Integer
    Dim ilWk As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slVehicle As String
    Dim slChar As String
    lgTotal = 0
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            slVehicle = lbcVehicle.List(ilLoop)
            slChar = Left$(slVehicle, 1)
            Do While slChar = "*"
                slVehicle = Mid$(slVehicle, 2)
                slChar = Left$(slVehicle, 1)
            Loop
            'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
            For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
                If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                    'For ilWk = LBound(tgBvfRec(ilBvf).tBvf.lGross) To UBound(tgBvfRec(ilBvf).tBvf.lGross) Step 1
                    For ilWk = imSDateSelectedIndex + 1 To imEDateSelectedIndex + 1 Step 1
                        If tgBvfRec(ilBvf).tBvf.lGross(ilWk) <> 0 Then
                            lgTotal = lgTotal + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                        End If
                    Next ilWk
                End If
            Next ilBvf
        End If
    Next ilLoop
    If ilDisplayTotal Then
        slStr = Trim$(Str$(lgTotal))
        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
        lacTotal.Caption = "(Current Total: " & slStr & ")"
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDatePop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Date from Corporate   *
'*                      calendar                       *
'*                                                     *
'*******************************************************
Private Sub mDatePop()
'
'   mMnfBudgetPop
'   Where:
'
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilNoWks As Integer
    Dim ilYear As Integer
    ilYear = 0
    'For ilLoop = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
    For ilLoop = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
        If tgBvfRec(ilLoop).tBvf.iYear > 0 Then
            ilYear = tgBvfRec(ilLoop).tBvf.iYear
            Exit For
        End If
    Next ilLoop
    cbcStartDate.Clear
    cbcEndDate.Clear
    If tgSpf.sRUseCorpCal = "Y" Then
        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
            If tgMCof(ilLoop).iYear = ilYear Then
                'gUnpackDate tgMCof(ilLoop).iStartDate(0, 1), tgMCof(ilLoop).iStartDate(1, 1), slStartDate
                'gUnpackDate tgMCof(ilLoop).iEndDate(0, 12), tgMCof(ilLoop).iEndDate(1, 12), slEndDate
                gUnpackDate tgMCof(ilLoop).iStartDate(0, 0), tgMCof(ilLoop).iStartDate(1, 0), slStartDate
                gUnpackDate tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iEndDate(1, 11), slEndDate
                ilNoWks = (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 7 + 1
                gDatePop slStartDate, ilNoWks, cbcStartDate
                slStartDate = gObtainNextSunday(slStartDate)
                ilNoWks = (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 7 + 1
                gDatePop slStartDate, ilNoWks, cbcEndDate
                Exit For
            End If
        Next ilLoop
    Else
        slStartDate = "1/15/" & Trim$(Str$(ilYear))
        slStartDate = gObtainYearStartDate(0, slStartDate)
        slEndDate = "12/15/" & Trim$(Str$(ilYear))
        slEndDate = gObtainYearEndDate(0, slEndDate)
        ilNoWks = (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 7 + 1
        gDatePop slStartDate, ilNoWks, cbcStartDate
        slStartDate = gObtainNextSunday(slStartDate)
        ilNoWks = (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 7 + 1
        gDatePop slStartDate, ilNoWks, cbcEndDate
    End If
    cbcStartDate.ListIndex = 0
    cbcEndDate.ListIndex = cbcEndDate.ListCount - 1
    Exit Sub

    On Error GoTo 0
    imTerminate = True
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
    imTerminate = False
    imFirstActivate = True
    igBDReturn = 0

    igLBBvfRec = 1  'Match definition in Budget

    Screen.MousePointer = vbHourglass
    BudScale.Height = cmcApply.Top + 5 * cmcApply.Height / 3
    gCenterStdAlone BudScale
    imChgMode = False
    imBSMode = False
    imDateChgMode = False
    imBypassSetting = False
    imAllClicked = False
    imSetAll = True
    'BudScale.Show
    Screen.MousePointer = vbHourglass
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    mDatePop
    If imTerminate Then
        Exit Sub
    End If
    'If lgTotal = 0 Then
    '    edcTotalGross.Enabled = False
    'Else
    '    edcTotalGross.Enabled = True
    'End If
    'slStr = Trim$(Str$(lgTotal))
    'gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
    'lacTotal.Caption = "(Current Total: " & slStr & ")"
    mComputeTotal True
    edcRound.Text = "1"
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    gCenterModalForm BudScale
    'plcScreen.Caption = "Scale for " & sgBAName
    Screen.MousePointer = vbDefault
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
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
    Dim ilLoop As Integer
    Dim slName As String

    'For ilLoop = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
    For ilLoop = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
        slName = Trim$(tgBvfRec(ilLoop).sVehicle)
        gFindMatch slName, 0, lbcVehicle
        If gLastFound(lbcVehicle) < 0 Then
            lbcVehicle.AddItem slName
        End If
    Next ilLoop

    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
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
    Dim ilLoop As Integer
    If (cbcStartDate.Text = "") Or (cbcEndDate.Text = "") Then
        cmcApply.Enabled = False
        Exit Sub
    End If
    If (edcIndexValue.Text = "") And (edcTotalGross.Text = "") Then
        cmcApply.Enabled = False
        Exit Sub
    End If
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            cmcApply.Enabled = True
            Exit Sub
        End If
    Next ilLoop
    cmcApply.Enabled = False
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
    Unload BudScale
    igManUnload = NO
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Scale for " & sgBAName
End Sub
