VERSION 5.00
Begin VB.Form BudTrend 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4305
   ClientLeft      =   375
   ClientTop       =   2940
   ClientWidth     =   9240
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
   ScaleHeight     =   4305
   ScaleWidth      =   9240
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Vehicles"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      TabIndex        =   18
      Top             =   3645
      Width           =   1365
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
      Height          =   3360
      Left            =   2835
      ScaleHeight     =   3300
      ScaleWidth      =   6150
      TabIndex        =   3
      Top             =   270
      Width           =   6210
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
         Left            =   1890
         TabIndex        =   16
         Top             =   2955
         Width           =   795
      End
      Begin VB.PictureBox pbcUpMove 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2775
         Picture         =   "Budtrend.frx":0000
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2040
         Width           =   180
      End
      Begin VB.PictureBox pbcDnMove 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3345
         Picture         =   "Budtrend.frx":00DA
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   990
         Width           =   180
      End
      Begin VB.ListBox lbcTo 
         Appearance      =   0  'Flat
         Height          =   1920
         Left            =   3750
         MultiSelect     =   2  'Extended
         TabIndex        =   10
         Top             =   555
         Width           =   2400
      End
      Begin VB.ListBox lbcFrom 
         Appearance      =   0  'Flat
         Height          =   1920
         Left            =   75
         TabIndex        =   11
         Top             =   555
         Width           =   2400
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
         Left            =   3690
         TabIndex        =   7
         Top             =   135
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
         Left            =   1050
         TabIndex        =   5
         Top             =   135
         Width           =   1560
      End
      Begin VB.CommandButton cmcMoveTo 
         Appearance      =   0  'Flat
         Caption         =   "M&ove   "
         Height          =   300
         Left            =   2655
         TabIndex        =   12
         Top             =   930
         Width           =   945
      End
      Begin VB.CommandButton cmcMoveBack 
         Appearance      =   0  'Flat
         Caption         =   "    Mo&ve"
         Height          =   300
         Left            =   2655
         TabIndex        =   15
         Top             =   1965
         Width           =   945
      End
      Begin VB.Label lacMsg 
         Appearance      =   0  'Flat
         Caption         =   "(Earliest to Latest)"
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
         Left            =   3750
         TabIndex        =   20
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0FFFF&
         X1              =   195
         X2              =   6060
         Y1              =   2865
         Y2              =   2865
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   195
         X2              =   6060
         Y1              =   2880
         Y2              =   2880
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
         Left            =   195
         TabIndex        =   17
         Top             =   3000
         Width           =   1560
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
         Left            =   2820
         TabIndex        =   6
         Top             =   180
         Width           =   840
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
         Width           =   930
      End
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   3255
      TabIndex        =   13
      Top             =   3945
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
      Top             =   -15
      Width           =   5640
   End
   Begin VB.CommandButton cmcApply 
      Appearance      =   0  'Flat
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4890
      TabIndex        =   14
      Top             =   3945
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
      Height          =   3300
      Left            =   150
      ScaleHeight     =   3240
      ScaleWidth      =   2505
      TabIndex        =   1
      Top             =   270
      Width           =   2565
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         Height          =   3180
         Left            =   30
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   30
         Width           =   2445
      End
   End
   Begin VB.Label lacNote 
      Appearance      =   0  'Flat
      Caption         =   "(* Vehicle Set)"
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
      Left            =   1605
      TabIndex        =   19
      Top             =   3630
      Width           =   1080
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   180
      Top             =   3915
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "BudTrend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Budtrend.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: BudTrend.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Trend input screen code
Option Explicit
Option Compare Text
'Office
Dim hmBvf As Integer    'Rate Card file handle
Dim tmBvfSrchKey As BVFKEY0    'Rcf key record image
Dim imBvfRecLen As Integer        'Rcf record length
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imSDateSelectedIndex As Integer
Dim imEDateSelectedIndex As Integer
Dim imDateComboBoxIndex As Integer
Dim imDateChgMode As Integer
Dim tmBvf() As BVF
Dim imLBBv As Integer
Dim imVefCode() As Integer
Dim lmDollars() As Long
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
        'mComputeTotal
    End If
    mSetCommands
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub
Private Sub cmcApply_Click()
    Dim ilLoop As Integer
    Dim ilBvf As Integer
    Dim slVehicle As String
    Dim ilWk As Integer
    Dim slStr As String
    Dim slAsterisk As String
    Dim slChar As String
    Dim slRound As String
    Dim ilIndex As Integer
    Dim ilCount As Integer
    Dim ilNoHits As Integer
    Dim ilRet As Integer
    Dim llGross As Long
    If imEDateSelectedIndex < imSDateSelectedIndex Then
        MsgBox "End Date is prior to Start Date", vbOKOnly + vbExclamation, "Trend"
        cbcEndDate.SetFocus
        Exit Sub
    End If
    slRound = edcRound.Text
    If slRound = "" Then
        slRound = "1"
    End If
    ilRet = MsgBox("This will Replace the Budget Values, Ok to Proceed", vbYesNo + vbQuestion, "Trend")
    If ilRet = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilRet = mReadBvfRec()
    ReDim lmDollars(0 To lbcTo.ListCount - 1) As Long
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
                        ilCount = 0
                        ilNoHits = 0
                        llGross = 0
                        ReDim lmDollars(0 To lbcTo.ListCount - 1) As Long
                        'For ilIndex = LBound(tmBvf) To UBound(tmBvf) - 1 Step 1
                        For ilIndex = imLBBv To UBound(tmBvf) - 1 Step 1
                            If (tgBvfRec(ilBvf).tBvf.iVefCode = tmBvf(ilIndex).iVefCode) And (tgBvfRec(ilBvf).tBvf.iSofCode = tmBvf(ilIndex).iSofCode) Then
                                lmDollars(ilCount) = tmBvf(ilIndex).lGross(ilWk)
                                llGross = tmBvf(ilIndex).lGross(ilWk)   'Save is one hit only
                                ilNoHits = ilNoHits + 1
                            End If
                            If (tmBvf(ilIndex).iMnfBudget <> tmBvf(ilIndex + 1).iMnfBudget) Or (tmBvf(ilIndex).iYear <> tmBvf(ilIndex + 1).iYear) Then
                                ilCount = ilCount + 1
                            End If
                        Next ilIndex
                        If ilNoHits > 1 Then
                            tgBvfRec(ilBvf).tBvf.lGross(ilWk) = mStraightLinePrediction(lmDollars())
                        Else
                            tgBvfRec(ilBvf).tBvf.lGross(ilWk) = llGross
                        End If
                        slStr = Trim$(Str$(tgBvfRec(ilBvf).tBvf.lGross(ilWk)))
                        slStr = gRoundStr(slStr, slRound, 0)
                        tgBvfRec(ilBvf).tBvf.lGross(ilWk) = gStrDecToLong(slStr, 0)
                    Next ilWk
                End If
            Next ilBvf
            lbcVehicle.List(ilLoop) = "*" & slAsterisk & slVehicle
        End If
    Next ilLoop
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
Private Sub cmcMoveBack_Click()
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim ilSortOrder As Integer
    Dim ilShowOrder As Integer
    Dim slNameYear As String
    Dim slYear As String
    Dim ilRet As Integer
    For ilIndex = lbcTo.ListCount - 1 To 0 Step -1
        If lbcTo.Selected(ilIndex) Then
            slName = lbcTo.List(ilIndex)
            tgBuyerCode(UBound(tgBuyerCode)) = tgBusCatCode(ilIndex)
            ReDim Preserve tgBuyerCode(0 To UBound(tgBuyerCode) + 1) As SORTCODE
            lbcTo.RemoveItem ilIndex
            'lbcFrom.AddItem slName
            For ilLoop = ilIndex To UBound(tgBusCatCode) - 1 Step 1
                tgBusCatCode(ilLoop) = tgBusCatCode(ilLoop + 1)
            Next ilLoop
            ReDim Preserve tgBusCatCode(0 To UBound(tgBusCatCode) - 1) As SORTCODE
        End If
    Next ilIndex
    If UBound(tgBuyerCode) - 1 > 0 Then
        ArraySortTyp fnAV(tgBuyerCode(), 0), UBound(tgBuyerCode), 0, LenB(tgBuyerCode(0)), 0, LenB(tgBuyerCode(0).sKey), 0
    End If
    ilSortOrder = 0
    ilShowOrder = 1
    lbcFrom.Clear
    For ilLoop = 0 To UBound(tgBuyerCode) - 1 Step 1
        slNameCode = tgBuyerCode(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slNameYear)
        If ilSortOrder = 0 Then
            ilRet = gParseItem(slNameYear, 1, "/", slYear)
            ilRet = gParseItem(slNameYear, 2, "/", slName)
        Else
            ilRet = gParseItem(slNameYear, 1, "/", slName)
            ilRet = gParseItem(slNameYear, 2, "/", slYear)
        End If
        If ilShowOrder = 0 Then
            slName = gSubStr("9999", slYear) & "/" & slName
        Else
            slName = slName & "/" & gSubStr("9999", slYear)
        End If
        slName = Trim$(slName)
        lbcFrom.AddItem slName  'Add ID to list box
    Next ilLoop
    mSetCommands
End Sub
Private Sub cmcMoveBack_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcMoveTo_Click()
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilIndex As Integer
    If lbcFrom.ListIndex >= 0 Then
        ilIndex = lbcFrom.ListIndex
        slName = lbcFrom.List(ilIndex)
        tgBusCatCode(UBound(tgBusCatCode)) = tgBuyerCode(ilIndex)
        ReDim Preserve tgBusCatCode(0 To UBound(tgBusCatCode) + 1) As SORTCODE
        lbcFrom.RemoveItem ilIndex
        lbcTo.AddItem slName
        For ilLoop = ilIndex To UBound(tgBuyerCode) - 1 Step 1
            tgBuyerCode(ilLoop) = tgBuyerCode(ilLoop + 1)
        Next ilLoop
        ReDim Preserve tgBuyerCode(0 To UBound(tgBuyerCode) - 1) As SORTCODE
    End If
    mSetCommands
End Sub
Private Sub cmcMoveTo_GotFocus()
    gCtrlGotFocus ActiveControl
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
    BudTrend.Refresh
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
    On Error Resume Next
    Erase tgBusCatCode
    Erase imVefCode
    Erase lmDollars
    btrDestroy hmBvf
    
    Set BudTrend = Nothing   'Remove data segment

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
        'mComputeTotal
    End If
    mSetCommands
End Sub
Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBudgetPop                      *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mBudgetPop()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slNameCode As String
    Dim slCode As String
    Dim slNameYear As String
    Dim slMnfName As String
    Dim slYear As String
    Dim ilMnfCode As Integer
    Dim ilYear As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilSLoop As Integer
    Dim ilELoop As Integer
    'Show Unique Budget names and year
    imPopReqd = False
    sgBuyerCodeTag = ""
    'ilRet = gPopVehBudgetBox(BudModel, 0, 1, lbcBudget, lbcBudgetCode)
    ilRet = gPopVehBudgetBox(BudTrend, 2, 0, 1, lbcFrom, tgBuyerCode(), sgBuyerCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mBudgetPopErr
        gCPErrorMsg ilRet, "mBudgetPop (gPopVehBudgetBox)", BudTrend
        On Error GoTo 0
        'lbcBudget.AddItem "[None]", 0  'Force as first item on list
        imPopReqd = True
        ilSLoop = LBound(tgBuyerCode)
        ilELoop = UBound(tgBuyerCode) - 1
        ilLoop = ilSLoop
        Do While ilLoop <= ilELoop
            slNameCode = tgBuyerCode(ilLoop).sKey  'lbcBudget.List(ilIndex - 1)
            ilRet = gParseItem(slNameCode, 1, "\", slNameYear)
            ilRet = gParseItem(slNameYear, 2, "/", slMnfName)
            ilRet = gParseItem(slNameYear, 1, "/", slYear)
            slYear = gSubStr("9999", slYear)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilMnfCode = Val(slCode)
            ilYear = Val(slYear)
            'If (tgBvfRec(LBound(tgBvfRec)).tBvf.iMnfBudget = ilMnfCode) And (tgBvfRec(LBound(tgBvfRec)).tBvf.iYear = ilYear) Then
            If (tgBvfRec(igLBBvfRec).tBvf.iMnfBudget = ilMnfCode) And (tgBvfRec(igLBBvfRec).tBvf.iYear = ilYear) Then
                For ilIndex = ilLoop To UBound(tgBuyerCode) - 1 Step 1
                    tgBuyerCode(ilIndex) = tgBuyerCode(ilIndex + 1)
                Next ilIndex
                lbcFrom.RemoveItem ilLoop - LBound(tgBuyerCode)
                ReDim Preserve tgBuyerCode(0 To UBound(tgBuyerCode) - 1)
                ilELoop = UBound(tgBuyerCode) - 1
                If ilLoop >= UBound(tgBuyerCode) - 1 Then
                    Exit Do
                End If
            Else
                ilLoop = ilLoop + 1
            End If
        Loop
    End If
    Exit Sub
mBudgetPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Dim ilRet As Integer
    imFirstActivate = True
    imTerminate = False
    igBDReturn = 0
    imLBBv = 1

    igLBBvfRec = 1  'Match definition in Budget

    Screen.MousePointer = vbHourglass
    hmBvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmBvf, "", sgDBPath & "Bvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Bvf.Btr)", BudTrend
    On Error GoTo 0
    'ReDim tmBvf(1 To 1) As BVF
    ReDim tmBvf(0 To 1) As BVF
    imBvfRecLen = Len(tmBvf(1))
    BudTrend.Height = cmcApply.Top + 5 * cmcApply.Height / 3
    gCenterStdAlone BudTrend
    imChgMode = False
    imBSMode = False
    imDateChgMode = False
    imBypassSetting = False
    imAllClicked = False
    imSetAll = True
    'BudTrend.Show
    Screen.MousePointer = vbHourglass
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    mDatePop
    If imTerminate Then
        Exit Sub
    End If
    lbcTo.Clear
    ReDim tgBusCatCode(0 To 0) As SORTCODE
    mBudgetPop
    If imTerminate Then
        Exit Sub
    End If
    edcRound.Text = "1"
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
'    gCenterModalForm BudTrend
    'plcScreen.Caption = "Trend for " & sgBAName
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
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
'*      Procedure Name:mReadBvfRec                     *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadBvfRec() As Integer
'
'   iRet = mReadBvfRec (iMnfCode, ilYears)
'   Where:
'       ilMnfCode(I)-Budget Name Code
'       ilYears(I)-Year to retrieve
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim ilOffSet As Integer
    Dim ilRecOK As Integer
    Dim ilVeh As Integer
    Dim ilFound As Integer
    Dim slNameYear As String
    Dim slMnfName As String
    Dim slYear As String
    Dim ilMnfCode As Integer
    Dim ilYear As Integer
    Dim ilIndex As Integer
    Dim ilBvf As Integer
    Dim slVehicle As String
    Dim ilSelect As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record

    'ReDim tmBvf(1 To 1) As BVF
    ReDim tmBvf(0 To 1) As BVF
    ilUpper = UBound(tmBvf)
    ReDim imVefCode(0 To 0) As Integer
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            slVehicle = lbcVehicle.List(ilLoop)
            'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
            For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
                If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                    ilFound = False
                    For ilIndex = 0 To UBound(imVefCode) - 1 Step 1
                        If tgBvfRec(ilBvf).tBvf.iVefCode = imVefCode(ilIndex) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilIndex
                    If Not ilFound Then
                        imVefCode(UBound(imVefCode)) = tgBvfRec(ilBvf).tBvf.iVefCode
                        ReDim Preserve imVefCode(0 To UBound(imVefCode) + 1) As Integer
                    End If
                    Exit For
                End If
            Next ilBvf
        End If
    Next ilLoop

    For ilSelect = 0 To lbcTo.ListCount - 1 Step 1
    'For ilSelect = lbcTo.ListCount - 1 To 0 Step -1
        btrExtClear hmBvf   'Clear any previous extend operation
        ilExtLen = Len(tmBvf(1))  'Extract operation record size
        slNameCode = tgBusCatCode(ilSelect).sKey  'lbcBudget.List(ilIndex - 1)
        ilRet = gParseItem(slNameCode, 1, "\", slNameYear)
        ilRet = gParseItem(slNameYear, 2, "/", slMnfName)
        ilRet = gParseItem(slNameYear, 1, "/", slYear)
        slYear = gSubStr("9999", slYear)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilMnfCode = Val(slCode)
        ilYear = Val(slYear)
        tmBvfSrchKey.iYear = ilYear
        tmBvfSrchKey.iSeqNo = 1
        tmBvfSrchKey.iMnfBudget = ilMnfCode
        ilRet = btrGetGreaterOrEqual(hmBvf, tmBvf(ilUpper), imBvfRecLen, tmBvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        'ilRet = btrGetFirst(hmBvf, tgBvfRec(1).tBvf, imBvfRecLen, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
            Call btrExtSetBounds(hmBvf, llNoRec, -1, "UC", "BVF", "") '"EG") 'Set extract limits (all records)
            ilOffSet = gFieldOffset("Bvf", "BvfMnfBudget")
            tlIntTypeBuff.iType = ilMnfCode
            ilRet = btrExtAddLogicConst(hmBvf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
            On Error GoTo mReadBvfRecErr
            gBtrvErrorMsg ilRet, "mReadBvfRec (btrExtAddLogicConst):" & "Bvf.Btr", BudTrend
            On Error GoTo 0
            ilOffSet = gFieldOffset("Bvf", "BvfYear")
            tlIntTypeBuff.iType = ilYear
            ilRet = btrExtAddLogicConst(hmBvf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
            On Error GoTo mReadBvfRecErr
            gBtrvErrorMsg ilRet, "mReadBvfRec (btrExtAddLogicConst):" & "Bvf.Btr", BudTrend
            On Error GoTo 0
            ilRet = btrExtAddField(hmBvf, 0, ilExtLen) 'Extract the whole record
            On Error GoTo mReadBvfRecErr
            gBtrvErrorMsg ilRet, "mReadBvfRec (btrExtAddField):" & "Bvf.Btr", BudTrend
            On Error GoTo 0
            'ilRet = btrExtGetNextExt(hmRpf)    'Extract record
            ilRet = btrExtGetNext(hmBvf, tmBvf(ilUpper), ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo mReadBvfRecErr
                gBtrvErrorMsg ilRet, "mReadBvfRec (btrExtGetNextExt):" & "Bvf.Btr", BudTrend
                On Error GoTo 0
                'ilRet = btrExtGetFirst(hmRpf, tlRpfExt, ilExtLen, llRecPos)
                ilExtLen = Len(tmBvf(1))  'Extract operation record size
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmBvf, tmBvf(ilUpper), ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    'User allow to see vehicle
                    ilRecOK = False
                    For ilVeh = 0 To UBound(imVefCode) - 1 Step 1
                        If imVefCode(ilVeh) = tmBvf(ilUpper).iVefCode Then
                            ilRecOK = True
                            Exit For
                        End If
                    Next ilVeh
                    If ilRecOK Then
                        ilUpper = ilUpper + 1
                        'ReDim Preserve tmBvf(1 To ilUpper) As BVF
                        ReDim Preserve tmBvf(0 To ilUpper) As BVF
                    End If
                    ilRet = btrExtGetNext(hmBvf, tmBvf(ilUpper), ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmBvf, tmBvf(ilUpper), ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
    Next ilSelect
    'mInitBudgetCtrls
    mReadBvfRec = True
    Exit Function
mReadBvfRecErr:
    On Error GoTo 0
    mReadBvfRec = False
    Exit Function
End Function
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
    If (cbcStartDate.Text = "") Or (cbcEndDate.Text = "") Or (lbcTo.ListCount < 2) Then
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
    Unload BudTrend
    igManUnload = NO
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Trend for " & sgBAName
End Sub
