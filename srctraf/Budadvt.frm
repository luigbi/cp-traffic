VERSION 5.00
Begin VB.Form BudAdvt 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4665
   ClientLeft      =   2220
   ClientTop       =   3210
   ClientWidth     =   9225
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
   ScaleHeight     =   4665
   ScaleWidth      =   9225
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Vehicles"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      TabIndex        =   20
      Top             =   3930
      Width           =   1500
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
      Height          =   3645
      Left            =   3255
      ScaleHeight     =   3585
      ScaleWidth      =   5640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   300
      Width           =   5700
      Begin VB.PictureBox pbcLbcAdvt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1890
         Left            =   195
         ScaleHeight     =   1890
         ScaleWidth      =   4335
         TabIndex        =   23
         Top             =   930
         Width           =   4335
      End
      Begin VB.ListBox lbcAdvt 
         Appearance      =   0  'Flat
         Height          =   1920
         Left            =   180
         MultiSelect     =   2  'Extended
         TabIndex        =   22
         Top             =   915
         Width           =   4365
      End
      Begin VB.TextBox edcIndex 
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
         Left            =   4680
         TabIndex        =   15
         Top             =   2220
         Width           =   795
      End
      Begin VB.CommandButton cmcGenerate 
         Appearance      =   0  'Flat
         Caption         =   "&Generate"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4215
         TabIndex        =   13
         Top             =   495
         Width           =   945
      End
      Begin VB.CheckBox ckcYear 
         Caption         =   "1997"
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
         Index           =   4
         Left            =   3255
         TabIndex        =   12
         Top             =   540
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CheckBox ckcYear 
         Caption         =   "1997"
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
         Index           =   3
         Left            =   2490
         TabIndex        =   11
         Top             =   540
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CheckBox ckcYear 
         Caption         =   "1997"
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
         Index           =   2
         Left            =   1725
         TabIndex        =   10
         Top             =   540
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CheckBox ckcYear 
         Caption         =   "1997"
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
         Index           =   1
         Left            =   960
         TabIndex        =   9
         Top             =   540
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CheckBox ckcYear 
         Caption         =   "1997"
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
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   540
         Width           =   780
      End
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
         Left            =   1950
         TabIndex        =   17
         Top             =   3210
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
         Height          =   315
         Left            =   3630
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
         Height          =   315
         Left            =   1050
         TabIndex        =   5
         Top             =   135
         Width           =   1560
      End
      Begin VB.Label lacIndex 
         Appearance      =   0  'Flat
         Caption         =   "Index"
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
         Left            =   4680
         TabIndex        =   14
         Top             =   1935
         Width           =   585
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   180
         X2              =   5565
         Y1              =   3135
         Y2              =   3135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0FFFF&
         X1              =   180
         X2              =   5565
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label lacRound 
         Appearance      =   0  'Flat
         Caption         =   "Round Actuals to"
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
         Left            =   255
         TabIndex        =   16
         Top             =   3270
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
         Left            =   2715
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
      Left            =   3345
      TabIndex        =   18
      Top             =   4320
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
      TabStop         =   0   'False
      Top             =   15
      Width           =   5640
   End
   Begin VB.CommandButton cmcApply 
      Appearance      =   0  'Flat
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4980
      TabIndex        =   19
      Top             =   4320
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
      Height          =   3510
      Left            =   165
      ScaleHeight     =   3450
      ScaleWidth      =   2880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   2940
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         Height          =   3390
         Left            =   30
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   30
         Width           =   2835
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
      Left            =   2025
      TabIndex        =   21
      Top             =   3945
      Width           =   1125
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   135
      Top             =   4260
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "BudAdvt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Budadvt.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: BudAdvt.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Trend input screen code
Option Explicit
Option Compare Text
'Contract
Dim hmCHF As Integer    'Contract Header file handle
Dim imCHFRecLen As Integer        'Chf record length
Dim tmChf As CHF
Dim hmClf As Integer    'Contract Header file handle
Dim imClfRecLen As Integer        'Chf record length
Dim tmClf As CLF
Dim hmCff As Integer    'Contract Header file handle
Dim imCffRecLen As Integer        'Chf record length
Dim tmCff As CFF
Dim hmAdf As Integer    'Salesperson file handle
Dim imAdfRecLen As Integer        'Chf record length
Dim tmAdf As ADF
Dim tmAdfSrchKey As INTKEY0
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

'Dim imListField(1 To 4) As Integer
Dim imListField(0 To 3) As Integer
'Dim imListFieldChar(1 To 3) As Integer


Dim tmChfAdvtExt() As CHFADVTEXT
Dim tmAdvtTotals() As ADVTTOTALS
Dim tmAdvtValues() As ADVTVALUES

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
        lbcAdvt.Clear
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
        lbcAdvt.Clear
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
Private Sub ckcYear_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcYear(Index).Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    lbcAdvt.Clear
    mSetCommands
End Sub
Private Sub cmcApply_Click()
    Dim slVehicle As String
    Dim ilWk As Integer
    Dim slStr As String
    Dim slAsterisk As String
    Dim slChar As String
    Dim slRound As String
    Dim ilBvf As Integer
    Dim ilLoop As Integer
    Dim slName As String
    Dim slIndex As String
    Dim slAdvt As String
    Dim slTotal As String
    Dim ilRet As Integer
    Dim llTotal As Long
    Dim ilCount As Integer
    Dim ilVLoop As Integer
    Dim ilYears As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    slRound = edcRound.Text
    If slRound = "" Then
        slRound = "1"
    End If
    ilRet = MsgBox("This will Replace the Budget Values, Ok to Proceed", vbYesNo + vbQuestion, "Advertiser")
    If ilRet = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    'Clear old values
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
                        tgBvfRec(ilBvf).tBvf.lGross(ilWk) = 0
                    Next ilWk
                End If
            Next ilBvf
        End If
    Next ilLoop
    'Replace values
    For ilLoop = 0 To lbcAdvt.ListCount - 1 Step 1
        slName = lbcAdvt.List(ilLoop)
        ilRet = gParseItem(slName, 1, "|", slAdvt)
        'slAdvt = Left$(slName, imListFieldChar(1) - 1)
        ilRet = gParseItem(slName, 2, "|", slTotal)
        gUnformatStr slTotal, UNFMTDEFAULT, slTotal
        'slTotal = Mid$(slName, imListFieldChar(1), imListFieldChar(2) - imListFieldChar(1))
        ilRet = gParseItem(slName, 3, "|", slIndex)
        'slIndex = Mid$(slName, imListFieldChar(2))
        'For ilVLoop = LBound(tmAdvtValues) To UBound(tmAdvtValues) - 1 Step 1
        '    If tmAdvtTotals(ilLoop).iPtAdvtTotals = tmAdvtValues(ilVLoop).iPtAdvtTotals Then
        ilVLoop = tmAdvtTotals(ilLoop).iFirstValue
        Do While ilVLoop <> -1
            'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
            For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
                If (tmAdvtValues(ilVLoop).iVefCode = tgBvfRec(ilBvf).tBvf.iVefCode) And (tmAdvtValues(ilVLoop).iSofCode = tgBvfRec(ilBvf).tBvf.iSofCode) Then
                    For ilWk = imSDateSelectedIndex + 1 To imEDateSelectedIndex + 1 Step 1
                        llTotal = 0
                        ilCount = 0
                        For ilYears = 0 To 4 Step 1
                            If ckcYear(ilYears).Value = vbChecked Then
                                If tmAdvtValues(ilVLoop).lDollars(ilYears, ilWk) > 0 Then
                                    llTotal = llTotal + tmAdvtValues(ilVLoop).lDollars(ilYears, ilWk)
                                    ilCount = ilCount + 1
                                End If
                            End If
                        Next ilYears
                        If ilCount > 1 Then
                            llTotal = llTotal / ilCount
                        End If
                        slStr = gLongToStrDec(llTotal, 0)
                        slStr = gMulStr(slIndex, slStr)
                        slStr = gRoundStr(slStr, slRound, 0)
                        tgBvfRec(ilBvf).tBvf.lGross(ilWk) = tgBvfRec(ilBvf).tBvf.lGross(ilWk) + gStrDecToLong(slStr, 0)
                    Next ilWk
                End If
            Next ilBvf
            ilVLoop = tmAdvtValues(ilVLoop).iNextValue
        Loop
            'End If
        'Next ilVLoop
    Next ilLoop
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
Private Sub cmcGenerate_Click()
    Dim ilRet As Integer
    ilRet = MsgBox("This will take some time, Ok to Proceed", vbYesNo + vbQuestion, "Advertiser")
    If ilRet = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ilRet = mBuildAdvt()
    mSetCommands
    Screen.MousePointer = vbDefault
End Sub
Private Sub edcIndex_Change()
    Dim ilLoop As Integer
    Dim slName As String
    Dim slIndex As String
    Dim slAdvt As String
    Dim slTotal As String
    Dim ilRet As Integer

    'Set index value from index to advertiser
    For ilLoop = 0 To lbcAdvt.ListCount - 1 Step 1
        If lbcAdvt.Selected(ilLoop) Then
            slName = lbcAdvt.List(ilLoop)
            ilRet = gParseItem(slName, 1, "|", slAdvt)
            'slAdvt = Left$(slName, imListFieldChar(1) - 1)
            ilRet = gParseItem(slName, 2, "|", slTotal)
            'slTotal = Mid$(slName, imListFieldChar(1), imListFieldChar(2) - imListFieldChar(1))
            slIndex = edcIndex.Text
            lbcAdvt.List(ilLoop) = slAdvt & "|" & slTotal & "|" & slIndex
        End If
    Next ilLoop
    pbcLbcAdvt_Paint
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
    If (igWinStatus(BUDGETSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    'gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    BudAdvt.Refresh
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
    
    Erase tmChfAdvtExt
    Erase tmAdvtTotals
    Erase tmAdvtValues
    Erase tgClfBud
    Erase tgCffBud

    btrDestroy hmCHF
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmAdf
    
    Set BudAdvt = Nothing   'Remove data segment
    
End Sub

Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcAdvt_Click()
    Dim ilLoop As Integer
    Dim slName As String
    Dim slIndex As String
    Dim slAdvt As String
    Dim slTotal As String
    Dim ilRet As Integer
    If lbcAdvt.SelCount = 0 Then
        edcIndex.Text = ""
    ElseIf lbcAdvt.SelCount = 1 Then
        'Set index from advertiser to Index
        For ilLoop = 0 To lbcAdvt.ListCount - 1 Step 1
            If lbcAdvt.Selected(ilLoop) Then
                slName = lbcAdvt.List(ilLoop)
                ilRet = gParseItem(slName, 3, "|", slIndex)
                'slIndex = Mid$(slName, imListFieldChar(2))
                edcIndex.Text = slIndex
            End If
        Next ilLoop
    Else
        'Set index value from index to advertiser
        For ilLoop = 0 To lbcAdvt.ListCount - 1 Step 1
            If lbcAdvt.Selected(ilLoop) Then
                slName = lbcAdvt.List(ilLoop)
                ilRet = gParseItem(slName, 1, "|", slAdvt)
                'slAdvt = Left$(slName, imListFieldChar(1) - 1)
                ilRet = gParseItem(slName, 2, "|", slTotal)
                'slTotal = Mid$(slName, imListFieldChar(1), imListFieldChar(2) - imListFieldChar(1))
                slIndex = edcIndex.Text
                lbcAdvt.List(ilLoop) = slAdvt & "|" & slTotal & "|" & slIndex
            End If
        Next ilLoop
        pbcLbcAdvt_Paint
    End If
    mSetCommands
End Sub
Private Sub lbcAdvt_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub lbcAdvt_Scroll()
    pbcLbcAdvt_Paint
End Sub

Private Sub lbcVehicle_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
        'mComputeTotal
    End If
    lbcAdvt.Clear
    mSetCommands
End Sub
Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBuildAdvt                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Build The Advertiser Arrays    *
'*                      with dollars                   *
'*                                                     *
'*******************************************************
Private Function mBuildAdvt() As Integer
    Dim ilLoop As Integer
    Dim ilBvf As Integer
    Dim slVehicle As String
    Dim ilWk As Integer
    Dim slStr As String
    Dim slRound As String
    Dim ilCount As Integer
    Dim ilRet As Integer
    Dim ilSpot As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStatus As String
    Dim slCntrType As String
    Dim ilHOType As Integer
    Dim ilCnt As Integer
    Dim ilClf As Integer
    Dim ilSlf As Integer
    Dim ilSofCode As Integer
    Dim ilIncludeHistory As Integer
    Dim llClfStartDate As Long
    Dim llClfEndDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llWkStartDate As Long
    Dim ilCff As Integer
    Dim llFlStartDate As Long
    Dim llFlEndDate As Long
    Dim ilNoSpots As Integer
    Dim llSpotPrice As Long
    Dim llPrice As Long
    Dim ilDay As Integer
    Dim ilYears As Integer
    Dim ilTIndex As Integer
    Dim ilVIndex As Integer
    Dim ilTLoop As Integer
    Dim ilVLoop As Integer
    Dim llTotals As Long
    Dim ilWkNo As Integer
    Dim ilFirstLastWk As Integer
    Dim ilVLast As Integer

    If imEDateSelectedIndex < imSDateSelectedIndex Then
        Screen.MousePointer = vbDefault
        MsgBox "End Date is prior to Start Date", vbOKOnly + vbExclamation, "Trend"
        cbcEndDate.SetFocus
        mBuildAdvt = False
        Exit Function
    End If
    ReDim tmAdvtTotals(0 To 0) As ADVTTOTALS
    ReDim tmAdvtValues(0 To 0) As ADVTVALUES
    'Get all contract active for dates requested
    For ilYears = 0 To 4 Step 1
        If ckcYear(ilYears).Value = vbChecked Then
            slStartDate = cbcStartDate.List(imSDateSelectedIndex)
            gObtainWkNo 5, slStartDate, ilWkNo, ilFirstLastWk
            slStr = ckcYear(ilYears).Caption
            slStartDate = gObtainStartDateForWkNo(5, ilWkNo, Val(slStr))
            slEndDate = cbcEndDate.List(imEDateSelectedIndex)
            slStr = ckcYear(ilYears).Caption
            If imEDateSelectedIndex < cbcEndDate.ListCount - 1 Then
                gObtainWkNo 5, slEndDate, ilWkNo, ilFirstLastWk
                slEndDate = gObtainStartDateForWkNo(5, ilWkNo, Val(slStr))
            Else
                slEndDate = gObtainStartDateForWkNo(5, -1, Val(slStr))
            End If
            slEndDate = gObtainNextSunday(slEndDate)
            llStartDate = gDateValue(slStartDate)
            llEndDate = gDateValue(slEndDate)
            slStatus = "HO"
            slCntrType = "CVTRQ"
            ilHOType = 1
            sgCntrForDateStamp = ""
            ilRet = gObtainCntrForDate(BudAdvt, slStartDate, slEndDate, slStatus, slCntrType, ilHOType, tmChfAdvtExt())
            For ilCnt = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
                ilIncludeHistory = False
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, tmChfAdvtExt(ilCnt).lCode, ilIncludeHistory, tgChfBud, tgClfBud(), tgCffBud())
                'Obtain selling office
                ilSofCode = -1
                For ilSlf = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
                    If tgChfBud.iSlfCode(0) = tgMSlf(ilSlf).iCode Then
                        ilSofCode = tgMSlf(ilSlf).iSofCode
                        Exit For
                    End If
                Next ilSlf

                For ilClf = LBound(tgClfBud) To UBound(tgClfBud) - 1 Step 1
                    If (tgClfBud(ilClf).ClfRec.sType = "S") Or (tgClfBud(ilClf).ClfRec.sType = "H") Then
                        gUnpackDateLong tgClfBud(ilClf).ClfRec.iStartDate(0), tgClfBud(ilClf).ClfRec.iStartDate(1), llClfStartDate
                        gUnpackDateLong tgClfBud(ilClf).ClfRec.iEndDate(0), tgClfBud(ilClf).ClfRec.iEndDate(1), llClfEndDate
                        If (llClfStartDate <= llClfEndDate) And (llEndDate >= llClfStartDate) And (llStartDate <= llClfEndDate) Then
                            For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
                                If lbcVehicle.Selected(ilLoop) Then
                                    slVehicle = lbcVehicle.List(ilLoop)
                                    'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
                                    For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
                                        If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                                            If (tgClfBud(ilClf).ClfRec.iVefCode = tgBvfRec(ilBvf).tBvf.iVefCode) And (ilSofCode = tgBvfRec(ilBvf).tBvf.iSofCode) Then
                                                ilTIndex = -1
                                                ilVIndex = -1
                                                For ilTLoop = LBound(tmAdvtTotals) To UBound(tmAdvtTotals) - 1 Step 1
                                                    If tgChfBud.iAdfCode = tmAdvtTotals(ilTLoop).iAdfCode Then
                                                        ilTIndex = ilTLoop
                                                        'For ilVLoop = LBound(tmAdvtValues) To UBound(tmAdvtValues) - 1 Step 1
                                                        '    If (ilTLoop = tmAdvtValues(ilVLoop).iPtAdvtTotals) And (tmAdvtValues(ilVLoop).iVefCode = tgClfBud(ilClf).ClfRec.iVefCode) And (tmAdvtValues(ilVLoop).iSofCode = ilSofCode) Then
                                                        '        ilVIndex = ilVLoop
                                                        '        Exit For
                                                        '    End If
                                                        'Next ilVLoop
                                                        ilVLoop = tmAdvtTotals(ilTLoop).iFirstValue
                                                        Do While ilVLoop <> -1
                                                            If (tmAdvtValues(ilVLoop).iVefCode = tgClfBud(ilClf).ClfRec.iVefCode) And (tmAdvtValues(ilVLoop).iSofCode = ilSofCode) Then
                                                                ilVIndex = ilVLoop
                                                                Exit Do
                                                            End If
                                                            ilVLast = ilVLoop
                                                            ilVLoop = tmAdvtValues(ilVLoop).iNextValue
                                                        Loop
                                                        Exit For
                                                    End If
                                                Next ilTLoop
                                                If ilVIndex >= 0 Then
                                                ElseIf ilTIndex >= 0 Then
                                                    ilVIndex = UBound(tmAdvtValues)
                                                    tmAdvtValues(ilVLast).iNextValue = ilVIndex
                                                    ReDim Preserve tmAdvtValues(0 To ilVIndex + 1) As ADVTVALUES
                                                    tmAdvtValues(ilVIndex).iPtAdvtTotals = ilTIndex
                                                    tmAdvtValues(ilVIndex).iVefCode = tgClfBud(ilClf).ClfRec.iVefCode
                                                    tmAdvtValues(ilVIndex).iSofCode = ilSofCode
                                                    tmAdvtValues(ilVIndex).iNextValue = -1
                                                Else
                                                    ilTIndex = UBound(tmAdvtTotals)
                                                    ReDim Preserve tmAdvtTotals(0 To ilTIndex + 1) As ADVTTOTALS
                                                    If tmAdf.iCode <> tgChfBud.iAdfCode Then
                                                        tmAdfSrchKey.iCode = tgChfBud.iAdfCode
                                                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                    End If
                                                    If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                                                        tmAdvtTotals(ilTIndex).sKey = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                                                    Else
                                                        tmAdvtTotals(ilTIndex).sKey = tmAdf.sName
                                                    End If
                                                    tmAdvtTotals(ilTIndex).iAdfCode = tgChfBud.iAdfCode
                                                    tmAdvtTotals(ilTIndex).sIndex = "1.0"
                                                    tmAdvtTotals(ilTIndex).iPtAdvtTotals = ilTIndex
                                                    ilVIndex = UBound(tmAdvtValues)
                                                    tmAdvtTotals(ilTIndex).iFirstValue = ilVIndex
                                                    ReDim Preserve tmAdvtValues(0 To ilVIndex + 1) As ADVTVALUES
                                                    tmAdvtValues(ilVIndex).iPtAdvtTotals = ilTIndex
                                                    tmAdvtValues(ilVIndex).iVefCode = tgClfBud(ilClf).ClfRec.iVefCode
                                                    tmAdvtValues(ilVIndex).iSofCode = ilSofCode
                                                    tmAdvtValues(ilVIndex).iNextValue = -1
                                                End If
                                                'For ilWk = LBound(tgBvfRec(ilBvf).tBvf.lGross) To UBound(tgBvfRec(ilBvf).tBvf.lGross) Step 1
                                                For ilWk = imSDateSelectedIndex + 1 To imEDateSelectedIndex + 1 Step 1
                                                    'slStr = cbcStartDate.List(ilWk - 1)
                                                    llWkStartDate = llStartDate + 7 * (ilWk - 1) 'gDateValue(slStr)
                                                    ilCff = tgClfBud(ilClf).iFirstCff
                                                    Do While ilCff <> -1
                                                        'Status: 0=New; 1=Old & retain; 2=Old & delete;-1=New, unused
                                                        gUnpackDateLong tgCffBud(ilCff).CffRec.iStartDate(0), tgCffBud(ilCff).CffRec.iStartDate(1), llFlStartDate
                                                        gUnpackDateLong tgCffBud(ilCff).CffRec.iEndDate(0), tgCffBud(ilCff).CffRec.iEndDate(1), llFlEndDate
                                                        ilDay = gWeekDayLong(llFlStartDate)
                                                        Do While ilDay <> 0
                                                            llFlStartDate = llFlStartDate - 1
                                                            ilDay = gWeekDayLong(llFlStartDate)
                                                        Loop
                                                        ilNoSpots = 0
                                                        If (llWkStartDate >= llFlStartDate) And (llWkStartDate <= llFlEndDate) Then
                                                            If tgCffBud(ilCff).CffRec.sDyWk = "D" Then
                                                                For ilSpot = 0 To 6 Step 1
                                                                    ilNoSpots = ilNoSpots + tgCffBud(ilCff).CffRec.iDay(ilSpot)
                                                                Next ilSpot
                                                            Else
                                                                ilNoSpots = tgCffBud(ilCff).CffRec.iSpotsWk + tgCffBud(ilCff).CffRec.iXSpotsWk
                                                            End If
                                                            Select Case tgCffBud(ilCff).CffRec.sPriceType
                                                                Case "T"
                                                                    llSpotPrice = tgCffBud(ilCff).CffRec.lActPrice
                                                                Case Else
                                                                    llSpotPrice = 0
                                                            End Select
                                                            llPrice = ilNoSpots * llSpotPrice
                                                            slStr = gLongToStrDec(llPrice, 2)
                                                            slStr = gRoundStr(slStr, slRound, 0)
                                                            tmAdvtValues(ilVIndex).lDollars(ilYears, ilWk) = tmAdvtValues(ilVIndex).lDollars(ilYears, ilWk) + gStrDecToLong(slStr, 0)
                                                        End If
                                                        If llWkStartDate < llFlStartDate Then
                                                            Exit Do
                                                        End If
                                                        ilCff = tgCffBud(ilCff).iNextCff
                                                    Loop
                                                Next ilWk
                                                Exit For
                                            End If
                                        End If
                                    Next ilBvf
                                End If
                            Next ilLoop
                        End If
                    End If
                Next ilClf
            Next ilCnt
        End If
    Next ilYears
    For ilTLoop = LBound(tmAdvtTotals) To UBound(tmAdvtTotals) - 1 Step 1
        'For ilVLoop = LBound(tmAdvtValues) To UBound(tmAdvtValues) - 1 Step 1
        '    If tmAdvtTotals(ilTLoop).iPtAdvtTotals = tmAdvtValues(ilVLoop).iPtAdvtTotals Then
        '        For ilYears = 0 To 4 Step 1
        '            If ckcYear(ilYears).Value Then
        '                For ilWk = 1 To 53 Step 1
        '                    tmAdvtTotals(ilTLoop).lTotal(ilYears) = tmAdvtTotals(ilTLoop).lTotal(ilYears) + tmAdvtValues(ilVLoop).lDollars(ilYears, ilWk)
        '                Next ilWk
        '            End If
        '        Next ilYears
        '    End If
        'Next ilVLoop
        ilVLoop = tmAdvtTotals(ilTLoop).iFirstValue
        Do While ilVLoop <> -1
            For ilYears = 0 To 4 Step 1
                If ckcYear(ilYears).Value = vbChecked Then
                    For ilWk = 1 To 53 Step 1
                        tmAdvtTotals(ilTLoop).lTotal(ilYears) = tmAdvtTotals(ilTLoop).lTotal(ilYears) + tmAdvtValues(ilVLoop).lDollars(ilYears, ilWk)
                    Next ilWk
                End If
            Next ilYears
            ilVLoop = tmAdvtValues(ilVLoop).iNextValue
        Loop
    Next ilTLoop
    If UBound(tmAdvtTotals) > LBound(tmAdvtTotals) Then
        ArraySortTyp fnAV(tmAdvtTotals(), 0), UBound(tmAdvtTotals), 0, LenB(tmAdvtTotals(0)), 0, LenB(tmAdvtTotals(0).sKey), 0
    End If
    lbcAdvt.Clear
    For ilTLoop = LBound(tmAdvtTotals) To UBound(tmAdvtTotals) - 1 Step 1
        ilCount = 0
        llTotals = 0
        For ilYears = 0 To 4 Step 1
            If ckcYear(ilYears).Value = vbChecked Then
                If tmAdvtTotals(ilTLoop).lTotal(ilYears) > 0 Then
                    ilCount = ilCount + 1
                    llTotals = llTotals + tmAdvtTotals(ilTLoop).lTotal(ilYears)
                End If
            End If
        Next ilYears
        If ilCount > 1 Then
            llTotals = llTotals / ilCount
        End If
        tmAdvtTotals(ilTLoop).sTotal = gLongToStrDec(llTotals, 0)
        slStr = tmAdvtTotals(ilTLoop).sTotal
        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr
        'lbcAdvt.AddItem gAlignStringByPixel(Trim$(tmAdvtTotals(ilTLoop).sKey) & "|" & Trim$(slStr) & "|" & tmAdvtTotals(ilTLoop).sIndex, "|", imListField(), imListFieldChar())
        lbcAdvt.AddItem Trim$(tmAdvtTotals(ilTLoop).sKey) & "|" & Trim$(slStr) & "|" & tmAdvtTotals(ilTLoop).sIndex
    Next ilTLoop
    pbcLbcAdvt_Paint
    mBuildAdvt = True
End Function
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
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slDate As String
    Dim ilLoop As Integer
    imTerminate = False
    imFirstActivate = True
    igBDReturn = 0
    
    igLBBvfRec = 1  'Match definition in Budget

    Screen.MousePointer = vbHourglass
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", BudAdvt
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", BudAdvt
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", BudAdvt
    On Error GoTo 0
    imCffRecLen = Len(tmCff)
    hmAdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", BudAdvt
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    BudAdvt.Height = cmcApply.Top + 5 * cmcApply.Height / 3
    gCenterStdAlone BudAdvt
    imChgMode = False
    imBSMode = False
    imDateChgMode = False
    imBypassSetting = False
    imAllClicked = False
    imSetAll = True
'    imListField(1) = 15
'    imListField(2) = 26 * igAlignCharWidth
'    imListField(3) = 38 * igAlignCharWidth
'    imListField(4) = 60 * igAlignCharWidth
    imListField(0) = 15
    imListField(1) = 26 * igAlignCharWidth
    imListField(2) = 38 * igAlignCharWidth
    imListField(3) = 60 * igAlignCharWidth

    'imListFieldChar(1) = 0
    'imListFieldChar(2) = 0
    'imListFieldChar(3) = 0
    'BudAdvt.Show
    Screen.MousePointer = vbHourglass
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    mDatePop
    If imTerminate Then
        Exit Sub
    End If
    ilRet = gObtainSalesperson()
    slDate = Format$(gNow(), "m/d/yy")
    slDate = gObtainNextSunday(slDate)
    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
    'If Val(slYear) < tgBvfRec(LBound(tgBvfRec)).tBvf.iYear Then
    If Val(slYear) < tgBvfRec(igLBBvfRec).tBvf.iYear Then
        'slYear = Trim$(Str$(tgBvfRec(LBound(tgBvfRec)).tBvf.iYear - 1))
        slYear = Trim$(Str$(tgBvfRec(igLBBvfRec).tBvf.iYear - 1))
    End If
    For ilLoop = 0 To 4 Step 1
        If Val(slYear) < 1995 Then
            Exit For
        End If
        ckcYear(ilLoop).Caption = slYear
        ckcYear(ilLoop).Visible = True
        slYear = Trim$(Str$(Val(slYear) - 1))
    Next ilLoop
    edcRound.Text = "1"
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    'plcScreen.Caption = "Advertisers for " & sgBAName
'    gCenterModalForm BudAdvt
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

    If (cbcStartDate.Text = "") Or (cbcEndDate.Text = "") Or (lbcAdvt.ListCount <= 0) Then
        cmcApply.Enabled = False
    Else
        If (lbcVehicle.SelCount > 0) And (lbcAdvt.ListCount > 0) Then
            cmcApply.Enabled = True
        Else
            cmcApply.Enabled = False
        End If
    End If
    If (cbcStartDate.Text = "") Or (cbcEndDate.Text = "") Or ((Not ckcYear(0).Value = vbChecked) And (Not ckcYear(1).Value = vbChecked) And (Not ckcYear(2).Value = vbChecked) And (Not ckcYear(3).Value = vbChecked) And (Not ckcYear(4).Value = vbChecked)) Then
        cmcGenerate.Enabled = False
    Else
        If lbcVehicle.SelCount > 0 Then
            cmcGenerate.Enabled = True
        Else
            cmcGenerate.Enabled = False
        End If
    End If
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
    Unload BudAdvt
    igManUnload = NO
End Sub

Private Sub pbcLbcAdvt_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilAdvtEnd As Integer
    Dim ilField As Integer
    Dim llWidth As Long
    'Dim slFields(1 To 3) As String
    Dim slFields(0 To 2) As String
    Dim llFgColor As Long
    Dim ilFieldIndex As Integer
    
    ilAdvtEnd = lbcAdvt.TopIndex + lbcAdvt.Height \ fgListHtArial825
    If ilAdvtEnd > lbcAdvt.ListCount Then
        ilAdvtEnd = lbcAdvt.ListCount
    End If
    If lbcAdvt.ListCount <= lbcAdvt.Height \ fgListHtArial825 Then
        llWidth = lbcAdvt.Width - 30
    Else
        llWidth = lbcAdvt.Width - igScrollBarWidth - 30
    End If
    pbcLbcAdvt.Width = llWidth
    pbcLbcAdvt.Cls
    llFgColor = pbcLbcAdvt.ForeColor
    For ilLoop = lbcAdvt.TopIndex To ilAdvtEnd - 1 Step 1
        pbcLbcAdvt.ForeColor = llFgColor
        If lbcAdvt.MultiSelect = 0 Then
            If lbcAdvt.ListIndex = ilLoop Then
                gPaintArea pbcLbcAdvt, CSng(0), CSng((ilLoop - lbcAdvt.TopIndex) * fgListHtArial825), CSng(pbcLbcAdvt.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcAdvt.ForeColor = vbWhite
            End If
        Else
            If lbcAdvt.Selected(ilLoop) Then
                gPaintArea pbcLbcAdvt, CSng(0), CSng((ilLoop - lbcAdvt.TopIndex) * fgListHtArial825), CSng(pbcLbcAdvt.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcAdvt.ForeColor = vbWhite
            End If
        End If
        slStr = lbcAdvt.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilFieldIndex = 0
        For ilField = LBound(imListField) To UBound(imListField) - 1 Step 1
            pbcLbcAdvt.CurrentX = imListField(ilField)
            pbcLbcAdvt.CurrentY = (ilLoop - lbcAdvt.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            gAdjShowLen pbcLbcAdvt, slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcAdvt.Print slStr
            ilFieldIndex = ilFieldIndex + 1
        Next ilField
        pbcLbcAdvt.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Advertisers for " & sgBAName
End Sub
