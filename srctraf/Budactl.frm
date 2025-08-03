VERSION 5.00
Begin VB.Form BudActl 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4305
   ClientLeft      =   975
   ClientTop       =   1020
   ClientWidth     =   6795
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
   ScaleWidth      =   6795
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Vehicles"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      TabIndex        =   12
      Top             =   3645
      Width           =   1260
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
      Height          =   1905
      Left            =   3420
      ScaleHeight     =   1845
      ScaleWidth      =   3060
      TabIndex        =   3
      Top             =   270
      Width           =   3120
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
         Left            =   1905
         TabIndex        =   10
         Top             =   1335
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
         Left            =   1050
         TabIndex        =   7
         Top             =   540
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
      Begin VB.Line Line1 
         BorderColor     =   &H00E0FFFF&
         X1              =   195
         X2              =   2820
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   195
         X2              =   2820
         Y1              =   1185
         Y2              =   1185
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
         Left            =   210
         TabIndex        =   11
         Top             =   1380
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
         Left            =   120
         TabIndex        =   6
         Top             =   600
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
      Left            =   1980
      TabIndex        =   8
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
      Left            =   3615
      TabIndex        =   9
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
      Height          =   3360
      Left            =   150
      ScaleHeight     =   3300
      ScaleWidth      =   2895
      TabIndex        =   1
      Top             =   270
      Width           =   2955
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         Height          =   3180
         Left            =   60
         TabIndex        =   2
         Top             =   45
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
      Left            =   1995
      TabIndex        =   13
      Top             =   3630
      Width           =   1110
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   180
      Top             =   3915
      Width           =   360
   End
End
Attribute VB_Name = "BudActl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: BudActl.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Trend input screen code
Option Explicit
Option Compare Text
'Contract
Dim hmChf As Integer    'Contract Header file handle
Dim imChfRecLen As Integer        'Chf record length
Dim tmChf As CHF
Dim hmClf As Integer    'Contract Header file handle
Dim imClfRecLen As Integer        'Chf record length
Dim tmClf As CLF
Dim hmCff As Integer    'Contract Header file handle
Dim imCffRecLen As Integer        'Chf record length
Dim tmCff As CFF
Dim hmSlf As Integer    'Salesperson file handle
Dim imSlfRecLen As Integer        'Chf record length
Dim tmSlf As SLF
'Program library dates Field Areas
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcSelection)
Dim imSDateSelectedIndex As Integer
Dim imEDateSelectedIndex As Integer
Dim imDateComboBoxIndex As Integer
Dim imDateChgMode As Integer
Dim tmChfAdvtExt() As CHFADVTEXT
Private Sub cbcEndDate_Change()
    Dim ilRet As Integer
    Dim slDate As String
    Dim ilResult As Integer
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
    Dim ilRet As Integer
    Dim slDate As String
    Dim ilResult As Integer
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
    Dim value As Integer
    value = False
    If ckcAll.value = vbChecked Then
        value = True
    End If
    'End of coded added
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    ilValue = value
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
    Dim slIndex As String
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
    Dim ilSpot As Integer
    Dim llGross As Long
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
    Dim llWkEndDate As Long
    Dim ilCff As Integer
    Dim llFlStartDate As Long
    Dim llFlEndDate As Long
    Dim ilNoSpots As Integer
    Dim llSpotPrice As Long
    Dim llPrice As Long
    Dim ilDay As Integer
    If imEDateSelectedIndex < imSDateSelectedIndex Then
        MsgBox "End Date is prior to Start Date", vbOkOnly + vbExclamation, "Trend"
        cbcEndDate.SetFocus
        Exit Sub
    End If
    slRound = edcRound.Text
    If slRound = "" Then
        slRound = "1"
    End If
    ilRet = MsgBox("This will Replace The Budget Values, Ok to Proceed", vbYesNo + vbQuestion, "Trend")
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
            For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
                If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                    'For ilWk = LBound(tgBvfRec(ilBvf).tBvf.lGross) To UBound(tgBvfRec(ilBvf).tBvf.lGross) Step 1
                    For ilWk = imSDateSelectedIndex + 1 To imEDateSelectedIndex + 1 Step 1
                        tgBvfRec(ilBvf).tBvf.lGross(ilWk) = 0
                    Next ilWk
                End If
            Next ilBvf
        End If
    Next ilLoop
    'Get all contract active for dates requested
    slStartDate = cbcStartDate.List(imSDateSelectedIndex)
    slEndDate = cbcEndDate.List(imEDateSelectedIndex)
    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    slStatus = "HO"
    slCntrType = "CVTRQ"
    ilHOType = 1
    ilRet = gObtainCntrForDate(BudActl, slStartDate, slEndDate, slStatus, slCntrType, ilHOType, tmChfAdvtExt())
    For ilCnt = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
        ilIncludeHistory = False
        ilRet = gObtainCntr(hmChf, hmClf, hmCff, tmChfAdvtExt(ilCnt).lcode, ilIncludeHistory, tgChf, tgClf(), tgCff())
        'Obtain selling office
        ilSofCode = -1
        For ilSlf = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
            If tgChf.islfCode(0) = tgMSlf(ilSlf).iCode Then
                ilSofCode = tgMSlf(ilSlf).iSofCode
                Exit For
            End If
        Next ilSlf
        For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
            If (tgClf(ilClf).ClfRec.sType = "S") Or (tgClf(ilClf).ClfRec.sType = "H") Then
                gUnpackDateLong tgClf(ilClf).ClfRec.iStartDate(0), tgClf(ilClf).ClfRec.iStartDate(1), llClfStartDate
                gUnpackDateLong tgClf(ilClf).ClfRec.iEndDate(0), tgClf(ilClf).ClfRec.iEndDate(1), llClfEndDate
                If (llClfStartDate <= llClfEndDate) And (llEndDate >= llClfStartDate) And (llStartDate <= llClfEndDate) Then
                    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
                        If lbcVehicle.Selected(ilLoop) Then
                            slVehicle = lbcVehicle.List(ilLoop)
                            slVehicle = lbcVehicle.List(ilLoop)
                            slChar = Left$(slVehicle, 1)
                            slAsterisk = ""
                            Do While slChar = "*"
                                slAsterisk = slAsterisk & slChar
                                slVehicle = Mid$(slVehicle, 2)
                                slChar = Left$(slVehicle, 1)
                            Loop
                            For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
                                If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                                    If (tgClf(ilClf).ClfRec.iVefCode = tgBvfRec(ilBvf).tBvf.iVefCode) And (ilSofCode = tgBvfRec(ilBvf).tBvf.iSofCode) Then
                                        'For ilWk = LBound(tgBvfRec(ilBvf).tBvf.lGross) To UBound(tgBvfRec(ilBvf).tBvf.lGross) Step 1
                                        For ilWk = imSDateSelectedIndex + 1 To imEDateSelectedIndex + 1 Step 1
                                            slStr = cbcStartDate.List(ilWk - 1)
                                            llWkStartDate = gDateValue(slStr)
                                            ilCff = tgClf(ilClf).iFirstCff
                                            Do While ilCff <> -1
                                                'Status: 0=New; 1=Old & retain; 2=Old & delete;-1=New, unused
                                                gUnpackDateLong tgCff(ilCff).CffRec.iStartDate(0), tgCff(ilCff).CffRec.iStartDate(1), llFlStartDate
                                                gUnpackDateLong tgCff(ilCff).CffRec.iEndDate(0), tgCff(ilCff).CffRec.iEndDate(1), llFlEndDate
                                                ilDay = gWeekDayLong(llFlStartDate)
                                                Do While ilDay <> 0
                                                    llFlStartDate = llFlStartDate - 1
                                                Loop
                                                ilNoSpots = 0
                                                If (llWkStartDate >= llFlStartDate) And (llWkStartDate <= llFlEndDate) Then
                                                    If tgCff(ilCff).CffRec.sDyWk = "D" Then
                                                        For ilSpot = 0 To 6 Step 1
                                                            ilNoSpots = ilNoSpots + tgCff(ilCff).CffRec.iDay(ilSpot)
                                                        Next ilSpot
                                                    Else
                                                        ilNoSpots = tgCff(ilCff).CffRec.iSpotsWk + tgCff(ilCff).CffRec.iXSpotsWk
                                                    End If
                                                    Select Case tgCff(ilCff).CffRec.sPriceType
                                                        Case "T"
                                                            llSpotPrice = tgCff(ilCff).CffRec.lActPrice
                                                        Case Else
                                                            llSpotPrice = 0
                                                    End Select
                                                    llPrice = ilNoSpots * llSpotPrice
                                                    slStr = gLongToStrDec(llPrice, 2)
                                                    slStr = gRoundStr(slStr, slRound, 0)
                                                    tgBvfRec(ilBvf).tBvf.lGross(ilWk) = tgBvfRec(ilBvf).tBvf.lGross(ilWk) + gStrDecToLong(slStr, 0)
                                                End If
                                                If llWkStartDate < llFlStartDate Then
                                                    Exit Do
                                                End If
                                                ilCff = tgCff(ilCff).iNextCff
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
Private Sub edcRound_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcRound_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
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
    'If (igWinStatus(RATECARDSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
    '    lbcBudget.Enabled = False
    'Else
    '    lbcBudget.Enabled = True
    'End If
'    gShowBranner
End Sub
Private Sub Form_Load()
    mInit
End Sub
Private Sub imcHelp_Click()
    Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    Traffic!cdcSetup.Action = 6
End Sub
Private Sub lbcVehicle_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.value = False
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
    For ilLoop = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
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
                gUnpackDate tgMCof(ilLoop).iStartDate(0, 1), tgMCof(ilLoop).iStartDate(1, 1), slStartDate
                gUnpackDate tgMCof(ilLoop).iEndDate(0, 12), tgMCof(ilLoop).iEndDate(1, 12), slEndDate
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
mDatePopErr:
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
    Dim slStr As String
    imTerminate = False
    igBDReturn = 0
    
    Screen.MousePointer = vbHourglass
    hmChf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", BudActl
    On Error GoTo 0
    imChfRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", BudActl
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", BudActl
    On Error GoTo 0
    imCffRecLen = Len(tmCff)
    BudActl.Height = cmcApply.Top + 5 * cmcApply.Height / 3
    gCenterStdAlone BudActl
    imChgMode = False
    imBSMode = False
    imDateChgMode = False
    imBypassSetting = False
    imAllClicked = False
    imSetAll = True
    'BudActl.Show
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
    edcRound.Text = "1"
'    mInitDDE
    imcHelp.Picture = Traffic!imcHelp.Picture
    plcScreen.Caption = "Actuals for " & sgBAName
'    gCenterModalForm BudActl
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
    Dim ilRet As Integer 'btrieve status
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilLoop As Integer
    Dim slName As String
    
    For ilLoop = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
        slName = Trim$(tgBvfRec(ilLoop).sVehicle)
        gFindMatch slName, 0, lbcVehicle
        If gLastFound(lbcVehicle) < 0 Then
            lbcVehicle.AddItem slName
        End If
    Next ilLoop
    
    Exit Sub
mPopulateErr:
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
    Dim slStr As String
    Dim ilLoop As Integer
    If (cbcStartDate.Text = "") Or (cbcEndDate.Text = "") Then
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
    Dim ilRet As Integer
    Erase tmChfAdvtExt
    btrDestroy hmChf
    btrDestroy hmClf
    btrDestroy hmCff
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload BudActl
    Set BudActl = Nothing   'Remove data segment
    igManUnload = NO
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Actuals"
End Sub
