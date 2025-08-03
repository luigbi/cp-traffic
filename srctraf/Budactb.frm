VERSION 5.00
Begin VB.Form BudActB 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4665
   ClientLeft      =   600
   ClientTop       =   2205
   ClientWidth     =   9270
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
   ScaleWidth      =   9270
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Vehicles"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      TabIndex        =   12
      Top             =   3945
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
      Height          =   3630
      Left            =   3300
      ScaleHeight     =   3570
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   300
      Width           =   5745
      Begin VB.TextBox edcSellout 
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
         Index           =   1
         Left            =   4275
         TabIndex        =   18
         Top             =   810
         Width           =   795
      End
      Begin VB.TextBox edcSellout 
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
         Index           =   0
         Left            =   1425
         TabIndex        =   16
         Top             =   810
         Width           =   795
      End
      Begin VB.ListBox lbcRateCard 
         Appearance      =   0  'Flat
         Height          =   1710
         Index           =   1
         Left            =   3075
         MultiSelect     =   2  'Extended
         TabIndex        =   15
         Top             =   1230
         Width           =   2400
      End
      Begin VB.ListBox lbcRateCard 
         Appearance      =   0  'Flat
         Height          =   1710
         Index           =   0
         Left            =   300
         MultiSelect     =   2  'Extended
         TabIndex        =   14
         Top             =   1230
         Width           =   2400
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
         Left            =   1920
         TabIndex        =   10
         Top             =   3225
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
         Left            =   3900
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
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   195
         X2              =   5565
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0FFFF&
         X1              =   195
         X2              =   5550
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Remaining Avails Price Determination:"
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
         Left            =   1245
         TabIndex        =   20
         Top             =   570
         Width           =   3345
      End
      Begin VB.Label lacSellout 
         Appearance      =   0  'Flat
         Caption         =   "Sellout %"
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
         Index           =   1
         Left            =   3375
         TabIndex        =   19
         Top             =   855
         Width           =   870
      End
      Begin VB.Label lacSellout 
         Appearance      =   0  'Flat
         Caption         =   "Sellout %"
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
         Index           =   0
         Left            =   495
         TabIndex        =   17
         Top             =   855
         Width           =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0FFFF&
         X1              =   210
         X2              =   5565
         Y1              =   3105
         Y2              =   3105
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   210
         X2              =   5580
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
         Left            =   225
         TabIndex        =   11
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
         Left            =   2970
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
      Left            =   3240
      TabIndex        =   8
      Top             =   4260
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
      Left            =   4875
      TabIndex        =   9
      Top             =   4260
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
      Height          =   3525
      Left            =   150
      ScaleHeight     =   3465
      ScaleWidth      =   2880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   2940
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         Height          =   3390
         Left            =   45
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   45
         Width           =   2790
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
      Top             =   3930
      Width           =   1110
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   180
      Top             =   4215
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "BudActB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Budactb.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: BudActB.Frm
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
'Log Calendar
Dim hmLcf As Integer            'Log Calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim imLcfRecLen As Integer        'LCF record length
Dim hmSsf As Integer
Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
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

Dim tmChfAdvtExt() As CHFADVTEXT
Dim tmBudAvails() As BUDAVAILS
Dim tmRcfInfo() As RCFINFO

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
    ReDim ilSellout(0 To 1) As Integer
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If imEDateSelectedIndex < imSDateSelectedIndex Then
        MsgBox "End Date is prior to Start Date", vbOKOnly + vbExclamation, "Actuals"
        cbcEndDate.SetFocus
        Exit Sub
    End If
    slStr = edcSellout(0).Text
    ilSellout(0) = Val(slStr)
    slStr = edcSellout(1).Text
    ilSellout(1) = Val(slStr)
    If ilSellout(0) + ilSellout(1) > 100 Then
        MsgBox "Sellout Percentage can't exceed 100", vbOKOnly + vbExclamation, "Actuals"
        edcSellout(0).SetFocus
        Exit Sub
    End If
    slRound = edcRound.Text
    If slRound = "" Then
        slRound = "1"
    End If
    ilRet = MsgBox("This will Replace the Budget Values, Ok to Proceed", vbYesNo + vbQuestion, "Actuals")
    If ilRet = vbNo Then
        Exit Sub
    End If
    ilRet = MsgBox("This will take some time, Ok to Proceed", vbYesNo + vbQuestion, "Actuals")
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
    'Get all contract active for dates requested
    slStartDate = cbcStartDate.List(imSDateSelectedIndex)
    slEndDate = cbcEndDate.List(imEDateSelectedIndex)
    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    slStatus = "HO"
    slCntrType = "CVTRQ"
    ilHOType = 1
    sgCntrForDateStamp = ""
    ilRet = gObtainCntrForDate(BudActB, slStartDate, slEndDate, slStatus, slCntrType, ilHOType, tmChfAdvtExt())
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
                                    If (tgClfBud(ilClf).ClfRec.iVefCode = tgBvfRec(ilBvf).tBvf.iVefCode) And (ilSofCode = tgBvfRec(ilBvf).tBvf.iSofCode) Then
                                        'For ilWk = LBound(tgBvfRec(ilBvf).tBvf.lGross) To UBound(tgBvfRec(ilBvf).tBvf.lGross) Step 1
                                        For ilWk = imSDateSelectedIndex + 1 To imEDateSelectedIndex + 1 Step 1
                                            slStr = cbcStartDate.List(ilWk - 1)
                                            llWkStartDate = gDateValue(slStr)
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
                                                    tgBvfRec(ilBvf).tBvf.lGross(ilWk) = tgBvfRec(ilBvf).tBvf.lGross(ilWk) + gStrDecToLong(slStr, 0)
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
    mGetAvailAdj
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
Private Sub edcSellout_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcSellout_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcSellout(Index).Text
    slStr = Left$(slStr, edcSellout(Index).SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcSellout(Index).SelStart - edcSellout(Index).SelLength)
    If gCompNumberStr(slStr, "100") > 0 Then
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
    BudActB.Refresh
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
    Erase tmBudAvails
    Erase tmRcfInfo
    Erase tgClfBud
    Erase tgCffBud

    btrDestroy hmCHF
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmLcf
    btrDestroy hmSsf
    
    Set BudActB = Nothing   'Remove data segment
    
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
Private Sub mGetAvailAdj()
    Dim ilVpfIndex As Integer
    Dim ilVefCode As Integer
    Dim ilMonth As Integer
    Dim ilYear As Integer
    ReDim ilSellout(0 To 1) As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilRcf As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilRet As Integer
    Dim ilBvf As Integer
    Dim ilLoop1 As Integer
    Dim slVehicle As String
    Dim ilWk As Integer
    Dim slAsterisk As String
    Dim slChar As String
    Dim llWkStartDate As Long
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCode As Integer
    Dim ilFound As Integer
    Dim llATotal As Long
    Dim llCTotal As Long
    Dim slATotal As String
    Dim slRound As String
    Dim ilCount As Integer
    ReDim tmRcfInfo(0 To 0) As RCFINFO
    slStr = edcSellout(0).Text
    ilSellout(0) = Val(slStr)
    slStr = edcSellout(1).Text
    ilSellout(1) = Val(slStr)
    slRound = edcRound.Text
    If slRound = "" Then
        slRound = "1"
    End If
    If ((ilSellout(0) > 0) And (lbcRateCard(0).SelCount > 0)) Or ((ilSellout(1) > 0) And (lbcRateCard(0).SelCount > 0)) Then
        For ilLoop = 0 To 1 Step 1
            If (ilSellout(ilLoop) > 0) And (lbcRateCard(ilLoop).SelCount > 0) Then
                For ilIndex = 0 To lbcRateCard(ilLoop).ListCount - 1 Step 1
                    If lbcRateCard(ilLoop).Selected(ilIndex) Then
                        slNameCode = tgRateCardCode(ilIndex).sKey
                        ilRet = gParseItem(slNameCode, 3, "\", slCode)
                        ilCode = Val(slCode)
                        For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
                            If tgMRcf(ilRcf).iCode = ilCode Then
                                'Disallow the same year
                                ilFound = False
                                For ilLoop1 = 0 To UBound(tmRcfInfo) - 1 Step 1
                                    If (tmRcfInfo(ilLoop1).iYear = tgMRcf(ilRcf).iYear) And (tmRcfInfo(ilLoop1).iSellout = ilLoop) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilLoop1
                                If Not ilFound Then
                                    tmRcfInfo(UBound(tmRcfInfo)).iRcfCode = tgMRcf(ilRcf).iCode
                                    tmRcfInfo(UBound(tmRcfInfo)).iYear = tgMRcf(ilRcf).iYear
                                    tmRcfInfo(UBound(tmRcfInfo)).iSellout = ilLoop
                                    ReDim Preserve tmRcfInfo(0 To UBound(tmRcfInfo) + 1) As RCFINFO
                                End If
                                Exit For
                            End If
                        Next ilRcf
                    End If
                Next ilIndex
            End If
        Next ilLoop
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
                'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
                For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
                    If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                        ilVefCode = tgBvfRec(ilBvf).tBvf.iVefCode
                        ilVpfIndex = gVpfFind(BudActB, ilVefCode)
                        Exit For
                    End If
                Next ilBvf
                ilCount = 0
                'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
                For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
                    If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                        ilCount = ilCount + 1
                    End If
                Next ilBvf
                For ilWk = imSDateSelectedIndex + 1 To imEDateSelectedIndex + 1 Step 1
                    slStr = cbcStartDate.List(ilWk - 1)
                    llWkStartDate = gDateValue(slStr)
                    gObtainMonthYear 0, slStr, ilMonth, ilYear
                    ReDim tmBudAvails(0 To 0) As BUDAVAILS
                    For ilRcf = 0 To UBound(tmRcfInfo) - 1 Step 1
                        If tmRcfInfo(ilRcf).iYear = ilYear Then
                            For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                                If (tgMRif(llRif).iRcfCode = tmRcfInfo(ilRcf).iRcfCode) And (tgMRif(llRif).iVefCode = ilVefCode) Then
                                    'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                                    '    If (tgMRdf(ilRdf).iCode = tgMRif(llRif).iRdfcode) And (tgMRdf(ilRdf).sBase = "Y") Then
                                        ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfcode)
                                        If ilRdf <> -1 Then
                                            If tgMRdf(ilRdf).sBase = "Y" Then
                                                ilFound = False
                                                For ilLoop1 = LBound(tmBudAvails) To UBound(tmBudAvails) - 1 Step 1
                                                    If tmBudAvails(ilLoop1).tRdf.iCode = tgMRdf(ilRdf).iCode Then
                                                        ilFound = True
                                                        Exit For
                                                    End If
                                                Next ilLoop1
                                                If Not ilFound Then
                                                    tmBudAvails(UBound(tmBudAvails)).tRdf = tgMRdf(ilRdf)
                                                    tmBudAvails(UBound(tmBudAvails)).iSellout = tmRcfInfo(ilRcf).iSellout
                                                    tmBudAvails(UBound(tmBudAvails)).lRifIndex = llRif
                                                    ReDim Preserve tmBudAvails(0 To UBound(tmBudAvails) + 1) As BUDAVAILS
                                                End If
                                            End If
                                    '        Exit For
                                        End If
                                    'Next ilRdf
                                End If
                            Next llRif
                        End If
                    Next ilRcf
                    mGetAvailCounts hmSsf, hmLcf, ilVefCode, ilVpfIndex, llWkStartDate, llWkStartDate + 6
                    'Get Avail Total for Vehicle
                    llATotal = 0
                    For ilLoop1 = LBound(tmBudAvails) To UBound(tmBudAvails) - 1 Step 1
                        llATotal = llATotal + (ilSellout(tmBudAvails(ilLoop1).iSellout) * tmBudAvails(ilLoop1).lRate * tmBudAvails(ilLoop1).i30Avails) / 100
                    Next ilLoop1
                    slATotal = gLongToStrDec(llATotal, 0) & ".00"
                    'Get Current Total
                    llCTotal = 0
                    'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
                    For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
                        If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                            llCTotal = llCTotal + tgBvfRec(ilBvf).tBvf.lGross(ilWk)
                        End If
                    Next ilBvf
                    'For ilBvf = LBound(tgBvfRec) To UBound(tgBvfRec) - 1 Step 1
                    For ilBvf = igLBBvfRec To UBound(tgBvfRec) - 1 Step 1
                        If StrComp(slVehicle, Trim$(tgBvfRec(ilBvf).sVehicle), 1) = 0 Then
                            If llCTotal > 0 Then
                                slStr = gMulStr(slATotal, gLongToStrDec(tgBvfRec(ilBvf).tBvf.lGross(ilWk), 0))
                                slStr = gDivStr(slStr, gLongToStrDec(llCTotal, 0))
                                slStr = gRoundStr(slStr, slRound, 0)
                                tgBvfRec(ilBvf).tBvf.lGross(ilWk) = tgBvfRec(ilBvf).tBvf.lGross(ilWk) + gStrDecToLong(slStr, 0)
                            Else
                                If ilCount > 0 Then
                                    slStr = gMulStr(slATotal, gLongToStrDec(tgBvfRec(ilBvf).tBvf.lGross(ilWk), 0))
                                    slStr = gDivStr(slStr, gIntToStrDec(ilCount, 0))
                                    slStr = gRoundStr(slStr, slRound, 0)
                                    tgBvfRec(ilBvf).tBvf.lGross(ilWk) = tgBvfRec(ilBvf).tBvf.lGross(ilWk) + gStrDecToLong(slStr, 0)
                                Else
                                    tgBvfRec(ilBvf).tBvf.lGross(ilWk) = 0
                                End If
                            End If
                        End If
                    Next ilBvf
                Next ilWk
            End If
        Next ilLoop
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gGetAvailCounts                 *
'*                                                     *
'*             Created:10/20/97      By:D. Hosaka      *
'*             Copy of mGetAvails Counts made into     *
'*             generalized subroutine                  *
''*                                                    *
'*                                                     *
'*            Comments:Obtain the Avail counts         *
'*                                                     *
'*******************************************************
Private Sub mGetAvailCounts(hlSsf As Integer, hlLcf As Integer, ilVefCode As Integer, ilVpfIndex As Integer, llWkSDate As Long, llWkEDate As Long)
'
'   Where:
'
'   hlSsf (I) - handle to SSF file
'   hlLcf (I) - handle to Lcf file
'   ilVefCode (I) - vehicle code to process
'   ilVpfIndex (I) - vehicle options pointer
'   llWkSDate (I) - week start date to begin searching Avails
'   llWkEDate (I) - week end date to stop searching avails
'
'   Note: Remnants; Direct Response; per Inquiry; PSA and Promos are not
'         saved with a miss status
'         For scheduled spots the rank is used to determine if it is one
'         of the above (Direct reponse=1010; Remnant=1020; per Inquiry= 1030;
'         PSA=1060; Promo=1050.
'

    Dim ilType As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilEvt As Integer
    Dim ilRet As Integer
    Dim ilSpot As Integer
    Dim llTime As Long
    Dim ilRdf As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilLen As Integer
    Dim ilNo30 As Integer
    Dim ilNo60 As Integer
    Dim ilDay As Integer
    Dim ilLtfCode As Integer
    Dim ilAvailOk As Integer
    Dim ilSpotOK As Integer
    Dim llLoopDate As Long
    Dim ilWeekDay As Integer
    Dim llLatestDate As Long
    Dim ilIndex As Integer
    Dim ilLo As Integer
    Dim ilWkNo As Integer               'week index to rate card
    Dim ilAdjAdd As Integer
    Dim ilVefIndex As Integer
    ReDim ilEvtType(0 To 14) As Integer

    ilType = 0
    'Exclude Sports
    ilVefIndex = gBinarySearchVef(ilVefCode)
    If ilVefIndex = -1 Then
        Exit Sub
    End If
    llLatestDate = gGetLatestLCFDate(hlLcf, "C", ilVefCode)
    'set the type of events to get fro the day (only Contract avails)
    For ilLoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(ilLoop) = False
    Next ilLoop
    ilEvtType(2) = True
    If tgVpf(ilVpfIndex).sSSellOut = "B" Then           'if units & seconds - add 2 to 30 sec unit and take away 1 fro 60
        ilAdjAdd = 2
    ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then       'if units only - take 1 away from 60 count and add 1 to 30 count
        ilAdjAdd = 1
    End If
    For llLoopDate = llWkSDate To llWkEDate Step 1
        slDate = Format$(llLoopDate, "m/d/yy")
        gPackDate slDate, ilDate0, ilDate1
        gObtainWkNo 0, slDate, ilWkNo, ilLo        'obtain the week bucket number
        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
        If tgMVef(ilVefIndex).sType <> "G" Then
            tmSsfSrchKey.iType = ilType
            tmSsfSrchKey.iVefCode = ilVefCode
            tmSsfSrchKey.iDate(0) = ilDate0
            tmSsfSrchKey.iDate(1) = ilDate1
            tmSsfSrchKey.iStartTime(0) = 0
            tmSsfSrchKey.iStartTime(1) = 0
            ilRet = gSSFGetGreaterOrEqual(hlSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        Else
            tmSsfSrchKey2.iVefCode = ilVefCode
            tmSsfSrchKey2.iDate(0) = ilDate0
            tmSsfSrchKey2.iDate(1) = ilDate1
            ilRet = gSSFGetGreaterOrEqualKey2(hlSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
            ilType = tmSsf.iType
        End If
        If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iType <> ilType) Or (tmSsf.iVefCode <> ilVefCode Or (tmSsf.iDate(0) <> ilDate0) Or (tmSsf.iDate(1) <> ilDate1)) Then
            If (llLoopDate > llLatestDate) Then
                ReDim tlLLC(0 To 0) As LLC  'Merged library names
                If tgMVef(ilVefIndex).sType <> "G" Then
                    ilWeekDay = gWeekDayStr(slDate) + 1 '1=Monday TFN; 2=Tues,...7=Sunday TFN
                    If ilWeekDay = 1 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNMO", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 2 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTU", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 3 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNWE", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 4 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTH", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 5 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNFR", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 6 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSA", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 7 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSU", "12M", "12M", ilEvtType(), tlLLC())
                    End If
                End If
                tmSsf.iType = ilType
                tmSsf.iVefCode = ilVefCode
                tmSsf.iDate(0) = ilDate0
                tmSsf.iDate(1) = ilDate1
                gPackTime tlLLC(0).sStartTime, tmSsf.iStartTime(0), tmSsf.iStartTime(1)
                tmSsf.iCount = 0
                'tmSsf.iNextTime(0) = 1  'Time not defined
                'tmSsf.iNextTime(1) = 0
                For ilIndex = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1

                    tmAvail.iRecType = Val(tlLLC(ilIndex).sType)
                    gPackTime tlLLC(ilIndex).sStartTime, tmAvail.iTime(0), tmAvail.iTime(1)
                    tmAvail.iLtfCode = tlLLC(ilIndex).iLtfCode
                    tmAvail.iAvInfo = tlLLC(ilIndex).iAvailInfo Or tlLLC(ilIndex).iUnits
                    tmAvail.iLen = CInt(gLengthToCurrency(tlLLC(ilIndex).sLength))
                    tmAvail.ianfCode = Val(tlLLC(ilIndex).sName)
                    tmAvail.iNoSpotsThis = 0
                    tmAvail.iOrigUnit = 0
                    tmAvail.iOrigLen = 0
                    tmSsf.iCount = tmSsf.iCount + 1
                    tmSsf.tPas(ADJSSFPASBZ + tmSsf.iCount) = tmAvail
                Next ilIndex
                ilRet = BTRV_ERR_NONE
            End If
        End If

        Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
            gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llDate
            ilDay = gWeekDayLong(llDate)
            ilEvt = 1
            Do While ilEvt <= tmSsf.iCount
               LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                If tmProg.iRecType = 1 Then    'Program (not working for nested prog)
                    ilLtfCode = tmProg.iLtfCode
                ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 2) Then 'Contract Avails only
                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                    'Determine which rate card program this is associated with
                    For ilRdf = LBound(tmBudAvails) To UBound(tmBudAvails) - 1 Step 1
                        ilAvailOk = False
                        If (tmBudAvails(ilRdf).tRdf.iLtfCode(0) <> 0) Or (tmBudAvails(ilRdf).tRdf.iLtfCode(1) <> 0) Or (tmBudAvails(ilRdf).tRdf.iLtfCode(2) <> 0) Then
                            If (ilLtfCode = tmBudAvails(ilRdf).tRdf.iLtfCode(0)) Or (ilLtfCode = tmBudAvails(ilRdf).tRdf.iLtfCode(1)) Or (ilLtfCode = tmBudAvails(ilRdf).tRdf.iLtfCode(1)) Then
                                ilAvailOk = False    'True- code later
                            End If
                        Else

                            For ilLoop = LBound(tmBudAvails(ilRdf).tRdf.iStartTime, 2) To UBound(tmBudAvails(ilRdf).tRdf.iStartTime, 2) Step 1 'Row
                                If (tmBudAvails(ilRdf).tRdf.iStartTime(0, ilLoop) <> 1) Or (tmBudAvails(ilRdf).tRdf.iStartTime(1, ilLoop) <> 0) Then
                                    gUnpackTimeLong tmBudAvails(ilRdf).tRdf.iStartTime(0, ilLoop), tmBudAvails(ilRdf).tRdf.iStartTime(1, ilLoop), False, llStartTime
                                    gUnpackTimeLong tmBudAvails(ilRdf).tRdf.iEndTime(0, ilLoop), tmBudAvails(ilRdf).tRdf.iEndTime(1, ilLoop), True, llEndTime
                                    'If (llTime >= llStartTime) And (llTime < llEndTime) And (tmBudAvails(ilRdf).tRdf.sWkDays(ilLoop, ilDay + 1) = "Y") Then
                                    If (llTime >= llStartTime) And (llTime < llEndTime) And (tmBudAvails(ilRdf).tRdf.sWkDays(ilLoop, ilDay) = "Y") Then
                                        ilAvailOk = True
                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                        End If
                        If ilAvailOk Then
                            If tmBudAvails(ilRdf).tRdf.sInOut = "I" Then   'Book into
                                If tmAvail.ianfCode <> tmBudAvails(ilRdf).tRdf.ianfCode Then
                                    ilAvailOk = False
                                End If
                            ElseIf tmBudAvails(ilRdf).tRdf.sInOut = "O" Then   'Exclude
                                If tmAvail.ianfCode = tmBudAvails(ilRdf).tRdf.ianfCode Then
                                    ilAvailOk = False
                                End If
                            End If
                        End If
                        If ilAvailOk Then
                            'Determine if Avr created
                            ilFound = False
                            tmBudAvails(ilRdf).lRate = tgMRif(tmBudAvails(ilRdf).lRifIndex).lRate(ilWkNo)
                            'Always gather inventory
                            ilLen = tmAvail.iLen
                            ilNo30 = 0
                            ilNo60 = 0
                            If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                                'Convert inventory to number of 30's and 60's
                                Do While ilLen >= 60
                                    ilNo60 = ilNo60 + 1
                                    ilLen = ilLen - 60
                                Loop
                                Do While ilLen >= 30
                                    ilNo30 = ilNo30 + 1
                                    ilLen = ilLen - 30
                                Loop
                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                'Count 30 or 60 and set flag if neither
                                If ilLen = 60 Then
                                    ilNo60 = 1
                                ElseIf ilLen = 30 Then
                                    ilNo30 = 1
                                Else
                                    If ilLen <= 30 Then
                                        ilNo30 = 1
                                    Else
                                        ilNo60 = 1
                                    End If
                                End If
                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                'Count 30 or 60 and set flag if neither
                                If ilLen = 60 Then
                                    ilNo60 = 1
                                ElseIf ilLen = 30 Then
                                    ilNo30 = 1
                                End If
                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                            End If
                            tmBudAvails(ilRdf).i30Avails = tmBudAvails(ilRdf).i30Avails + ilNo30 + ilAdjAdd * ilNo60
                            'Always calculate Avails
                            For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                               LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                                ilSpotOK = True                             'assume spot is OK to include

                                If (tmSpot.iRank And RANKMASK) = REMNANTRANK Then
                                    ilSpotOK = False
                                End If
                                If (tmSpot.iRank And RANKMASK) = PERINQUIRYRANK Then
                                    ilSpotOK = False
                                End If
                                'If tmSpot.iRank = 1040 And Not tlCntTypes.iTrade Then
                                '    ilSpotOK = False
                                'End If
                                If (tmSpot.iRank And RANKMASK) = EXTRARANK Then
                                    ilSpotOK = False
                                End If
                                If (tmSpot.iRank And RANKMASK) = PROMORANK Then
                                    ilSpotOK = False
                                End If
                                If (tmSpot.iRank And RANKMASK) = PSARANK Then
                                    ilSpotOK = False
                                End If
                                If (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                    ilSpotOK = False
                                End If
                                ilLen = tmSpot.iPosLen And &HFFF
                                If ilSpotOK Then                            'continue testing other filters
                                    'ilLen = tmSdf.iLen
                                    ilNo30 = 0
                                    ilNo60 = 0
                                    If tgVpf(ilVpfIndex).sSSellOut = "B" Then                   'both units and seconds
                                    'Convert inventory to number of 30's and 60's
                                        Do While ilLen >= 60
                                            ilNo60 = ilNo60 + 1
                                            ilLen = ilLen - 60
                                        Loop
                                        Do While ilLen >= 30
                                            ilNo30 = ilNo30 + 1
                                            ilLen = ilLen - 30
                                        Loop
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then               'units sold
                                        'Count 30 or 60 and set flag if neither
                                        If ilLen = 60 Then
                                            ilNo60 = 1
                                        ElseIf ilLen = 30 Then
                                            ilNo30 = 1
                                        Else
                                            If ilLen <= 30 Then
                                                ilNo30 = 1
                                            Else
                                                ilNo60 = 1
                                            End If
                                        End If
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then               'matching units
                                        'Count 30 or 60 and set flag if neither
                                        If ilLen = 60 Then
                                            ilNo60 = 1
                                        ElseIf ilLen = 30 Then
                                            ilNo30 = 1
                                        End If
                                    ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                    End If
                                    tmBudAvails(ilRdf).i30Avails = tmBudAvails(ilRdf).i30Avails - ilNo30 - ilAdjAdd * ilNo60
                                End If
                            Next ilSpot                             'loop from ssf file for # spots in avail
                        End If                                          'Avail OK
                    Next ilRdf                                          'ilRdf = lBound(tlAvRdf)
                    ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                End If
                ilEvt = ilEvt + 1   'Increment to next event
            Loop                                                        'do while ilEvt <= tmSsf.iCount
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            ilRet = gSSFGetNext(hlSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            If tgMVef(ilVefIndex).sType = "G" Then
                ilType = tmSsf.iType
            End If
        Loop
    Next llLoopDate

    'Get missed
    'Key 2: VefCode; SchStatus; AdfCode; Date, Time
    'For ilPass = 0 To 2 Step 1
    '    slDate = Format$(llWkSDate, "m/d/yy")
    '    gPackDate slDate, ilDate0, ilDate1
    '    tmSdfSrchKey2.iVefCode = ilVefCode
    '    If ilPass = 0 Then
    '        slType = "M"
    '    ElseIf ilPass = 1 Then
    '        slType = "R"
    '    ElseIf ilPass = 2 Then
    '        slType = "U"
    '    End If
    '    tmSdfSrchKey2.sSchStatus = slType
    '    tmSdfSrchKey2.iAdfCode = 0
    '    tmSdfSrchKey2.iDate(0) = ilSAvailsDates(0)
    '    tmSdfSrchKey2.iDate(1) = ilSAvailsDates(1)
    '    tmSdfSrchKey2.iTime(0) = 0
    '    tmSdfSrchKey2.iTime(1) = 0
    '    ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
    '    'This code added as replacement for Ext operation
    '    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.sSchStatus = slType)
    '       gUnPackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
    '       If (llDate > llWkEDate) Then
    '           Exit Do
    '       End If
    '       slDate = Format$(llDate, "m/d/yy")
    '       gPackDate slDate, ilDate0, ilDate1
    '       gObtainWkNo 0, slDate, ilWkNo, ilLo        'obtain the week bucket number
    '       ilDay = gWeekDayLong(llDate)
    '       gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
    '       For ilRdf = LBound(tmBudAvails) To UBound(tmBudAvails) - 1 Step 1
    '                    ilAvailOk = False
    '                    If (tmBudAvails(ilRdf).tRdf.iLtfCode(0) <> 0) Or (tmBudAvails(ilRdf).tRdf.iLtfCode(1) <> 0) Or (tmBudAvails(ilRdf).tRdf.iLtfCode(2) <> 0) Then
    '                        If (ilLtfCode = tmBudAvails(ilRdf).tRdf.iLtfCode(0)) Or (ilLtfCode = tmBudAvails(ilRdf).tRdf.iLtfCode(1)) Or (ilLtfCode = tmBudAvails(ilRdf).tRdf.iLtfCode(1)) Then
    '                            ilAvailOk = False    'True- code later
    '                        End If
    '                    Else
    '                        For ilLoop = LBound(tmBudAvails(ilRdf).tRdf.iStartTime, 2) To UBound(tmBudAvails(ilRdf).tRdf.iStartTime, 2) Step 1 'Row
    '                            If (tmBudAvails(ilRdf).tRdf.iStartTime(0, ilLoop) <> 1) Or (tmBudAvails(ilRdf).tRdf.iStartTime(1, ilLoop) <> 0) Then
    '                                gUnpackTimeLong tmBudAvails(ilRdf).tRdf.iStartTime(0, ilLoop), tmBudAvails(ilRdf).tRdf.iStartTime(1, ilLoop), False, llStartTime
    '                                gUnpackTimeLong tmBudAvails(ilRdf).tRdf.iEndTime(0, ilLoop), tmBudAvails(ilRdf).tRdf.iEndTime(1, ilLoop), True, llEndTime
    '                                'Don't include the end time i.e. 10a-3p is 10a thru 2:59:59p
    '                                If (llTime >= llStartTime) And (llTime < llEndTime) And (tmBudAvails(ilRdf).tRdf.sWkDays(ilLoop, ilDay + 1) = "Y") Then
    '                                    ilAvailOk = True
    '                                    Exit For
    '                                End If
    '                            End If
    '                        Next ilLoop
    '                    End If
    '                    If ilAvailOk Then
    '                        tmBudAvails(ilRdf).lRate = tgRif(tmBudAvails(ilRdf).iRifIndex).lRate(ilWkNo)
    '                        ilNo30 = 0
    '                        ilNo60 = 0
    '                        ilLen = tmSdf.iLen
    '                        If tgVpf(ilVpfIndex).sSSellOut = "B" Then
    '                        'Convert inventory to number of 30's and 60's
    '                            Do While ilLen >= 60
    '                                ilNo60 = ilNo60 + 1
    '                                ilLen = ilLen - 60
    '                            Loop
    '                            Do While ilLen >= 30
    '                                ilNo30 = ilNo30 + 1
    '                                ilLen = ilLen - 30
    '                            Loop
    '                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
    '                            'Count 30 or 60 and set flag if neither
    '                            If ilLen = 60 Then
    '                                ilNo60 = 1
    '                            ElseIf ilLen = 30 Then
    '                                ilNo30 = 1
    '                            Else
    '                                If ilLen <= 30 Then
    '                                    ilNo30 = 1
    '                                Else
    '                                    ilNo60 = 1
    '                                End If
    '                            End If
    '                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
    '                            'Count 30 or 60 and set flag if neither
    '                            If ilLen = 60 Then
    '                                ilNo60 = 1
    '                            ElseIf ilLen = 30 Then
    '                                ilNo30 = 1
    '                            End If
    '                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
    '                        End If
    '                        tmBudAvails(ilRdf).i30Avail = tmBudAvails(ilRdf).i30Avail - ilNo30 - ilAdjAdd * ilNo60
    '                    End If
    '                Next ilRdf
    '        ilRet = btrGetNext(hlSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    '    Loop
    'Next ilPass
    Erase ilEvtType
    Erase tlLLC
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
    igBDReturn = 0
    
    igLBBvfRec = 1  'Match definition in Budget

    Screen.MousePointer = vbHourglass
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", BudActB
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", BudActB
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", BudActB
    On Error GoTo 0
    imCffRecLen = Len(tmCff)
    hmLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", BudActB
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)
    hmSsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ssf.Btr)", BudActB
    On Error GoTo 0
    imSsfRecLen = Len(tmSsf)
    BudActB.Height = cmcApply.Top + 5 * cmcApply.Height / 3
    gCenterStdAlone BudActB
    imChgMode = False
    imBSMode = False
    imDateChgMode = False
    imBypassSetting = False
    imAllClicked = False
    imSetAll = True
    'BudActB.Show
    Screen.MousePointer = vbHourglass
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    mDatePop
    If imTerminate Then
        Exit Sub
    End If
    mRateCardPop
    If imTerminate Then
        Exit Sub
    End If
    ilRet = gObtainSalesperson()
    edcRound.Text = "1"
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    'plcScreen.Caption = "Actuals for " & sgBAName
'    gCenterModalForm BudActB
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
'*      Procedure Name:mRateCardPop                    *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Rate Card list        *
'*                      box if required                *
'*                                                     *
'*******************************************************
Private Sub mRateCardPop()
'
'   mRateCardPop
'   Where:
'
    Dim ilRet As Integer
    Dim llDate As Long
    Dim ilLoop As Integer
    llDate = 0  'Get all
    'Repopulate if required- if rate card changed by another user while in this screen
    'ilRet = gPopRateCardBox(Contract, llDate, lbcRateCard, Traffic!lbcRateCardCode, -1)
    ilRet = gPopRateCardBox(BudActB, llDate, lbcRateCard(0), tgRateCardCode(), sgRateCardCodeTag, -1)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mRateCardErr
        gCPErrorMsg ilRet, "mRateCardPop (gPopRateCardBox)", BudActB
        On Error GoTo 0
        For ilLoop = 0 To lbcRateCard(0).ListCount - 1 Step 1
            lbcRateCard(1).AddItem lbcRateCard(0).List(ilLoop)
        Next ilLoop
    End If
    Exit Sub
mRateCardErr:
    On Error GoTo 0
    imTerminate = True
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
    If (lbcRateCard(0).SelCount > 2) Or (lbcRateCard(1).SelCount > 2) Then
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
    Unload BudActB
    igManUnload = NO
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Actuals for " & sgBAName
End Sub
